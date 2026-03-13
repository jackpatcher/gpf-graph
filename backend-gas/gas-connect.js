/*
	Google Apps Script (GAS) connector for GPF NAV data
	---------------------------------------------------
	Deploy as Web App and use these endpoints:

	GET  ?action=init
	GET  ?action=health
	GET  ?action=data&since=YYYY-MM-DD&limit=5000
	GET  ?action=sync&token=YOUR_SYNC_TOKEN&startYear=1998&startMonth=1
	POST { action: "init" | "sync" | "data", token?, since?, limit?, startYear?, startMonth?, endYear?, endMonth? }

	Required Script Properties:
	- SHEET_ID: target Google Sheet ID

	Optional Script Properties:
	- SHEET_NAME: default "NAV_DATA"
	- CONFIG_SHEET_NAME: default "CONFIG"
	- USERS_SHEET_NAME: default "USERS"
	- PORTFOLIO_SHEET_NAME: default "PORTFOLIO_CHANGES"
	- SYNC_TOKEN: shared secret for sync endpoint
*/

const GAS_CONNECT_CONFIG = {
	sheetId: PropertiesService.getScriptProperties().getProperty('SHEET_ID') || '',
	sheetName: PropertiesService.getScriptProperties().getProperty('SHEET_NAME') || 'NAV_DATA',
	configSheetName:
		PropertiesService.getScriptProperties().getProperty('CONFIG_SHEET_NAME') || 'CONFIG',
	usersSheetName:
		PropertiesService.getScriptProperties().getProperty('USERS_SHEET_NAME') || 'USERS',
	portfolioSheetName:
		PropertiesService.getScriptProperties().getProperty('PORTFOLIO_SHEET_NAME') || 'PORTFOLIO_CHANGES',
	syncToken: PropertiesService.getScriptProperties().getProperty('SYNC_TOKEN') || '',
	mailSenderName: PropertiesService.getScriptProperties().getProperty('MAIL_SENDER_NAME') || 'GPF Graph',
	googleClientId: PropertiesService.getScriptProperties().getProperty('GOOGLE_CLIENT_ID') || '',
	timeoutMs: 30000,
	apiUrlTemplate:
		'https://www.gpf.or.th/thai2019/About/memberfund-api.php?pageName=NAVBottom_{MM}_{YYYY}',
	unitCostCount: 14
};

const GAS_CONNECT_PROJECT_CONFIG_DEFAULTS = [
	{
		key: 'SHEET_VERSION',
		value: '4',
		description: 'Schema version for NAV + users + Google email change OTP flow'
	},
	{
		key: 'PROJECT_NAME',
		value: 'gpf-graph',
		description: 'Project identifier'
	},
	{
		key: 'API_URL_TEMPLATE',
		value:
			'https://www.gpf.or.th/thai2019/About/memberfund-api.php?pageName=NAVBottom_{MM}_{YYYY}',
		description: 'Upstream API template'
	},
	{
		key: 'START_YEAR',
		value: '1998',
		description: 'Default full-sync start year'
	},
	{
		key: 'START_MONTH',
		value: '1',
		description: 'Default full-sync start month'
	},
	{
		key: 'SYNC_TIMEZONE',
		value: 'Asia/Bangkok',
		description: 'Project timezone'
	},
	{
		key: 'LAST_SYNC_AT',
		value: '',
		description: 'Last successful sync timestamp'
	},
	{
		key: 'LAST_SYNC_STATUS',
		value: '',
		description: 'Last sync status summary'
	},
	{
		key: 'USERS_SHEET_NAME',
		value: 'USERS',
		description: 'Sheet name for user records'
	},
	{
		key: 'PORTFOLIO_SHEET_NAME',
		value: 'PORTFOLIO_CHANGES',
		description: 'Sheet name for portfolio batch JSON records'
	}
];

const GAS_CONNECT_USER_HEADERS = [
	'userId',
	'username',
	'passwordHash',
	'displayName',
	'role',
	'status',
	'createdAt',
	'updatedAt',
	'googleId',
	'emailChangeOtpHash',
	'emailChangeOtpExpiresAt',
	'pendingEmail',
	'pendingEmailLinkExpiresAt',
	'emailChangeLastSentAt'
];

function doGet(e) {
	return handleRequest_(e, null);
}

function doPost(e) {
	let payload = {};
	try {
		payload = e && e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
	} catch (error) {
		return jsonOut_({ ok: false, error: 'Invalid JSON body' });
	}
	return handleRequest_(e, payload);
}

// Run this once from Apps Script editor to grant required scopes (including mail send scope).
function authorizeRequiredScopes() {
	MailApp.getRemainingDailyQuota();
	return {
		ok: true,
		message: 'Scopes authorized. You can deploy/redeploy Web App and retry OTP email.'
	};
}

function handleRequest_(e, body) {
	try {
		validateConfig_();

		const params = (e && e.parameter) || {};
		const action = String((body && body.action) || params.action || 'data').toLowerCase();

		if (action === 'init') {
			const initResult = ensureProjectInitialized_();
			return jsonOut_({ ok: true, action: 'init', result: initResult });
		}

		ensureProjectInitialized_();

		if (action === 'health') {
			return jsonOut_({
				ok: true,
				service: 'gas-nav-connector',
				sheetName: GAS_CONNECT_CONFIG.sheetName,
				configSheetName: GAS_CONNECT_CONFIG.configSheetName,
				usersSheetName: GAS_CONNECT_CONFIG.usersSheetName,
				portfolioSheetName: GAS_CONNECT_CONFIG.portfolioSheetName,
				timestamp: new Date().toISOString()
			});
		}

		if (action === 'sync') {
			const incomingToken = String((body && body.token) || params.token || '');
			if (!isAuthorizedSync_(incomingToken)) {
				return jsonOut_({ ok: false, error: 'Unauthorized sync request' });
			}

			const now = new Date();
			const syncStart = resolveSyncStart_(body, params);
			const startYear = syncStart.startYear;
			const startMonth = syncStart.startMonth;
			const endYear = toInt_((body && body.endYear) || params.endYear, now.getFullYear());
			const endMonth = toInt_((body && body.endMonth) || params.endMonth, now.getMonth() + 1);
			const reconcile = toBoolean_((body && body.reconcile) || params.reconcile);

			const result = syncToSheet_(startYear, startMonth, endYear, endMonth, { reconcile: reconcile });
			return jsonOut_({ ok: true, action: 'sync', result: result });
		}

		if (action === 'monthsummary' || action === 'month_summary') {
			const result = buildMonthSummary_();
			return jsonOut_({ ok: true, action: 'monthSummary', result: result });
		}

		if (action === 'datamonth' || action === 'data_month') {
			const year = toInt_((body && body.year) || params.year, 0);
			const month = toInt_((body && body.month) || params.month, 0);
			const result = readMonthData_(year, month);
			return jsonOut_({ ok: true, action: 'dataMonth', result: result });
		}

		if (action === 'registeruser' || action === 'register_user') {
			const username = String((body && body.username) || params.username || '').trim();
			const passwordHash = String((body && body.passwordHash) || params.passwordHash || '').trim();
			const displayName = String((body && body.displayName) || params.displayName || '').trim();
			const result = registerUser_(username, passwordHash, displayName);
			return jsonOut_({ ok: true, action: 'registerUser', result: result });
		}

		if (action === 'loginuser' || action === 'login_user') {
			const username = String((body && body.username) || params.username || '').trim();
			const passwordHash = String((body && body.passwordHash) || params.passwordHash || '').trim();
			const result = loginUser_(username, passwordHash);
			return jsonOut_({ ok: true, action: 'loginUser', result: result });
		}

		if (action === 'getuserprofile' || action === 'get_user_profile') {
			const userId = String((body && body.userId) || params.userId || '').trim();
			const result = getUserProfile_(userId);
			return jsonOut_({ ok: true, action: 'getUserProfile', result: result });
		}

		if (action === 'loginwithgoogletoken' || action === 'login_with_google_token') {
			const idToken = String((body && body.idToken) || params.idToken || '').trim();
			const pendingUserId = String((body && body.pendingUserId) || params.pendingUserId || '').trim();
			const result = loginWithGoogleToken_(idToken, pendingUserId);
			return jsonOut_({ ok: true, action: 'loginWithGoogleToken', result: result });
		}

		if (action === 'getadminapprovaldashboard' || action === 'get_admin_approval_dashboard') {
			const adminUserId = String((body && body.adminUserId) || params.adminUserId || '').trim();
			const result = getAdminApprovalDashboard_(adminUserId);
			return jsonOut_({ ok: true, action: 'getAdminApprovalDashboard', result: result });
		}

		if (action === 'approveuserregistration' || action === 'approve_user_registration') {
			const adminUserId = String((body && body.adminUserId) || params.adminUserId || '').trim();
			const targetUserId = String((body && body.targetUserId) || params.targetUserId || '').trim();
			const result = approveUserRegistration_(adminUserId, targetUserId);
			return jsonOut_({ ok: true, action: 'approveUserRegistration', result: result });
		}

		if (action === 'requestemailchangeotp' || action === 'request_email_change_otp') {
			const userId = String((body && body.userId) || params.userId || '').trim();
			const result = requestEmailChangeOtp_(userId);
			return jsonOut_({ ok: true, action: 'requestEmailChangeOtp', result: result });
		}

		if (action === 'verifyemailchangeotp' || action === 'verify_email_change_otp') {
			const userId = String((body && body.userId) || params.userId || '').trim();
			const otpCode = String((body && body.otpCode) || params.otpCode || '').trim();
			const newEmail = String((body && body.newEmail) || params.newEmail || '').trim();
			const result = verifyEmailChangeOtp_(userId, otpCode, newEmail);
			return jsonOut_({ ok: true, action: 'verifyEmailChangeOtp', result: result });
		}

		if (action === 'saveportfoliobatch' || action === 'save_portfolio_batch') {
			const userId = String((body && body.userId) || params.userId || '').trim();
			const effectiveDate = String((body && body.effectiveDate) || params.effectiveDate || '').trim();
			const entriesJson = String((body && body.entriesJson) || params.entriesJson || '[]');
			const note = String((body && body.note) || params.note || '').trim();
			const batchId = String((body && body.batchId) || params.batchId || '').trim();
			const result = savePortfolioBatch_(userId, effectiveDate, entriesJson, note, batchId);
			return jsonOut_({ ok: true, action: 'savePortfolioBatch', result: result });
		}

		if (action === 'updateportfoliobatch' || action === 'update_portfolio_batch') {
			const changeId = String((body && body.changeId) || params.changeId || '').trim();
			const userId = String((body && body.userId) || params.userId || '').trim();
			const effectiveDate = String((body && body.effectiveDate) || params.effectiveDate || '').trim();
			const entriesJson = String((body && body.entriesJson) || params.entriesJson || '[]');
			const note = String((body && body.note) || params.note || '').trim();
			const batchId = String((body && body.batchId) || params.batchId || '').trim();
			const result = updatePortfolioBatch_(changeId, userId, effectiveDate, entriesJson, note, batchId);
			return jsonOut_({ ok: true, action: 'updatePortfolioBatch', result: result });
		}

		if (action === 'deleteportfoliobatch' || action === 'delete_portfolio_batch') {
			const changeId = String((body && body.changeId) || params.changeId || '').trim();
			const userId = String((body && body.userId) || params.userId || '').trim();
			const result = deletePortfolioBatch_(changeId, userId);
			return jsonOut_({ ok: true, action: 'deletePortfolioBatch', result: result });
		}

		if (action === 'getportfoliohistory' || action === 'get_portfolio_history') {
			const userId = String((body && body.userId) || params.userId || '').trim();
			const limit = toInt_((body && body.limit) || params.limit, 500);
			const result = getPortfolioHistory_(userId, limit);
			return jsonOut_({ ok: true, action: 'getPortfolioHistory', result: result });
		}

		const since = String((body && body.since) || params.since || '').trim();
		const limit = toInt_((body && body.limit) || params.limit, 5000);
		const result = readData_(since, limit);
		return jsonOut_({ ok: true, action: 'data', result: result });
	} catch (error) {
		return jsonOut_({ ok: false, error: error.message || String(error) });
	}
}

function validateConfig_() {
	if (!GAS_CONNECT_CONFIG.sheetId) {
		throw new Error('Missing script property: SHEET_ID');
	}
}

function resolveSyncStart_(body, params) {
	const hasExplicitStartYear = (body && body.startYear != null) || params.startYear != null;
	const hasExplicitStartMonth = (body && body.startMonth != null) || params.startMonth != null;

	if (hasExplicitStartYear || hasExplicitStartMonth) {
		return {
			startYear: toInt_((body && body.startYear) || params.startYear, 1998),
			startMonth: toInt_((body && body.startMonth) || params.startMonth, 1)
		};
	}

	const latest = getLatestDataMonthYear_();
	if (latest) {
		return latest;
	}

	return { startYear: 1998, startMonth: 1 };
}

function getLatestDataMonthYear_() {
	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const sheet = ensureDataSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, buildHeaders_());
	const dateCol = requireHeaderColumn_(schema.headerMap, 'date');
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return null;

	const values = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();
	let latestDate = null;

	for (let i = 0; i < values.length; i++) {
		const parsed = parseDateCellValue_(values[i][0]);
		if (!parsed) continue;

		if (!latestDate || parsed.dateObj.getTime() > latestDate.getTime()) {
			latestDate = parsed.dateObj;
		}
	}

	if (!latestDate) return null;
	return {
		startYear: latestDate.getFullYear(),
		startMonth: latestDate.getMonth() + 1
	};
}

function isAuthorizedSync_(incomingToken) {
	if (!GAS_CONNECT_CONFIG.syncToken) return true;
	return incomingToken && incomingToken === GAS_CONNECT_CONFIG.syncToken;
}

function syncToSheet_(startYear, startMonth, endYear, endMonth, options) {
	if (startYear > endYear || (startYear === endYear && startMonth > endMonth)) {
		throw new Error('Invalid start/end month-year range');
	}

	const opts = options || {};
	const reconcile = !!opts.reconcile;

	const sheet = getOrCreateSheet_();
	let existingIndex = buildDateIndex_(sheet);

	let month = startMonth;
	let year = startYear;

	let fetchedMonths = 0;
	let failedMonths = 0;
	let emptyMonths = 0;
	let insertedRows = 0;
	let updatedRows = 0;
	let deletedRows = 0;
	const failedSamples = [];

	while (year < endYear || (year === endYear && month <= endMonth)) {
		try {
			const records = fetchMonthData_(month, year);
			if (!records || !records.length) {
				emptyMonths += 1;
			}
			const upserted = upsertRows_(sheet, existingIndex, records);

			if (reconcile && records && records.length) {
				const validDateMap = {};
				for (let i = 0; i < records.length; i++) {
					const dateText = String(records[i].date || '').trim();
					if (dateText) validDateMap[dateText] = true;
				}
				const removed = deleteStaleRowsForMonth_(sheet, year, month, validDateMap);
				const removedDuplicates = deleteDuplicateRowsForMonth_(sheet, year, month);
				deletedRows += removed + removedDuplicates;
				if (removed > 0 || removedDuplicates > 0) {
					existingIndex = buildDateIndex_(sheet);
				}
			}

			fetchedMonths += 1;
			insertedRows += upserted.inserted;
			updatedRows += upserted.updated;
		} catch (error) {
			failedMonths += 1;
			if (failedSamples.length < 10) {
				failedSamples.push({
					year: year,
					month: month,
					error: error && error.message ? String(error.message) : String(error)
				});
			}
		}

		month += 1;
		if (month > 12) {
			month = 1;
			year += 1;
		}
	}

	const syncSummary =
		'OK:' +
		fetchedMonths +
		', FAIL:' +
		failedMonths +
		', EMPTY:' +
		emptyMonths +
		', INSERT:' +
		insertedRows +
		', DELETE:' +
		deletedRows +
		', UPDATE:' +
		updatedRows;
	setConfigValue_('LAST_SYNC_AT', new Date().toISOString(), 'Last successful sync timestamp');
	setConfigValue_('LAST_SYNC_STATUS', syncSummary, 'Last sync status summary');

	return {
		fetchedMonths: fetchedMonths,
		failedMonths: failedMonths,
		emptyMonths: emptyMonths,
		insertedRows: insertedRows,
		updatedRows: updatedRows,
		deletedRows: deletedRows,
		reconcile: reconcile,
		failedSamples: failedSamples,
		updatedAt: new Date().toISOString()
	};
}

function deleteStaleRowsForMonth_(sheet, year, month, validDateMap) {
	if (!validDateMap) return 0;
	const schema = ensureSheetHeaders_(sheet, buildHeaders_());
	const dateCol = requireHeaderColumn_(schema.headerMap, 'date');
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return 0;

	const values = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();
	const rowsToDelete = [];

	for (let i = 0; i < values.length; i++) {
		const parsed = parseDateCellValue_(values[i][0]);
		if (!parsed) continue;
		if (parsed.year !== year || parsed.month !== month) continue;
		if (!validDateMap[parsed.label]) {
			rowsToDelete.push(i + 2);
		}
	}

	rowsToDelete.sort(function (a, b) {
		return b - a;
	});

	for (let i = 0; i < rowsToDelete.length; i++) {
		sheet.deleteRow(rowsToDelete[i]);
	}

	return rowsToDelete.length;
}

function deleteDuplicateRowsForMonth_(sheet, year, month) {
	const schema = ensureSheetHeaders_(sheet, buildHeaders_());
	const dateCol = requireHeaderColumn_(schema.headerMap, 'date');
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return 0;

	const values = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();
	const seenDate = {};
	const rowsToDelete = [];

	// Traverse bottom-up and keep latest row for each date.
	for (let i = values.length - 1; i >= 0; i--) {
		const parsed = parseDateCellValue_(values[i][0]);
		if (!parsed) continue;
		if (parsed.year !== year || parsed.month !== month) continue;

		const key = parsed.label;
		if (seenDate[key]) {
			rowsToDelete.push(i + 2);
		} else {
			seenDate[key] = true;
		}
	}

	rowsToDelete.sort(function (a, b) {
		return b - a;
	});

	for (let i = 0; i < rowsToDelete.length; i++) {
		sheet.deleteRow(rowsToDelete[i]);
	}

	return rowsToDelete.length;
}

function ensureProjectInitialized_() {
	if (typeof initializeProjectSheets === 'function') {
		// If gas-sheet.js exists in the same GAS project, use its canonical initializer.
		return initializeProjectSheets();
	}

	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const dataSheet = ensureDataSheetInitialized_(spreadsheet);
	const configSheet = ensureConfigSheetInitialized_(spreadsheet);
	const usersSheet = ensureUsersSheetInitialized_(spreadsheet);
	const portfolioSheet = ensurePortfolioSheetInitialized_(spreadsheet);
	ensureDefaultProjectConfig_(configSheet);

	return {
		ok: true,
		spreadsheetId: spreadsheet.getId(),
		dataSheet: dataSheet.getName(),
		configSheet: configSheet.getName(),
		usersSheet: usersSheet.getName(),
		portfolioSheet: portfolioSheet.getName(),
		initializedAt: new Date().toISOString()
	};
}

function ensureDataSheetInitialized_(spreadsheet) {
	let sheet = spreadsheet.getSheetByName(GAS_CONNECT_CONFIG.sheetName);
	if (!sheet) {
		sheet = spreadsheet.insertSheet(GAS_CONNECT_CONFIG.sheetName);
	}

	const schema = ensureSheetHeaders_(sheet, buildHeaders_());
	gcApplyDataSheetStyle_(sheet, schema.headers, schema.headerMap);
	return sheet;
}

function ensureConfigSheetInitialized_(spreadsheet) {
	let sheet = spreadsheet.getSheetByName(GAS_CONNECT_CONFIG.configSheetName);
	if (!sheet) {
		sheet = spreadsheet.insertSheet(GAS_CONNECT_CONFIG.configSheetName);
	}

	const schema = ensureSheetHeaders_(sheet, ['key', 'value', 'description', 'updatedAt']);
	gcApplyConfigSheetStyle_(sheet, schema.headers, schema.headerMap);
	return sheet;
}

function ensureUsersSheetInitialized_(spreadsheet) {
	let sheet = spreadsheet.getSheetByName(GAS_CONNECT_CONFIG.usersSheetName);
	if (!sheet) {
		sheet = spreadsheet.insertSheet(GAS_CONNECT_CONFIG.usersSheetName);
	}

	const schema = ensureSheetHeaders_(sheet, GAS_CONNECT_USER_HEADERS);
	sheet.setFrozenRows(1);
	sheet.getRange(1, 1, 1, schema.headers.length).setFontWeight('bold').setBackground('#dcfce7');
	return sheet;
}

function ensurePortfolioSheetInitialized_(spreadsheet) {
	let sheet = spreadsheet.getSheetByName(GAS_CONNECT_CONFIG.portfolioSheetName);
	if (!sheet) {
		sheet = spreadsheet.insertSheet(GAS_CONNECT_CONFIG.portfolioSheetName);
	}

	const schema = ensureSheetHeaders_(sheet, [
		'changeId',
		'batchId',
		'userId',
		'username',
		'effectiveDate',
		'entriesJson',
		'note',
		'createdAt',
		'updatedAt'
	]);
	sheet.setFrozenRows(1);
	sheet.getRange(1, 1, 1, schema.headers.length).setFontWeight('bold').setBackground('#ffedd5');
	return sheet;
}

function syncPortfolioUsernameByUserId_(spreadsheet, userId, username) {
	const normalizedUserId = String(userId || '').trim();
	const normalizedUsername = String(username || '').trim();
	if (!normalizedUserId || !normalizedUsername) return 0;

	const sheet = ensurePortfolioSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, [
		'changeId',
		'batchId',
		'userId',
		'username',
		'effectiveDate',
		'entriesJson',
		'note',
		'createdAt',
		'updatedAt'
	]);
	const headerMap = schema.headerMap;
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return 0;

	const rows = sheet.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();
	let changed = 0;
	for (let i = 0; i < rows.length; i++) {
		const row = rows[i];
		const rowUserId = String(getCellByHeader_(row, headerMap, 'userId') || '').trim();
		if (rowUserId !== normalizedUserId) continue;
		const rowUsername = String(getCellByHeader_(row, headerMap, 'username') || '').trim();
		if (rowUsername === normalizedUsername) continue;
		setCellByHeader_(row, headerMap, 'username', normalizedUsername);
		setCellByHeader_(row, headerMap, 'updatedAt', new Date().toISOString());
		rows[i] = row;
		changed++;
	}

	if (changed > 0) {
		sheet.getRange(2, 1, rows.length, schema.headers.length).setValues(rows);
	}

	return changed;
}

function reassignPortfolioUserIdByUsername_(spreadsheet, oldUserId, newUserId, username) {
	const normalizedOldUserId = String(oldUserId || '').trim();
	const normalizedNewUserId = String(newUserId || '').trim();
	const normalizedUsername = String(username || '').trim().toLowerCase();
	if (!normalizedOldUserId || !normalizedNewUserId || !normalizedUsername) return 0;
	if (normalizedOldUserId === normalizedNewUserId) return 0;

	const sheet = ensurePortfolioSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, [
		'changeId',
		'batchId',
		'userId',
		'username',
		'effectiveDate',
		'entriesJson',
		'note',
		'createdAt',
		'updatedAt'
	]);
	const headerMap = schema.headerMap;
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return 0;

	const rows = sheet.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();
	let changed = 0;
	for (let i = 0; i < rows.length; i++) {
		const row = rows[i];
		const rowUserId = String(getCellByHeader_(row, headerMap, 'userId') || '').trim();
		const rowUsername = String(getCellByHeader_(row, headerMap, 'username') || '').trim().toLowerCase();
		if (rowUserId !== normalizedOldUserId || rowUsername !== normalizedUsername) {
			continue;
		}

		setCellByHeader_(row, headerMap, 'userId', normalizedNewUserId);
		setCellByHeader_(row, headerMap, 'updatedAt', new Date().toISOString());
		rows[i] = row;
		changed++;
	}

	if (changed > 0) {
		sheet.getRange(2, 1, rows.length, schema.headers.length).setValues(rows);
	}

	return changed;
}

function generateGoogleUserIdBase_(googleId) {
	return 'G_' + sha256Text_(String(googleId || '').trim()).slice(0, 20);
}

function collectExistingUserIds_(sheet, schema) {
	const headerMap = schema.headerMap;
	const out = {};
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return out;

	const rows = sheet.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();
	for (let i = 0; i < rows.length; i++) {
		const id = String(getCellByHeader_(rows[i], headerMap, 'userId') || '').trim();
		if (id) out[id] = true;
	}
	return out;
}

function makeUniqueUserId_(existingIdMap, preferredId) {
	const base = String(preferredId || '').trim() || ('U' + String(Date.now()));
	if (!existingIdMap[base]) {
		existingIdMap[base] = true;
		return base;
	}

	let index = 2;
	while (existingIdMap[base + '_' + index]) {
		index++;
	}
	const uniqueId = base + '_' + index;
	existingIdMap[uniqueId] = true;
	return uniqueId;
}

function repairDuplicateUserIds_(spreadsheet) {
	const usersSheet = ensureUsersSheetInitialized_(spreadsheet);
	const usersSchema = ensureSheetHeaders_(usersSheet, GAS_CONNECT_USER_HEADERS);
	const headerMap = usersSchema.headerMap;
	const lastRow = usersSheet.getLastRow();
	if (lastRow <= 1) return 0;

	const rows = usersSheet.getRange(2, 1, lastRow - 1, usersSchema.headers.length).getValues();
	const existingIdMap = {};
	let maxNumericUserId = 0;
	for (let i = 0; i < rows.length; i++) {
		const id = String(getCellByHeader_(rows[i], headerMap, 'userId') || '').trim();
		if (id) existingIdMap[id] = true;
		const m = id.match(/^U(\d+)$/i);
		if (m) {
			const n = parseInt(m[1], 10);
			if (Number.isFinite(n) && n > maxNumericUserId) maxNumericUserId = n;
		}
	}

	const seen = {};
	let changed = 0;
	for (let i = 0; i < rows.length; i++) {
		const row = rows[i];
		const currentId = String(getCellByHeader_(row, headerMap, 'userId') || '').trim();
		const username = String(getCellByHeader_(row, headerMap, 'username') || '').trim();
		const googleId = String(getCellByHeader_(row, headerMap, 'googleId') || '').trim();

		if (currentId && !seen[currentId]) {
			seen[currentId] = true;
			continue;
		}

		let preferredId = '';
		if (googleId) {
			preferredId = generateGoogleUserIdBase_(googleId);
		} else {
			maxNumericUserId += 1;
			preferredId = 'U' + String(maxNumericUserId).padStart(4, '0');
		}

		const nextId = makeUniqueUserId_(existingIdMap, preferredId);
		setCellByHeader_(row, headerMap, 'userId', nextId);
		setCellByHeader_(row, headerMap, 'updatedAt', new Date().toISOString());
		rows[i] = row;
		seen[nextId] = true;
		changed++;

		if (currentId && username) {
			reassignPortfolioUserIdByUsername_(spreadsheet, currentId, nextId, username);
		}
	}

	if (changed > 0) {
		usersSheet.getRange(2, 1, rows.length, usersSchema.headers.length).setValues(rows);
	}

	return changed;
}

function buildUsersByUsernameIndex_(sheet) {
	const schema = ensureSheetHeaders_(sheet, GAS_CONNECT_USER_HEADERS);
	const headerMap = schema.headerMap;
	const lastRow = sheet.getLastRow();
	const out = {};
	if (lastRow <= 1) return out;

	const values = sheet.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();
	for (let i = 0; i < values.length; i++) {
		const row = values[i];
		const username = String(getCellByHeader_(row, headerMap, 'username') || '').trim();
		if (!username) continue;
		out[username.toLowerCase()] = {
			rowIndex: i + 2,
			userId: String(getCellByHeader_(row, headerMap, 'userId') || '').trim(),
			username: username,
			passwordHash: String(getCellByHeader_(row, headerMap, 'passwordHash') || '').trim(),
			displayName: String(getCellByHeader_(row, headerMap, 'displayName') || '').trim(),
			googleId: String(getCellByHeader_(row, headerMap, 'googleId') || '').trim(),
			role: normalizeUserRole_(getCellByHeader_(row, headerMap, 'role')),
			status: normalizeUserStatus_(getCellByHeader_(row, headerMap, 'status'))
		};
	}

	return out;
}

function getNextUserId_(sheet) {
	const schema = ensureSheetHeaders_(sheet, GAS_CONNECT_USER_HEADERS);
	const headerMap = schema.headerMap;
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) {
		return 'U0001';
	}

	const values = sheet.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();
	let maxId = 0;
	for (let i = 0; i < values.length; i++) {
		const text = String(getCellByHeader_(values[i], headerMap, 'userId') || '').trim();
		const m = text.match(/^U(\d+)$/i);
		if (!m) continue;
		const n = parseInt(m[1], 10);
		if (Number.isFinite(n) && n > maxId) {
			maxId = n;
		}
	}

	return 'U' + String(maxId + 1).padStart(4, '0');
}

function getUserById_(userId) {
	const normalizedId = String(userId || '').trim();
	if (!normalizedId) return null;

	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const sheet = ensureUsersSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, GAS_CONNECT_USER_HEADERS);
	const headerMap = schema.headerMap;
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return null;

	const values = sheet.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();
	for (let i = 0; i < values.length; i++) {
		const row = values[i];
		if (String(getCellByHeader_(row, headerMap, 'userId') || '').trim() !== normalizedId) {
			continue;
		}
		return {
			userId: normalizedId,
			username: String(getCellByHeader_(row, headerMap, 'username') || '').trim(),
			displayName: String(getCellByHeader_(row, headerMap, 'displayName') || '').trim(),
			googleId: String(getCellByHeader_(row, headerMap, 'googleId') || '').trim(),
			role: normalizeUserRole_(getCellByHeader_(row, headerMap, 'role')),
			status: normalizeUserStatus_(getCellByHeader_(row, headerMap, 'status'))
		};
	}

	return null;
}

function registerUser_(username, passwordHash, displayName) {
	const normalizedUsername = String(username || '').trim();
	const normalizedHash = String(passwordHash || '').trim();
	const normalizedDisplay = String(displayName || '').trim();

	if (normalizedUsername.length < 3) {
		throw new Error('Username must be at least 3 characters');
	}
	if (!normalizedHash) {
		throw new Error('passwordHash is required');
	}

	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const sheet = ensureUsersSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, GAS_CONNECT_USER_HEADERS);
	const headerMap = schema.headerMap;
	const indexByUsername = buildUsersByUsernameIndex_(sheet);
	if (indexByUsername[normalizedUsername.toLowerCase()]) {
		throw new Error('Username already exists');
	}

	const userId = getNextUserId_(sheet);
	const payload = createEmptyRowByHeaders_(schema.headers);
	setCellByHeader_(payload, headerMap, 'userId', userId);
	setCellByHeader_(payload, headerMap, 'username', normalizedUsername);
	setCellByHeader_(payload, headerMap, 'passwordHash', normalizedHash);
	setCellByHeader_(payload, headerMap, 'displayName', normalizedDisplay || normalizedUsername);
	setCellByHeader_(payload, headerMap, 'role', 'user');
	setCellByHeader_(payload, headerMap, 'status', 'approved');
	setCellByHeader_(payload, headerMap, 'createdAt', new Date().toISOString());
	setCellByHeader_(payload, headerMap, 'updatedAt', new Date().toISOString());
	sheet.appendRow(payload);

	return {
		userId: userId,
		username: normalizedUsername,
		displayName: normalizedDisplay || normalizedUsername,
		googleId: '',
		role: 'user',
		status: 'approved',
		authType: 'password',
		createdAt: new Date().toISOString()
	};
}

function loginUser_(username, passwordHash) {
	const normalizedUsername = String(username || '').trim();
	const normalizedHash = String(passwordHash || '').trim();
	if (!normalizedUsername || !normalizedHash) {
		throw new Error('username and passwordHash are required');
	}

	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const sheet = ensureUsersSheetInitialized_(spreadsheet);
	const indexByUsername = buildUsersByUsernameIndex_(sheet);
	const found = indexByUsername[normalizedUsername.toLowerCase()];
	if (!found) {
		throw new Error('User not found');
	}
	if (String(found.passwordHash || '') !== normalizedHash) {
		throw new Error('Invalid password');
	}
	if (found.status !== 'approved') {
		throw new Error('Account is pending admin approval');
	}

	return {
		userId: found.userId,
		username: found.username,
		displayName: found.displayName || found.username,
		googleId: String(found.googleId || '').trim(),
		role: found.role,
		status: found.status,
		authType: 'password'
	};
}

function getUserProfile_(userId) {
	const user = requireApprovedUserById_(userId);
	return {
		userId: String(user.userId || '').trim(),
		username: String(user.username || '').trim(),
		displayName: String(user.displayName || user.username || '').trim(),
		googleId: String(user.googleId || '').trim(),
		role: normalizeUserRole_(user.role),
		status: normalizeUserStatus_(user.status)
	};
}

/**
 * Verify a Google ID token and upsert the user in USERS sheet.
 * GAS Script Property: GOOGLE_CLIENT_ID (optional but recommended for audience validation).
 */
function loginWithGoogleToken_(idToken, pendingUserId) {
	const identity = verifyGoogleIdToken_(idToken);
	const googleId = identity.googleId;
	const email = identity.email;
	const displayName = identity.displayName;
	const normalizedPendingUserId = String(pendingUserId || '').trim();
	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	repairDuplicateUserIds_(spreadsheet);
	const sheet = ensureUsersSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, GAS_CONNECT_USER_HEADERS);
	const headerMap = schema.headerMap;
	const nowIso = new Date().toISOString();
	const indexByUsername = buildUsersByUsernameIndex_(sheet);

	const existing = findUserByGoogleId_(sheet, schema, googleId);
	if (existing) {
		const row = sheet.getRange(existing.rowIndex, 1, 1, schema.headers.length).getValues()[0];
		const conflictByEmail = indexByUsername[email.toLowerCase()];
		if (conflictByEmail && conflictByEmail.userId !== existing.userId) {
			throw new Error('Email already belongs to another user');
		}
		const pendingState = getActivePendingEmailState_(row, headerMap);
		setCellByHeader_(row, headerMap, 'username', email);
		setCellByHeader_(row, headerMap, 'displayName', displayName);
		setCellByHeader_(row, headerMap, 'role', normalizeUserRole_(getCellByHeader_(row, headerMap, 'role')));
		setCellByHeader_(row, headerMap, 'status', normalizeUserStatus_(getCellByHeader_(row, headerMap, 'status')));
		setCellByHeader_(row, headerMap, 'updatedAt', nowIso);
		if (!pendingState.active) {
			clearEmailChangeStateInRow_(row, headerMap);
		}
		sheet.getRange(existing.rowIndex, 1, 1, schema.headers.length).setValues([row]);
		const existingStatus = normalizeUserStatus_(getCellByHeader_(row, headerMap, 'status'));
		const existingRole = normalizeUserRole_(getCellByHeader_(row, headerMap, 'role'));
		if (existingStatus !== 'approved') {
			throw new Error('Account is pending admin approval');
		}
		syncPortfolioUsernameByUserId_(spreadsheet, existing.userId, email);
		return {
			userId: existing.userId,
			username: email,
			displayName: displayName,
			googleId: googleId,
			email: email,
			role: existingRole,
			status: existingStatus,
			authType: 'google',
			emailChangePending: pendingState.active,
			pendingEmail: pendingState.active ? pendingState.pendingEmail : '',
			pendingEmailLinkExpiresAt: pendingState.active ? pendingState.expiresAt : ''
		};
	}

	let pendingUser = findUserByPendingEmail_(sheet, schema, email);
	if (!pendingUser && normalizedPendingUserId) {
		const pendingByUserId = findUserByIdInSheet_(sheet, schema, normalizedPendingUserId);
		if (pendingByUserId) {
			const pendingRow = sheet.getRange(pendingByUserId.rowIndex, 1, 1, schema.headers.length).getValues()[0];
			const pendingState = getActivePendingEmailState_(pendingRow, headerMap);
			if (pendingState.active) {
				pendingUser = pendingByUserId;
			}
		}
	}
	if (pendingUser) {
		const row = sheet.getRange(pendingUser.rowIndex, 1, 1, schema.headers.length).getValues()[0];
		const conflictByEmail = indexByUsername[email.toLowerCase()];
		if (conflictByEmail && conflictByEmail.userId !== pendingUser.userId) {
			throw new Error('Email already belongs to another user');
		}
		setCellByHeader_(row, headerMap, 'username', email);
		setCellByHeader_(row, headerMap, 'displayName', displayName);
		setCellByHeader_(row, headerMap, 'googleId', googleId);
		setCellByHeader_(row, headerMap, 'role', normalizeUserRole_(getCellByHeader_(row, headerMap, 'role')));
		setCellByHeader_(row, headerMap, 'status', 'approved');
		setCellByHeader_(row, headerMap, 'updatedAt', nowIso);
		clearEmailChangeStateInRow_(row, headerMap);
		sheet.getRange(pendingUser.rowIndex, 1, 1, schema.headers.length).setValues([row]);
		syncPortfolioUsernameByUserId_(spreadsheet, pendingUser.userId, email);
		return {
			userId: pendingUser.userId,
			username: email,
			displayName: displayName,
			googleId: googleId,
			email: email,
			role: normalizeUserRole_(getCellByHeader_(row, headerMap, 'role')),
			status: 'approved',
			authType: 'google',
			linkedByEmailChange: true
		};
	}

	const foundByEmail = indexByUsername[email.toLowerCase()];
	if (foundByEmail) {
		if (String(foundByEmail.googleId || '').trim()) {
			throw new Error('Email already belongs to another Google account');
		}

		const row = sheet.getRange(foundByEmail.rowIndex, 1, 1, schema.headers.length).getValues()[0];
		setCellByHeader_(row, headerMap, 'displayName', displayName);
		setCellByHeader_(row, headerMap, 'googleId', googleId);
		setCellByHeader_(row, headerMap, 'role', normalizeUserRole_(getCellByHeader_(row, headerMap, 'role')));
		setCellByHeader_(row, headerMap, 'status', normalizeUserStatus_(getCellByHeader_(row, headerMap, 'status')));
		setCellByHeader_(row, headerMap, 'updatedAt', nowIso);
		clearEmailChangeStateInRow_(row, headerMap);
		sheet.getRange(foundByEmail.rowIndex, 1, 1, schema.headers.length).setValues([row]);
		const foundStatus = normalizeUserStatus_(getCellByHeader_(row, headerMap, 'status'));
		const foundRole = normalizeUserRole_(getCellByHeader_(row, headerMap, 'role'));
		if (foundStatus !== 'approved') {
			throw new Error('Account is pending admin approval');
		}
		syncPortfolioUsernameByUserId_(spreadsheet, foundByEmail.userId, email);
		return {
			userId: foundByEmail.userId,
			username: email,
			displayName: displayName,
			googleId: googleId,
			email: email,
			role: foundRole,
			status: foundStatus,
			authType: 'google'
		};
	}

	const existingIdMap = collectExistingUserIds_(sheet, schema);
	const userId = makeUniqueUserId_(existingIdMap, generateGoogleUserIdBase_(googleId));
	const payload = createEmptyRowByHeaders_(schema.headers);
	setCellByHeader_(payload, headerMap, 'userId', userId);
	setCellByHeader_(payload, headerMap, 'username', email);
	setCellByHeader_(payload, headerMap, 'passwordHash', '');
	setCellByHeader_(payload, headerMap, 'displayName', displayName);
	setCellByHeader_(payload, headerMap, 'role', 'user');
	setCellByHeader_(payload, headerMap, 'status', 'pending');
	setCellByHeader_(payload, headerMap, 'createdAt', nowIso);
	setCellByHeader_(payload, headerMap, 'updatedAt', nowIso);
	setCellByHeader_(payload, headerMap, 'googleId', googleId);
	sheet.appendRow(payload);
	throw new Error('Account created and waiting for admin approval');
}

function requireAdminUserById_(userId) {
	const approvedUser = requireApprovedUserById_(userId);
	if (normalizeUserRole_(approvedUser.role) !== 'admin') {
		throw new Error('Admin permission required');
	}
	return approvedUser;
}

function getAdminApprovalDashboard_(adminUserId) {
	requireAdminUserById_(adminUserId);

	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const sheet = ensureUsersSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, GAS_CONNECT_USER_HEADERS);
	const headerMap = schema.headerMap;
	const lastRow = sheet.getLastRow();

	if (lastRow <= 1) {
		return {
			summary: { totalUsers: 0, pending: 0, approved: 0, admins: 0 },
			pendingUsers: [],
			updatedAt: new Date().toISOString()
		};
	}

	const values = sheet.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();
	const pendingUsers = [];
	let totalUsers = 0;
	let approvedCount = 0;
	let pendingCount = 0;
	let adminCount = 0;

	for (let i = 0; i < values.length; i++) {
		const row = values[i];
		const userId = String(getCellByHeader_(row, headerMap, 'userId') || '').trim();
		if (!userId) continue;
		totalUsers++;

		const role = normalizeUserRole_(getCellByHeader_(row, headerMap, 'role'));
		const status = normalizeUserStatus_(getCellByHeader_(row, headerMap, 'status'));
		if (role === 'admin') adminCount++;
		if (status === 'approved') approvedCount++;
		if (status === 'pending') {
			pendingCount++;
			pendingUsers.push({
				userId: userId,
				username: String(getCellByHeader_(row, headerMap, 'username') || '').trim(),
				displayName: String(getCellByHeader_(row, headerMap, 'displayName') || '').trim(),
				role: role,
				status: status,
				createdAt: String(getCellByHeader_(row, headerMap, 'createdAt') || '').trim(),
				updatedAt: String(getCellByHeader_(row, headerMap, 'updatedAt') || '').trim()
			});
		}
	}

	pendingUsers.sort(function (a, b) {
		const aTime = Date.parse(String(a.createdAt || ''));
		const bTime = Date.parse(String(b.createdAt || ''));
		if (Number.isFinite(aTime) && Number.isFinite(bTime) && aTime !== bTime) {
			return aTime - bTime;
		}
		return String(a.userId || '').localeCompare(String(b.userId || ''));
	});

	return {
		summary: {
			totalUsers: totalUsers,
			pending: pendingCount,
			approved: approvedCount,
			admins: adminCount
		},
		pendingUsers: pendingUsers,
		updatedAt: new Date().toISOString()
	};
}

function approveUserRegistration_(adminUserId, targetUserId) {
	const adminUser = requireAdminUserById_(adminUserId);
	const normalizedTargetUserId = String(targetUserId || '').trim();
	if (!normalizedTargetUserId) {
		throw new Error('targetUserId is required');
	}

	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const sheet = ensureUsersSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, GAS_CONNECT_USER_HEADERS);
	const headerMap = schema.headerMap;
	const target = findUserByIdInSheet_(sheet, schema, normalizedTargetUserId);
	if (!target) {
		throw new Error('Target user not found');
	}

	const row = sheet.getRange(target.rowIndex, 1, 1, schema.headers.length).getValues()[0];
	const currentStatus = normalizeUserStatus_(getCellByHeader_(row, headerMap, 'status'));
	if (currentStatus === 'approved') {
		return {
			userId: normalizedTargetUserId,
			username: String(getCellByHeader_(row, headerMap, 'username') || '').trim(),
			displayName: String(getCellByHeader_(row, headerMap, 'displayName') || '').trim(),
			status: 'approved',
			alreadyApproved: true,
			approvedAt: String(getCellByHeader_(row, headerMap, 'updatedAt') || '').trim() || new Date().toISOString(),
			approvedBy: adminUser.userId,
			notification: { attempted: false, sent: false, reason: 'Already approved' }
		};
	}

	setCellByHeader_(row, headerMap, 'status', 'approved');
	setCellByHeader_(row, headerMap, 'updatedAt', new Date().toISOString());
	sheet.getRange(target.rowIndex, 1, 1, schema.headers.length).setValues([row]);

	const approvedUsername = String(getCellByHeader_(row, headerMap, 'username') || '').trim();
	const approvedDisplayName = String(getCellByHeader_(row, headerMap, 'displayName') || '').trim();
	const approvedAt = String(getCellByHeader_(row, headerMap, 'updatedAt') || '').trim();

	let notification = { attempted: false, sent: false, reason: '' };
	if (isValidEmail_(approvedUsername)) {
		notification.attempted = true;
		try {
			sendApprovalNotificationMail_(approvedUsername, approvedDisplayName || approvedUsername);
			notification.sent = true;
		} catch (mailError) {
			notification.sent = false;
			notification.reason = String((mailError && mailError.message) || mailError || 'Email notification failed');
		}
	} else {
		notification.reason = 'Username is not a valid email address';
	}

	return {
		userId: normalizedTargetUserId,
		username: approvedUsername,
		displayName: approvedDisplayName,
		status: 'approved',
		approvedAt: approvedAt,
		approvedBy: adminUser.userId,
		notification: notification
	};
}

function sendApprovalNotificationMail_(toEmail, displayName) {
	const recipient = String(toEmail || '').trim();
	if (!recipient || !isValidEmail_(recipient)) {
		throw new Error('Valid recipient email is required');
	}

	const safeName = String(displayName || recipient).trim();
	const subject = 'บัญชีของคุณได้รับการอนุมัติแล้ว';
	const textBody = [
		'สวัสดี ' + safeName + ',',
		'',
		'บัญชีของคุณได้รับการอนุมัติเรียบร้อยแล้ว',
		'ขณะนี้คุณสามารถเข้าสู่ระบบและใช้งานระบบได้ทันที',
		'',
		'ขอบคุณที่ใช้งานระบบ'
	].join('\n');
	const htmlBody =
		'<p>สวัสดี ' + safeName + ',</p>' +
		'<p>บัญชีของคุณได้รับการอนุมัติเรียบร้อยแล้ว</p>' +
		'<p>ขณะนี้คุณสามารถเข้าสู่ระบบและใช้งานระบบได้ทันที</p>' +
		'<p>ขอบคุณที่ใช้งานระบบ</p>';

	try {
		GmailApp.sendEmail(recipient, subject, textBody, {
			htmlBody: htmlBody,
			name: GAS_CONNECT_CONFIG.mailSenderName
		});
	} catch (_gmailError) {
		MailApp.sendEmail({
			to: recipient,
			subject: subject,
			htmlBody: htmlBody,
			body: textBody,
			name: GAS_CONNECT_CONFIG.mailSenderName
		});
	}
}

function findUserByGoogleId_(sheet, schema, googleId) {
	const headerMap = schema.headerMap;
	if (!headerMap.googleId) return null;

	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return null;

	const values = sheet.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();
	for (let i = 0; i < values.length; i++) {
		const row = values[i];
		const rowGoogleId = String(getCellByHeader_(row, headerMap, 'googleId') || '').trim();
		if (rowGoogleId && rowGoogleId === googleId) {
			return {
				rowIndex: i + 2,
				userId: String(getCellByHeader_(row, headerMap, 'userId') || '').trim(),
				username: String(getCellByHeader_(row, headerMap, 'username') || '').trim(),
				displayName: String(getCellByHeader_(row, headerMap, 'displayName') || '').trim()
			};
		}
	}
	return null;
}

function findUserByPendingEmail_(sheet, schema, email) {
	const normalizedEmail = String(email || '').trim().toLowerCase();
	if (!normalizedEmail) return null;

	const headerMap = schema.headerMap;
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return null;

	const values = sheet.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();
	const nowMs = Date.now();
	for (let i = 0; i < values.length; i++) {
		const row = values[i];
		const pendingEmail = String(getCellByHeader_(row, headerMap, 'pendingEmail') || '').trim().toLowerCase();
		const linkExpiresAt = String(getCellByHeader_(row, headerMap, 'pendingEmailLinkExpiresAt') || '').trim();
		if (!pendingEmail || pendingEmail !== normalizedEmail) continue;
		const expiry = Date.parse(linkExpiresAt);
		if (!Number.isFinite(expiry) || expiry < nowMs) continue;
		return {
			rowIndex: i + 2,
			userId: String(getCellByHeader_(row, headerMap, 'userId') || '').trim(),
			username: String(getCellByHeader_(row, headerMap, 'username') || '').trim(),
			displayName: String(getCellByHeader_(row, headerMap, 'displayName') || '').trim()
		};
	}

	return null;
}

function requestEmailChangeOtp_(userId) {
	const normalizedUserId = String(userId || '').trim();
	if (!normalizedUserId) {
		throw new Error('userId is required');
	}

	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const sheet = ensureUsersSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, GAS_CONNECT_USER_HEADERS);
	const headerMap = schema.headerMap;
	const user = findUserByIdInSheet_(sheet, schema, normalizedUserId);
	if (!user) {
		throw new Error('User not found');
	}
	if (normalizeUserStatus_(user.status) !== 'approved') {
		throw new Error('Account is pending admin approval');
	}

	const currentEmail = String(user.username || '').trim();
	if (!isValidEmail_(currentEmail)) {
		throw new Error('Current account email is invalid or unavailable');
	}

	const row = sheet.getRange(user.rowIndex, 1, 1, schema.headers.length).getValues()[0];
	const lastSentAtText = String(getCellByHeader_(row, headerMap, 'emailChangeLastSentAt') || '').trim();
	const lastSentAtMs = Date.parse(lastSentAtText);
	if (Number.isFinite(lastSentAtMs) && Date.now() - lastSentAtMs < 60000) {
		throw new Error('Please wait at least 60 seconds before requesting another OTP');
	}

	const otpCode = generateOtpCode_();
	const nowIso = new Date().toISOString();
	const expiresAt = new Date(Date.now() + 10 * 60 * 1000).toISOString();
	setCellByHeader_(row, headerMap, 'emailChangeOtpHash', sha256Text_(otpCode));
	setCellByHeader_(row, headerMap, 'emailChangeOtpExpiresAt', expiresAt);
	setCellByHeader_(row, headerMap, 'pendingEmail', '');
	setCellByHeader_(row, headerMap, 'pendingEmailLinkExpiresAt', '');
	setCellByHeader_(row, headerMap, 'emailChangeLastSentAt', nowIso);
	setCellByHeader_(row, headerMap, 'updatedAt', nowIso);
	sheet.getRange(user.rowIndex, 1, 1, schema.headers.length).setValues([row]);

	try {
		sendEmailChangeOtpMail_(currentEmail, otpCode, expiresAt);
	} catch (error) {
		throw buildMailAuthorizationError_(error);
	}

	return {
		userId: normalizedUserId,
		sentTo: maskEmail_(currentEmail),
		expiresAt: expiresAt,
		sentAt: nowIso
	};
}

function sendEmailChangeOtpMail_(toEmail, otpCode, expiresAtIso) {
	const recipient = String(toEmail || '').trim();
	if (!recipient) {
		throw new Error('OTP recipient email is required');
	}

	const expiresAtText = String(expiresAtIso || '').trim();
	const subject = 'รหัสยืนยันการเปลี่ยนอีเมลบัญชีของคุณ';
	const textBody = [
		'คำขอเปลี่ยนอีเมลบัญชี',
		'',
		'รหัสยืนยันของคุณ: ' + String(otpCode || ''),
		'รหัสนี้ใช้ได้ถึง: ' + (expiresAtText || '-'),
		'',
		'หากคุณไม่ได้เป็นผู้ดำเนินการ กรุณาเพิกเฉยอีเมลฉบับนี้'
	].join('\n');
	const htmlBody =
		'<p>คำขอเปลี่ยนอีเมลบัญชี</p>' +
		'<p>รหัสยืนยันของคุณ:</p>' +
		'<p style="font-size:24px;font-weight:700;letter-spacing:4px;margin:4px 0 8px 0">' + String(otpCode || '') + '</p>' +
		'<p>รหัสนี้ใช้ได้ถึง: <strong>' + (expiresAtText || '-') + '</strong></p>' +
		'<p>หากคุณไม่ได้เป็นผู้ดำเนินการ กรุณาเพิกเฉยอีเมลฉบับนี้</p>';

	// Prefer GmailApp for more sender metadata; fallback to MailApp for compatibility.
	try {
		GmailApp.sendEmail(recipient, subject, textBody, {
			htmlBody: htmlBody,
			name: GAS_CONNECT_CONFIG.mailSenderName
		});
	} catch (_gmailError) {
		MailApp.sendEmail({
			to: recipient,
			subject: subject,
			htmlBody: htmlBody,
			body: textBody,
			name: GAS_CONNECT_CONFIG.mailSenderName
		});
	}
}

function buildMailAuthorizationError_(error) {
	const rawMessage = String((error && error.message) || error || 'Mail send failed');
	const authHints = [
		'https://www.googleapis.com/auth/script.send_mail',
		'MailApp.sendEmail',
		'ไม่ได้รับอนุญาต',
		'Authorization is required'
	];
	const isAuthError = authHints.some((hint) => rawMessage.indexOf(hint) >= 0);
	if (!isAuthError) {
		return new Error(rawMessage);
	}

	return new Error(
		'OTP send failed: missing mail permission. In Apps Script editor run function authorizeRequiredScopes() and approve permissions, then redeploy Web App (Execute as: Me). Original error: ' + rawMessage
	);
}

function verifyEmailChangeOtp_(userId, otpCode, newEmail) {
	const normalizedUserId = String(userId || '').trim();
	const normalizedOtpCode = String(otpCode || '').trim();
	if (!normalizedUserId || !normalizedOtpCode) {
		throw new Error('userId and otpCode are required');
	}

	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const sheet = ensureUsersSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, GAS_CONNECT_USER_HEADERS);
	const headerMap = schema.headerMap;
	const user = findUserByIdInSheet_(sheet, schema, normalizedUserId);
	if (!user) {
		throw new Error('User not found');
	}
	if (normalizeUserStatus_(user.status) !== 'approved') {
		throw new Error('Account is pending admin approval');
	}

	const row = sheet.getRange(user.rowIndex, 1, 1, schema.headers.length).getValues()[0];
	const otpHash = String(getCellByHeader_(row, headerMap, 'emailChangeOtpHash') || '').trim();
	const otpExpiresAt = String(getCellByHeader_(row, headerMap, 'emailChangeOtpExpiresAt') || '').trim();
	if (!otpHash || sha256Text_(normalizedOtpCode) !== otpHash) {
		throw new Error('OTP is invalid');
	}
	const otpExpiresAtMs = Date.parse(otpExpiresAt);
	if (!Number.isFinite(otpExpiresAtMs) || otpExpiresAtMs < Date.now()) {
		throw new Error('OTP has expired');
	}

	const nowIso = new Date().toISOString();
	const linkExpiresAt = new Date(Date.now() + 30 * 60 * 1000).toISOString();
	setCellByHeader_(row, headerMap, 'pendingEmail', '__ANY_GOOGLE_EMAIL__');
	setCellByHeader_(row, headerMap, 'pendingEmailLinkExpiresAt', linkExpiresAt);
	setCellByHeader_(row, headerMap, 'emailChangeOtpHash', '');
	setCellByHeader_(row, headerMap, 'emailChangeOtpExpiresAt', '');
	setCellByHeader_(row, headerMap, 'updatedAt', nowIso);
	sheet.getRange(user.rowIndex, 1, 1, schema.headers.length).setValues([row]);

	return {
		userId: normalizedUserId,
		pendingEmail: '',
		pendingByGoogleLoginOnly: true,
		linkExpiresAt: linkExpiresAt,
		message: 'OTP verified. Please log out and sign in with the target Google account before the link expires.'
	};
}

function verifyGoogleIdToken_(idToken) {
	const normalizedToken = String(idToken || '').trim();
	if (!normalizedToken) {
		throw new Error('idToken is required');
	}

	const tokeninfoUrl =
		'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(normalizedToken);
	const response = UrlFetchApp.fetch(tokeninfoUrl, {
		method: 'get',
		muteHttpExceptions: true
	});

	const code = response.getResponseCode();
	if (code !== 200) {
		throw new Error('Invalid Google ID token (status ' + code + ')');
	}

	let tokenInfo;
	try {
		tokenInfo = JSON.parse(response.getContentText('utf-8'));
	} catch (_error) {
		throw new Error('Failed to parse token info from Google');
	}

	if (GAS_CONNECT_CONFIG.googleClientId && tokenInfo.aud !== GAS_CONNECT_CONFIG.googleClientId) {
		throw new Error('Token audience mismatch');
	}
	if (String(tokenInfo.email_verified || '') !== 'true') {
		throw new Error('Google email not verified');
	}

	const googleId = String(tokenInfo.sub || '').trim();
	const email = String(tokenInfo.email || '').trim().toLowerCase();
	const displayName = String(tokenInfo.name || tokenInfo.given_name || email).trim();
	if (!googleId || !email) {
		throw new Error('Missing required fields (sub, email) from Google token info');
	}

	return {
		googleId: googleId,
		email: email,
		displayName: displayName
	};
}

function findUserByIdInSheet_(sheet, schema, userId) {
	const normalizedUserId = String(userId || '').trim();
	if (!normalizedUserId) return null;

	const headerMap = schema.headerMap;
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return null;

	const values = sheet.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();
	for (let i = 0; i < values.length; i++) {
		const row = values[i];
		if (String(getCellByHeader_(row, headerMap, 'userId') || '').trim() !== normalizedUserId) {
			continue;
		}
		return {
			rowIndex: i + 2,
			userId: normalizedUserId,
			username: String(getCellByHeader_(row, headerMap, 'username') || '').trim(),
			displayName: String(getCellByHeader_(row, headerMap, 'displayName') || '').trim(),
			googleId: String(getCellByHeader_(row, headerMap, 'googleId') || '').trim(),
			role: normalizeUserRole_(getCellByHeader_(row, headerMap, 'role')),
			status: normalizeUserStatus_(getCellByHeader_(row, headerMap, 'status'))
		};
	}

	return null;
}

function clearEmailChangeStateInRow_(row, headerMap) {
	setCellByHeader_(row, headerMap, 'emailChangeOtpHash', '');
	setCellByHeader_(row, headerMap, 'emailChangeOtpExpiresAt', '');
	setCellByHeader_(row, headerMap, 'pendingEmail', '');
	setCellByHeader_(row, headerMap, 'pendingEmailLinkExpiresAt', '');
}

function getActivePendingEmailState_(row, headerMap) {
	const pendingEmail = String(getCellByHeader_(row, headerMap, 'pendingEmail') || '').trim().toLowerCase();
	const expiresAt = String(getCellByHeader_(row, headerMap, 'pendingEmailLinkExpiresAt') || '').trim();
	const expiresAtMs = Date.parse(expiresAt);
	const active = !!pendingEmail && Number.isFinite(expiresAtMs) && expiresAtMs >= Date.now();
	return {
		active: active,
		pendingEmail: pendingEmail,
		expiresAt: expiresAt
	};
}

function generateOtpCode_() {
	return String(Math.floor(100000 + Math.random() * 900000));
}

function sha256Text_(text) {
	const digest = Utilities.computeDigest(
		Utilities.DigestAlgorithm.SHA_256,
		String(text || ''),
		Utilities.Charset.UTF_8
	);
	let out = '';
	for (let i = 0; i < digest.length; i++) {
		const v = (digest[i] + 256) % 256;
		out += ('0' + v.toString(16)).slice(-2);
	}
	return out;
}

function isValidEmail_(email) {
	return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || '').trim());
}

function normalizeUserRole_(value) {
	const role = String(value || '').trim().toLowerCase();
	if (role === 'admin') return 'admin';
	return 'user';
}

function normalizeUserStatus_(value) {
	const status = String(value || '').trim().toLowerCase();
	if (status === 'pending') return 'pending';
	if (status === 'approve' || status === 'approved') return 'approved';
	return 'approved';
}

function requireApprovedUserById_(userId) {
	const user = getUserById_(userId);
	if (!user) {
		throw new Error('User not found for userId: ' + String(userId || '').trim());
	}
	if (normalizeUserStatus_(user.status) !== 'approved') {
		throw new Error('Account is pending admin approval');
	}
	return user;
}

function maskEmail_(email) {
	const text = String(email || '').trim();
	const parts = text.split('@');
	if (parts.length !== 2) return text;
	const local = parts[0];
	if (local.length <= 2) {
		return local.charAt(0) + '*@' + parts[1];
	}
	return local.slice(0, 2) + '***@' + parts[1];
}

function normalizePortfolioEntries_(entriesJsonText) {
	let parsed;
	try {
		parsed = JSON.parse(String(entriesJsonText || '[]'));
	} catch (_error) {
		throw new Error('Invalid entriesJson');
	}

	if (!Array.isArray(parsed) || parsed.length === 0) {
		throw new Error('entriesJson must be a non-empty array');
	}

	const out = [];
	for (let i = 0; i < parsed.length; i++) {
		const item = parsed[i] || {};
		const key = String(item.unitCostKey || '').trim();
		if (!/^unitCost\d+$/i.test(key)) {
			throw new Error('Invalid unitCostKey at index ' + i);
		}
		const units = toNumberOrNull_(item.units);
		if (!Number.isFinite(units) || units < 0) {
			throw new Error('Invalid units at index ' + i);
		}
		out.push({
			unitCostKey: key,
			planName: String(item.planName || '').trim(),
			units: units
		});
	}

	return out;
}

function savePortfolioBatch_(userId, effectiveDate, entriesJson, note, batchId) {
	const normalizedUserId = String(userId || '').trim();
	const normalizedDate = String(effectiveDate || '').trim();
	const normalizedNote = String(note || '').trim();
	if (!normalizedUserId) {
		throw new Error('userId is required');
	}
	if (!/^\d{4}-\d{2}-\d{2}$/.test(normalizedDate)) {
		throw new Error('effectiveDate must be YYYY-MM-DD');
	}

	const user = requireApprovedUserById_(normalizedUserId);

	const normalizedEntries = normalizePortfolioEntries_(entriesJson);
	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const sheet = ensurePortfolioSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, [
		'changeId',
		'batchId',
		'userId',
		'username',
		'effectiveDate',
		'entriesJson',
		'note',
		'createdAt',
		'updatedAt'
	]);
	const headerMap = schema.headerMap;

	const nowIso = new Date().toISOString();
	const payload = createEmptyRowByHeaders_(schema.headers);
	const finalBatchId = String(batchId || '').trim() || ('B' + normalizedDate.replace(/-/g, '') + '-' + String(Date.now()).slice(-5));
	setCellByHeader_(payload, headerMap, 'changeId', 'PC' + String(Date.now()));
	setCellByHeader_(payload, headerMap, 'batchId', finalBatchId);
	setCellByHeader_(payload, headerMap, 'userId', normalizedUserId);
	setCellByHeader_(payload, headerMap, 'username', user.username);
	setCellByHeader_(payload, headerMap, 'effectiveDate', normalizedDate);
	setCellByHeader_(payload, headerMap, 'entriesJson', JSON.stringify(normalizedEntries));
	setCellByHeader_(payload, headerMap, 'note', normalizedNote);
	setCellByHeader_(payload, headerMap, 'createdAt', nowIso);
	setCellByHeader_(payload, headerMap, 'updatedAt', nowIso);
	sheet.appendRow(payload);

	return {
		changeId: getCellByHeader_(payload, headerMap, 'changeId'),
		batchId: finalBatchId,
		userId: normalizedUserId,
		effectiveDate: normalizedDate,
		entriesJson: JSON.stringify(normalizedEntries),
		entriesCount: normalizedEntries.length,
		note: normalizedNote,
		createdAt: nowIso
	};
}

function findPortfolioChangeRow_(sheet, headerMap, changeId, userId, batchId) {
	const normalizedChangeId = String(changeId || '').trim();
	const normalizedUserId = String(userId || '').trim();
	const normalizedBatchId = String(batchId || '').trim();
	if ((!normalizedChangeId && !normalizedBatchId) || !normalizedUserId) return null;

	const schema = ensureSheetHeaders_(sheet, [
		'changeId',
		'batchId',
		'userId',
		'username',
		'effectiveDate',
		'entriesJson',
		'note',
		'createdAt',
		'updatedAt'
	]);
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return null;

	const values = sheet.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();
	for (let i = 0; i < values.length; i++) {
		const row = values[i];
		const rowChangeId = String(getCellByHeader_(row, headerMap, 'changeId') || '').trim();
		const rowBatchId = String(getCellByHeader_(row, headerMap, 'batchId') || '').trim();
		const rowUserId = String(getCellByHeader_(row, headerMap, 'userId') || '').trim();
		const matchedByChangeId = normalizedChangeId && rowChangeId === normalizedChangeId;
		const matchedByBatchId = !normalizedChangeId && normalizedBatchId && rowBatchId === normalizedBatchId;
		if ((matchedByChangeId || matchedByBatchId) && rowUserId === normalizedUserId) {
			return {
				rowIndex: i + 2,
				rowValues: row
			};
		}
	}

	return null;
}

function updatePortfolioBatch_(changeId, userId, effectiveDate, entriesJson, note, batchId) {
	const normalizedChangeId = String(changeId || '').trim();
	const normalizedUserId = String(userId || '').trim();
	const normalizedDate = String(effectiveDate || '').trim();
	const normalizedNote = String(note || '').trim();
	const normalizedBatchId = String(batchId || '').trim();
	if ((!normalizedChangeId && !normalizedBatchId) || !normalizedUserId) {
		throw new Error('changeId or batchId and userId are required');
	}
	if (!/^\d{4}-\d{2}-\d{2}$/.test(normalizedDate)) {
		throw new Error('effectiveDate must be YYYY-MM-DD');
	}

	const normalizedEntries = normalizePortfolioEntries_(entriesJson);
	const user = requireApprovedUserById_(normalizedUserId);
	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const sheet = ensurePortfolioSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, [
		'changeId',
		'batchId',
		'userId',
		'username',
		'effectiveDate',
		'entriesJson',
		'note',
		'createdAt',
		'updatedAt'
	]);
	const headerMap = schema.headerMap;

	const found = findPortfolioChangeRow_(sheet, headerMap, normalizedChangeId, normalizedUserId, normalizedBatchId);
	if (!found) {
		throw new Error('Portfolio batch not found');
	}

	const row = found.rowValues.slice();
	if (!String(getCellByHeader_(row, headerMap, 'changeId') || '').trim()) {
		setCellByHeader_(row, headerMap, 'changeId', 'PC' + String(Date.now()));
	}
	setCellByHeader_(row, headerMap, 'batchId', String(batchId || '').trim() || String(getCellByHeader_(row, headerMap, 'batchId') || ''));
	if (user && String(user.username || '').trim()) {
		setCellByHeader_(row, headerMap, 'username', String(user.username || '').trim());
	}
	setCellByHeader_(row, headerMap, 'effectiveDate', normalizedDate);
	setCellByHeader_(row, headerMap, 'entriesJson', JSON.stringify(normalizedEntries));
	setCellByHeader_(row, headerMap, 'note', normalizedNote);
	setCellByHeader_(row, headerMap, 'updatedAt', new Date().toISOString());
	sheet.getRange(found.rowIndex, 1, 1, schema.headers.length).setValues([row]);

	return {
		changeId: String(getCellByHeader_(row, headerMap, 'changeId') || ''),
		batchId: String(getCellByHeader_(row, headerMap, 'batchId') || ''),
		userId: normalizedUserId,
		effectiveDate: normalizedDate,
		entriesJson: JSON.stringify(normalizedEntries),
		note: normalizedNote,
		updatedAt: String(getCellByHeader_(row, headerMap, 'updatedAt') || '')
	};
}

function deletePortfolioBatch_(changeId, userId) {
	const normalizedChangeId = String(changeId || '').trim();
	const normalizedUserId = String(userId || '').trim();
	if (!normalizedChangeId || !normalizedUserId) {
		throw new Error('changeId and userId are required');
	}
	requireApprovedUserById_(normalizedUserId);

	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const sheet = ensurePortfolioSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, [
		'changeId',
		'batchId',
		'userId',
		'username',
		'effectiveDate',
		'entriesJson',
		'note',
		'createdAt',
		'updatedAt'
	]);
	const headerMap = schema.headerMap;

	const found = findPortfolioChangeRow_(sheet, headerMap, normalizedChangeId, normalizedUserId);
	if (!found) {
		throw new Error('Portfolio batch not found');
	}

	sheet.deleteRow(found.rowIndex);
	return {
		changeId: normalizedChangeId,
		userId: normalizedUserId,
		deleted: true,
		deletedAt: new Date().toISOString()
	};
}

function getPortfolioHistory_(userId, limit) {
	const normalizedUserId = String(userId || '').trim();
	if (!normalizedUserId) {
		throw new Error('userId is required');
	}
	requireApprovedUserById_(normalizedUserId);

	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const sheet = ensurePortfolioSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(sheet, [
		'changeId',
		'batchId',
		'userId',
		'username',
		'effectiveDate',
		'entriesJson',
		'note',
		'createdAt',
		'updatedAt'
	]);
	const headerMap = schema.headerMap;
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) {
		return { userId: normalizedUserId, count: 0, batches: [] };
	}

	const rows = sheet.getRange(2, 1, lastRow - 1, schema.headers.length).getValues();
	const out = [];
	for (let i = 0; i < rows.length; i++) {
		const row = rows[i];
		if (String(getCellByHeader_(row, headerMap, 'userId') || '').trim() !== normalizedUserId) {
			continue;
		}

		const entriesText = String(getCellByHeader_(row, headerMap, 'entriesJson') || '[]');
		let parsedEntries = [];
		try {
			parsedEntries = JSON.parse(entriesText);
		} catch (_error) {
			parsedEntries = [];
		}

		out.push({
			changeId: String(getCellByHeader_(row, headerMap, 'changeId') || ''),
			batchId: String(getCellByHeader_(row, headerMap, 'batchId') || ''),
			userId: normalizedUserId,
			username: String(getCellByHeader_(row, headerMap, 'username') || ''),
			effectiveDate: String(getCellByHeader_(row, headerMap, 'effectiveDate') || ''),
			entriesJson: entriesText,
			entries: Array.isArray(parsedEntries) ? parsedEntries : [],
			note: String(getCellByHeader_(row, headerMap, 'note') || ''),
			createdAt: String(getCellByHeader_(row, headerMap, 'createdAt') || ''),
			updatedAt: String(getCellByHeader_(row, headerMap, 'updatedAt') || '')
		});
	}

	out.sort(function (a, b) {
		if (a.effectiveDate !== b.effectiveDate) {
			return String(b.effectiveDate).localeCompare(String(a.effectiveDate));
		}
		return String(b.createdAt).localeCompare(String(a.createdAt));
	});

	const normalizedLimit = toInt_(limit, 500);
	const sliced = normalizedLimit > 0 ? out.slice(0, normalizedLimit) : out;
	return {
		userId: normalizedUserId,
		count: sliced.length,
		batches: sliced,
		generatedAt: new Date().toISOString()
	};
}

function ensureDefaultProjectConfig_(configSheet) {
	const schema = ensureSheetHeaders_(configSheet, ['key', 'value', 'description', 'updatedAt']);
	const headers = schema.headers;
	const headerMap = schema.headerMap;
	const keyIndex = gcBuildConfigKeyIndex_(configSheet);

	for (let i = 0; i < GAS_CONNECT_PROJECT_CONFIG_DEFAULTS.length; i++) {
		const item = GAS_CONNECT_PROJECT_CONFIG_DEFAULTS[i];
		const rowIndex = keyIndex[item.key] || -1;
		const payload = createEmptyRowByHeaders_(headers);
		setCellByHeader_(payload, headerMap, 'key', item.key);
		setCellByHeader_(payload, headerMap, 'description', item.description);
		setCellByHeader_(payload, headerMap, 'updatedAt', new Date().toISOString());

		if (rowIndex > 0) {
			const existingRow = configSheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
			setCellByHeader_(
				payload,
				headerMap,
				'value',
				getCellByHeader_(existingRow, headerMap, 'value')
			);
			configSheet.getRange(rowIndex, 1, 1, headers.length).setValues([payload]);
		} else {
			setCellByHeader_(payload, headerMap, 'value', item.value);
			configSheet.appendRow(payload);
		}
	}
}

function fetchMonthData_(month, year) {
	const monthStr = String(month).padStart(2, '0');
	const yearStr = String(year);
	const url = GAS_CONNECT_CONFIG.apiUrlTemplate.replace('{MM}', monthStr).replace('{YYYY}', yearStr);

	const response = UrlFetchApp.fetch(url, {
		method: 'get',
		muteHttpExceptions: true,
		followRedirects: true,
		headers: {
			'User-Agent':
				'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36',
			Accept: 'application/json,text/plain,*/*',
			Referer: 'https://www.gpf.or.th/'
		}
	});

	const code = response.getResponseCode();
	if (code < 200 || code >= 300) {
		throw new Error('Upstream status ' + code + ' for ' + monthStr + '/' + yearStr);
	}

	const text = response.getContentText('utf-8');
	if (!text || !text.trim()) {
		return [];
	}

	const payload = parseApiPayload_(text);
	return mapApiRows_(payload);
}

function parseApiPayload_(text) {
	try {
		return JSON.parse(text);
	} catch (error) {
		const match = text.match(/\[\s*\{[\s\S]*\}\s*\]/);
		if (!match) {
			throw new Error('Unexpected API payload format');
		}
		return JSON.parse(match[0]);
	}
}

function mapApiRows_(rows) {
	if (!Array.isArray(rows)) return [];

	const out = [];
	for (let i = 0; i < rows.length; i++) {
		const row = rows[i] || {};
		const parsed = parseDateString_(String(row.LAUNCH_DATE || ''));
		if (!parsed) continue;

		const unitCosts = [];
		for (let n = 1; n <= GAS_CONNECT_CONFIG.unitCostCount; n++) {
			unitCosts.push(toNumberOrNull_(row['UNIT_COST' + n]));
		}

		out.push({
			date: parsed.label,
			year: parsed.year,
			month: parsed.month,
			unitCosts: unitCosts,
			updatedAt: new Date().toISOString()
		});
	}

	// Deduplicate by date and keep last occurrence.
	const byDate = {};
	for (let i = 0; i < out.length; i++) {
		byDate[out[i].date] = out[i];
	}

	const result = Object.keys(byDate)
		.map(function (date) {
			return byDate[date];
		})
		.sort(function (a, b) {
			return toDateObj_(a.date).getTime() - toDateObj_(b.date).getTime();
		});

	return result;
}

function getOrCreateSheet_() {
	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	ensureConfigSheetInitialized_(spreadsheet);
	const configSheet = spreadsheet.getSheetByName(GAS_CONNECT_CONFIG.configSheetName);
	if (configSheet) {
		ensureDefaultProjectConfig_(configSheet);
	}
	return ensureDataSheetInitialized_(spreadsheet);
}

function gcBuildConfigKeyIndex_(sheet) {
	const index = {};
	const schema = ensureSheetHeaders_(sheet, ['key', 'value', 'description', 'updatedAt']);
	const keyCol = requireHeaderColumn_(schema.headerMap, 'key');
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return index;

	const values = sheet.getRange(2, keyCol, lastRow - 1, 1).getValues();
	for (let i = 0; i < values.length; i++) {
		const key = String(values[i][0] || '').trim();
		if (key) {
			index[key] = i + 2;
		}
	}

	return index;
}

function setConfigValue_(key, value, description) {
	const spreadsheet = SpreadsheetApp.openById(GAS_CONNECT_CONFIG.sheetId);
	const configSheet = ensureConfigSheetInitialized_(spreadsheet);
	const schema = ensureSheetHeaders_(configSheet, ['key', 'value', 'description', 'updatedAt']);
	const headers = schema.headers;
	const headerMap = schema.headerMap;
	const index = gcBuildConfigKeyIndex_(configSheet);
	const rowIndex = index[key] || -1;

	const payload = createEmptyRowByHeaders_(headers);
	setCellByHeader_(payload, headerMap, 'key', key);
	setCellByHeader_(payload, headerMap, 'value', value == null ? '' : String(value));
	setCellByHeader_(payload, headerMap, 'description', description == null ? '' : String(description));
	setCellByHeader_(payload, headerMap, 'updatedAt', new Date().toISOString());

	if (rowIndex > 0) {
		configSheet.getRange(rowIndex, 1, 1, payload.length).setValues([payload]);
	} else {
		configSheet.appendRow(payload);
	}
}

function gcApplyDataSheetStyle_(sheet, headers, headerMap) {
	const headerWidth = headers.length;
	sheet.setFrozenRows(1);
	sheet.getRange(1, 1, 1, headerWidth).setFontWeight('bold').setBackground('#e2e8f0');

	setColumnWidthIfPresent_(sheet, headerMap, 'date', 105);
	setColumnWidthIfPresent_(sheet, headerMap, 'year', 60);
	setColumnWidthIfPresent_(sheet, headerMap, 'month', 70);

	const unitHeaders = buildUnitCostHeaderNames_();
	for (let i = 0; i < unitHeaders.length; i++) {
		setColumnWidthIfPresent_(sheet, headerMap, unitHeaders[i], 92);
	}

	setColumnWidthIfPresent_(sheet, headerMap, 'updatedAt', 190);
}

function gcApplyConfigSheetStyle_(sheet, headers, headerMap) {
	const headerWidth = headers.length;
	sheet.setFrozenRows(1);
	sheet.getRange(1, 1, 1, headerWidth).setFontWeight('bold').setBackground('#dbeafe');
	setColumnWidthIfPresent_(sheet, headerMap, 'key', 220);
	setColumnWidthIfPresent_(sheet, headerMap, 'value', 260);
	setColumnWidthIfPresent_(sheet, headerMap, 'description', 500);
	setColumnWidthIfPresent_(sheet, headerMap, 'updatedAt', 190);
}

function buildHeaders_() {
	const headers = ['date', 'year', 'month'];
	for (let n = 1; n <= GAS_CONNECT_CONFIG.unitCostCount; n++) {
		headers.push('unitCost' + n);
	}
	headers.push('updatedAt');
	return headers;
}

function buildUnitCostHeaderNames_() {
	const names = [];
	for (let n = 1; n <= GAS_CONNECT_CONFIG.unitCostCount; n++) {
		names.push('unitCost' + n);
	}
	return names;
}

function buildDateIndex_(sheet) {
	const index = {};
	const schema = ensureSheetHeaders_(sheet, buildHeaders_());
	const dateCol = requireHeaderColumn_(schema.headerMap, 'date');
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return index;

	const values = sheet.getRange(2, dateCol, lastRow - 1, 1).getValues();
	for (let i = 0; i < values.length; i++) {
		const parsed = parseDateCellValue_(values[i][0]);
		if (parsed) {
			index[parsed.label] = i + 2;
		}
	}
	return index;
}

function upsertRows_(sheet, dateIndex, records) {
	if (!records || !records.length) {
		return { inserted: 0, updated: 0 };
	}

	let inserted = 0;
	let updated = 0;
	const schema = ensureSheetHeaders_(sheet, buildHeaders_());
	const headers = schema.headers;
	const headerMap = schema.headerMap;
	const width = headers.length;
	const rowsToAppend = [];

	for (let i = 0; i < records.length; i++) {
		const record = records[i];
		const rowArray = toSheetRow_(record, headers, headerMap);
		const existingRow = dateIndex[record.date];

		if (existingRow) {
			const existingValues = sheet.getRange(existingRow, 1, 1, width).getValues()[0];
			const mergedValues = mergeRowValues_(existingValues, rowArray, headerMap, buildHeaders_());
			sheet.getRange(existingRow, 1, 1, width).setValues([mergedValues]);
			updated += 1;
		} else {
			rowsToAppend.push(rowArray);
			inserted += 1;
		}
	}

	if (rowsToAppend.length) {
		const start = sheet.getLastRow() + 1;
		sheet.getRange(start, 1, rowsToAppend.length, width).setValues(rowsToAppend);

		// Refresh index for subsequent operations in the same run.
		const dateCol = requireHeaderColumn_(headerMap, 'date');
		for (let i = 0; i < rowsToAppend.length; i++) {
			const dateText = String(rowsToAppend[i][dateCol - 1] || '');
			if (!dateText) continue;
			dateIndex[dateText] = start + i;
		}
	}

	return { inserted: inserted, updated: updated };
}

function toSheetRow_(record, headers, headerMap) {
	const row = createEmptyRowByHeaders_(headers);
	setCellByHeader_(row, headerMap, 'date', record.date || '');
	setCellByHeader_(row, headerMap, 'year', record.year || '');
	setCellByHeader_(row, headerMap, 'month', record.month || '');

	const unitCosts = record.unitCosts || [];
	const unitHeaders = buildUnitCostHeaderNames_();
	for (let n = 0; n < unitHeaders.length; n++) {
		setCellByHeader_(row, headerMap, unitHeaders[n], unitCosts[n] == null ? '' : unitCosts[n]);
	}

	setCellByHeader_(row, headerMap, 'updatedAt', record.updatedAt || new Date().toISOString());
	return row;
}

function readData_(since, limit) {
	const sheet = getOrCreateSheet_();
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) {
		return { count: 0, records: [] };
	}

	const schema = ensureSheetHeaders_(sheet, buildHeaders_());
	const headers = schema.headers;
	const headerMap = schema.headerMap;
	const width = headers.length;
	const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();

	const sinceDate = since ? parseLooseDate_(since) : null;
	const out = [];
	const unitHeaders = buildUnitCostHeaderNames_();

	for (let i = 0; i < values.length; i++) {
		const row = values[i];
		const parsed = parseDateCellValue_(getCellByHeader_(row, headerMap, 'date'));
		if (!parsed) continue;

		if (sinceDate && parsed.dateObj.getTime() <= sinceDate.getTime()) {
			continue;
		}

		const unitCosts = [];
		for (let n = 0; n < unitHeaders.length; n++) {
			unitCosts.push(toNumberOrNull_(getCellByHeader_(row, headerMap, unitHeaders[n])));
		}

		out.push({
			date: parsed.label,
			unitCosts: unitCosts,
			updatedAt: String(getCellByHeader_(row, headerMap, 'updatedAt') || '')
		});

		if (out.length >= limit) {
			break;
		}
	}

	return {
		count: out.length,
		records: out,
		generatedAt: new Date().toISOString()
	};
}

function buildMonthSummary_() {
	const sheet = getOrCreateSheet_();
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) {
		return { count: 0, months: [] };
	}

	const schema = ensureSheetHeaders_(sheet, buildHeaders_());
	const headers = schema.headers;
	const headerMap = schema.headerMap;
	const width = headers.length;
	const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
	const monthMap = {};

	for (let i = 0; i < values.length; i++) {
		const row = values[i];
		const parsed = parseDateCellValue_(getCellByHeader_(row, headerMap, 'date'));
		if (!parsed) continue;

		const monthKey = Utilities.formatString('%04d-%02d', parsed.year, parsed.month);
		let item = monthMap[monthKey];
		if (!item) {
			item = {
				year: parsed.year,
				month: parsed.month,
				monthKey: monthKey,
				recordCount: 0,
				lastDate: null,
				lastDateObj: null,
				lastUpdatedAt: ''
			};
			monthMap[monthKey] = item;
		}

		item.recordCount += 1;
		if (!item.lastDateObj || parsed.dateObj.getTime() > item.lastDateObj.getTime()) {
			item.lastDateObj = parsed.dateObj;
			item.lastDate = parsed.label;
		}

		const updatedAt = String(getCellByHeader_(row, headerMap, 'updatedAt') || '').trim();
		if (updatedAt && updatedAt > item.lastUpdatedAt) {
			item.lastUpdatedAt = updatedAt;
		}
	}

	const months = Object.keys(monthMap)
		.map(function (key) {
			const item = monthMap[key];
			return {
				year: item.year,
				month: item.month,
				monthKey: item.monthKey,
				recordCount: item.recordCount,
				lastDate: item.lastDate || '',
				lastUpdatedAt: item.lastUpdatedAt || ''
			};
		})
		.sort(function (a, b) {
			if (a.year !== b.year) return b.year - a.year;
			return b.month - a.month;
		});

	return {
		count: months.length,
		months: months,
		generatedAt: new Date().toISOString()
	};
}

function readMonthData_(year, month) {
	if (!Number.isInteger(year) || !Number.isInteger(month) || year < 1900 || month < 1 || month > 12) {
		throw new Error('Invalid year/month for dataMonth');
	}

	const sheet = getOrCreateSheet_();
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) {
		return { year: year, month: month, count: 0, records: [] };
	}

	const schema = ensureSheetHeaders_(sheet, buildHeaders_());
	const headers = schema.headers;
	const headerMap = schema.headerMap;
	const width = headers.length;
	const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();
	const out = [];
	const unitHeaders = buildUnitCostHeaderNames_();

	for (let i = 0; i < values.length; i++) {
		const row = values[i];
		const parsed = parseDateCellValue_(getCellByHeader_(row, headerMap, 'date'));
		if (!parsed) continue;
		if (parsed.year !== year || parsed.month !== month) continue;

		const unitCosts = [];
		for (let n = 0; n < unitHeaders.length; n++) {
			unitCosts.push(toNumberOrNull_(getCellByHeader_(row, headerMap, unitHeaders[n])));
		}

		out.push({
			date: parsed.label,
			unitCosts: unitCosts,
			updatedAt: String(getCellByHeader_(row, headerMap, 'updatedAt') || '')
		});
	}

	out.sort(function (a, b) {
		return toDateObj_(a.date).getTime() - toDateObj_(b.date).getTime();
	});

	return {
		year: year,
		month: month,
		count: out.length,
		records: out,
		generatedAt: new Date().toISOString()
	};
}

function ensureSheetHeaders_(sheet, requiredHeaders) {
	const existingHeaders = getSheetHeaders_(sheet);
	const headerSet = {};

	for (let i = 0; i < existingHeaders.length; i++) {
		const h = existingHeaders[i];
		if (h) headerSet[h] = true;
	}

	const nextHeaders = existingHeaders.slice();
	for (let i = 0; i < requiredHeaders.length; i++) {
		const name = requiredHeaders[i];
		if (!headerSet[name]) {
			nextHeaders.push(name);
			headerSet[name] = true;
		}
	}

	const hasHeaderRow = sheet.getLastRow() >= 1;
	if (!hasHeaderRow || nextHeaders.length !== existingHeaders.length) {
		sheet.getRange(1, 1, 1, nextHeaders.length).setValues([nextHeaders]);
	} else if (nextHeaders.length > 0) {
		sheet.getRange(1, 1, 1, nextHeaders.length).setValues([nextHeaders]);
	}

	return {
		headers: nextHeaders,
		headerMap: buildHeaderMap_(nextHeaders)
	};
}

function getSheetHeaders_(sheet) {
	if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return [];
	return sheet
		.getRange(1, 1, 1, sheet.getLastColumn())
		.getValues()[0]
		.map(function (value) {
			return String(value || '').trim();
		})
		.filter(function (name) {
			return !!name;
		});
}

function buildHeaderMap_(headers) {
	const map = {};
	for (let i = 0; i < headers.length; i++) {
		const name = String(headers[i] || '').trim();
		if (!name) continue;
		map[name] = i + 1;
	}
	return map;
}

function requireHeaderColumn_(headerMap, columnName) {
	const col = headerMap[columnName] || 0;
	if (!col) {
		throw new Error('Missing required column: ' + columnName);
	}
	return col;
}

function getCellByHeader_(row, headerMap, columnName) {
	const col = headerMap[columnName] || 0;
	if (!col) return '';
	return row[col - 1];
}

function setCellByHeader_(row, headerMap, columnName, value) {
	const col = headerMap[columnName] || 0;
	if (!col) return;
	row[col - 1] = value;
}

function createEmptyRowByHeaders_(headers) {
	return new Array(headers.length).fill('');
}

function mergeRowValues_(baseRow, patchRow, headerMap, requiredHeaders) {
	const output = baseRow.slice();
	for (let i = 0; i < requiredHeaders.length; i++) {
		const name = requiredHeaders[i];
		const col = headerMap[name] || 0;
		if (!col) continue;
		output[col - 1] = patchRow[col - 1];
	}
	return output;
}

function setColumnWidthIfPresent_(sheet, headerMap, columnName, width) {
	const col = headerMap[columnName] || 0;
	if (!col) return;
	sheet.setColumnWidth(col, width);
}

function parseDateString_(text) {
	const datePart = String(text || '').trim().split(' ')[0];
	const bits = datePart.split('/');
	if (bits.length !== 3) return null;

	const day = parseInt(bits[0], 10);
	const month = parseInt(bits[1], 10);
	const year = parseInt(bits[2], 10);
	if (!day || !month || !year) return null;

	return {
		day: day,
		month: month,
		year: year,
		label: Utilities.formatString('%02d/%02d/%d', day, month, year),
		dateObj: new Date(year, month - 1, day)
	};
}

function parseDateCellValue_(value) {
	if (value instanceof Date && !isNaN(value.getTime())) {
		const day = value.getDate();
		const month = value.getMonth() + 1;
		const year = value.getFullYear();
		return {
			day: day,
			month: month,
			year: year,
			label: Utilities.formatString('%02d/%02d/%d', day, month, year),
			dateObj: new Date(year, month - 1, day)
		};
	}

	const text = String(value == null ? '' : value).trim();
	if (!text) return null;

	const bySlash = parseDateString_(text);
	if (bySlash) return bySlash;

	const loose = parseLooseDate_(text);
	if (!loose) return null;
	const day = loose.getDate();
	const month = loose.getMonth() + 1;
	const year = loose.getFullYear();
	return {
		day: day,
		month: month,
		year: year,
		label: Utilities.formatString('%02d/%02d/%d', day, month, year),
		dateObj: new Date(year, month - 1, day)
	};
}

function parseLooseDate_(text) {
	const v = String(text || '').trim();
	if (!v) return null;

	if (v.indexOf('-') > -1) {
		const bits = v.split('-');
		if (bits.length === 3) {
			const y = parseInt(bits[0], 10);
			const m = parseInt(bits[1], 10);
			const d = parseInt(bits[2], 10);
			if (y && m && d) return new Date(y, m - 1, d);
		}
	}

	const parsed = parseDateString_(v);
	return parsed ? parsed.dateObj : null;
}

function toDateObj_(ddmmyyyy) {
	const parsed = parseDateString_(ddmmyyyy);
	return parsed ? parsed.dateObj : new Date(0);
}

function toInt_(value, fallback) {
	const n = parseInt(value, 10);
	return Number.isFinite(n) ? n : fallback;
}

function toNumberOrNull_(value) {
	const n = parseFloat(value);
	return Number.isFinite(n) ? n : null;
}

function toBoolean_(value) {
	if (value === true || value === false) return value;
	const text = String(value == null ? '' : value).trim().toLowerCase();
	if (!text) return false;
	return text === '1' || text === 'true' || text === 'yes' || text === 'y' || text === 'on';
}

function installDailySyncTriggerAt10() {
	deleteDailySyncTriggerAt10();

	ScriptApp.newTrigger('dailySyncAt10Job_')
		.timeBased()
		.everyDays(1)
		.atHour(10)
		.inTimezone('Asia/Bangkok')
		.create();

	return {
		ok: true,
		message: 'Installed daily sync trigger at 10:00 (Asia/Bangkok)',
		installedAt: new Date().toISOString()
	};
}

function deleteDailySyncTriggerAt10() {
	const all = ScriptApp.getProjectTriggers();
	for (let i = 0; i < all.length; i++) {
		const trigger = all[i];
		if (trigger.getHandlerFunction() === 'dailySyncAt10Job_') {
			ScriptApp.deleteTrigger(trigger);
		}
	}

	return {
		ok: true,
		message: 'Deleted all dailySyncAt10Job_ triggers',
		deletedAt: new Date().toISOString()
	};
}

function dailySyncAt10Job_() {
	validateConfig_();
	ensureProjectInitialized_();

	const now = new Date();
	const latest = getLatestDataMonthYear_() || { startYear: 1998, startMonth: 1 };

	return syncToSheet_(
		latest.startYear,
		latest.startMonth,
		now.getFullYear(),
		now.getMonth() + 1,
		{ reconcile: true }
	);
}

function jsonOut_(payload) {
	return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(
		ContentService.MimeType.JSON
	);
}


