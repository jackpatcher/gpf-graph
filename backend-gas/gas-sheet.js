/*
	GAS Spreadsheet schema utilities for GPF NAV project
	----------------------------------------------------
	What this file does:
	- Creates/updates data sheet structure (NAV_DATA)
	- Creates/updates config sheet structure (CONFIG)
	- Creates/updates users sheet (USERS)
	- Creates/updates portfolio change history sheet (PORTFOLIO_CHANGES)
	- Provides key-value config helpers for project settings

	Usage in Apps Script editor:
	1) Set SCRIPT PROPERTY: SHEET_ID
	2) Run setupProjectReady()
*/

const GAS_SHEET_SCHEMA = {
	dataSheetName: 'NAV_DATA',
	configSheetName: 'CONFIG',
	usersSheetName: 'USERS',
	portfolioSheetName: 'PORTFOLIO_CHANGES',
	unitCostCount: 14,
	dataHeaders: buildDataHeaders_(),
	configHeaders: ['key', 'value', 'description', 'updatedAt'],
	usersHeaders: [
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
	],
	portfolioHeaders: [
		'changeId',
		'batchId',
		'userId',
		'username',
		'effectiveDate',
		'entriesJson',
		'note',
		'createdAt',
		'updatedAt'
	],
	sampleUsers: [
		{
			userId: 'U001',
			username: 'demo_user_1',
			passwordHash: 'demo_hash_amp',
			displayName: 'Demo User 1',
			role: 'admin',
			status: 'approved'
		},
		{
			userId: 'U002',
			username: 'demo_user_2',
			passwordHash: 'demo_hash_demo',
			displayName: 'Demo User 2',
			role: 'user',
			status: 'approved'
		}
	],
	samplePortfolioChanges: [
		{
			changeId: 'PC0001',
			batchId: 'B20260115-01',
			userId: 'U001',
			username: 'demo_user_1',
			effectiveDate: '2026-01-15',
			entriesJson: JSON.stringify([
				{ unitCostKey: 'unitCost3', planName: 'bond', units: 800.0 },
				{ unitCostKey: 'unitCost4', planName: 'ฝากBond สั้น', units: 450.0 }
			]),
			note: 'บันทึกรอบเดียวหลายแผน'
		},
		{
			changeId: 'PC0002',
			batchId: 'B20260210-01',
			userId: 'U001',
			username: 'demo_user_1',
			effectiveDate: '2026-02-10',
			entriesJson: JSON.stringify([
				{ unitCostKey: 'unitCost6', planName: 'หุ้น 65', units: 350.0 }
			]),
			note: 'เพิ่มสัดส่วนหุ้น'
		},
		{
			changeId: 'PC0003',
			batchId: 'B20260120-01',
			userId: 'U002',
			username: 'demo_user_2',
			effectiveDate: '2026-01-20',
			entriesJson: JSON.stringify([
				{ unitCostKey: 'unitCost3', planName: 'bond', units: 900.0 }
			]),
			note: 'พอร์ตตัวอย่างผู้ใช้ที่สอง'
		}
	],
	configDefaults: [
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
			value: 'https://www.gpf.or.th/thai2019/About/memberfund-api.php?pageName=NAVBottom_{MM}_{YYYY}',
			description: 'Upstream API template'
		},
		{
			key: 'START_YEAR',
			value: '1998',
			description: 'First year for full sync'
		},
		{
			key: 'START_MONTH',
			value: '1',
			description: 'First month for full sync (1-12)'
		},
		{
			key: 'SYNC_TIMEZONE',
			value: 'Asia/Bangkok',
			description: 'Project timezone for scheduling and logs'
		},
		{
			key: 'SYNC_LIMIT_MONTHS_PER_RUN',
			value: '0',
			description: '0 = no limit, otherwise max months per sync run'
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
			description: 'Sheet name for portfolio change history'
		}
	]
};

/**
 * Initialize NAV, CONFIG, USERS and PORTFOLIO_CHANGES sheets with readable table structure.
 */
function initializeProjectSheets() {
	const spreadsheet = openSpreadsheetFromScriptProperty_();

	const dataSheet = ensureDataSheet_(spreadsheet);
	const configSheet = ensureConfigSheet_(spreadsheet);
	const usersSheet = ensureUsersSheet_(spreadsheet);
	const portfolioSheet = ensurePortfolioSheet_(spreadsheet);

	return {
		ok: true,
		spreadsheetId: spreadsheet.getId(),
		dataSheet: dataSheet.getName(),
		configSheet: configSheet.getName(),
		usersSheet: usersSheet.getName(),
		portfolioSheet: portfolioSheet.getName(),
		dataHeaderCount: GAS_SHEET_SCHEMA.dataHeaders.length,
		configHeaderCount: GAS_SHEET_SCHEMA.configHeaders.length,
		usersHeaderCount: GAS_SHEET_SCHEMA.usersHeaders.length,
		portfolioHeaderCount: GAS_SHEET_SCHEMA.portfolioHeaders.length,
		initializedAt: new Date().toISOString()
	};
}

/**
 * Return all config key-value entries as object.
 */
function getAllProjectConfig() {
	const sheet = ensureConfigSheet_(openSpreadsheetFromScriptProperty_());
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return {};

	const rows = sheet.getRange(2, 1, lastRow - 1, GAS_SHEET_SCHEMA.configHeaders.length).getValues();
	const result = {};

	for (let i = 0; i < rows.length; i++) {
		const key = String(rows[i][0] || '').trim();
		if (!key) continue;
		result[key] = rows[i][1] == null ? '' : String(rows[i][1]);
	}

	return result;
}

/**
 * Get one config value by key.
 */
function getProjectConfigValue(key) {
	const normalizedKey = String(key || '').trim();
	if (!normalizedKey) return null;

	const sheet = ensureConfigSheet_(openSpreadsheetFromScriptProperty_());
	const rowIndex = findConfigRowIndex_(sheet, normalizedKey);
	if (rowIndex < 0) return null;

	return sheet.getRange(rowIndex, 2).getValue();
}

/**
 * Upsert one config value by key.
 */
function setProjectConfigValue(key, value, description) {
	const normalizedKey = String(key || '').trim();
	if (!normalizedKey) {
		throw new Error('Config key is required');
	}

	const sheet = ensureConfigSheet_(openSpreadsheetFromScriptProperty_());
	const rowIndex = findConfigRowIndex_(sheet, normalizedKey);
	const payload = [
		normalizedKey,
		value == null ? '' : String(value),
		description == null ? '' : String(description),
		new Date().toISOString()
	];

	if (rowIndex > 0) {
		sheet.getRange(rowIndex, 1, 1, payload.length).setValues([payload]);
	} else {
		sheet.appendRow(payload);
	}

	return { ok: true, key: normalizedKey };
}

function ensureDataSheet_(spreadsheet) {
	let sheet = spreadsheet.getSheetByName(GAS_SHEET_SCHEMA.dataSheetName);
	if (!sheet) {
		sheet = spreadsheet.insertSheet(GAS_SHEET_SCHEMA.dataSheetName);
	}

	const headers = GAS_SHEET_SCHEMA.dataHeaders;
	const expectedWidth = headers.length;

	const shouldResetHeader =
		sheet.getLastRow() < 1 ||
		sheet.getLastColumn() < expectedWidth ||
		String(sheet.getRange(1, 1).getValue() || '').trim() !== headers[0];

	if (shouldResetHeader) {
		sheet.clear();
		sheet.getRange(1, 1, 1, expectedWidth).setValues([headers]);
	} else {
		// Keep existing data but ensure header is exactly as expected.
		sheet.getRange(1, 1, 1, expectedWidth).setValues([headers]);
	}

	applyDataSheetStyle_(sheet, expectedWidth);
	return sheet;
}

function ensureConfigSheet_(spreadsheet) {
	let sheet = spreadsheet.getSheetByName(GAS_SHEET_SCHEMA.configSheetName);
	if (!sheet) {
		sheet = spreadsheet.insertSheet(GAS_SHEET_SCHEMA.configSheetName);
	}

	const headers = GAS_SHEET_SCHEMA.configHeaders;
	const expectedWidth = headers.length;

	const shouldResetHeader =
		sheet.getLastRow() < 1 ||
		sheet.getLastColumn() < expectedWidth ||
		String(sheet.getRange(1, 1).getValue() || '').trim() !== headers[0];

	if (shouldResetHeader) {
		sheet.clear();
		sheet.getRange(1, 1, 1, expectedWidth).setValues([headers]);
	} else {
		sheet.getRange(1, 1, 1, expectedWidth).setValues([headers]);
	}

	applyConfigSheetStyle_(sheet, expectedWidth);
	upsertDefaultConfigRows_(sheet);
	return sheet;
}

function ensureUsersSheet_(spreadsheet) {
	let sheet = spreadsheet.getSheetByName(GAS_SHEET_SCHEMA.usersSheetName);
	if (!sheet) {
		sheet = spreadsheet.insertSheet(GAS_SHEET_SCHEMA.usersSheetName);
	}

	const headers = GAS_SHEET_SCHEMA.usersHeaders;
	const expectedWidth = headers.length;

	// Only hard-reset (clear) when the sheet is empty or the first header is wrong.
	// Do NOT clear on column count mismatch — new columns (e.g. googleId) should be
	// appended without wiping existing user data.
	const shouldResetHeader =
		sheet.getLastRow() < 1 ||
		String(sheet.getRange(1, 1).getValue() || '').trim() !== headers[0];

	if (shouldResetHeader) {
		sheet.clear();
		sheet.getRange(1, 1, 1, expectedWidth).setValues([headers]);
	} else {
		sheet.getRange(1, 1, 1, expectedWidth).setValues([headers]);
	}

	applyUsersSheetStyle_(sheet, expectedWidth);
	upsertSampleUsers_(sheet);
	return sheet;
}

function ensurePortfolioSheet_(spreadsheet) {
	let sheet = spreadsheet.getSheetByName(GAS_SHEET_SCHEMA.portfolioSheetName);
	if (!sheet) {
		sheet = spreadsheet.insertSheet(GAS_SHEET_SCHEMA.portfolioSheetName);
	}

	const headers = GAS_SHEET_SCHEMA.portfolioHeaders;
	const expectedWidth = headers.length;

	const shouldResetHeader =
		sheet.getLastRow() < 1 ||
		sheet.getLastColumn() < expectedWidth ||
		String(sheet.getRange(1, 1).getValue() || '').trim() !== headers[0];

	if (shouldResetHeader) {
		sheet.clear();
		sheet.getRange(1, 1, 1, expectedWidth).setValues([headers]);
	} else {
		sheet.getRange(1, 1, 1, expectedWidth).setValues([headers]);
	}

	applyPortfolioSheetStyle_(sheet, expectedWidth);
	upsertSamplePortfolioRows_(sheet);
	return sheet;
}

function upsertDefaultConfigRows_(sheet) {
	const existingMap = buildConfigKeyIndex_(sheet);

	for (let i = 0; i < GAS_SHEET_SCHEMA.configDefaults.length; i++) {
		const item = GAS_SHEET_SCHEMA.configDefaults[i];
		const rowIndex = existingMap[item.key] || -1;

		if (rowIndex > 0) {
			// Keep user value if already set, only refresh description and timestamp.
			const currentValue = sheet.getRange(rowIndex, 2).getValue();
			sheet
				.getRange(rowIndex, 1, 1, GAS_SHEET_SCHEMA.configHeaders.length)
				.setValues([[item.key, currentValue, item.description, new Date().toISOString()]]);
		} else {
			sheet.appendRow([item.key, item.value, item.description, new Date().toISOString()]);
		}
	}
}

function upsertSampleUsers_(sheet) {
	if (sheet.getLastRow() > 1) return;

	const nowIso = new Date().toISOString();
	const values = GAS_SHEET_SCHEMA.sampleUsers.map(function (item) {
		return [
			item.userId,
			item.username,
			item.passwordHash,
			item.displayName,
			item.role || 'user',
			item.status || 'approved',
			nowIso,
			nowIso,
			'',
			'',
			'',
			'',
			'',
			''
		];
	});

	if (values.length) {
		sheet.getRange(2, 1, values.length, GAS_SHEET_SCHEMA.usersHeaders.length).setValues(values);
	}
}

function upsertSamplePortfolioRows_(sheet) {
	if (sheet.getLastRow() > 1) return;

	const nowIso = new Date().toISOString();
	const values = GAS_SHEET_SCHEMA.samplePortfolioChanges.map(function (item) {
		return [
			item.changeId,
			item.batchId,
			item.userId,
			item.username,
			item.effectiveDate,
			item.entriesJson,
			item.note,
			nowIso,
			nowIso
		];
	});

	if (values.length) {
		sheet
			.getRange(2, 1, values.length, GAS_SHEET_SCHEMA.portfolioHeaders.length)
			.setValues(values);
	}
}

function buildConfigKeyIndex_(sheet) {
	const map = {};
	const lastRow = sheet.getLastRow();
	if (lastRow <= 1) return map;

	const keys = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
	for (let i = 0; i < keys.length; i++) {
		const key = String(keys[i][0] || '').trim();
		if (!key) continue;
		map[key] = i + 2;
	}
	return map;
}

function findConfigRowIndex_(sheet, key) {
	const map = buildConfigKeyIndex_(sheet);
	return map[key] || -1;
}

function applyDataSheetStyle_(sheet, headerWidth) {
	sheet.setFrozenRows(1);
	sheet.getRange(1, 1, 1, headerWidth).setFontWeight('bold').setBackground('#e2e8f0');

	// Layout for readability.
	sheet.setColumnWidth(1, 105); // date
	sheet.setColumnWidth(2, 60); // year
	sheet.setColumnWidth(3, 70); // month

	for (let col = 4; col <= 3 + GAS_SHEET_SCHEMA.unitCostCount; col++) {
		sheet.setColumnWidth(col, 92);
	}

	sheet.setColumnWidth(4 + GAS_SHEET_SCHEMA.unitCostCount, 190); // updatedAt

	if (sheet.getLastRow() > 1) {
		const dataRows = sheet.getLastRow() - 1;
		sheet
			.getRange(2, 1, dataRows, 1)
			.setNumberFormat('dd/MM/yyyy');
		sheet
			.getRange(2, 4, dataRows, GAS_SHEET_SCHEMA.unitCostCount)
			.setNumberFormat('0.0000');
	}
}

function applyConfigSheetStyle_(sheet, headerWidth) {
	sheet.setFrozenRows(1);
	sheet.getRange(1, 1, 1, headerWidth).setFontWeight('bold').setBackground('#dbeafe');

	sheet.setColumnWidth(1, 220); // key
	sheet.setColumnWidth(2, 260); // value
	sheet.setColumnWidth(3, 500); // description
	sheet.setColumnWidth(4, 190); // updatedAt
}

function applyUsersSheetStyle_(sheet, headerWidth) {
	sheet.setFrozenRows(1);
	sheet.getRange(1, 1, 1, headerWidth).setFontWeight('bold').setBackground('#dcfce7');

	sheet.setColumnWidth(1, 100); // userId
	sheet.setColumnWidth(2, 220); // username / email
	sheet.setColumnWidth(3, 240); // passwordHash
	sheet.setColumnWidth(4, 220); // displayName
	sheet.setColumnWidth(5, 100); // role
	sheet.setColumnWidth(6, 110); // status
	sheet.setColumnWidth(7, 190); // createdAt
	sheet.setColumnWidth(8, 190); // updatedAt
	sheet.setColumnWidth(9, 220); // googleId
	sheet.setColumnWidth(10, 220); // emailChangeOtpHash
	sheet.setColumnWidth(11, 190); // emailChangeOtpExpiresAt
	sheet.setColumnWidth(12, 220); // pendingEmail
	sheet.setColumnWidth(13, 190); // pendingEmailLinkExpiresAt
	sheet.setColumnWidth(14, 190); // emailChangeLastSentAt
}

function applyPortfolioSheetStyle_(sheet, headerWidth) {
	sheet.setFrozenRows(1);
	sheet.getRange(1, 1, 1, headerWidth).setFontWeight('bold').setBackground('#ffedd5');

	sheet.setColumnWidth(1, 110); // changeId
	sheet.setColumnWidth(2, 140); // batchId
	sheet.setColumnWidth(3, 100); // userId
	sheet.setColumnWidth(4, 170); // username
	sheet.setColumnWidth(5, 120); // effectiveDate
	sheet.setColumnWidth(6, 520); // entriesJson
	sheet.setColumnWidth(7, 280); // note
	sheet.setColumnWidth(8, 190); // createdAt
	sheet.setColumnWidth(9, 190); // updatedAt

	if (sheet.getLastRow() > 1) {
		const dataRows = sheet.getLastRow() - 1;
		sheet.getRange(2, 5, dataRows, 1).setNumberFormat('yyyy-mm-dd');
	}
}

function buildDataHeaders_() {
	const headers = ['date', 'year', 'month'];
	for (let i = 1; i <= 14; i++) {
		headers.push('unitCost' + i);
	}
	headers.push('updatedAt');
	return headers;
}

function openSpreadsheetFromScriptProperty_() {
	const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
	if (!sheetId) {
		throw new Error('Missing script property: SHEET_ID');
	}
	return SpreadsheetApp.openById(sheetId);
}

/**
 * One-click setup for first-time project initialization.
 * - Creates/updates NAV_DATA, CONFIG, USERS, PORTFOLIO_CHANGES
 * - Installs daily sync trigger at 10:00 (if function exists in project)
 */
function setupProjectReady() {
	const initResult = initializeProjectSheets();

	let triggerResult = {
		ok: false,
		message: 'installDailySyncTriggerAt10() not found in this GAS project'
	};

	if (typeof installDailySyncTriggerAt10 === 'function') {
		triggerResult = installDailySyncTriggerAt10();
	}

	return {
		ok: true,
		action: 'setupProjectReady',
		init: initResult,
		trigger: triggerResult,
		completedAt: new Date().toISOString()
	};
}

/**
 * Backward-compatible alias for typo usage.
 */
function setupRrojectReady() {
	return setupProjectReady();
}
