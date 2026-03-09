# Setup GAS (Google Apps Script) For `gpf-graph`

เอกสารนี้คือขั้นตอนตั้งค่าแบบ end-to-end เพื่อทดสอบระบบ `GAS-only` (Frontend -> GAS -> Google Sheet -> IndexedDB)

## 1) เตรียม Google Sheet

1. สร้าง Google Sheet ใหม่ 1 ไฟล์
2. คัดลอก `Spreadsheet ID` จาก URL

ตัวอย่าง URL:
`https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit`

## 2) สร้างโปรเจค Apps Script

1. เข้า `script.new`
2. สร้าง project ใหม่
3. สร้างไฟล์สคริปต์ 2 ไฟล์ แล้ววางโค้ดจาก repo นี้
- `gas-sheet.js`
- `gas-connect.js`

หมายเหตุ:
- ให้อยู่ใน GAS project เดียวกัน
- อย่าเปลี่ยนชื่อฟังก์ชันหลัก (`doGet`, `doPost`, `initializeProjectSheets`)

## 3) ตั้ง Script Properties

ไปที่ `Project Settings` -> `Script properties` แล้วเพิ่มค่า:

- `SHEET_ID` = `<SPREADSHEET_ID>` (จำเป็น)
- `SHEET_NAME` = `NAV_DATA` (optional)
- `CONFIG_SHEET_NAME` = `CONFIG` (optional)
- `SYNC_TOKEN` = `<YOUR_SECRET_TOKEN>` (แนะนำ)

## 4) Deploy เป็น Web App

1. กด `Deploy` -> `New deployment`
2. Type = `Web app`
3. Execute as = `Me`
4. Who has access = `Anyone` (หรือ `Anyone with link`)
5. Deploy แล้วคัดลอก Web App URL

ตัวอย่าง:
`https://script.google.com/macros/s/AKfycb.../exec`

## 5) Initialize schema ครั้งแรก

เรียก URL นี้ 1 ครั้ง (ใน browser):

`<WEB_APP_URL>?action=init`

ผลที่คาดหวัง:
- สร้างชีท `NAV_DATA` และ `CONFIG`
- ตั้ง header และ default config อัตโนมัติ

## 6) ทดสอบ endpoint พื้นฐาน

1. Health check:
`<WEB_APP_URL>?action=health`

2. Sync ข้อมูลเข้า Google Sheet:
`<WEB_APP_URL>?action=sync&token=<SYNC_TOKEN>&startYear=1998&startMonth=1`

3. อ่านข้อมูล:
`<WEB_APP_URL>?action=data&limit=100`

## 7) ตั้งค่า Frontend (`index.html`)

แก้ค่าต่อไปนี้ใน `index.html`:

```javascript
const GAS_WEB_APP_URL = 'YOUR_WEB_APP_URL';
const GAS_SYNC_TOKEN = 'YOUR_SYNC_TOKEN_OR_EMPTY';
const GAS_PAGE_LIMIT = 5000;
```

## 8) ทดสอบหน้าเว็บ

1. เปิดหน้า `index.html`
2. เลือกช่วงวัน แล้วดูว่ากราฟโหลดข้อมูลได้
3. ไปที่แท็บ `จัดการ DB`
4. กดปุ่ม `ซิงก์ข้อมูลล่าสุดจาก GAS`
5. ตรวจสถานะ DB ว่าจำนวนเดือน/จำนวนแถวเพิ่มขึ้น

## 9) ยืนยันว่า IndexedDB ถูกใช้งาน

ตรวจใน browser devtools:
- Application -> IndexedDB -> `gpf_nav_cache_db` -> `monthly_nav_cache`

ควรเห็นข้อมูลรายเดือนถูกเก็บไว้

## 10) Troubleshooting (สั้นๆ)

- Error: `Missing script property: SHEET_ID`
  - ยังไม่ได้ตั้ง `SHEET_ID` ใน Script properties

- Error: `Unauthorized sync request`
  - ค่า `SYNC_TOKEN` ที่ส่งมาไม่ตรงกับใน GAS

- Frontend ขึ้นว่า `ยังไม่ได้ตั้งค่า GAS_WEB_APP_URL`
  - ยังไม่ได้แก้ค่า constant ใน `index.html`

- เรียก Web App ไม่ได้
  - ตรวจสิทธิ์ deployment ว่าอนุญาตให้เข้าถึง URL ได้

## Quick Checklist

- [ ] สร้าง Google Sheet และได้ `SHEET_ID`
- [ ] วาง `gas-sheet.js` + `gas-connect.js` ใน GAS project เดียวกัน
- [ ] ตั้ง Script properties ครบ
- [ ] Deploy Web App สำเร็จ
- [ ] เรียก `?action=init` สำเร็จ
- [ ] เรียก `?action=sync` สำเร็จ
- [ ] ตั้งค่า `GAS_WEB_APP_URL` ใน `index.html`
- [ ] หน้าเว็บโหลดข้อมูลและเขียน IndexedDB ได้
