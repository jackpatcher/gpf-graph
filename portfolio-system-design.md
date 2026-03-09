# Portfolio + P/L System Design (Login + Units)

เอกสารนี้ออกแบบระบบเพิ่มจาก `gpf-graph` เดิม เพื่อให้ผู้ใช้ล็อกอิน และบันทึกว่าในแต่ละวัน/วันที่ปรับแผน มีจำนวนหน่วย (`units`) ของแต่ละแผนเท่าไร แล้วคำนวณผลกำไร/ขาดทุน (+/-) ได้

## 1) เป้าหมาย

- ผู้ใช้ล็อกอินเข้าใช้งานระบบส่วนตัว
- ผู้ใช้บันทึกพอร์ตเป็นรายเหตุการณ์ (ส่วนใหญ่คือวันเปลี่ยนแผน)
- ระบบคำนวณมูลค่าพอร์ตย้อนหลัง/ปัจจุบันจาก `NAV_DATA`
- แสดงผลกำไร/ขาดทุนทั้งแบบรวมและรายวัน
- แยกข้อมูลตามผู้ใช้ (row-level isolation)

## 2) สถาปัตยกรรม

- Frontend (`index.html`): UI + IndexedDB cache + chart rendering
- GAS (`gas-connect.js`): API layer + auth verify + portfolio CRUD + calculation
- Google Sheets:
  - `NAV_DATA` (มีแล้ว)
  - `CONFIG` (มีแล้ว)
  - `USERS` (ใหม่)
  - `PORTFOLIO_SNAPSHOTS` (ใหม่)
  - `PORTFOLIO_EVENTS` (ใหม่, optional)

แนวคิดหลัก:
- ใช้ `Snapshot` เป็นแหล่งจริงของจำนวนหน่วยในแต่ละแผน ณ วันที่ผู้ใช้บันทึก
- วันที่ระหว่าง snapshot ให้ถือจำนวนหน่วยคงเดิมจนถึง snapshot ถัดไป
- มูลค่าพอร์ตคำนวณจาก `units * nav`

## 3) Login Strategy (แนะนำ)

## 3.1 วิธีที่แนะนำ

ใช้ Google Identity (GIS) ฝั่ง frontend แล้วส่ง `id_token` ไป GAS เพื่อตรวจสอบ และ map กับ user ในชีท `USERS`

เหตุผล:
- ไม่ต้องเก็บรหัสผ่านเอง
- ใช้กับ GH Pages ได้
- เหมาะกับ GAS

## 3.2 USERS schema

Sheet: `USERS`
Columns:
- `userId` (internal UUID)
- `googleSub` (unique)
- `email`
- `displayName`
- `role` (`user` | `admin`)
- `status` (`active` | `disabled`)
- `createdAt`
- `lastLoginAt`

Rule:
- ทุก API portfolio ต้องรับ `idToken`
- GAS verify token -> หา `googleSub` -> map เป็น `userId`

## 4) Portfolio Data Model

## 4.1 PORTFOLIO_SNAPSHOTS (หลัก)

Sheet: `PORTFOLIO_SNAPSHOTS`
Columns:
- `snapshotId` (UUID)
- `userId`
- `effectiveDate` (`YYYY-MM-DD`)
- `note` (เช่น "rebalance")
- `units2`
- `units3`
- `units4`
- `units5`
- `units6`
- `units7`
- `units8`
- `units9`
- `units10`
- `units12`
- `units13`
- `units14`
- `units15`
- `cashFlow` (เงินเข้า/ออกในวันนั้น, optional; rebalance ปกติ = 0)
- `updatedAt`

ข้อกำหนด:
- `(userId, effectiveDate)` unique
- units ต้อง >= 0
- ค่าว่างตีเป็น 0

## 4.2 PORTFOLIO_EVENTS (optional, สำหรับ audit)

Sheet: `PORTFOLIO_EVENTS`
Columns:
- `eventId`
- `userId`
- `eventType` (`REBALANCE`, `BUY`, `SELL`, `DEPOSIT`, `WITHDRAW`)
- `eventDate`
- `payloadJson`
- `createdAt`

## 5) Calculation Logic

นิยาม:
- `U[p,d]` = หน่วยแผน `p` ณ วันที่ `d` (จาก snapshot ล่าสุดที่ <= d)
- `NAV[p,d]` = NAV แผน `p` วันที่ `d` จาก `NAV_DATA`
- `MV[d]` = มูลค่าพอร์ต ณ วันที่ `d` = sum(U[p,d] * NAV[p,d])
- `CF[d]` = cash flow วัน `d` (เงินเข้า +, เงินออก -)

สูตรหลัก:
- Cumulative net cash in ถึงวัน `d`:
  - `NCF[d] = sum(CF[t]) for t <= d`
- Total P/L ณ วัน `d`:
  - `PL_total[d] = MV[d] - NCF[d]`
- Daily P/L:
  - `PL_day[d] = MV[d] - MV[d-1] - CF[d]`
- Daily Return (ถ้าต้องการ):
  - `R[d] = PL_day[d] / max(MV[d-1], eps)`

กรณี rebalance ไม่มีเงินเข้าออก:
- ใส่ `CF[d] = 0`
- P/L จะสะท้อนจากการเปลี่ยน NAV เท่านั้น

## 6) API Endpoints (เพิ่มใน gas-connect.js)

Base: `?action=...`

1. `authLogin`
- Input: `idToken`
- Output: `{ userId, email, displayName, role }`

2. `portfolioUpsertSnapshot`
- Input:
  - `idToken`
  - `effectiveDate`
  - `units{plan}` (2,3,4,5,6,7,8,9,10,12,13,14,15)
  - `cashFlow` (optional)
  - `note` (optional)
- Output: `{ snapshotId, effectiveDate }`

3. `portfolioGetSnapshots`
- Input: `idToken`, `startDate`, `endDate`
- Output: snapshots list

4. `portfolioCalc`
- Input: `idToken`, `startDate`, `endDate`
- Output:
  - `timeline[]`: date, marketValue, netCashFlow, totalPL, dayPL
  - `latest`: latest date summary
  - `breakdown`: per-plan value at latest date

5. `portfolioDeleteSnapshot`
- Input: `idToken`, `snapshotId`

## 7) Frontend UX Flow

1. ผู้ใช้กด `Login with Google`
2. เก็บ session token ใน memory (`localStorage` optional)
3. แสดง section `พอร์ตของฉัน`
4. ฟอร์มบันทึก snapshot:
- วันที่มีผล
- หน่วยแต่ละแผน
- เงินเข้าออก (optional)
- หมายเหตุ
5. ปุ่ม `บันทึก` -> เรียก `portfolioUpsertSnapshot`
6. หลังบันทึก:
- เรียก `portfolioCalc`
- update chart + cards (`MV`, `P/L`, `%Return`)

## 8) IndexedDB Integration

- คง `NAV_DATA` cache ใน IndexedDB เหมือนเดิม
- เพิ่ม store ใหม่: `user_portfolio_cache`
  - key: `userId:dateRange`
  - value: timeline + latest summary
- invalidation:
  - เมื่อมี `portfolioUpsertSnapshot`/delete ให้ล้าง cache ของ user

## 9) Security / Data Isolation

- ทุก endpoint พอร์ตต้องตรวจ `idToken`
- ห้ามรับ `userId` จาก frontend โดยตรงเป็น source of truth
- บนชีททุกแถวผูก `userId` ชัดเจน
- validate input:
  - date format
  - units numeric and >= 0
  - cashFlow numeric

## 10) Rollout Plan (แนะนำ)

Phase 1:
- เพิ่มชีท `USERS`, `PORTFOLIO_SNAPSHOTS`
- เพิ่ม endpoint `authLogin`, `portfolioUpsertSnapshot`, `portfolioCalc`
- เพิ่ม UI form บันทึก snapshot

Phase 2:
- เพิ่ม metrics (TWR, MWR)
- เพิ่ม event log + audit
- เพิ่ม export portfolio CSV

## 11) ตัวอย่าง Snapshot Payload

```json
{
  "action": "portfolioUpsertSnapshot",
  "idToken": "<google-id-token>",
  "effectiveDate": "2026-03-09",
  "units2": 1234.5678,
  "units3": 456.789,
  "units4": 0,
  "units5": 0,
  "units6": 0,
  "units7": 321.12,
  "units8": 0,
  "units9": 0,
  "units10": 0,
  "units12": 0,
  "units13": 15.5,
  "units14": 0,
  "units15": 0,
  "cashFlow": 0,
  "note": "rebalance"
}
```

## 12) สิ่งที่ต้องตัดสินใจก่อนเริ่มโค้ดจริง

- จะใช้ Google Login (แนะนำ) หรือ custom login
- จะนับ P/L แบบรวมเงินเข้าออกหรือแบบ Time-weighted เป็นค่า default
- จะบังคับ snapshot ทุกครั้งที่เปลี่ยนหน่วย หรือให้แก้ย้อนหลังได้แค่ admin
