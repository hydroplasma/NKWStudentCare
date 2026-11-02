# สรุปโค้ด Google Apps Script — เวอร์ชัน 95

- **โปรเจกต์:** ระบบเช็คนักเรียน
- **Script ID:** `1I7tP4XLiYemtW0L5ddOu16qmvuLt4bXabVm53BDgkr0FFDuezhd_Ik34`
- **Revision ที่ตรวจสอบ:** เวอร์ชัน 95
- **ผู้แก้ไขในประวัติ:** Thawatchai Kanjak
- **เวลาในประวัติ:** 2 พฤศจิกายน 04:47 (ตามหน้า Project history)

## ภาพรวม

เวอร์ชันนี้เป็น Google Apps Script Web App สำหรับระบบเช็กและติดตามการเข้าร่วมกิจกรรม/การมาเรียนของนักเรียน โดยมีฝั่ง server-side สำหรับ routing หน้าเว็บ การอ่านและประมวลผลข้อมูลจาก Google Sheets การยืนยันตัวตนครู และฝั่ง HTML/JavaScript สำหรับหน้าบันทึกข้อมูล สรุปผล สถิติ และเครื่องมือทำความสะอาดข้อมูล

## ไฟล์ที่พบใน Project history

`appsscript.json`, `รหัส.gs`, `index.html`, `lib.html`, `script_main.html`, `script_summary.html`, `script_period_stats.html`, `stats.html`, `cleaning.html`, `ไม่มีชื่อ.html`, `subject_attendance.html` และ `stats2.html`

## Routing ที่ตรวจสอบจากหน้าจอของเวอร์ชัน 95

ไฟล์ `รหัส.gs` กำหนด `doGet(e)` และเลือก template ตาม query parameter:

- `?page=stats` → `stats.html` พร้อมชื่อหน้า “วิเคราะห์รายชั้น”
- `?page=cleaning` → `cleaning.html` พร้อมชื่อหน้า “เขตพื้นที่”
- `?page=subject_attendance` → `subject_attendance.html` พร้อมชื่อหน้า “เช็คชื่อรายวิชา” และอนุญาตให้ฝังใน iframe/modal
- ไม่ระบุหน้า → `index.html` เป็นหน้าเริ่มต้น

ทุกหน้าที่ route ตรวจสอบมี viewport meta tag และตั้ง `XFrameOptionsMode.ALLOWALL`; หากเกิดข้อผิดพลาด `doGet` จะคืน HTML ที่แสดงข้อความผิดพลาดแทนการล้มเหลวเงียบ ๆ

ฟังก์ชันช่วยที่เห็นในไฟล์ server คือ `getScriptUrl()` สำหรับหา URL ของ Web App และ `include(filename)` สำหรับฝัง HTML partial พร้อมบันทึก error ลง Logger เมื่อหาไฟล์ไม่พบ

## การตั้งค่า manifest

`appsscript.json` ที่ตรวจสอบมีสาระสำคัญดังนี้:

- timezone: `Asia/Bangkok`
- runtime: `V8`
- exception logging: `STACKDRIVER`
- Web App execute as: `USER_DEPLOYING`
- Web App access: `ANYONE_ANONYMOUS`

## หน้าที่หลักของส่วนติดต่อผู้ใช้

- `index.html` — shell/หน้าแรกของระบบ ธีมสีส้ม-ขาว-ดำ และฝัง `lib.html`
- `script_main.html` — หน้าหลักสำหรับ login, เลือกวันที่/คาบ/ห้องเรียน โหลดรายชื่อนักเรียน และบันทึกสถานะ
- `script_summary.html` — สรุปผลรายห้องและภาพรวม พร้อมตาราง/กราฟ
- `script_period_stats.html` — สถิติแยกตามช่วงเวลาและนักเรียนที่ขาดบ่อย
- `stats.html`, `stats2.html` — dashboard วิเคราะห์ข้อมูล กรองข้อมูล ตาราง และสร้างหน้าพิมพ์เอกสาร
- `cleaning.html` — เครื่องมือทำความสะอาด/จัดกลุ่มข้อมูลพื้นที่หรือห้องเรียน
- `subject_attendance.html` — หน้าตรวจสอบการเข้าเรียนรายวิชาใน modal/iframe
- `lib.html` — CSS/utility ที่ใช้ร่วมกัน
- `appsscript.json` — manifest และ deployment configuration
- `รหัส.gs` — API ฝั่ง server สำหรับ Google Sheets, config, authentication, attendance และสถิติ

## การเชื่อมต่อและข้อมูล

ฝั่ง client เรียก server ผ่าน `google.script.run` และใช้ `localStorage` สำหรับ state บางส่วน เช่น session/ตัวเลือกหน้าจอ ส่วนฝั่ง server ใช้ `SpreadsheetApp` เป็นแหล่งข้อมูลหลัก และคืนผลลัพธ์ในรูปแบบ response object สำหรับการแสดงผลในหน้าเว็บ กราฟในหน้าสถิติใช้ไลบรารี chart ที่ฝัง/เรียกจาก HTML template

## ข้อควรระวังด้านความปลอดภัยและการ deploy

1. Manifest เปิด Web App เป็น `ANYONE_ANONYMOUS`; ควรตรวจสอบ authentication ในทุก server function ที่เข้าถึงข้อมูลนักเรียน
2. ควรจำกัดและตรวจสอบค่าที่รับจาก `e.parameter` ก่อนเลือก template หรือ query ข้อมูล
3. ควรทบทวนการใช้ `ALLOWALL` เพราะทำให้หน้าเว็บสามารถถูกฝังใน iframe จาก origin อื่นได้
4. ไม่ควร commit secret, spreadsheet ID ที่เป็นความลับ, token หรือข้อมูลนักเรียนจริงลง repository
5. ก่อน deploy revision ใหม่ควรทดสอบ routing ทั้ง 4 กรณี, login, การบันทึกสถานะ, dashboard และการพิมพ์เอกสาร

## หมายเหตุเกี่ยวกับการนำขึ้น GitHub

Repository นี้มี source snapshot ของระบบอยู่แล้ว และเอกสารฉบับนี้ถูกเพิ่มเพื่อสรุปสิ่งที่ตรวจสอบจากหน้า Project history ของเวอร์ชัน 95 โดยตรง การดึง source ผ่าน API ถูกงดใช้ตามข้อกำหนดของผู้ใช้ให้คัดลอกจากหน้าจอเท่านั้น
