# NKWStudentCare

ระบบเช็กและติดตามนักเรียนเข้าร่วมกิจกรรม/การมาเรียนบน Google Apps Script สำหรับโรงเรียนน้ำคำวิทยา

Repository นี้เก็บ source snapshot ของระบบและเอกสารสรุปการตรวจสอบ Google Apps Script Project history เวอร์ชัน 95 ไว้ที่ [docs/version-95-summary.md](docs/version-95-summary.md)

## โครงสร้างหลัก

ไฟล์ `รหัส.gs` เป็น server-side Apps Script ที่จัดการ Web App routing, Google Sheets, configuration, authentication, attendance และสถิติ ส่วนไฟล์ HTML ประกอบด้วยหน้าแรก หน้าบันทึกข้อมูล หน้าสรุป หน้าสถิติ หน้าทำความสะอาดข้อมูล และหน้าเช็กชื่อรายวิชา โดย client เรียก server ผ่าน `google.script.run`

## หมายเหตุ

ก่อน deploy ควรตรวจสอบสิทธิ์ของ Web App, authentication ของทุก server function, การใช้ `ALLOWALL` และไม่ commit secret หรือข้อมูลนักเรียนจริงลง repository ดูรายละเอียดเพิ่มเติมในเอกสารสรุปเวอร์ชัน 95
