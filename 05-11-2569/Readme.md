# 📚 NKW Student Care 4.0 - ระบบติดตามดูแลและช่วยเหลือนักเรียน

ระบบบริหารจัดการ เช็คชื่อ ประเมินความเสี่ยง และออกเอกสารราชการสำหรับสถานศึกษา พัฒนาด้วย **Google Apps Script (GAS)** ทำหน้าที่เป็นทั้ง Backend และ Frontend ทำงานร่วมกับ **Google Sheets** เป็นฐานข้อมูล

---

## 🏗️ 1. สถาปัตยกรรมและเทคโนโลยี (Tech Stack)
*   **Backend:** Google Apps Script (GAS) `Code.gs`
*   **Frontend:** HTML5, Vanilla JS, jQuery 3.6.0
*   **Styling:** Tailwind CSS (ผ่าน CDN), Custom CSS (Mobile-first, Responsive)
*   **Libraries:** SweetAlert2 (Alerts), Flatpickr (Date Picker), Chart.js (Graphs), DataTables
*   **Database:** Google Sheets
*   **Architecture:** Client-Server ผ่าน `google.script.run` มีการใช้ `Iframe` สำหรับหน้าต่างย่อย (Cleaning & Subject)

---

## 🗂️ 2. โครงสร้างไฟล์ (File Structure)
*   `Code.gs`: รวม API ทั้งหมด (CRUD, Auth, Risk Logic, Doc Generation, Sync)
*   `index.html`: โครงสร้างหน้าหลัก (Shell), Navigation, และ HTML Modals ทั้งหมด
*   `lib.html`: รวม CDN Links (Tailwind, jQuery, ฯลฯ) และ CSS สไตล์หลักของระบบ (Responsive)
*   `script_main.html`: จัดการ Login, Navigation Routing, ระบบเช็คแถว (Attendance) และจัดการสิทธิ์
*   `script_summary.html`: จัดการหน้าสถิติ, Risk Assessment, และระบบวินัย (Discipline)
*   `cleaning.html`: หน้าจอ Iframe สำหรับเช็คเขตพื้นที่ (บันทึกทำเวร/ให้คะแนน)
*   `subject_attendance.html`: หน้าจอ Iframe สำหรับเช็คชื่อเข้าเรียนรายวิชา

---

## 🗄️ 3. โครงสร้างฐานข้อมูล (Database Schema)
ระบบจะไม่ทำงานหากโครงสร้าง Sheet ไม่ตรงกับที่กำหนด:

### `Teachers` (ครู)
*   A: `id` (รหัสครู)
*   B: `name` (ชื่อ-สกุล)
*   C: `username` (ชื่อผู้ใช้)
*   D: `password` (รหัสผ่าน)
*   E: `status` (1 = ใช้งาน, 0 = ระงับ)
*   F: `classroom_ids` (ห้องที่ปรึกษา คั่นด้วย comma เช่น `m1-1,m1-2`)
*   G: `imageUrl` (ลิงก์รูปโปรไฟล์)
*   H: `phone` (เบอร์โทร)
*   I: `role` (`director`, `discipline`, `teacher`)

### `Students` (นักเรียน)
*   A: `classroom_id` (ห้องเรียน)
*   B: `student_id` (รหัสนักเรียน)
*   C: `no` (เลขที่)
*   D: `name` (ชื่อ-สกุล)
*   E: `status` (1 = ใช้งาน)
*   F: `password` (ถ้าว่าง ระบบจะใช้ `student_id` เป็นรหัสผ่าน)
*   G: `role` (`student`, `council`)
*   H: `council_classroom_ids` (ห้องที่สภาฯ ดูแล)

### ตารางอื่นๆ ที่สำคัญ
*   `Config`: เก็บค่า `schoolName`, `director_name`, `discipline_teacher`, `garuda_url` (ดึงจากเซลล์ B14), น้ำหนัก Risk
*   `Timetable`: ตารางสอนของครู (เชื่อมกับ `TeacherEmail` และ `ClassroomId`)
*   `Attendance`: เก็บ JSON สถานะรายวัน (แถว)
*   `Cleaning_Attendance`: เก็บ JSON สถานะการทำเวร
*   `Attendance_Subjects`: เก็บสถานะรายวิชา
*   `Discipline_Cases`: เก็บสถานะเคสวินัย (Stage 1-4, ลายเซ็น, สถานะอนุมัติ)

---

## 🔐 4. ระบบสิทธิ์การใช้งาน (RBAC - Role-Based Access Control)
ระบบมี 5 บทบาท (ต้องพิมพ์ตัวพิมพ์เล็กใน Sheet เท่านั้น):
1.  **`director` (ผอ.):** ดูได้ทุกหน้า, อนุมัติเอกสารได้
2.  **`discipline` (งานกิจการฯ):** ทำได้ทุกอย่างเหมือนครู, ดูวินัยทุกห้อง, **แต่ไม่อนุมัติเอกสาร**
3.  **`teacher` (ครู):** เช็คแถว/เขตพื้นที่ได้เฉพาะห้องตนเอง, เช็คเรียนเฉพาะวิชาตนเอง
4.  **`council` (สภานักเรียน):** ช่วยเช็คแถว/เขตพื้นที่ได้, ห้ามเช็คเรียน/วินัย, มีหน้าดู "ผลของฉัน"
5.  **`student` (นักเรียน):** เข้าดูได้แค่หน้า Dashboard ของตนเอง 100%

---

## ⚙️ 5. เงื่อนไขทางธุรกิจที่สำคัญ (Core Business Logic)

### A. ระบบประเมินความเสี่ยง (Risk Assessment)
*   คำนวณจาก 3 กิจกรรม: แถว (40%), เขตพื้นที่ (30%), เรียน (30%)
*   **Dynamic Weights:** ถ้าสัปดาห์ไหนไม่มีข้อมูลเขตพื้นที่เลย ระบบจะเกลี่ยน้ำหนักที่เหลือให้ 2 กิจกรรม (เช่น กลายเป็น 57% / 43%)
*   จัดกลุ่ม: 🟢 ไว้ใจ (≥85%), 🟡 ห่วงใย (60-84%), 🔴 ใกล้ชิด (<60%) หรือมี Streak ขาดติดกัน ≥ 3 วัน

### B. ระบบวินัย (Discipline Workflow) - *ต้องรันตามลำดับ*
*   **Stage 1 (ตักเตือน):** สร้างเคสอัตโนมัติเมื่อได้ Yellow/Red
*   **Stage 2 (ทัณฑ์บน):** ยกระดับเมื่อผ่านไป 7 วันและยัง Yellow/Red ต้องเซ็น 3 ฝ่าย (ผอ., ผปค., นร.)
*   **Stage 3 (กิจกรรม):** ยกระดับเมื่ออยู่ใน Stage 2 เกิน 14 วันและยัง Red อยู่
*   **Stage 4 (แจ้ง ผปค.):** ครูต้องกด "ไม่ผ่านกิจกรรม" ระบบจะยกระดับเป็น Stage 4 อัตโนมัติ

---

## 🚨 6. หลุมพรางและข้อควรระวัง (CRITICAL GOTCHAS - ห้ามลบ/ห้ามแก้!)
นี่คือจุดที่ Developer เก่าแก้ไขมาแล้ว หากแก้โค้ดส่วนนี้ต้องระวังเป็นพิเศษ:

1.  **Timezone หายนะ:** ห้ามใช้ `new Date().toISOString().split('T')[0]` เด็ดขาด! เพราะจะได้เวลา UTC (ช้ากว่าไทย 7 ชม.) ต้องใช้ฟังก์ชัน `getLocalDateStr()` ใน `script_main` หรือ `Utilities.formatDate(..., 'Asia/Bangkok')` ใน `Code.gs` เสมอ
2.  **เลข 0 หาย (Leading Zeros):** Google Sheets มักแปลง `01496` เป็น `1496` ให้ใช้ `getDisplayValues()` ในการอ่านข้อมูล และใช้ `.replace(/^0+/, '')` ในการเทียบรหัสนักเรียน
3.  **Timetable คาบเรียน `1-2` กลายเป็นวันที่:** Google Sheets จะแปลง "1-2" เป็น 1 ก.พ. ต้องมีฟังก์ชัน `normalizePeriod()` ใน `Code.gs` เพื่อแปลงวันที่กลับเป็นข้อความ "1-2"
4.  **Iframe Cache:** การเปิด Iframe (cleaning.html) ต้องเติม `?v=Date.now()` ต่อท้าย URL เสมอ ไม่งั้นจะโหลดโค้ดเก่าจาก Cache
5.  **Timeout ในปุ่ม Sync:** ฟังก์ชัน Sync วินัยห้ามใช้ `runGoogleScript` ที่มี Timeout 15 วินาที ให้ใช้ `google.script.run` ดั้งเดิม (ใน `syncDisciplineNow`) เพราะบางทีคำนวณเคสเยอะอาจใช้เวลาเกิน
6.  **เอกสารราชการ:** ต้องเป็นสีดำ 100% ห้ามมีตัวอักษรสี (เช่น สีแดง) เด็ดขาด ต้องใช้ฟอนต์ `TH Sarabun PSK` ขนาด 16pt
7.  **ตราครุฑ:** ดึงจาก `Config` เซลล์ `B14` โดยระบบมีฟังก์ชันแปลงลิงก์ Google Drive ให้เป็น Direct Link อัตโนมัติ

---

## 🚀 7. การติดตั้งและ Deploy (Installation Guide)

1.  สร้าง Google Sheets และสร้าง Sheet ตาม Schema ในข้อ 3
2.  เปิด Apps Script (Extensions -> Apps Script)
3.  สร้างไฟล์ HTML และ GS ให้ครบตาม File Structure
4.  แก้ไขตัวแปร `SPREADSHEET_ID` ใน `Code.gs` ให้เป็น ID ของ Sheet คุณ
5.  **รันฟังก์ชัน Setup ครั้งแรก (สำคัญมาก):**
    *   เลือก `setupDisciplineConfig` แล้วกด Run
    *   เลือก `migrateSprint1to4` แล้วกด Run
6.  **Deploy:**
    *   Deploy -> New deployment -> Web app
    *   Execute as: **Me**
    *   Who has access: **Anyone** (หรือตามนโยบาย)
7.  **อัปเดตโค้ดครั้งต่อไป:** ต้องกด Deploy -> Manage Deployments -> สร้าง **New version** ทุกครั้ง (ห้าม Save เฉยๆ)

ออกแบบและพัฒนาโดย © 2025 Waiwai IT & KruBlackThawatchai | Thawatchai Kanjak (เวอร์ชัน 2-ปัจจุบัน)