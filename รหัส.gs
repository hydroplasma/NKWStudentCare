// --- ส่วนที่ 1: ปรับ Routing ใน doGet ---
function doGet(e) {
  try {
    const configRes = getConfig();
    const site_config = configRes.data ? configRes.data.site_config : {};
    const systemName = site_config['systemName'] || 'Student Attendance System';

// หน้า stats
if (e.parameter.page === 'stats') {
  return HtmlService.createTemplateFromFile('stats').evaluate()
    .setTitle(systemName + " - วิเคราะห์รายชั้น")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
} 
// หน้า cleaning (ตัวอย่างการเพิ่มหน้าใหม่)
else if (e.parameter.page === 'cleaning') {
  return HtmlService.createTemplateFromFile('cleaning').evaluate()
    .setTitle(systemName + " - เขตพื้นที่")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// หน้า subject_attendance (สำหรับเปิดใน iframe modal)
else if (e.parameter.page === 'subject_attendance') {
  return HtmlService.createTemplateFromFile('subject_attendance').evaluate()
    .setTitle(systemName + " - เช็คชื่อรายวิชา")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// หน้าหลัก (โหลด index เป็นหน้าแรก)
    return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle(systemName) // หรือเปลี่ยนเป็น systemName + " - วิเคราะห์รายชั้น" ตามที่ต้องการ
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput('เกิดข้อผิดพลาด: ' + error.message);
  }
}

// เพิ่มฟังก์ชันนี้เพื่อดึง URL ของ Web App เรา
function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    Logger.log('Error in include ' + filename + ': ' + error.message + '\nStack: ' + error.stack);
    return '';
  }
}

// Spreadsheet Configuration
const SPREADSHEET_ID = '1BiNWBtHfg7wnZmLwbOSTJS1bfVO5Mnfb6LMGdQcoOLk';
const SHEET_NAMES = {
  CONFIG: 'Config',
  STATUS_LIST: 'StatusList',
  TEACHERS: 'Teachers',
  PERIODS: 'Periods',
  CLASSROOMS: 'Classrooms',
  STUDENTS: 'Students',
  ATTENDANCE: 'Attendance'
};

// 1. เพิ่มตัวแปรชื่อชีตสำหรับระบบเขตพื้นที่
const CLEANING_SHEET_NAMES = {
  ATTENDANCE: 'Cleaning_Attendance',
  SCORES: 'Cleaning_Scores',
  CRITERIA: 'Cleaning_Criteria'
};

const RISK_ACTION_LOG_SHEET_NAME = 'Risk_Action_Log';

// Headers for each sheet
const HEADERS = {
  [SHEET_NAMES.CONFIG]: ['key', 'value'],
  [SHEET_NAMES.STATUS_LIST]: ['label', 'color', 'hexColor', 'status'],
  [SHEET_NAMES.TEACHERS]: ['id', 'name', 'username', 'password', 'status', 'classroom_ids'],
  [SHEET_NAMES.PERIODS]: ['id', 'name', 'status'],
  [SHEET_NAMES.CLASSROOMS]: ['id', 'name', 'status'],
  [SHEET_NAMES.STUDENTS]: ['classroom_id', 'student_id', 'no', 'name', 'status'],
  [SHEET_NAMES.ATTENDANCE]: ['date', 'period', 'room_id', 'teacher', 'students', 'updated_at'],
  
  // 2. เพิ่มหัวตารางสำหรับระบบเขตพื้นที่ (ป้องกัน Error จำนวนคอลัมน์เป็น 0)
  [CLEANING_SHEET_NAMES.ATTENDANCE]: ['Date', 'Classroom ID', 'Checker ID', 'Student Data', 'Updated At'],
  [CLEANING_SHEET_NAMES.SCORES]: ['Date', 'Classroom ID', 'Checker ID', 'Scores', 'Total Score', 'Updated At'],
  [CLEANING_SHEET_NAMES.CRITERIA]: ['ID', 'Name', 'Max Score'],
  [RISK_ACTION_LOG_SHEET_NAME]: ['timestamp', 'action', 'document_type', 'student_id', 'assessment_start', 'assessment_end', 'document_no', 'actor_name']
};

// API Version
const API_VERSION = '1.0';

// Default mode for accessing Spreadsheet
const DEFAULT_SPREADSHEET_MODE = 'byId'; // Can be 'byId' or 'active'

/**
 * Utility function to get Spreadsheet based on mode
 * @param {string} mode - 'byId' or 'active'
 * @returns {Object} Spreadsheet object
 */
function getSpreadsheet(mode = DEFAULT_SPREADSHEET_MODE) {
  try {
    if (mode === 'byId') {
      if (!SPREADSHEET_ID) {
        throw new Error('Spreadsheet ID is not set.');
      }
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } else if (mode === 'active') {
      const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      if (!activeSpreadsheet) {
        throw new Error('No active spreadsheet found.');
      }
      return activeSpreadsheet;
    } else {
      throw new Error('Invalid Spreadsheet mode. Use "byId" or "active".');
    }
  } catch (error) {
    Logger.log('Error in getSpreadsheet: ' + error.message + '\nStack: ' + error.stack);
    throw error;
  }
}

function getStudentAnalysisUrl() {
  const sheet = getSpreadsheet().getSheetByName('Config');
  return sheet.getRange('B8').getValue(); // ดึงค่า URL จาก B8
}

// ========================================
// Sheet setup and data initialization
// Related sheet helpers, setup routines, sample-data initializer, and risk log.
// ========================================

/**
 * Utility function to get or create sheet
 */
function getOrCreateSheet(sheetName, mode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const ss = getSpreadsheet(mode);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = HEADERS[sheetName];
      if (headers) {
        sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
      if (sheetName === SHEET_NAMES.ATTENDANCE) {
        sheet.getRange('A:A').setNumberFormat('@');
      }
    } else {
      const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (existingHeaders.filter(h => h).length === 0) {
        const headers = HEADERS[sheetName];
        if (headers) {
          sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        }
      }
      if (sheetName === SHEET_NAMES.ATTENDANCE) {
        sheet.getRange('A:A').setNumberFormat('@');
      }
    }
    return sheet;
  } catch (error) {
    Logger.log('Error in getOrCreateSheet: ' + error.message + '\nStack: ' + error.stack);
    throw error;
  }
}

function setupCleaningSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheetNames = {
    ATTENDANCE: 'Cleaning_Attendance', 
    SCORES: 'Cleaning_Scores',         
    CRITERIA: 'Cleaning_Criteria'      
  };

  // 1. จัดการชีต Attendance
  // ใช้ || (OR) ถ้าหาชีตไม่เจอให้สร้างใหม่ทันที เพื่อเลี่ยงการเขียนทับตัวแปร
  const sheetAtt = ss.getSheetByName(sheetNames.ATTENDANCE) || ss.insertSheet(sheetNames.ATTENDANCE);
  const headersAtt = ['Date', 'Classroom ID', 'Checker ID', 'Student Data', 'Updated At'];
  sheetAtt.getRange(1, 1, 1, headersAtt.length).setValues([headersAtt])
          .setFontWeight("bold").setBackground("#d9ead3");

  // 2. จัดการชีต Scores
  const sheetSco = ss.getSheetByName(sheetNames.SCORES) || ss.insertSheet(sheetNames.SCORES);
  const headersSco = ['Date', 'Classroom ID', 'Checker ID', 'Scores', 'Total Score', 'Updated At'];
  sheetSco.getRange(1, 1, 1, headersSco.length).setValues([headersSco])
          .setFontWeight("bold").setBackground("#c9daf8");

  // 3. จัดการชีต Criteria
  let sheetCri = ss.getSheetByName(sheetNames.CRITERIA);
  if (!sheetCri) {
    sheetCri = ss.insertSheet(sheetNames.CRITERIA);
    // สร้างข้อมูลเกณฑ์ตั้งต้น
    const criteriaData = [
      ['ID', 'Name', 'Max Score'],
      ['C01', 'กวาดพื้น/ถูพื้น', 10],
      ['C02', 'ทิ้งขยะ', 5],
      ['C03', 'ความเป็นระเบียบ', 5]
    ];
    sheetCri.getRange(1, 1, criteriaData.length, criteriaData[0].length).setValues(criteriaData);
    sheetCri.getRange(1, 1, 1, criteriaData[0].length).setFontWeight("bold").setBackground("#fff2cc");
  } else {
    // ถ้ามีอยู่แล้วแค่เขียนหัวตาราง
    const headersCri = ['ID', 'Name', 'Max Score'];
    sheetCri.getRange(1, 1, 1, headersCri.length).setValues([headersCri])
            .setFontWeight("bold").setBackground("#fff2cc");
  }

  console.log('✅ สร้างชีตและหัวตารางสำเร็จ!');
}

// ฟังก์ชันสำหรับสร้างฐานข้อมูลระบบเช็คชื่อรายวิชา (รันเพียงครั้งเดียว)
function setupSubjectDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  try {
    // ---------------------------------------------------------
    // 1. สร้างชีต Timetable (ตารางสอน)
    // ---------------------------------------------------------
    var timetableSheet = ss.getSheetByName("Timetable");
    if (!timetableSheet) {
      timetableSheet = ss.insertSheet("Timetable");
      var timetableHeaders = [
        "TimetableID", "DayOfWeek", "Period", "SubjectCode", "SubjectName", 
        "TeacherEmail", "GradeLevel", "Room", "ClassroomId"
      ];
      timetableSheet.appendRow(timetableHeaders);
      
      // จัดรูปแบบหัวตารางให้สวยงาม
      timetableSheet.getRange(1, 1, 1, timetableHeaders.length)
        .setFontWeight("bold")
        .setBackground("#f9cb9c") // สีส้มอ่อน
        .setHorizontalAlignment("center");
        
      // ตรึงแถวบนสุด
      timetableSheet.setFrozenRows(1);
      
      // เพิ่มข้อมูลตัวอย่าง 1 บรรทัด เพื่อให้ครูดูเป็นแนวทาง
      var exampleTeacherEmail = Session.getActiveUser().getEmail() || "teacher@namkham.ac.th";
      timetableSheet.appendRow([
        "T001", "จันทร์", 1, "ว23101", "วิทยาศาสตร์", 
        exampleTeacherEmail, "ม.3", "1", "m3_1"
      ]);
      
      // ปรับขนาดคอลัมน์อัตโนมัติ
      timetableSheet.autoResizeColumns(1, timetableHeaders.length);
    }

    // ---------------------------------------------------------
    // 2. สร้างชีต Attendance_Subjects (บันทึกผลการเข้าเรียน)
    // ---------------------------------------------------------
    var attendanceSheet = ss.getSheetByName("Attendance_Subjects");
    if (!attendanceSheet) {
      attendanceSheet = ss.insertSheet("Attendance_Subjects");
      var attendanceHeaders = [
        "Timestamp", "Date", "SubjectCode", "SubjectName", "TeacherEmail", 
        "GradeLevel", "Room", "StudentID", "StudentName", "Status", "Note"
      ];
      attendanceSheet.appendRow(attendanceHeaders);
      
      // จัดรูปแบบหัวตารางให้สวยงาม
      attendanceSheet.getRange(1, 1, 1, attendanceHeaders.length)
        .setFontWeight("bold")
        .setBackground("#d9ead3") // สีเขียวอ่อน
        .setHorizontalAlignment("center");
        
      // ตรึงแถวบนสุด
      attendanceSheet.setFrozenRows(1);
      
      // ปรับขนาดคอลัมน์อัตโนมัติ
      attendanceSheet.autoResizeColumns(1, attendanceHeaders.length);
    }
    
    ui.alert('✅ สำเร็จ!', 'สร้างชีต "Timetable" และ "Attendance_Subjects" เรียบร้อยแล้ว\n\nกรุณาไปที่ Google Sheets เพื่อแก้ไขข้อมูลตารางสอนตัวอย่างให้เป็นข้อมูลจริงครับ', ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('❌ เกิดข้อผิดพลาด', error.toString(), ui.ButtonSet.OK);
  }
}

// Create sample data
function createSampleData(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  // ฟังก์ชันเดิม ไม่เปลี่ยนแปลง
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return createResponse('error', 'Server is busy, please try again', null, { code: 'LOCK_TIMEOUT' });
    }

    try {
      // Config
      const configSheet = getOrCreateSheet(SHEET_NAMES.CONFIG, spreadsheetMode);
      configSheet.getDataRange().clearContent();
      configSheet.getRange(1, 1, 1, HEADERS[SHEET_NAMES.CONFIG].length).setValues([HEADERS[SHEET_NAMES.CONFIG]]);
      const configData = [
        ['logoUrl', 'https://i.postimg.cc/RZ0PCqVy/NKW-LOGO.png'],
        ['systemName', 'ระบบเช็คชื่อนักเรียน โรงเรียนน้ำคำวิทยา'],
        ['systemAdmin', '© 2025 Waiwai IT & KruBlackThawatchai'],
        ['schoolName', 'โรงเรียนน้ำคำวิทยา'],
        ['semester', 'ภาคเรียนที่ 1'],
        ['schoolYear', '2569']
      ];
      configSheet.getRange(2, 1, configData.length, configData[0].length).setValues(configData);

      // StatusList
      const statusSheet = getOrCreateSheet(SHEET_NAMES.STATUS_LIST, spreadsheetMode);
      statusSheet.getDataRange().clearContent();
      statusSheet.getRange(1, 1, 1, HEADERS[SHEET_NAMES.STATUS_LIST].length).setValues([HEADERS[SHEET_NAMES.STATUS_LIST]]);
      const statusData = [
        ['มา', 'bg-green-100', '#4ade80', 1],
        ['ขาด', 'bg-red-100', '#f87171', 1],
        ['สาย', 'bg-yellow-100', '#facc15', 1],
        ['ลา', 'bg-blue-100', '#60a5fa', 1],
        ['ป่วย', 'bg-purple-100', '#c084fc', 1]
      ];
      statusSheet.getRange(2, 1, statusData.length, statusData[0].length).setValues(statusData);

      // Teachers
const teachersSheet = getOrCreateSheet(SHEET_NAMES.TEACHERS, spreadsheetMode);
teachersSheet.getDataRange().clearContent();
teachersSheet.getRange(1, 1, 1, HEADERS[SHEET_NAMES.TEACHERS].length).setValues([HEADERS[SHEET_NAMES.TEACHERS]]);
// ✅ ไม่สร้างครูตัวอย่าง — ให้แอดมินเพิ่มครูจริงในชีตเอง

      // Periods
      const periodsSheet = getOrCreateSheet(SHEET_NAMES.PERIODS, spreadsheetMode);
      periodsSheet.getDataRange().clearContent();
      periodsSheet.getRange(1, 1, 1, HEADERS[SHEET_NAMES.PERIODS].length).setValues([HEADERS[SHEET_NAMES.PERIODS]]);
      const periodsData = [
        ['morning', 'เช้า', 1],
        ['noon', 'เที่ยง', 1],
        ['evening', 'เย็น', 1]
      ];
      periodsSheet.getRange(2, 1, periodsData.length, periodsData[0].length).setValues(periodsData);

      // Classrooms
      const classroomsSheet = getOrCreateSheet(SHEET_NAMES.CLASSROOMS, spreadsheetMode);
      classroomsSheet.getDataRange().clearContent();
      classroomsSheet.getRange(1, 1, 1, HEADERS[SHEET_NAMES.CLASSROOMS].length).setValues([HEADERS[SHEET_NAMES.CLASSROOMS]]);
      const classroomsData = [
        ['m1-1', 'ม.1/1', 1],
        ['m2-1', 'ม.2/1', 1]
      ];
      classroomsSheet.getRange(2, 1, classroomsData.length, classroomsData[0].length).setValues(classroomsData);

      // Students
      const studentsSheet = getOrCreateSheet(SHEET_NAMES.STUDENTS, spreadsheetMode);
      studentsSheet.getDataRange().clearContent();
      studentsSheet.getRange(1, 1, 1, HEADERS[SHEET_NAMES.STUDENTS].length).setValues([HEADERS[SHEET_NAMES.STUDENTS]]);
      const studentsData = [
        ['m1-1', '11201', 1, 'เด็กชายภูวดล ใจดี', 1],
        ['m1-1', '11202', 2, 'เด็กหญิงนลิน ใจงาม', 1],
        ['m2-1', '12201', 1, 'เด็กชายสมศักดิ์ รักเรียน', 1],
        ['m2-1', '12202', 2, 'เด็กหญิงสมหญิง เรียนดี', 1]
      ];
      studentsSheet.getRange(2, 1, studentsData.length, studentsData[0].length).setValues(studentsData);

      return createResponse('success', 'Sample data created successfully');
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    Logger.log('Error in createSampleData: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to create sample data: ' + error.message, null, error);
  }
}

/**
 * บันทึกเฉพาะการสร้างเอกสารพิมพ์ ไม่ได้บันทึกการออกหนังสือหรือการติดต่อผู้ปกครอง
 * เพิ่มแถวใหม่เท่านั้น; ไม่แก้ไขหรือลบรายการเดิม
 */
function logRiskDocumentGeneration(documentType, studentId, startDate, endDate, documentNo, actorName) {
  let lock;
  let lockAcquired = false;
  try {
    const allowedTypes = ['warning', 'activity', 'summary'];
    const type = String(documentType || '').trim();
    if (allowedTypes.indexOf(type) === -1) return { status: 'error', message: 'ประเภทเอกสารไม่ถูกต้อง' };

    const start = String(startDate || '').trim();
    const end = String(endDate || '').trim();
    const validDate = function(value) { return !value || /^\d{4}-\d{2}-\d{2}$/.test(value); };
    if (!validDate(start) || !validDate(end) || (start && end && start > end)) {
      return { status: 'error', message: 'ช่วงวันที่ประเมินไม่ถูกต้อง' };
    }

    const lockService = LockService.getScriptLock();
    lock = lockService;
    if (!lock.tryLock(10000)) return { status: 'error', message: 'Server busy' };
    lockAcquired = true;

    const sheet = getOrCreateSheet(RISK_ACTION_LOG_SHEET_NAME);
    sheet.appendRow([
      new Date(),
      'สร้างเอกสาร',
      type,
      String(studentId || '').trim().slice(0, 100),
      start,
      end,
      String(documentNo || '').trim().slice(0, 150),
      String(actorName || '').trim().slice(0, 150)
    ]);
    return { status: 'success', message: 'บันทึกการสร้างเอกสารแล้ว' };
  } catch (err) {
    Logger.log('Error in logRiskDocumentGeneration: ' + (err.message || err));
    return { status: 'error', message: err.message || String(err) };
  } finally {
    if (lock && lockAcquired) lock.releaseLock();
  }
}

// Utility function to create standardized response
function createResponse(status, message, data = null, error = null) {
  const response = {
    status,
    message,
    data,
    timestamp: new Date().toISOString(),
    version: API_VERSION
  };
  if (error) {
    response.error = {
      code: error.code || 'UNKNOWN_ERROR',
      details: error.message
    };
  }
  return response;
}

// ========================================
// ฟังก์ชันล้างข้อมูลซ้ำ (ดึงเฉพาะข้อมูลล่าสุด)
// ========================================

function getCleanedAttendanceRows(data) {
  if (!data || data.length <= 1) return [];
  const uniqueRecords = {};
  
  data.slice(1).forEach(row => {
    let rDate = '';
    if (row[0] instanceof Date) {
      rDate = Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd');
    } else {
      rDate = String(row[0] || '').trim().substring(0, 10);
    }
    const rPeriod = String(row[1] || '').trim();
    const rRoom = String(row[2] || '').trim();
    
    // สร้าง Key จาก วันที่_ช่วงเวลา_ห้อง หากซ้ำ ข้อมูลล่างสุด(ล่าสุด) จะทับข้อมูลเก่า
    if (rDate && rPeriod && rRoom) {
      const key = `${rDate}_${rPeriod}_${rRoom}`;
      uniqueRecords[key] = row; 
    }
  });
  return Object.values(uniqueRecords);
}

function getCleanedCleaningRows(data) {
  if (!data || data.length <= 1) return [];
  const uniqueRecords = {};
  data.slice(1).forEach(row => {
    let rDate = (row[0] instanceof Date) ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(row[0] || '').trim().substring(0, 10);
    const rRoom = String(row[1] || '').trim();
    if (rDate && rRoom) {
      uniqueRecords[`${rDate}_${rRoom}`] = row;
    }
  });
  return Object.values(uniqueRecords);
}

// ========================================

// Authenticate teacher
function authenticateTeacher(username, password, spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    if (!username || !password) {
      throw new Error('Username or password is missing');
    }
    const sheet = getOrCreateSheet(SHEET_NAMES.TEACHERS, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    
    const teacher = data.slice(1).find(row =>
      `${row[2]}` === `${username}` &&
      `${row[3]}` === `${password}` &&
      `${row[4]}` === '1'
    );

    if (teacher) {
      return createResponse('success', 'Authentication successful', {
        id: teacher[0],
        name: teacher[1],
        username: teacher[2],
        classroom_ids: teacher[5] ? teacher[5].split(',') : []
      });
    } else {
      return createResponse('error', 'Invalid username or password', null, { code: 'AUTH_FAILED' });
    }
  } catch (error) {
    Logger.log('Error in authenticateTeacher: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to authenticate: ' + error.message, null, error);
  }
}

function testGetConfig() {
  const data = getConfig().data;
  const site_config = data.site_config['systemName'];
  Logger.log(site_config);
}

// Get configuration
function getConfig(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.CONFIG, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const config = {
      site_config: {},
      school_config: {}
    };
    data.slice(1).forEach(row => {
      const key = row[0];
      const value = row[1];
      if (['logoUrl', 'systemName', 'systemAdmin'].includes(key)) {
        config.site_config[key] = value;
      } else if (['schoolName', 'semester','schoolYear'].includes(key)) {
        config.school_config[key] = value;
      }
    });
    return createResponse('success', 'Config retrieved successfully', config);
  } catch (error) {
    Logger.log('Error in getConfig: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to get config: ' + error.message, null, error);
  }
}

// Get status list
function getStatusList(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.STATUS_LIST, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const statusList = data.length <= 1 ? [] : data.slice(1)
      .filter(row => row[3] === 1)
      .map(row => ({
        label: row[0],
        color: row[1],
        hexColor: row[2],
        status: row[3]
      }));
    return createResponse('success', 'Status list retrieved successfully', statusList);
  } catch (error) {
    Logger.log('Error in getStatusList: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to get status list: ' + error.message, null, error);
  }
}

// Get all teachers
function getTeachers(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.TEACHERS, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const teachers = data.length <= 1 ? [] : data.slice(1).map(row => ({
      id: row[0],
      name: row[1],
      username: row[2],
      password: row[3],
      status: row[4],
      classroom_ids: row[5] ? row[5].split(',') : []
    }));
    return createResponse('success', 'Teachers retrieved successfully', teachers);
  } catch (error) {
    Logger.log('Error in getTeachers: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to get teachers: ' + error.message, null, error);
  }
}

// Get all periods
function getPeriods(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.PERIODS, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const periods = data.length <= 1 ? [] : data.slice(1)
      .filter(row => row[2] === 1) 
      .map(row => ({
        id: row[0],
        name: row[1],
        status: row[2]
      }));
    return createResponse('success', 'Periods retrieved successfully', periods);
  } catch (error) {
    Logger.log('Error in getPeriods: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to get periods: ' + error.message, null, error);
  }
}

// Get classrooms
function getClassrooms(teacherClassroomIds = [], spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    if (!teacherClassroomIds || !Array.isArray(teacherClassroomIds)) {
      return createResponse('error', 'Invalid classroom_ids', null, { code: 'INVALID_INPUT' });
    }
    const sheet = getOrCreateSheet(SHEET_NAMES.CLASSROOMS, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const classrooms = data.length <= 1 ? [] : data.slice(1)
      .filter(row => row[2] === 1 && (teacherClassroomIds.length === 0 || teacherClassroomIds.includes(row[0].toString().trim())))
      .map(row => ({
        id: row[0].toString().trim(),
        name: row[1].toString().trim(),
        status: row[2]
      }));
    return createResponse('success', 'Classrooms retrieved successfully', classrooms);
  } catch (error) {
    Logger.log('Error in getClassrooms: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to get classrooms: ' + error.message, null, error);
  }
}

// Get students by classroom IDs
function getStudents(classroomIds, spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    if (!classroomIds || !Array.isArray(classroomIds) || !classroomIds.length) {
      return createResponse('error', 'Invalid or empty classroom_ids', null, { code: 'INVALID_INPUT' });
    }
    const sheet = getOrCreateSheet(SHEET_NAMES.STUDENTS, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const studentsByClassroom = {};
    classroomIds.forEach(id => {
      studentsByClassroom[id] = data.length <= 1 ? [] : data.slice(1)
        .filter(row => row[0].toString().trim() === id && row[4] === 1)
        .map(row => ({
          classroom_id: row[0].toString().trim(),
          student_id: row[1].toString().trim(),
          no: row[2].toString().trim(),
          name: row[3].toString().trim(),
          status: row[4]
        }));
    });
    return createResponse('success', 'Students retrieved successfully', studentsByClassroom);
  } catch (error) {
    Logger.log('Error in getStudents: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to get students: ' + error.message, null, error);
  }
}

// Get all students grouped by classroom
function getAllStudents(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.STUDENTS, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const studentsByClassroom = {};
    if (data.length > 1) {
      data.slice(1)
        .filter(row => row[4] === 1)
        .forEach(row => {
          const classroomId = row[0].toString().trim();
          if (!studentsByClassroom[classroomId]) {
            studentsByClassroom[classroomId] = [];
          }
          studentsByClassroom[classroomId].push({
            classroom_id: row[0].toString().trim(),
            student_id: row[1].toString().trim(),
            no: row[2].toString().trim(),
            name: row[3].toString().trim(),
            status: row[4]
          });
        });
    }
    return createResponse('success', 'All students retrieved successfully', studentsByClassroom);
  } catch (error) {
    Logger.log('Error in getAllStudents: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to get all students: ' + error.message, null, error);
  }
}

// Get attendance records
function getAttendance(date, period, roomId, spreadsheetMode) {
  if (spreadsheetMode === undefined) spreadsheetMode = DEFAULT_SPREADSHEET_MODE;
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.ATTENDANCE, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    
    // กรองข้อมูลซ้ำก่อน
    const cleanData = getCleanedAttendanceRows(data);
    
    const attendance = cleanData.filter((row, index) => {
        const rowDate = typeof row[0] === 'string' ? row[0].trim() : 
                        row[0] instanceof Date ? row[0].toISOString().split('T')[0] : 
                        String(row[0] || '').trim();
        const rowPeriod = typeof row[1] === 'string' ? row[1].trim() : String(row[1] || '').trim();
        const rowRoomId = typeof row[2] === 'string' ? row[2].trim() : String(row[2] || '').trim();

        if (!rowDate || !rowPeriod || !rowRoomId || !/^\d{4}-\d{2}-\d{2}$/.test(rowDate)) {
          return false;
        }

        return (!date || rowDate === date.trim()) && 
               (!period || rowPeriod === period.trim()) && 
               (!roomId || rowRoomId === roomId.trim());
      })
      .map(row => ({
        date: typeof row[0] === 'string' ? row[0].trim() : 
              row[0] instanceof Date ? row[0].toISOString().split('T')[0] : 
              String(row[0] || '').trim(),
        period: typeof row[1] === 'string' ? row[1].trim() : String(row[1] || '').trim(),
        room_id: typeof row[2] === 'string' ? row[2].trim() : String(row[2] || '').trim(),
        teacher: typeof row[3] === 'string' ? row[3].trim() : String(row[3] || '').trim(),
        students: JSON.parse(row[4] || "[]"),
        updated_at: typeof row[5] === 'string' ? row[5].trim() : String(row[5] || '').trim()
      }));
    return createResponse('success', 'Attendance retrieved successfully', attendance);
  } catch (error) {
    Logger.log('Error in getAttendance: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to get attendance: ' + error.message, null, error);
  }
}

// Save or update attendance
function saveAttendance(jsonData, spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return createResponse('error', 'Server is busy, please try again', null, { code: 'LOCK_TIMEOUT' });
    }

    try {
      const entry = JSON.parse(jsonData);
      let { date, period, room_id, teacher, students, updated_at } = entry;

      if (!date || !period || !room_id || !teacher || !students || !updated_at) {
        return createResponse('error', 'Missing required fields', null, { code: 'INVALID_INPUT' });
      }

      date = date.trim();
      period = period.trim();
      room_id = room_id.trim();
      teacher = teacher.trim();
      updated_at = updated_at.trim();

      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
        return createResponse('error', 'Invalid date format', null, { code: 'INVALID_DATE' });
      }

      if (!Array.isArray(students) || students.length === 0 || students.some(s => !s || !s.student_id || !s.status)) {
        return createResponse('error', 'Invalid students data', null, { code: 'INVALID_STUDENTS' });
      }

      // Server-side permission check against the active teacher and assigned classrooms.
      const teachersSheet = getOrCreateSheet(SHEET_NAMES.TEACHERS, spreadsheetMode);
      const teachersData = teachersSheet.getDataRange().getValues();
      const teacherRow = teachersData.slice(1).find(row =>
        String(row[2] || '').trim() === String(teacher).trim() && String(row[4] || '') === '1'
      );
      if (!teacherRow) {
        Logger.log(`❌ [Attendance] Auth failed: teacher="${teacher}" not found or inactive`);
        return createResponse('error', 'ไม่พบข้อมูลครูผู้บันทึก หรือบัญชีถูกปิดใช้งาน', null, { code: 'AUTH_FAILED' });
      }
      const allowedRooms = String(teacherRow[5] || '').split(',').map(r => r.trim()).filter(Boolean);
      if (!allowedRooms.includes(room_id)) {
        Logger.log(`❌ [Attendance] Permission denied: "${teacher}" tried room "${room_id}". Allowed: [${allowedRooms.join(', ') || 'NONE'}]`);
        return createResponse('error', 'คุณไม่มีสิทธิ์เช็คชื่อห้องนี้ (เฉพาะครูที่ปรึกษาเท่านั้น)', null, { code: 'PERMISSION_DENIED' });
      }
      Logger.log(`✅ [Attendance] Permission granted: ${teacher} → ${room_id} (${date}/${period})`);

      const sheet = getOrCreateSheet(SHEET_NAMES.ATTENDANCE, spreadsheetMode);
      const data = sheet.getDataRange().getValues();
      
      const existingIndex = data.slice(1).findIndex(row => {
        const rowDate = typeof row[0] === 'string' ? row[0].trim() : 
                        row[0] instanceof Date ? row[0].toISOString().split('T')[0] : 
                        String(row[0] || '').trim();
        const rowPeriod = typeof row[1] === 'string' ? row[1].trim() : String(row[1] || '').trim();
        const rowRoomId = typeof row[2] === 'string' ? row[2].trim() : String(row[2] || '').trim();
        return rowDate === date && rowPeriod === period && rowRoomId === room_id;
      }) + 2;

      const rowData = [
        date,
        period,
        room_id,
        teacher,
        JSON.stringify(students),
        updated_at
      ];

      sheet.getRange('A:A').setNumberFormat('@');

      if (existingIndex > 1) {
        sheet.getRange(existingIndex, 1, 1, rowData.length).setValues([rowData]);
        return createResponse('success', 'Updated attendance successfully');
      } else {
        sheet.appendRow(rowData);
        return createResponse('success', 'Saved attendance successfully');
      }
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    Logger.log('Error in saveAttendance: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to save attendance: ' + error.message, null, error);
  }
}

// Get initial data
function getInitialData(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const ss = getSpreadsheet(spreadsheetMode);
    const sheets = {
      config: getOrCreateSheet(SHEET_NAMES.CONFIG, spreadsheetMode),
      statusList: getOrCreateSheet(SHEET_NAMES.STATUS_LIST, spreadsheetMode),
      teachers: getOrCreateSheet(SHEET_NAMES.TEACHERS, spreadsheetMode),
      periods: getOrCreateSheet(SHEET_NAMES.PERIODS, spreadsheetMode)
    };

    const data = {};
    for (const [key, sheet] of Object.entries(sheets)) {
      data[key] = sheet.getDataRange().getValues();
    }

    const config = { site_config: {}, school_config: {} };
    if (data.config.length <= 1) {
      Logger.log('Warning: Config sheet is empty or has no data');
    } else {
      data.config.slice(1).forEach(row => {
        if (row.length < 2) return; 
        const key = row[0]?.toString().trim();
        const value = row[1]?.toString().trim();
        if (['logoUrl', 'systemName', 'systemAdmin'].includes(key)) {
          config.site_config[key] = value;
        } else if (['schoolName', 'semester', 'schoolYear'].includes(key)) {
          config.school_config[key] = value;
        }
      });
    }

    const statusList = data.statusList.length <= 1 ? [] : data.statusList.slice(1)
      .filter(row => row.length >= 4 && row[3] === 1)
      .map(row => ({
        label: row[0]?.toString().trim(),
        color: row[1]?.toString().trim(),
        hexColor: row[2]?.toString().trim(),
        status: row[3]
      }));

    const teachers = data.teachers.length <= 1 ? [] : data.teachers.slice(1)
      .filter(row => row.length >= 6)
      .map(row => ({
        id: row[0]?.toString().trim(),
        name: row[1]?.toString().trim(),
        username: row[2]?.toString().trim(),
        password: row[3]?.toString().trim(),
        status: row[4],
        classroom_ids: row[5] ? row[5].toString().trim().split(',').map(id => id.trim()) : []
      }));

    const periods = data.periods.length <= 1 ? [] : data.periods.slice(1)
      .filter(row => row.length >= 3 && row[2] === 1)
      .map(row => ({
        id: row[0]?.toString().trim(),
        name: row[1]?.toString().trim(),
        status: row[2]
      }));

    const responseData = { config, statusList, teachers, periods };
    
    return createResponse('success', 'Initial data retrieved successfully', responseData);
  } catch (error) {
    Logger.log('Error in getInitialData: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to get initial data: ' + error.message, null, error);
  }
}

function getLoginData(username, password, spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const ss = getSpreadsheet(spreadsheetMode);
    const sheets = {
      teachers: getOrCreateSheet(SHEET_NAMES.TEACHERS, spreadsheetMode),
      classrooms: getOrCreateSheet(SHEET_NAMES.CLASSROOMS, spreadsheetMode),
      students: getOrCreateSheet(SHEET_NAMES.STUDENTS, spreadsheetMode)
    };
    const data = {};
    for (const [key, sheet] of Object.entries(sheets)) {
      data[key] = sheet.getDataRange().getValues();
    }

    const teacher = data.teachers.slice(1).find(row =>
      row[2]?.toString().trim() === username &&
      row[3]?.toString().trim() === password &&
      row[4] === 1
    );
    if (!teacher) {
      return createResponse('error', 'Invalid username or password', null, { code: 'AUTH_FAILED' });
    }
    const teacherData = {
      id: teacher[0]?.toString().trim(),
      name: teacher[1]?.toString().trim(),
      username: teacher[2]?.toString().trim(),
      classroom_ids: teacher[5] ? teacher[5].toString().trim().split(',').map(id => id.trim()) : []
    };

    const classrooms = data.classrooms.length <= 1 ? [] : data.classrooms.slice(1)
      .filter(row => row[2] === 1 && teacherData.classroom_ids.includes(row[0]?.toString().trim()))
      .map(row => ({
        id: row[0]?.toString().trim(),
        name: row[1]?.toString().trim(),
        status: row[2]
      }));

    const studentsByClassroom = {};
    teacherData.classroom_ids.forEach(id => {
      studentsByClassroom[id] = data.students.length <= 1 ? [] : data.students.slice(1)
        .filter(row => row[0]?.toString().trim() === id && row[4] === 1)
        .map(row => ({
          classroom_id: row[0]?.toString().trim(),
          student_id: row[1]?.toString().trim(),
          no: row[2]?.toString().trim(),
          name: row[3]?.toString().trim(),
          status: row[4]
        }));
    });
    const studentCounts = [];
    for (const roomId in studentsByClassroom) {
      studentCounts.push({ room_id: roomId, count: studentsByClassroom[roomId].length });
    }

    const responseData = {
      teacher: teacherData,
      classrooms,
      studentCounts
    };

    return createResponse('success', 'Login successful', responseData);
  } catch (error) {
    Logger.log('Error in getLoginData: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to get login data: ' + error.message, null, error);
  }
}

function getStudentsAndAttendance(classroomId, date, period, spreadsheetMode) {
  if (spreadsheetMode === undefined) spreadsheetMode = DEFAULT_SPREADSHEET_MODE;
  try {
    const ss = getSpreadsheet(spreadsheetMode);
    const sheets = {
      students: getOrCreateSheet(SHEET_NAMES.STUDENTS, spreadsheetMode),
      attendance: getOrCreateSheet(SHEET_NAMES.ATTENDANCE, spreadsheetMode)
    };
    const data = {};
    for (const [key, sheet] of Object.entries(sheets)) {
      data[key] = sheet.getDataRange().getValues();
    }

    const students = data.students.length <= 1 ? [] : data.students.slice(1)
      .filter(row => row[0]?.toString().trim() === classroomId && row[4] === 1)
      .map(row => ({
        classroom_id: row[0]?.toString().trim(),
        student_id: row[1]?.toString().trim(),
        no: row[2]?.toString().trim(),
        name: row[3]?.toString().trim(),
        status: row[4]
      }));

    // กรองข้อมูลซ้ำก่อน
    const cleanData = getCleanedAttendanceRows(data.attendance);

    const attendance = cleanData.filter(row => {
        const rowDate = typeof row[0] === 'string' ? row[0].trim() : 
                        row[0] instanceof Date ? row[0].toISOString().split('T')[0] : 
                        String(row[0] || '').trim();
        const rowPeriod = typeof row[1] === 'string' ? row[1].trim() : String(row[1] || '').trim();
        const rowRoomId = typeof row[2] === 'string' ? row[2].trim() : String(row[2] || '').trim();
        return rowDate === date && rowPeriod === period && rowRoomId === classroomId;
      })
      .map(row => ({
        date: typeof row[0] === 'string' ? row[0].trim() : 
              row[0] instanceof Date ? row[0].toISOString().split('T')[0] : 
              String(row[0] || '').trim(),
        period: typeof row[1] === 'string' ? row[1].trim() : String(row[1] || '').trim(),
        room_id: typeof row[2] === 'string' ? row[2].trim() : String(row[2] || '').trim(),
        teacher: typeof row[3] === 'string' ? row[3].trim() : String(row[3] || '').trim(),
        students: JSON.parse(row[4] || "[]"),
        updated_at: typeof row[5] === 'string' ? row[5].trim() : String(row[5] || '').trim()
      }));

    const responseData = {
      students,
      attendance: attendance[0] || null
    };

    return createResponse('success', 'Students and attendance retrieved', responseData);
  } catch (error) {
    Logger.log('Error in getStudentsAndAttendance: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to get data: ' + error.message, null, error);
  }
}

// Get attendance records for date range
function getAttendanceRange(startDate, endDate, period, spreadsheetMode) {
  if (spreadsheetMode === undefined) spreadsheetMode = DEFAULT_SPREADSHEET_MODE;
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.ATTENDANCE, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    
    // กรองข้อมูลซ้ำก่อน
    const cleanData = getCleanedAttendanceRows(data);
    
    const attendance = cleanData.filter((row, index) => {
        const rowDate = typeof row[0] === 'string' ? row[0].trim() : 
          row[0] instanceof Date ? row[0].toISOString().split('T')[0] : 
          String(row[0] || '').trim();
        const rowPeriod = typeof row[1] === 'string' ? row[1].trim() : String(row[1] || '').trim();
        
        if (!rowDate || !rowPeriod || !/^\d{4}-\d{2}-\d{2}$/.test(rowDate)) return false;
        
        const isInDateRange = rowDate >= startDate.trim() && rowDate <= endDate.trim();
        const isPeriodMatch = !period || rowPeriod === period.trim();
        
        return isInDateRange && isPeriodMatch;
      })
      .map(row => ({
        date: typeof row[0] === 'string' ? row[0].trim() : 
          row[0] instanceof Date ? row[0].toISOString().split('T')[0] : 
          String(row[0] || '').trim(),
        period: typeof row[1] === 'string' ? row[1].trim() : String(row[1] || '').trim(),
        room_id: typeof row[2] === 'string' ? row[2].trim() : String(row[2] || '').trim(),
        teacher: typeof row[3] === 'string' ? row[3].trim() : String(row[3] || '').trim(),
        students: JSON.parse(row[4] || "[]"),
        updated_at: typeof row[5] === 'string' ? row[5].trim() : String(row[5] || '').trim()
      }));
    
    return createResponse('success', 'Attendance range retrieved successfully', attendance);
  } catch (error) {
    Logger.log('Error in getAttendanceRange: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to get attendance range: ' + error.message, null, error);
  }
}

// Get top absent students in date range
function getTopAbsentStudents(startDate, endDate, period, limit, spreadsheetMode) {
  if (spreadsheetMode === undefined) spreadsheetMode = DEFAULT_SPREADSHEET_MODE;
  
  try {
    if (!limit || limit <= 0) limit = 20;
    
    // Get clean attendance array
    const attendanceResponse = getAttendanceRange(startDate, endDate, period, spreadsheetMode);
    if (attendanceResponse.status !== 'success') {
      throw new Error(attendanceResponse.error?.details || 'Failed to get attendance data');
    }
    const attendanceData = attendanceResponse.data;
    if (attendanceData.length === 0) {
      return createResponse('success', 'No attendance data found', []);
    }
    
    const classroomResponse = getClassrooms([], spreadsheetMode);
    const allClassrooms = classroomResponse.data;
    const studentResponse = getStudents(allClassrooms.map(c => c.id), spreadsheetMode);
    const studentsByClassroom = studentResponse.data;
    
    const latestByRoomAndDate = {};
    attendanceData.forEach(entry => {
      // 📌 จุดแก้ไข: รวม Period เข้าไปใน Key เพื่อแยกการเช็คแต่ละรอบให้ชัดเจน
      const key = entry.room_id + '_' + entry.date + '_' + entry.period;
      if (!latestByRoomAndDate[key] || 
          new Date(entry.updated_at || entry.date) > new Date(latestByRoomAndDate[key].updated_at || entry.date)) {
        latestByRoomAndDate[key] = entry;
      }
    });
    
    const absentCounts = {};
    Object.values(latestByRoomAndDate).forEach(entry => {
      if (Array.isArray(entry.students)) {
        entry.students.forEach(student => {
          if (student.status === 'ขาด') {
            const studentId = String(student.student_id);
            if (!absentCounts[studentId]) {
              absentCounts[studentId] = {
                student_id: studentId,
                room_id: entry.room_id,
                count: 0
              };
            }
            absentCounts[studentId].count++;
          }
        });
      }
    });
    
    const absentList = [];
    for (const studentId in absentCounts) {
      const item = absentCounts[studentId];
      let studentInfo = null;
      let foundRoomId = null;
      
      for (const roomId in studentsByClassroom) {
        if (studentsByClassroom.hasOwnProperty(roomId)) {
          const found = studentsByClassroom[roomId].find(s => String(s.student_id) === studentId);
          if (found) {
            studentInfo = found;
            foundRoomId = roomId;
            break;
          }
        }
      }
      
      if (studentInfo) {
        const classroom = allClassrooms.find(c => c.id === foundRoomId);
        absentList.push({
          student_id: studentId,
          no: studentInfo.no || '',
          name: studentInfo.name || '',
          classroom_id: foundRoomId,
          classroom_name: classroom ? classroom.name : foundRoomId,
          absent_count: item.count
        });
      }
    }
    
    absentList.sort((a, b) => {
      if (b.absent_count !== a.absent_count) return b.absent_count - a.absent_count;
      if (a.classroom_name !== b.classroom_name) return a.classroom_name.localeCompare(b.classroom_name);
      return (a.no || '').localeCompare(b.no || '');
    });
    
    const topAbsent = absentList.slice(0, limit);
    return createResponse('success', 'Top absent students retrieved successfully', topAbsent);
  } catch (error) {
    Logger.log('Error in getTopAbsentStudents: ' + error.message + '\nStack: ' + error.stack);
    return createResponse('error', 'Failed to get top absent students: ' + error.message, null, error);
  }
}

// ========================================
// CLEANING SYSTEM - SERVER SIDE CODE
// ========================================


const CLEANING_HEADERS = {
  [CLEANING_SHEET_NAMES.CRITERIA]: ['id', 'name', 'max_score'],
  [CLEANING_SHEET_NAMES.ATTENDANCE]: ['date', 'classroom_id', 'checker_id', 'student_data', 'updated_at'],
  [CLEANING_SHEET_NAMES.SCORES]: ['date', 'classroom_id', 'checker_id', 'score_json', 'total_score', 'updated_at']
};

function getCleaningConfig() {
  try {
    // 1. ดึงข้อมูลห้องเรียนทั้งหมดโดยอาศัยฟังก์ชัน getClassrooms เดิมในระบบ
    let classrooms = [];
    try {
      const crRes = getClassrooms([]); // ส่ง array ว่างเพื่อดึงห้องทั้งหมด
      classrooms = crRes.data || [];
    } catch(e) {
      throw new Error("ไม่สามารถเรียกใช้งานฟังก์ชัน getClassrooms ได้");
    }

    // 2. ดึงข้อมูลนักเรียนทั้งหมดจัดกลุ่มตามห้องโดยอาศัยฟังก์ชัน getStudents เดิมในระบบ
    let students = {};
    try {
      const roomIds = classrooms.map(c => c.id);
      const stRes = getStudents(roomIds);
      students = stRes.data || {};
    } catch(e) {
      throw new Error("ไม่สามารถเรียกใช้งานฟังก์ชัน getStudents ได้");
    }

    // 3. กำหนดหัวข้อเกณฑ์การให้คะแนน (คุณครูสามารถแก้ไขชื่อเกณฑ์หรือคะแนนเต็มได้ที่นี่)
    const criteria = [
      { name: "ความสะอาดพื้นและบริเวณโดยรอบ", max_score: 10 },
      { name: "การจัดการขยะและถังขยะ", max_score: 10 },
      { name: "ความเรียบร้อยของการจัดเก็บอุปกรณ์", max_score: 10 }
    ];

    return createResponse('success', 'Config loaded successfully', {
      classrooms: classrooms,
      students: students,
      criteria: criteria
    });
  } catch (error) {
    return createResponse('error', error.message, null, error);
  }
}
function saveCleaningData(jsonData, spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return createResponse('error', 'Server is busy, please try again', null, { code: 'LOCK_TIMEOUT' });
    }
    try {
      const data = JSON.parse(jsonData);
      const { date, classroom_id, checker_id, student_data, scores } = data;

      if (!date || !classroom_id || !checker_id || !student_data || !scores) {
        return createResponse('error', 'Missing required fields', null, { code: 'INVALID_INPUT' });
      }
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
        return createResponse('error', 'Invalid date format', null, { code: 'INVALID_DATE' });
      }

      // Server-side permission check against the active teacher and assigned classrooms.
      const teachersSheet = getOrCreateSheet(SHEET_NAMES.TEACHERS, spreadsheetMode);
      const teachersData = teachersSheet.getDataRange().getValues();
      const teacherRow = teachersData.slice(1).find(row =>
        String(row[2] || '').trim() === String(checker_id).trim() && String(row[4] || '') === '1'
      );
      if (!teacherRow) {
        Logger.log(`❌ [Cleaning] Permission denied: teacher "${checker_id}" not found or inactive`);
        return createResponse('error', 'ไม่พบข้อมูลครูผู้บันทึก หรือบัญชีถูกปิดใช้งาน', null, { code: 'AUTH_FAILED' });
      }
      const allowedRooms = String(teacherRow[5] || '').split(',').map(r => r.trim()).filter(Boolean);
      const targetRoom = String(classroom_id).trim();
      if (!allowedRooms.includes(targetRoom)) {
        Logger.log(`❌ [Cleaning] Permission denied: "${checker_id}" tried to save room "${targetRoom}". Allowed: [${allowedRooms.join(', ') || 'NONE'}]`);
        return createResponse('error', 'คุณไม่มีสิทธิ์เช็คเขตพื้นที่ห้องนี้ (เฉพาะครูที่ปรึกษาเท่านั้น)', null, { code: 'PERMISSION_DENIED' });
      }
      Logger.log(`✅ [Cleaning] Permission granted: ${checker_id} → ${targetRoom}`);

      const updated_at = new Date().toISOString();

      const attendanceSheet = getOrCreateSheet(CLEANING_SHEET_NAMES.ATTENDANCE, spreadsheetMode);
      const attendanceData = attendanceSheet.getDataRange().getValues();
      
      const existingAttendanceIndex = attendanceData.slice(1).findIndex(row => {
        const rowDate = String(row[0]).trim();
        const rowClassroom = String(row[1]).trim();
        return rowDate === date && rowClassroom === classroom_id;
      }) + 2;

      const attendanceRow = [ date, classroom_id, checker_id, JSON.stringify(student_data), updated_at ];
      if (existingAttendanceIndex > 1) {
        attendanceSheet.getRange(existingAttendanceIndex, 1, 1, attendanceRow.length).setValues([attendanceRow]);
      } else {
        attendanceSheet.appendRow(attendanceRow);
      }

      let total = 0;
      let maxTotal = 0;
      for (const criterionName in scores) { total += scores[criterionName] || 0; }
      
      const criteriaSheet = getOrCreateSheet(CLEANING_SHEET_NAMES.CRITERIA, spreadsheetMode);
      const criteriaData = criteriaSheet.getDataRange().getValues();
      criteriaData.slice(1).forEach(row => { maxTotal += row[2] || 0; });

      const totalScore = `${total}/${maxTotal}`;
      const scoresSheet = getOrCreateSheet(CLEANING_SHEET_NAMES.SCORES, spreadsheetMode);
      const scoresData = scoresSheet.getDataRange().getValues();
      
      const existingScoreIndex = scoresData.slice(1).findIndex(row => {
        const rowDate = String(row[0]).trim();
        const rowClassroom = String(row[1]).trim();
        return rowDate === date && rowClassroom === classroom_id;
      }) + 2;

      const scoreRow = [ date, classroom_id, checker_id, JSON.stringify(scores), totalScore, updated_at ];
      if (existingScoreIndex > 1) {
        scoresSheet.getRange(existingScoreIndex, 1, 1, scoreRow.length).setValues([scoreRow]);
      } else {
        scoresSheet.appendRow(scoreRow);
      }

      return createResponse('success', existingAttendanceIndex > 1 ? 'Updated successfully' : 'Saved successfully', { total_score: totalScore });
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return createResponse('error', 'Failed to save: ' + error.message, null, error);
  }
}

function getCleaningStats(filter, spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const { startDate, endDate, classroom_id } = filter;
    
    // ดึงข้อมูลชีตต่างๆ
    const attendanceSheet = getOrCreateSheet(CLEANING_SHEET_NAMES.ATTENDANCE, spreadsheetMode);
    const attendanceData = attendanceSheet.getDataRange().getValues();
    const scoresSheet = getOrCreateSheet(CLEANING_SHEET_NAMES.SCORES, spreadsheetMode);
    const scoresData = scoresSheet.getDataRange().getValues();

    // 💡 [เพิ่มใหม่] ดึงข้อมูลจากชีต Teachers เพื่อแปลง Username เป็น ชื่อจริง
    const teachersSheet = getOrCreateSheet(SHEET_NAMES.TEACHERS, spreadsheetMode);
    const teachersData = teachersSheet.getDataRange().getValues();
    const teacherMap = {};
    // โครงสร้างชีต Teachers: ['id', 'name', 'username', ...] -> name = row[1], username = row[2]
    teachersData.slice(1).forEach(row => {
      const name = String(row[1]).trim();
      const username = String(row[2]).trim();
      if (username) {
        teacherMap[username] = name;
      }
    });

    // กรองข้อมูลซ้ำ
    const cleanAttendance = getCleanedCleaningRows(attendanceData);
    const cleanScores = getCleanedCleaningRows(scoresData);

    const getFormattedDate = (dateVal) => {
      if (!dateVal) return "";
      if (dateVal instanceof Date) {
        return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      return String(dateVal).trim();
    };

    // กรองข้อมูล Attendance
    const filteredAttendance = cleanAttendance.filter(row => {
      if (row[0] === 'Date') return false; 
      const rowDate = getFormattedDate(row[0]);
      const rowClassroom = String(row[1]).trim();
      const dateMatch = rowDate >= startDate && rowDate <= endDate;
      const classroomMatch = !classroom_id || rowClassroom === classroom_id;
      return dateMatch && classroomMatch;
    });

    // กรองข้อมูล Scores
    const filteredScores = cleanScores.filter(row => {
      if (row[0] === 'Date') return false;
      const rowDate = getFormattedDate(row[0]);
      const rowClassroom = String(row[1]).trim();
      const dateMatch = rowDate >= startDate && rowDate <= endDate;
      const classroomMatch = !classroom_id || rowClassroom === classroom_id;
      return dateMatch && classroomMatch;
    });

    // คำนวณสถิตินักเรียน
    const studentStats = {};
    filteredAttendance.forEach(row => {
      const roomId = String(row[1] || '').trim();
      const students = JSON.parse(row[3] || '[]');
      students.forEach(student => {
        const sid = student.student_id;
        if (!studentStats[sid]) {
          studentStats[sid] = {
            student_id: sid,
            name: student.name || sid,
            classroom_id: roomId,
            total: 0, done: 0, not_done: 0, leave: 0
          };
        }
        studentStats[sid].total++;
        if (student.status === 'ทำ') studentStats[sid].done++;
        else if (student.status === 'ไม่ทำ') studentStats[sid].not_done++;
        else if (student.status === 'ลา' || student.status === 'หนีเวร') studentStats[sid].leave++;
      });
    });

    // เรียงตามห้อง + เลขที่
    const allStudents = Object.values(studentStats).sort((a, b) => {
      if (a.classroom_id !== b.classroom_id) return a.classroom_id.localeCompare(b.classroom_id);
      return String(a.student_id).localeCompare(String(b.student_id));
    });

    const topStudents = Object.values(studentStats).sort((a, b) => b.done - a.done).slice(0, 5);

    // คำนวณคะแนนเฉลี่ย
    let totalScore = 0;
    let count = 0;
    filteredScores.forEach(row => {
      const scoreStr = String(row[4]); 
      if (scoreStr.includes('/')) {
        const [score, max] = scoreStr.split('/').map(Number);
        if (!isNaN(score) && !isNaN(max) && max > 0) {
          totalScore += (score / max) * 100;
          count++;
        }
      }
    });
    const avgScore = count > 0 ? (totalScore / count).toFixed(2) : 0;

    // จัดเรียงรายละเอียดคะแนนสำหรับตาราง
    const scoreDetails = filteredScores.map(row => {
      const checkerUsername = String(row[2]).trim();
      // 💡 [แก้ไข] นำ username ไปค้นหาชื่อจริงใน teacherMap ถ้าไม่เจอให้แสดง username ไปก่อน
      const checkerName = teacherMap[checkerUsername] || checkerUsername; 

      return {
        date: getFormattedDate(row[0]), 
        classroom_id: row[1],
        checker: checkerName, 
        scores: JSON.parse(row[3] || '{}'),
        total_score: row[4]
      };
    });

    return createResponse('success', 'Stats retrieved successfully', {
      studentStats: allStudents,
      allStudents: allStudents,
      topStudents,
      avgScore,
      scoreDetails
    });
  } catch (error) {
    return createResponse('error', 'Failed to get stats: ' + error.message, null, error);
  }
}

/**
 * รายละเอียดรายวันของนักเรียน 1 คน (สำหรับ cleaning)
 */
function getCleaningStudentDetail(studentId, startDate, endDate) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(CLEANING_SHEET_NAMES.ATTENDANCE);
    if (!sheet) return { status: "success", data: [] };

    const cleaned = getCleanedCleaningRows(sheet.getDataRange().getValues());
    const result = [];
    const targetStu = String(studentId || "").trim();

    cleaned.forEach(row => {
      const rowDate = formatDateStr(row[0]);
      const rowRoom = String(row[1] || "").trim();
      if (startDate && rowDate < startDate) return;
      if (endDate && rowDate > endDate) return;

      let students = [];
      try { students = JSON.parse(row[3] || '[]'); } catch (e) {}
      if (!Array.isArray(students)) return;

      const found = students.find(s => String(s.student_id) === targetStu);
      if (found) {
        result.push({
          date: rowDate,
          classroom_id: rowRoom,
          status: found.status,
          role: found.role || '',
          checker: String(row[2] || '').trim()
        });
      }
    });

    result.sort((a, b) => b.date.localeCompare(a.date));
    return { status: "success", data: result };
  } catch (err) {
    Logger.log('Error in getCleaningStudentDetail: ' + err.message);
    return { status: "error", message: err.toString(), data: [] };
  }
}

// ========================================
// 5. สถิติแดชบอร์ดหลัก (เพิ่มการกรองข้อมูลซ้ำแล้ว)
// ========================================

function getDashboardStatsSummary(startDate, endDate, selectedPeriods, selectedStatuses, filterLevel) {
  const SHEET_NAMES = { CLASSROOMS: 'Classrooms', STUDENTS: 'Students', ATTENDANCE: 'Attendance' };
  
  try {
    const ss = getSpreadsheet();
    const attSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
    const attData = attSheet ? attSheet.getDataRange().getValues() : [];
    const stdData = ss.getSheetByName(SHEET_NAMES.STUDENTS).getDataRange().getValues();
    const roomData = ss.getSheetByName(SHEET_NAMES.CLASSROOMS).getDataRange().getValues();
    
    const roomNameMap = {};
    roomData.slice(1).forEach(r => { if(r[0]) roomNameMap[r[0].toString().trim()] = r[1].toString().trim(); });
    
    const stdMap = {};
    stdData.slice(1).forEach(row => {
      if(row[1]) {
        const sId = row[1].toString().trim();
        stdMap[sId] = { classroom_id: row[0].toString().trim(), name: row[3] };
      }
    });

    const studentStats = {}; 
    const roomStats = {}; 
    const uniqueDates = new Set();
    
    const targetStatuses = Array.isArray(selectedStatuses) ? selectedStatuses : [selectedStatuses];
    const targetPeriods = Array.isArray(selectedPeriods) ? selectedPeriods.filter(p => p !== "") : [];

    // 📌 นำข้อมูลที่ได้ไปกรองตัวที่ซ้ำกันออกก่อน (ยึดบรรทัดล่างสุด)
    const cleanData = getCleanedAttendanceRows(attData);

    if (cleanData.length > 0) {
      cleanData.forEach(row => {
        let rDate = (row[0] instanceof Date) ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(row[0]).trim().substring(0, 10);
        const rPeriod = String(row[1] || "").trim();
        const rRoom = String(row[2] || "").trim();
        const rJson = row[4];

        let isLevelMatch = false;
        if (filterLevel === "all" || !filterLevel || filterLevel === "") isLevelMatch = true;
        else if (filterLevel === "m1-3") isLevelMatch = ["m1", "m2", "m3"].some(l => rRoom.toLowerCase().startsWith(l));
        else if (filterLevel === "m4-6") isLevelMatch = ["m4", "m5", "m6"].some(l => rRoom.toLowerCase().startsWith(l));
        else isLevelMatch = rRoom.toLowerCase().startsWith(filterLevel.toLowerCase());

        if (rDate >= startDate && rDate <= endDate && isLevelMatch) {
          uniqueDates.add(rDate);
          let students = [];
          try { students = JSON.parse(rJson); } catch (e) {}

          if (Array.isArray(students)) {
            if (!roomStats[rRoom]) {
              roomStats[rRoom] = { classroom: roomNameMap[rRoom] || rRoom, periods: {}, total: 0 };
            }

            if (targetPeriods.length === 0 || targetPeriods.includes(rPeriod)) {
              const periodKey = rPeriod || "ไม่ระบุเวลา";
              if (!roomStats[rRoom].periods[periodKey]) {
                roomStats[rRoom].periods[periodKey] = 0;
              }
              roomStats[rRoom].periods[periodKey]++; 
              roomStats[rRoom].total++;
            }

            students.forEach(s => {
              const sId = String(s.student_id).trim();
              if (!studentStats[sId]) {
                const info = stdMap[sId] || { name: 'ไม่พบชื่อ ('+sId+')', classroom_id: rRoom };
                studentStats[sId] = { 
                  student_id: sId, 
                  name: info.name, 
                  classroom_name: roomNameMap[info.classroom_id] || info.classroom_id, 
                  total_checks: 0, 
                  filtered_count: 0 
                };
              }
              
              if (targetPeriods.length === 0 || targetPeriods.includes(rPeriod)) {
                studentStats[sId].total_checks++;
                if (targetStatuses.includes(s.status)) {
                  studentStats[sId].filtered_count++;
                }
              }
            });
          }
        }
      });
    }

    const sortedPeriodSummary = Object.values(roomStats)
      .filter(r => r.total > 0)
      .sort((a, b) => a.classroom.localeCompare(b.classroom, 'th'));

    return {
      status: 'success',
      data: {
        startDate, endDate,
        schoolDays: uniqueDates.size,
        students: Object.values(studentStats).filter(s => s.filtered_count > 0),
        periodSummary: sortedPeriodSummary, 
        grandTotalFiltered: Object.values(studentStats).reduce((a, b) => a + b.filtered_count, 0),
        grandTotalChecks: Object.values(studentStats).reduce((a, b) => a + b.total_checks, 0),
        statuses: targetStatuses
      }
    };
  } catch (e) { 
    return { status: 'error', message: e.message }; 
  }
}

// ========================================
// 6. ดึงข้อมูลพิมพ์ใบงาน (อัปเดตให้กรองตามช่วงเวลา)
// ========================================
function getStudentReportData(studentId, startDate, endDate, statusFilters, classroomId, periodFilter) {
  try {
    const ss = getSpreadsheet();
    const attData = ss.getSheetByName('Attendance').getDataRange().getValues();
    const cleanData = getCleanedAttendanceRows(attData); // ข้อมูลที่กรองซ้ำแล้ว
    
    let specificDates = [];
    const uniqueDateStrings = new Set();
    const targetStatuses = Array.isArray(statusFilters) ? statusFilters : [statusFilters];

    cleanData.forEach(row => {
      let rDateStr = (row[0] instanceof Date) ? Utilities.formatDate(row[0], Session.getScriptTimeZone(), "yyyy-MM-dd") : String(row[0]).substring(0, 10);
      let rPeriod = String(row[1] || "").trim();
      let rRoom = String(row[2] || "").trim(); // รหัสห้องในชีต
      
      // เงื่อนไข: วันที่อยู่ในช่วง, ช่วงเวลาตรงกัน (หรือว่าง), สถานะตรงกัน
      let dateMatch = (rDateStr >= startDate && rDateStr <= endDate);
      let periodMatch = (!periodFilter || periodFilter === "" || rPeriod === periodFilter);
      
      if (dateMatch && periodMatch) {
         let students = JSON.parse(row[4] || "[]");
         let studentRecord = students.find(s => String(s.student_id) === String(studentId));
         
         if (studentRecord && targetStatuses.includes(studentRecord.status)) {
           if (!uniqueDateStrings.has(rDateStr)) {
             uniqueDateStrings.add(rDateStr);
             specificDates.push(rDateStr);
           }
         }
      }
    });

    return { status: 'success', data: { rawDates: specificDates } };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}
function getWeeklyAttendanceByRoom(weekDates, period, spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.ATTENDANCE, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const statusMap = {};

    // 💡 ฟังก์ชันช่วยแปลงวันที่จาก Google Sheets ให้เป็น YYYY-MM-DD เสมอ
    const getFormattedDate = (dateVal) => {
      if (!dateVal) return "";
      if (dateVal instanceof Date) {
        return Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy-MM-dd");
      }
      return String(dateVal).trim();
    };

    for (let i = 1; i < data.length; i++) {
      const rowDate = getFormattedDate(data[i][0]);
      const rowPeriod = String(data[i][1]).trim();
      const rowRoom = String(data[i][2]).trim();

      // ตรวจสอบว่าตรงกับวันที่ในสัปดาห์ และช่วงเวลาที่เลือกหรือไม่
      if (weekDates.includes(rowDate) && rowPeriod === period) {
        if (!statusMap[rowRoom]) statusMap[rowRoom] = {};
        statusMap[rowRoom][rowDate] = true; // บันทึกว่าวันนั้นเช็คแล้วเป็นเครื่องหมายถูก
      }
    }
    
    return createResponse('success', 'Weekly status retrieved', statusMap);
  } catch (error) {
    return createResponse('error', 'Failed to get weekly status: ' + error.message, null, error);
  }
}
/**
 * Helper: แปลงวันที่ให้เป็น yyyy-MM-dd เสมอ
 */
function formatDateStr(d) {
  if (!d) return "";
  if (d instanceof Date) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return String(d).trim().substring(0, 10);
}

/**
 * ดึงผลการเข้าเรียนที่เคยบันทึกไว้ของ (วันที่ + วิชา + ห้อง)
 * ใช้ pre-fill ฟอร์มตอนแก้ไขย้อนหลัง
 */
function getSubjectAttendanceForEdit(date, subjectCode, subjectName, gradeLevel, room) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Attendance_Subjects");
    if (!sheet) return { status: "success", data: [] };

    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { status: "success", data: [] };

    var targetDate  = formatDateStr(date);
    var targetCode  = String(subjectCode || "").trim();
    var targetName  = String(subjectName || "").trim();
    var targetGrade = String(gradeLevel  || "").trim();
    var targetRoom  = String(room        || "").trim();

    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowDate  = formatDateStr(row[1]);
      var rowCode  = String(row[2] || "").trim();
      var rowName  = String(row[3] || "").trim();
      var rowGrade = String(row[5] || "").trim();
      var rowRoom  = String(row[6] || "").trim();

      // ถ้ามีรหัสวิชา → match ด้วยรหัส, ถ้าไม่มี → match ด้วยชื่อ
      var codeMatch = targetCode !== ""
        ? (rowCode === targetCode)
        : (rowCode === "" && rowName === targetName);

      if (rowDate === targetDate && codeMatch &&
          rowGrade === targetGrade && rowRoom === targetRoom) {
        result.push({
          student_id: String(row[7]  || "").trim(),
          name:       String(row[8]  || "").trim(),
          status:     String(row[9]  || "").trim(),
          note:       String(row[10] || "").trim()
        });
      }
    }

    return { status: "success", data: result };
  } catch (err) {
    return { status: "error", message: err.toString(), data: [] };
  }
}

function getTeacherSubjectSchedule(teacherEmail, targetDate) {
  try {
    Logger.log('📅 getTeacherSubjectSchedule: email="' + teacherEmail + '", date="' + targetDate + '"');

    // ✅ ใช้ getSpreadsheet() ให้เหมือนระบบส่วนอื่น
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Timetable");
    if (!sheet) {
      Logger.log('⚠️ ไม่พบชีต Timetable');
      return { status: "success", data: [], message: "ยังไม่มีชีต Timetable" };
    }

    var data = sheet.getDataRange().getValues();
    Logger.log('📊 Timetable rows (รวม header): ' + data.length);

    if (data.length <= 1) {
      return { status: "success", data: [], message: "ชีต Timetable ว่าง" };
    }

    // ✅ แปลงวันที่แบบ local (กัน timezone เพี้ยน)
    var dateObj;
    if (targetDate) {
      var p = String(targetDate).split('-');
      dateObj = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
    } else {
      dateObj = new Date();
    }

    var days = ["อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"];
    var dayName = days[dateObj.getDay()];
    var normEmail = String(teacherEmail || '').trim().toLowerCase();

    Logger.log('🗓️ Day = ' + dayName + ', looking for email = "' + normEmail + '"');

    var result = [];
    var emailsInSheet = {};

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var rowDay = String(row[1] || '').trim();
      var rowEmail = String(row[5] || '').trim();
      var rowEmailNorm = rowEmail.toLowerCase();

      if (rowEmailNorm) emailsInSheet[rowEmailNorm] = true;

      if (rowDay === dayName && rowEmailNorm === normEmail) {
        result.push({
          timetableId: String(row[0] || ''),
          period:      row[2],
          subjectCode: String(row[3] || ''),
          subjectName: String(row[4] || ''),
          gradeLevel:  String(row[6] || ''),
          room:        String(row[7] || ''),
          classroomId: String(row[8] || '')
        });
      }
    }

    Logger.log('✅ พบ ' + result.length + ' คาบ | Emails ที่มีใน Timetable: ' + Object.keys(emailsInSheet).join(', '));

    return { status: "success", data: result };

  } catch (e) {
    Logger.log('❌ Error in getTeacherSubjectSchedule: ' + e.message + '\n' + e.stack);
    return { status: "error", message: e.message, data: [] };
  }
}

// 2. ฟังก์ชันบันทึกข้อมูลการเช็คชื่อรายวิชาลงชีต Attendance_Subjects
function saveSubjectAttendance(jsonPayload) {
  try {
    var payload = JSON.parse(jsonPayload);
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName("Attendance_Subjects");
    
    // ถ้ายังไม่มีชีต ให้สร้างให้อัตโนมัติพร้อมหัวตาราง
    if (!sheet) {
      sheet = ss.insertSheet("Attendance_Subjects");
      sheet.appendRow(["Timestamp", "Date", "SubjectCode", "SubjectName", "TeacherEmail", "GradeLevel", "Room", "StudentID", "StudentName", "Status", "Note"]);
    }
    
    var timestamp = new Date();
    var rowsToAdd = [];
    
    payload.students.forEach(function(stu) {
      rowsToAdd.push([
        timestamp,
        payload.date,
        payload.subjectCode,
        payload.subjectName,
        payload.teacherEmail,
        payload.gradeLevel,
        payload.room,
        stu.student_id,
        stu.name,
        stu.status, // มา, ขาด, ลา, สาย
        stu.note || ""
      ]);
    });
    
    if (rowsToAdd.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rowsToAdd.length, rowsToAdd[0].length).setValues(rowsToAdd);
      return { status: "success", message: "บันทึกผลการเข้าเรียนสำเร็จ" };
    }
    
    return { status: "error", message: "ไม่พบข้อมูลรายชื่อนักเรียนสำหรับบันทึก" };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

// ========================================
// SUBJECT SUMMARY V2 — Date Range + Charts + Student Detail
// ========================================

/** สรุปสถิติรายวิชา รองรับช่วงวันที่และตัวกรองห้อง */
function getSubjectAttendanceSummaryV2(teacherIdentifier, subjectKey, startDate, endDate, classroomFilter) {
  try {
    const emptyTotals = { present: 0, absent: 0, leave: 0, late: 0, sick: 0, total: 0 };
    const sheet = getSpreadsheet().getSheetByName("Attendance_Subjects");
    if (!sheet) {
      return { status: "success", data: { students: [], totals: emptyTotals, dailyTrend: {} } };
    }

    const data = sheet.getDataRange().getValues();
    const summary = {};
    const dailyTrend = {};
    const normTeacher = String(teacherIdentifier || "").trim().toLowerCase();
    const subjKey = String(subjectKey || "").trim();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDate = formatDateStr(row[1]);
      const rowCode = String(row[2] || "").trim();
      const rowName = String(row[3] || "").trim();
      const rowTeacher = String(row[4] || "").trim().toLowerCase();
      const rowGrade = String(row[5] || "").trim();
      const rowRoom = String(row[6] || "").trim();
      const studentId = String(row[7] || "").trim();
      const studentName = String(row[8] || "").trim();
      const status = String(row[9] || "").trim();

      if (!rowDate || !studentId || rowTeacher !== normTeacher) continue;
      if (startDate && rowDate < startDate) continue;
      if (endDate && rowDate > endDate) continue;

      // The subject list uses its code as key, or the name when no code exists.
      const rowSubjectKey = rowCode || rowName;
      const matchesSubject = rowSubjectKey === subjKey;
      if (!matchesSubject) continue;

      if (classroomFilter) {
        const gradeRoom = rowGrade + "/" + rowRoom;
        if (gradeRoom !== classroomFilter && rowRoom !== classroomFilter && rowGrade !== classroomFilter) continue;
      }

      if (!summary[studentId]) {
        summary[studentId] = {
          id: studentId, name: studentName, gradeLevel: rowGrade, room: rowRoom,
          present: 0, absent: 0, leave: 0, late: 0, sick: 0, total: 0
        };
      }
      const student = summary[studentId];
      if (status === "มา") student.present++;
      else if (status === "ขาด") student.absent++;
      else if (status === "ลา") student.leave++;
      else if (status === "สาย") student.late++;
      else if (status === "ป่วย") student.sick++;
      student.total++;

      if (!dailyTrend[rowDate]) {
        dailyTrend[rowDate] = { present: 0, absent: 0, leave: 0, late: 0, sick: 0 };
      }
      if (status === "มา") dailyTrend[rowDate].present++;
      else if (status === "ขาด") dailyTrend[rowDate].absent++;
      else if (status === "ลา") dailyTrend[rowDate].leave++;
      else if (status === "สาย") dailyTrend[rowDate].late++;
      else if (status === "ป่วย") dailyTrend[rowDate].sick++;
    }

    const students = Object.values(summary);
    students.sort((a, b) => {
      const roomA = a.gradeLevel + "/" + a.room;
      const roomB = b.gradeLevel + "/" + b.room;
      return roomA.localeCompare(roomB, "th") || a.id.localeCompare(b.id);
    });

    const totals = { present: 0, absent: 0, leave: 0, late: 0, sick: 0, total: 0 };
    students.forEach(student => {
      totals.present += student.present;
      totals.absent += student.absent;
      totals.leave += student.leave;
      totals.late += student.late;
      totals.sick += student.sick;
      totals.total += student.total;
    });

    return { status: "success", data: { students, totals, dailyTrend } };
  } catch (err) {
    Logger.log("Error in getSubjectAttendanceSummaryV2: " + err.message);
    return { status: "error", message: err.toString() };
  }
}

/** รายละเอียดรายวันของนักเรียนหนึ่งคน สำหรับ modal */
function getSubjectStudentDetail(teacherIdentifier, subjectKey, studentId, startDate, endDate) {
  try {
    const sheet = getSpreadsheet().getSheetByName("Attendance_Subjects");
    if (!sheet) return { status: "success", data: [] };

    const data = sheet.getDataRange().getValues();
    const result = [];
    const normTeacher = String(teacherIdentifier || "").trim().toLowerCase();
    const subjKey = String(subjectKey || "").trim();
    const targetStudent = String(studentId || "").trim();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDate = formatDateStr(row[1]);
      const rowCode = String(row[2] || "").trim();
      const rowName = String(row[3] || "").trim();
      const rowTeacher = String(row[4] || "").trim().toLowerCase();
      const rowStudent = String(row[7] || "").trim();

      if (!rowDate || rowStudent !== targetStudent || rowTeacher !== normTeacher) continue;
      if (startDate && rowDate < startDate) continue;
      if (endDate && rowDate > endDate) continue;

      const rowSubjectKey = rowCode || rowName;
      const matchesSubject = rowSubjectKey === subjKey;
      if (!matchesSubject) continue;

      result.push({
        date: rowDate,
        status: String(row[9] || "").trim(),
        note: String(row[10] || "").trim()
      });
    }

    result.sort((a, b) => b.date.localeCompare(a.date));
    return { status: "success", data: result };
  } catch (err) {
    Logger.log("Error in getSubjectStudentDetail: " + err.message);
    return { status: "error", message: err.toString(), data: [] };
  }
}

// ฟังก์ชันดึงรายชื่อวิชาทั้งหมดที่ครูสอน (ปรับรูปแบบข้อความตามเงื่อนไขใหม่)
function getTeacherSubjectsList(teacherIdentifier) {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName("Timetable");
  if (!sheet) return { status: "success", data: [] };
  
  var data = sheet.getDataRange().getValues();
  var subjectsMap = {};
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowTeacher = String(row[5]).trim();
    
    if (rowTeacher.toLowerCase() === String(teacherIdentifier).toLowerCase()) {
      var subjCode = String(row[3]).trim();
      var subjName = String(row[4]).trim();
      var grade = String(row[6]).trim();
      var room = String(row[7]).trim();
      
      // ✅ ข้อ 3: วิชาไหนไม่ระบุ ชั้น ไม่ต้องแสดง /1 (แสดงเฉพาะชั้น หรือเฉพาะห้องถ้ามีแค่อย่างเดียว)
      var gradeRoom = "";
      if (grade !== "" && room !== "") {
        gradeRoom = grade + "/" + room;
      } else if (grade !== "") {
        gradeRoom = grade;
      } else if (room !== "") {
        gradeRoom = room;
      }
      
      // ถ้าไม่มีรหัสวิชา ให้ใช้ชื่อวิชาเป็น Key แทนป้องกันบัค
      var mapKey = subjCode !== "" ? subjCode : subjName;
      
      if(!subjectsMap[mapKey]) {
        subjectsMap[mapKey] = { code: subjCode, name: subjName, classes: gradeRoom !== "" ? [gradeRoom] : [] };
      } else if (gradeRoom !== "" && subjectsMap[mapKey].classes.indexOf(gradeRoom) === -1) {
        subjectsMap[mapKey].classes.push(gradeRoom);
      }
    }
  }
  
  var result = Object.keys(subjectsMap).map(function(key) {
    var subj = subjectsMap[key];
    
    var displayName = subj.name;
    // ✅ ข้อ 4: วิชาไหนไม่มีรหัส ไม่ต้องมีวงเล็บ ()
    if (subj.code !== "") {
      displayName += " (" + subj.code + ")";
    }
    
    // ✅ ข้อ 2: หน้าสรุปรายวิชา ไม่ต้องมีคำว่า "สอนห้อง"
    if (subj.classes.length > 0) {
      displayName += " - " + subj.classes.join(', ');
    }
    
    return { key: key, displayName: displayName };
  });
  
  return { status: "success", data: result };
}
/**
 * ดึงรายชื่อนักเรียนของห้องเรียนหนึ่งๆ (ใช้โดย subject_attendance.html)
 * @param {string} classroomId - รหัสห้องเรียน เช่น "m1-1", "m3_1"
 * @returns {{students: Array}}
 */
function getClassroomStudents(classroomId) {
  try {
    if (!classroomId) return { students: [] };

    const key = String(classroomId).trim();
    const res = getStudents([key]);
    if (!res || res.status !== 'success') return { students: [] };

    const students = (res.data && res.data[key]) ? res.data[key] : [];
    return { students: students };
  } catch (e) {
    Logger.log('Error in getClassroomStudents: ' + e.message);
    return { students: [], error: e.message };
  }
}

// CLEANING — Tab-based UI Extensions
// ========================================

/**
 * ดึงสถานะรายสัปดาห์ (จ.-ศ.) ของ cleaning
 * ใช้สร้างแถบ ✓ ⚡ ในตาราง Cleaning Tab
 */
function getCleaningWeeklyStatus(weekDates) {
  try {
    const sheet = getOrCreateSheet(CLEANING_SHEET_NAMES.ATTENDANCE);
    const data = sheet.getDataRange().getValues();
    const statusMap = {};

    const fmt = (d) => {
      if (!d) return "";
      if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
      return String(d).trim();
    };

    // clean duplicates (เก็บแถวล่างสุด = ล่าสุด)
    const cleaned = getCleanedCleaningRows(data);
    cleaned.forEach(row => {
      const rowDate = fmt(row[0]);
      const rowRoom = String(row[1]).trim();
      if (weekDates.includes(rowDate) && rowRoom) {
        if (!statusMap[rowRoom]) statusMap[rowRoom] = {};
        statusMap[rowRoom][rowDate] = true;
      }
    });

    return createResponse('success', 'Cleaning weekly status retrieved', statusMap);
  } catch (error) {
    Logger.log('Error in getCleaningWeeklyStatus: ' + error.message);
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

/**
 * ดึงข้อมูล cleaning ของ (date + classroom) เพื่อ pre-fill modal
 */
function getCleaningDataForEdit(date, classroomId) {
  try {
    const ss = getSpreadsheet();
    const fmt = (d) => {
      if (!d) return "";
      if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
      return String(d).trim();
    };

    let studentData = [];
    let scores = {};
    let hasExisting = false;

    const attSheet = ss.getSheetByName(CLEANING_SHEET_NAMES.ATTENDANCE);
    if (attSheet) {
      const cleaned = getCleanedCleaningRows(attSheet.getDataRange().getValues());
      const row = cleaned.find(r => fmt(r[0]) === date && String(r[1]).trim() === classroomId);
      if (row) {
        try { studentData = JSON.parse(row[3] || '[]'); } catch(e) {}
        hasExisting = true;
      }
    }

    const scoreSheet = ss.getSheetByName(CLEANING_SHEET_NAMES.SCORES);
    if (scoreSheet) {
      const cleaned = getCleanedCleaningRows(scoreSheet.getDataRange().getValues());
      const row = cleaned.find(r => fmt(r[0]) === date && String(r[1]).trim() === classroomId);
      if (row) {
        try { scores = JSON.parse(row[3] || '{}'); } catch(e) {}
        hasExisting = true;
      }
    }

    return createResponse('success', 'Cleaning data', { studentData, scores, hasExisting });
  } catch (error) {
    Logger.log('Error in getCleaningDataForEdit: ' + error.message);
    return createResponse('error', error.message, { studentData: [], scores: {}, hasExisting: false }, error);
  }
}

// ========================================
// PHASE 3: Attendance Stats V2
// ========================================

function getAttendanceSummaryV2(startDate, endDate, period, classroomFilter) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.ATTENDANCE);
    const cleanData = getCleanedAttendanceRows(sheet.getDataRange().getValues());
    const summary = {}, dailyTrend = {};
    const targetPeriod = String(period || '').trim();
    const targetClassroom = String(classroomFilter || '').trim();

    cleanData.forEach(row => {
      const rowDate = formatDateStr(row[0]);
      const rowPeriod = String(row[1] || '').trim();
      const rowRoom = String(row[2] || '').trim();
      if (startDate && rowDate < startDate) return;
      if (endDate && rowDate > endDate) return;
      if (targetPeriod && rowPeriod !== targetPeriod) return;
      if (targetClassroom && rowRoom !== targetClassroom) return;

      let students = [];
      try { students = JSON.parse(row[4] || '[]'); } catch (e) {}

      students.forEach(s => {
        const sid = String(s.student_id || '').trim();
        if (!sid) return;
        const status = String(s.status || '').trim();
        if (!summary[sid]) {
          summary[sid] = {
            id: sid, name: s.name || sid, classroom_id: rowRoom,
            present: 0, absent: 0, late: 0, leave: 0, sick: 0, total: 0
          };
        }
        const rec = summary[sid];
        if (status === 'มา') rec.present++;
        else if (status === 'ขาด') rec.absent++;
        else if (status === 'สาย') rec.late++;
        else if (status === 'ลา') rec.leave++;
        else if (status === 'ป่วย') rec.sick++;
        rec.total++;

        if (!dailyTrend[rowDate]) dailyTrend[rowDate] = { present: 0, absent: 0, late: 0, leave: 0, sick: 0 };
        if (status === 'มา') dailyTrend[rowDate].present++;
        else if (status === 'ขาด') dailyTrend[rowDate].absent++;
        else if (status === 'สาย') dailyTrend[rowDate].late++;
        else if (status === 'ลา') dailyTrend[rowDate].leave++;
        else if (status === 'ป่วย') dailyTrend[rowDate].sick++;
      });
    });

    const students = Object.values(summary).sort((a, b) => {
      if (a.classroom_id !== b.classroom_id) return a.classroom_id.localeCompare(b.classroom_id);
      return a.id.localeCompare(b.id);
    });

    const totals = { present: 0, absent: 0, late: 0, leave: 0, sick: 0, total: 0 };
    students.forEach(s => {
      totals.present += s.present;
      totals.absent += s.absent;
      totals.late += s.late;
      totals.leave += s.leave;
      totals.sick += s.sick;
      totals.total += s.total;
    });

    return { status: 'success', data: { students, totals, dailyTrend } };
  } catch (err) {
    Logger.log('Error in getAttendanceSummaryV2: ' + err.message);
    return { status: 'error', message: err.toString() };
  }
}

function getAttendanceStudentDetail(studentId, startDate, endDate, period) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.ATTENDANCE);
    const cleanData = getCleanedAttendanceRows(sheet.getDataRange().getValues());
    const result = [];
    const targetStu = String(studentId || '').trim();
    const targetPeriod = String(period || '').trim();

    cleanData.forEach(row => {
      const rowDate = formatDateStr(row[0]);
      const rowPeriod = String(row[1] || '').trim();
      if (startDate && rowDate < startDate) return;
      if (endDate && rowDate > endDate) return;
      if (targetPeriod && rowPeriod !== targetPeriod) return;

      let students = [];
      try { students = JSON.parse(row[4] || '[]'); } catch (e) {}
      const found = students.find(s => String(s.student_id || '').trim() === targetStu);
      if (found) {
        result.push({
          date: rowDate, period: rowPeriod,
          status: found.status || '', reason: found.reason || ''
        });
      }
    });

    result.sort((a, b) => b.date.localeCompare(a.date));
    return { status: 'success', data: result };
  } catch (err) {
    Logger.log('Error in getAttendanceStudentDetail: ' + err.message);
    return { status: 'error', message: err.toString(), data: [] };
  }
}

// ========================================
// PHASE 3: Weekly Check Report (3 modes)
// ========================================

function getWeeklyCheckReport(startDate, endDate, mode, options) {
  try {
    const opts = options || {};

    // Build weekDates (Mon-Fri)
    const weekDates = [];
    const [sy, sm, sd] = String(startDate).split('-').map(Number);
    for (let i = 0; i < 7; i++) {
      const d = new Date(sy, sm - 1, sd + i);
      const ds = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (ds >= startDate && ds <= endDate) weekDates.push(ds);
    }

    const rows = [];

    // ===== ATTENDANCE =====
    if (mode === 'attendance') {
      const targetPeriod = String(opts.period || '').trim();
      const sheet = getOrCreateSheet(SHEET_NAMES.ATTENDANCE);
      const cleanData = getCleanedAttendanceRows(sheet.getDataRange().getValues());
      const checkMap = {};
      cleanData.forEach(row => {
        const rowDate = formatDateStr(row[0]);
        const rowPeriod = String(row[1] || '').trim();
        const rowRoom = String(row[2] || '').trim();
        if (!weekDates.includes(rowDate)) return;
        if (targetPeriod && rowPeriod !== targetPeriod) return;
        if (!checkMap[rowRoom]) checkMap[rowRoom] = {};
        checkMap[rowRoom][rowDate] = true;
      });

      const crRes = getClassrooms([]);
      const classrooms = (crRes.data || []).sort((a, b) => a.name.localeCompare(b.name, 'th'));

      // Student counts
      const studentCounts = {};
      const allStu = getAllStudents();
      if (allStu.status === 'success') {
        Object.keys(allStu.data).forEach(rid => { studentCounts[rid] = allStu.data[rid].length; });
      }

      classrooms.forEach(c => {
        const dates = {};
        weekDates.forEach(d => { dates[d] = !!(checkMap[c.id] && checkMap[c.id][d]); });
        rows.push({
          id: c.id, name: c.name,
          sublabel: `${studentCounts[c.id] || 0} คน`,
          dates, clickable: true, mode: 'attendance'
        });
      });

    // ===== CLEANING =====
    } else if (mode === 'cleaning') {
      const sheet = getOrCreateSheet(CLEANING_SHEET_NAMES.ATTENDANCE);
      const cleanData = getCleanedCleaningRows(sheet.getDataRange().getValues());
      const checkMap = {};
      cleanData.forEach(row => {
        const rowDate = formatDateStr(row[0]);
        const rowRoom = String(row[1] || '').trim();
        if (!weekDates.includes(rowDate)) return;
        if (!checkMap[rowRoom]) checkMap[rowRoom] = {};
        checkMap[rowRoom][rowDate] = true;
      });

      const crRes = getClassrooms([]);
      const classrooms = (crRes.data || []).sort((a, b) => a.name.localeCompare(b.name, 'th'));
      const studentCounts = {};
      const allStu = getAllStudents();
      if (allStu.status === 'success') {
        Object.keys(allStu.data).forEach(rid => { studentCounts[rid] = allStu.data[rid].length; });
      }

      classrooms.forEach(c => {
        const dates = {};
        weekDates.forEach(d => { dates[d] = !!(checkMap[c.id] && checkMap[c.id][d]); });
        rows.push({
          id: c.id, name: c.name,
          sublabel: `${studentCounts[c.id] || 0} คน`,
          dates, clickable: true, mode: 'cleaning'
        });
      });

    // ===== SUBJECT =====
    } else if (mode === 'subject') {
      const teacherEmail = String(opts.teacherEmail || '').trim().toLowerCase();
      const ss = getSpreadsheet();
      const ttSheet = ss.getSheetByName('Timetable');
      const attSheet = ss.getSheetByName('Attendance_Subjects');
      if (!ttSheet) return { status: 'success', data: { weekDates, rows: [] } };

      const ttData = ttSheet.getDataRange().getValues();
      const days = ["อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"];

      const scheduleByDate = {};
      weekDates.forEach(d => {
        const dayOfWeek = new Date(d + 'T00:00:00Z').getUTCDay();
        const dayName = days[dayOfWeek];
        scheduleByDate[d] = [];
        for (let i = 1; i < ttData.length; i++) {
          const r = ttData[i];
          const rowDay = String(r[1] || '').trim();
          const rowTeacher = String(r[5] || '').trim().toLowerCase();
          if (rowDay === dayName && rowTeacher === teacherEmail) {
            scheduleByDate[d].push({
              code: String(r[3] || '').trim(),
              name: String(r[4] || '').trim(),
              grade: String(r[6] || '').trim(),
              room: String(r[7] || '').trim(),
              classroomId: String(r[8] || '').trim()
            });
          }
        }
      });

      const checkMap = {};
      if (attSheet) {
        const attData = attSheet.getDataRange().getValues();
        for (let i = 1; i < attData.length; i++) {
          const r = attData[i];
          const rowDate = formatDateStr(r[1]);
          const rowCode = String(r[2] || '').trim();
          const rowName = String(r[3] || '').trim();
          const rowTeacher = String(r[4] || '').trim().toLowerCase();
          const rowGrade = String(r[5] || '').trim();
          const rowRoom = String(r[6] || '').trim();
          if (!weekDates.includes(rowDate) || rowTeacher !== teacherEmail) continue;
          const key = `${rowDate}_${rowCode || rowName}_${rowGrade}_${rowRoom}`;
          checkMap[key] = true;
        }
      }

      const rowMap = {};
      weekDates.forEach(d => {
        scheduleByDate[d].forEach(sch => {
          const gradeRoom = sch.grade + (sch.room ? '/' + sch.room : '');
          const key = `${sch.code || sch.name}_${gradeRoom}`;
          if (!rowMap[key]) {
            rowMap[key] = {
              id: key,
              name: sch.name + (sch.code ? ` (${sch.code})` : ''),
              sublabel: gradeRoom,
              dates: {},
              meta: { code: sch.code, name: sch.name, grade: sch.grade, room: sch.room, classroomId: sch.classroomId },
              clickable: true,
              mode: 'subject'
            };
          }
          const checkKey = `${d}_${sch.code || sch.name}_${sch.grade}_${sch.room}`;
          rowMap[key].dates[d] = !!checkMap[checkKey];
        });
      });

      Object.values(rowMap).sort((a, b) => a.name.localeCompare(b.name, 'th')).forEach(r => rows.push(r));
    }

    return { status: 'success', data: { weekDates, rows } };
  } catch (err) {
    Logger.log('Error in getWeeklyCheckReport: ' + err.message);
    return { status: 'error', message: err.toString() };
  }
}

function diagnoseSubjectAttendance() {
  var ss = getSpreadsheet();
  var out = [];

  // 1. Timetable
  var tSheet = ss.getSheetByName("Timetable");
  if (!tSheet) {
    out.push("❌ ไม่พบชีต Timetable");
  } else {
    var tData = tSheet.getDataRange().getValues();
    out.push("✅ Timetable: " + tData.length + " rows");
    if (tData.length > 0) out.push("   Headers: " + JSON.stringify(tData[0]));
    if (tData.length > 1) out.push("   Sample : " + JSON.stringify(tData[1]));

    var emails = {};
    var days = {};
    for (var i = 1; i < tData.length; i++) {
      var e = String(tData[i][5] || '').trim();
      var d = String(tData[i][1] || '').trim();
      if (e) emails[e] = (emails[e] || 0) + 1;
      if (d) days[d] = (days[d] || 0) + 1;
    }
    out.push("   TeacherEmails ใน Timetable: " + JSON.stringify(emails));
    out.push("   Days ใน Timetable: " + JSON.stringify(days));
  }

  // 2. Teachers
  var teachSheet = ss.getSheetByName("Teachers");
  if (teachSheet) {
    var tData2 = teachSheet.getDataRange().getValues();
    var users = tData2.slice(1).map(function(r){ return String(r[2]||'').trim(); });
    out.push("✅ Teachers (usernames): " + JSON.stringify(users));
  }

  // 3. ทดสอบเรียกจริง
  var testEmail = "somchai"; // ⬅️ เปลี่ยนเป็น username ของครูที่จะทดสอบ
  var testDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var res = getTeacherSubjectSchedule(testEmail, testDate);
  out.push("🧪 Test: email=" + testEmail + ", date=" + testDate);
  out.push("   → " + JSON.stringify(res));

  Logger.log(out.join('\n'));
  return out.join('\n');
}


// TODO(Security): Require verified approved-account identity on these server APIs before deployment.
// The current USER_DEPLOYING / ANYONE_ANONYMOUS setup may not expose the active user's email;
// client-side filters are not an authorization boundary.
// ========================================
// PHASE 3.5: Student Risk Assessment
// ========================================
const RISK_WEIGHT_KEYS = {
  ATTENDANCE: 'riskWeight_attendance',
  CLEANING: 'riskWeight_cleaning',
  SUBJECT: 'riskWeight_subject'
};
const RISK_DEFAULT_WEIGHTS = { attendance: 40, cleaning: 30, subject: 30 };

function riskNormalizeWeights(value) {
  let input = value || {};
  if (typeof input === 'string') input = JSON.parse(input);
  const weights = {};
  ['attendance', 'cleaning', 'subject'].forEach(function(key) {
    const n = Number(input[key]);
    weights[key] = Number.isFinite(n) && n >= 0 && n <= 100 ? n : RISK_DEFAULT_WEIGHTS[key];
  });
  if (weights.attendance + weights.cleaning + weights.subject <= 0) {
    return Object.assign({}, RISK_DEFAULT_WEIGHTS);
  }
  return weights;
}

function riskValidateWeightInput(value) {
  let input = value;
  if (typeof input === 'string') input = JSON.parse(input);
  if (!input || typeof input !== 'object') throw new Error('ข้อมูลน้ำหนักไม่ถูกต้อง');
  const weights = {};
  ['attendance', 'cleaning', 'subject'].forEach(function(key) {
    const n = Number(input[key]);
    if (!Number.isFinite(n) || n < 0 || n > 100) throw new Error('น้ำหนักแต่ละกิจกรรมต้องอยู่ระหว่าง 0 ถึง 100');
    weights[key] = n;
  });
  if (weights.attendance + weights.cleaning + weights.subject <= 0) throw new Error('ผลรวมต้องมากกว่า 0');
  return weights;
}

function getRiskWeights() {
  try {
    const weights = Object.assign({}, RISK_DEFAULT_WEIGHTS);
    const sheet = getOrCreateSheet(SHEET_NAMES.CONFIG);
    const data = sheet.getDataRange().getValues();
    data.slice(1).forEach(function(row) {
      const key = String(row[0] || '').trim();
      const value = Number(row[1]);
      if (!Number.isFinite(value) || value < 0 || value > 100) return;
      if (key === RISK_WEIGHT_KEYS.ATTENDANCE) weights.attendance = value;
      if (key === RISK_WEIGHT_KEYS.CLEANING) weights.cleaning = value;
      if (key === RISK_WEIGHT_KEYS.SUBJECT) weights.subject = value;
    });
    return { status: 'success', data: riskNormalizeWeights(weights) };
  } catch (err) {
    return { status: 'error', message: err.message || String(err), data: Object.assign({}, RISK_DEFAULT_WEIGHTS) };
  }
}

function saveRiskWeights(weightsJson) {
  let lock;
  let lockAcquired = false;
  try {
    const weights = riskValidateWeightInput(weightsJson);
    const lockService = LockService.getScriptLock();
    lock = lockService;
    if (!lock.tryLock(10000)) return { status: 'error', message: 'Server busy' };
    lockAcquired = true;
    const sheet = getOrCreateSheet(SHEET_NAMES.CONFIG);
    const data = sheet.getDataRange().getValues();
    const settings = [
      [RISK_WEIGHT_KEYS.ATTENDANCE, weights.attendance],
      [RISK_WEIGHT_KEYS.CLEANING, weights.cleaning],
      [RISK_WEIGHT_KEYS.SUBJECT, weights.subject]
    ];
    settings.forEach(function(setting) {
      let found = false;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0] || '').trim() === setting[0]) {
          sheet.getRange(i + 1, 2).setValue(setting[1]);
          found = true;
          break;
        }
      }
      if (!found) sheet.appendRow(setting);
    });
    return { status: 'success', message: 'บันทึกน้ำหนักสำเร็จ', data: weights };
  } catch (err) {
    return { status: 'error', message: err.message || String(err) };
  } finally {
    if (lock && lockAcquired) lock.releaseLock();
  }
}

function riskGetDateList(startDate, endDate) {
  const valid = function(value) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
    const p = value.split('-').map(Number);
    const d = new Date(Date.UTC(p[0], p[1] - 1, p[2]));
    return d.getUTCFullYear() === p[0] && d.getUTCMonth() === p[1] - 1 && d.getUTCDate() === p[2];
  };
  if (!valid(startDate) || !valid(endDate) || startDate > endDate) {
    throw new Error('ช่วงวันที่ไม่ถูกต้อง');
  }
  const start = startDate.split('-').map(Number);
  const end = endDate.split('-').map(Number);
  const cursor = new Date(Date.UTC(start[0], start[1] - 1, start[2]));
  const finish = new Date(Date.UTC(end[0], end[1] - 1, end[2]));
  if ((finish.getTime() - cursor.getTime()) / 86400000 > 366) {
    throw new Error('เลือกช่วงเวลาได้ไม่เกิน 367 วัน');
  }
  const dates = [];
  while (cursor <= finish) {
    dates.push([cursor.getUTCFullYear(), String(cursor.getUTCMonth() + 1).padStart(2, '0'), String(cursor.getUTCDate()).padStart(2, '0')].join('-'));
    cursor.setUTCDate(cursor.getUTCDate() + 1);
  }
  return dates;
}

function riskScoreStatus(kind, status) {
  const value = String(status || '').trim();
  if (kind === 'cleaning') {
    if (value === 'ทำ') return { points: 1, count: 'done' };
    if (value === 'ลา') return { points: 0.5, count: 'excused' };
    if (value === 'ไม่ทำ' || value === 'หนีเวร') return { points: 0, count: 'notDone' };
    return null;
  }
  if (value === 'มา') return { points: 1, count: 'present' };
  if (value === 'สาย') return { points: 0.75, count: 'late' };
  if (value === 'ลา' || value === 'ป่วย') return { points: 0.5, count: value === 'ลา' ? 'leave' : 'sick' };
  if (value === 'ขาด') return { points: 0, count: 'absent' };
  return null;
}

function riskStudentKey(classroomId, studentId) {
  return String(classroomId) + '::' + String(studentId);
}

function riskAddEntry(map, key, date, kind, status, detail) {
  const scored = riskScoreStatus(kind, status);
  if (!scored) return;
  if (!map[key]) map[key] = {};
  if (!map[key][date]) {
    map[key][date] = { total: 0, points: 0, counts: {}, events: [] };
  }
  const bucket = map[key][date];
  bucket.total++;
  bucket.points += scored.points;
  bucket.counts[scored.count] = (bucket.counts[scored.count] || 0) + 1;
  const event = Object.assign({ status: String(status || '').trim(), points: scored.points }, detail || {});
  bucket.events.push(event);
}

function riskSubjectStudentKey(studentId, grade, room, classroomFilter, byId) {
  const candidates = (byId[String(studentId)] || []).filter(function(student) {
    return !classroomFilter || student.classroom_id === classroomFilter;
  });
  if (!candidates.length) return '';
  const gradeRoom = (String(grade || '').trim() + '/' + String(room || '').trim()).trim();
  const matches = candidates.filter(function(student) {
    return student.classroom_id === String(room || '').trim() ||
      student.classroom_id === gradeRoom ||
      student.classroom_name === gradeRoom ||
      student.classroom_name === String(room || '').trim();
  });
  if (matches.length === 1) return matches[0]._riskKey;
  return candidates.length === 1 ? candidates[0]._riskKey : '';
}

function riskCollectDataset(startDate, endDate, classroomFilter) {
  const dates = riskGetDateList(startDate, endDate);
  const targetRoom = String(classroomFilter || '').trim();
  const allStudents = getAllStudents();
  if (!allStudents || allStudents.status !== 'success') throw new Error('โหลดรายชื่อนักเรียนไม่สำเร็จ');
  const classroomResponse = getClassrooms([]);
  const classrooms = classroomResponse && classroomResponse.status === 'success' ? classroomResponse.data || [] : [];
  const roomNameMap = {};
  classrooms.forEach(function(room) { roomNameMap[String(room.id)] = String(room.name || room.id); });

  const students = {};
  const byId = {};
  Object.keys(allStudents.data || {}).forEach(function(roomId) {
    const room = String(roomId).trim();
    if (targetRoom && room !== targetRoom) return;
    (allStudents.data[roomId] || []).forEach(function(student) {
      const id = String(student.student_id || '').trim();
      if (!id) return;
      const key = riskStudentKey(room, id);
      if (students[key]) return;
      const record = {
        id: id,
        name: String(student.name || '').trim(),
        classroom_id: room,
        classroom_name: roomNameMap[room] || room,
        _riskKey: key
      };
      students[key] = record;
      if (!byId[id]) byId[id] = [];
      byId[id].push(record);
    });
  });

  const attendance = {};
  const cleaning = {};
  const subject = {};
  const ss = getSpreadsheet();
  const attendanceSheet = ss.getSheetByName(SHEET_NAMES.ATTENDANCE);
  if (attendanceSheet) {
    getCleanedAttendanceRows(attendanceSheet.getDataRange().getValues()).forEach(function(row) {
      const date = formatDateStr(row[0]);
      const room = String(row[2] || '').trim();
      if (!date || date < startDate || date > endDate || (targetRoom && room !== targetRoom)) return;
      let entries = [];
      try { entries = JSON.parse(row[4] || '[]'); } catch (err) { return; }
      entries.forEach(function(entry) {
        const id = String(entry.student_id || '').trim();
        const key = riskStudentKey(room, id);
        if (!students[key]) return;
        riskAddEntry(attendance, key, date, 'attendance', entry.status, { period: String(row[1] || '').trim() });
      });
    });
  }

  const cleaningSheet = ss.getSheetByName(CLEANING_SHEET_NAMES.ATTENDANCE);
  if (cleaningSheet) {
    getCleanedCleaningRows(cleaningSheet.getDataRange().getValues()).forEach(function(row) {
      const date = formatDateStr(row[0]);
      const room = String(row[1] || '').trim();
      if (!date || date < startDate || date > endDate || (targetRoom && room !== targetRoom)) return;
      let entries = [];
      try { entries = JSON.parse(row[3] || '[]'); } catch (err) { return; }
      entries.forEach(function(entry) {
        const id = String(entry.student_id || '').trim();
        const key = riskStudentKey(room, id);
        if (!students[key]) return;
        riskAddEntry(cleaning, key, date, 'cleaning', entry.status, {});
      });
    });
  }

  const subjectSheet = ss.getSheetByName('Attendance_Subjects');
  if (subjectSheet) {
    const data = subjectSheet.getDataRange().getValues();
    const latest = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = formatDateStr(row[1]);
      if (!date || date < startDate || date > endDate) continue;
      const id = String(row[7] || '').trim();
      const grade = String(row[5] || '').trim();
      const room = String(row[6] || '').trim();
      const key = riskSubjectStudentKey(id, grade, room, targetRoom, byId);
      if (!key) continue;
      const code = String(row[2] || '').trim();
      const name = String(row[3] || '').trim();
      const subjectKey = code || name;
      const uniqueKey = [key, date, subjectKey].join('|');
      latest[uniqueKey] = { key: key, date: date, status: String(row[9] || '').trim(), subject: name || code, subjectCode: code };
    }
    Object.keys(latest).forEach(function(uniqueKey) {
      const item = latest[uniqueKey];
      riskAddEntry(subject, item.key, item.date, 'subject', item.status, { subject: item.subject, subjectCode: item.subjectCode });
    });
  }

  return { dates: dates, students: students, byId: byId, attendance: attendance, cleaning: cleaning, subject: subject, weights: null };
}

function riskAggregate(dayMap) {
  const total = { total: 0, points: 0, counts: {} };
  Object.keys(dayMap || {}).forEach(function(date) {
    const day = dayMap[date];
    total.total += day.total;
    total.points += day.points;
    Object.keys(day.counts || {}).forEach(function(key) {
      total.counts[key] = (total.counts[key] || 0) + day.counts[key];
    });
  });
  return total;
}

function riskBreakdown(total) {
  if (!total || total.total <= 0) return null;
  return {
    score: Math.round(total.points / total.total * 100),
    total: total.total,
    points: Number(total.points.toFixed(2)),
    present: total.counts.present || 0,
    absent: total.counts.absent || 0,
    late: total.counts.late || 0,
    leave: total.counts.leave || 0,
    sick: total.counts.sick || 0,
    done: total.counts.done || 0,
    notDone: total.counts.notDone || 0,
    excused: total.counts.excused || 0
  };
}

function riskComputeStudent(student, dataset, weights) {
  const key = student._riskKey;
  const attDays = dataset.attendance[key] || {};
  const cleaningDays = dataset.cleaning[key] || {};
  const subjectDays = dataset.subject[key] || {};
  const att = riskAggregate(attDays);
  const cleaning = riskAggregate(cleaningDays);
  const subject = riskAggregate(subjectDays);
  const percentages = {
    attendance: att.total ? att.points / att.total * 100 : null,
    cleaning: cleaning.total ? cleaning.points / cleaning.total * 100 : null,
    subject: subject.total ? subject.points / subject.total * 100 : null
  };
  let weighted = 0;
  let availableWeight = 0;
  ['attendance', 'cleaning', 'subject'].forEach(function(kind) {
    if (percentages[kind] !== null && weights[kind] > 0) {
      weighted += percentages[kind] * weights[kind];
      availableWeight += weights[kind];
    }
  });
  const composite = availableWeight ? Math.round(weighted / availableWeight) : null;

  let currentStreak = 0;
  let maxStreak = 0;
  dataset.dates.forEach(function(date) {
    const day = attDays[date];
    if (!day) return; // วันที่ไม่มี check ถูกข้าม ไม่ตีความว่าเป็นการขาด
    if (day.points === 0 && day.total > 0) {
      currentStreak++;
      maxStreak = Math.max(maxStreak, currentStreak);
    } else {
      currentStreak = 0;
    }
  });

  const dailyScores = {};
  dataset.dates.forEach(function(date) {
    const items = {
      attendance: attDays[date] || null,
      cleaning: cleaningDays[date] || null,
      subject: subjectDays[date] || null
    };
    let scoreSum = 0;
    let weightSum = 0;
    Object.keys(items).forEach(function(kind) {
      const item = items[kind];
      if (item && item.total > 0 && weights[kind] > 0) {
        scoreSum += item.points / item.total * 100 * weights[kind];
        weightSum += weights[kind];
      }
    });
    if (!weightSum) {
      dailyScores[date] = null;
      return;
    }
    const score = Math.round(scoreSum / weightSum);
    dailyScores[date] = {
      score: score,
      category: score >= 85 ? 'green' : score >= 60 ? 'yellow' : 'red',
      breakdown: {
        attendance: items.attendance ? { points: Number(items.attendance.points.toFixed(2)), total: items.attendance.total } : null,
        cleaning: items.cleaning ? { points: Number(items.cleaning.points.toFixed(2)), total: items.cleaning.total } : null,
        subject: items.subject ? { points: Number(items.subject.points.toFixed(2)), total: items.subject.total } : null
      }
    };
  });

  let category = 'yellow';
  if (maxStreak >= 3 || (composite !== null && composite < 60)) category = 'red';
  else if (composite === null || composite < 85 || maxStreak === 2) category = 'yellow';
  else category = 'green';

  return {
    id: student.id,
    name: student.name,
    classroom_id: student.classroom_id,
    classroom_name: student.classroom_name,
    composite: composite,
    hasData: composite !== null,
    category: category,
    streak: maxStreak,
    breakdown: {
      attendance: riskBreakdown(att),
      cleaning: riskBreakdown(cleaning),
      subject: riskBreakdown(subject)
    },
    dailyScores: dailyScores
  };
}

function getStudentRiskAssessment(startDate, endDate, classroomFilter, weightsJson) {
  try {
    const dataset = riskCollectDataset(String(startDate || ''), String(endDate || ''), classroomFilter);
    const weightResponse = weightsJson ? { data: riskNormalizeWeights(weightsJson) } : getRiskWeights();
    const weights = riskNormalizeWeights(weightResponse.data || RISK_DEFAULT_WEIGHTS);
    const results = Object.keys(dataset.students).map(function(key) {
      return riskComputeStudent(dataset.students[key], dataset, weights);
    });
    const order = { red: 0, yellow: 1, green: 2 };
    results.sort(function(a, b) {
      if (a.category !== b.category) return order[a.category] - order[b.category];
      if (a.composite === null) return b.composite === null ? a.name.localeCompare(b.name, 'th') : 1;
      if (b.composite === null) return -1;
      return a.composite - b.composite;
    });
    const summary = { red: 0, yellow: 0, green: 0, avgScore: 0, totalStudents: results.length, scoredStudents: 0 };
    let scoreTotal = 0;
    results.forEach(function(student) {
      summary[student.category]++;
      if (student.composite !== null) {
        scoreTotal += student.composite;
        summary.scoredStudents++;
      }
    });
    if (summary.scoredStudents) summary.avgScore = Math.round(scoreTotal / summary.scoredStudents);
    return {
      status: 'success',
      data: {
        period: { startDate: String(startDate), endDate: String(endDate), dates: dataset.dates },
        weights: weights,
        summary: summary,
        students: results
      }
    };
  } catch (err) {
    Logger.log('Error in getStudentRiskAssessment: ' + (err.message || err));
    return { status: 'error', message: err.message || String(err) };
  }
}

function getStudentRiskDetail(studentId, startDate, endDate, classroomFilter, weightsJson) {
  try {
    const dataset = riskCollectDataset(String(startDate || ''), String(endDate || ''), classroomFilter);
    const candidates = dataset.byId[String(studentId || '').trim()] || [];
    const student = candidates.find(function(item) { return !classroomFilter || item.classroom_id === String(classroomFilter); });
    if (!student) return { status: 'error', message: 'ไม่พบนักเรียนในช่วงข้อมูลที่เลือก' };
    const weights = riskNormalizeWeights(weightsJson ? weightsJson : (getRiskWeights().data || RISK_DEFAULT_WEIGHTS));
    const assessment = riskComputeStudent(student, dataset, weights);
    const key = student._riskKey;
    const rows = dataset.dates.map(function(date) {
      const att = (dataset.attendance[key] || {})[date] || null;
      const cleaning = (dataset.cleaning[key] || {})[date] || null;
      const subject = (dataset.subject[key] || {})[date] || null;
      const score = assessment.dailyScores[date] || null;
      return {
        date: date,
        attendance: att ? att.events : [],
        cleaning: cleaning ? cleaning.events : [],
        subject: subject ? subject.events : [],
        score: score ? score.score : null,
        category: score ? score.category : null
      };
    });
    return { status: 'success', data: { student: assessment, period: { startDate: String(startDate), endDate: String(endDate) }, days: rows } };
  } catch (err) {
    Logger.log('Error in getStudentRiskDetail: ' + (err.message || err));
    return { status: 'error', message: err.message || String(err) };
  }
}

function getRedFlagStudents(startDate, endDate, classroomFilter) {
  const result = getStudentRiskAssessment(startDate, endDate, classroomFilter);
  if (result.status !== 'success') return result;
  const students = result.data.students.filter(function(student) { return student.category === 'red'; });
  return { status: 'success', data: { period: result.data.period, total: students.length, students: students } };
}

// ========================================
// PHASE 3.6: Risk History + Print Support
// ========================================

/**
 * ดึงประวัติ Red/Yellow ของนักเรียนย้อนหลังตามจำนวนสัปดาห์
 * เพื่อประกอบการเลือกเอกสารสำหรับพิมพ์
 */
function getStudentRiskHistory(studentId, weeksBack, classroomFilter) {
  try {
    weeksBack = weeksBack || 8;
    const today = new Date();
    const results = [];

    for (let i = weeksBack - 1; i >= 0; i--) {
      const d = new Date(today);
      d.setDate(d.getDate() - i * 7);
      const dow = d.getDay();
      const monday = new Date(d);
      monday.setDate(d.getDate() - dow + (dow === 0 ? -6 : 1));
      const friday = new Date(monday);
      friday.setDate(monday.getDate() + 4);

      const start = Utilities.formatDate(monday, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const end = Utilities.formatDate(friday, Session.getScriptTimeZone(), 'yyyy-MM-dd');

      const weights = getRiskWeights().data;
      const assessment = getStudentRiskAssessment(start, end, classroomFilter || '', JSON.stringify(weights));
      if (assessment.status === 'success') {
        const student = (assessment.data.students || []).find(function(item) {
          return item.id === String(studentId) && (!classroomFilter || item.classroom_id === String(classroomFilter));
        });
        if (student) {
          results.push({
            startDate: start,
            endDate: end,
            composite: student.composite,
            category: student.category,
            streak: student.streak,
            breakdown: student.breakdown
          });
        }
      }
    }

    const redCount = results.filter(function(item) { return item.category === 'red'; }).length;
    const yellowCount = results.filter(function(item) { return item.category === 'yellow'; }).length;
    return {
      status: 'success',
      data: {
        weeks: results,
        redCount: redCount,
        yellowCount: yellowCount,
        recommendation: redCount >= 3 ? 'activity' : redCount >= 1 ? 'warning' : 'none'
      }
    };
  } catch (err) {
    Logger.log('Error in getStudentRiskHistory: ' + err.message);
    return { status: 'error', message: err.toString() };
  }
}
