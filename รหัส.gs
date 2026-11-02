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
  [SHEET_NAMES.TEACHERS]: ['id', 'name', 'username', 'password', 'status', 'classroom_ids', 'phone'],
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

// Cache timetable rows for 5 minutes to avoid reading the spreadsheet on every request.
function getCachedTimetableData() {
  var cacheKey = 'TIMETABLE_V1';
  var cache = null;

  try {
    cache = CacheService.getScriptCache();
    var cached = cache.get(cacheKey);
    if (cached) {
      var cachedData = JSON.parse(cached);
      if (Array.isArray(cachedData)) {
        Logger.log('✅ [Cache HIT] Timetable from cache');
        return cachedData;
      }
      cache.remove(cacheKey);
    }
  } catch (error) {
    Logger.log('⚠️ Timetable cache read failed: ' + error.message);
  }

  Logger.log('🔄 [Cache MISS] Loading Timetable from sheet...');
  var startedAt = Date.now();
  var sheet = getSpreadsheet().getSheetByName('Timetable');
  if (!sheet) {
    Logger.log('⚠️ ไม่พบชีต Timetable');
    return [];
  }

  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return [];
  var data = sheet.getRange(1, 1, lastRow, 9).getValues();
  Logger.log('📊 Timetable rows loaded: ' + data.length + ' in ' + (Date.now() - startedAt) + 'ms');

  if (cache) {
    try {
      var json = JSON.stringify(data);
      var byteSize = Utilities.newBlob(json, 'application/json').getBytes().length;
      if (byteSize < 90000) {
        cache.put(cacheKey, json, 300);
      } else {
        Logger.log('⚠️ Timetable too large for cache (' + byteSize + ' bytes); skipping cache');
      }
    } catch (error) {
      Logger.log('⚠️ Timetable cache write failed: ' + error.message);
    }
  }

  return data;
}

// Run manually after editing the Timetable sheet to invalidate its cached copy.
function clearTimetableCache() {
  try {
    CacheService.getScriptCache().remove('TIMETABLE_V1');
    return { status: 'success', message: 'ล้าง cache แล้ว' };
  } catch (error) {
    Logger.log('❌ Error clearing Timetable cache: ' + error.message);
    return { status: 'error', message: error.message || 'Unable to clear cache' };
  }
}

function refreshTimetableCache() {
  var result = clearTimetableCache();
  Logger.log(result.status === 'success' ? '✅ Timetable cache cleared' : '❌ Timetable cache clear failed');
  return result;
}

// Safe diagnostic for Apps Script Editor: reports row count and elapsed time only.
function debugSubjectSchedule() {
  var startedAt = Date.now();
  try {
    var data = getCachedTimetableData();
    var result = {
      status: 'success',
      rows: Math.max(0, data.length - 1),
      elapsedMs: Date.now() - startedAt
    };
    Logger.log('📅 debugSubjectSchedule: ' + JSON.stringify(result));
    return result;
  } catch (error) {
    var failure = {
      status: 'error',
      message: error && error.message ? String(error.message) : 'Unknown server error',
      elapsedMs: Date.now() - startedAt
    };
    Logger.log('❌ debugSubjectSchedule: ' + JSON.stringify(failure));
    return failure;
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
        classroom_ids: teacher[5] ? teacher[5].split(',') : [],
        phone: teacher[6] ? String(teacher[6]).trim() : ''
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
      classroom_ids: row[5] ? row[5].split(',') : [],
      phone: row[6] ? String(row[6]).trim() : ''
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

      // Server-side membership check against the active teacher's assigned classrooms.
      const teachersSheet = getOrCreateSheet(SHEET_NAMES.TEACHERS, spreadsheetMode);
      const teachersData = teachersSheet.getDataRange().getValues();
      const teacherRow = teachersData.slice(1).find(row =>
        String(row[2] || '').trim() === teacher && String(row[4] || '') === '1'
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
        classroom_ids: row[5] ? row[5].toString().trim().split(',').map(id => id.trim()) : [],
        phone: row[6]?.toString().trim() || ''
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
      classroom_ids: teacher[5] ? teacher[5].toString().trim().split(',').map(id => id.trim()) : [],
      phone: teacher[6]?.toString().trim() || ''
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

      // Server-side membership check against the active teacher's assigned classrooms.
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
  var startedAt = Date.now();
  try {
    Logger.log('📅 getTeacherSubjectSchedule START: date="' + targetDate + '"');

    var data = getCachedTimetableData();
    Logger.log('⏱️ After cache: ' + (Date.now() - startedAt) + 'ms, rows: ' + data.length);
    if (!data || data.length <= 1) {
      return { status: 'success', data: [], message: 'ไม่มีข้อมูลในชีต Timetable' };
    }

    // Parse the requested date in the script's local timezone.
    var dateObj;
    if (targetDate) {
      var parts = String(targetDate).split('-');
      dateObj = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    } else {
      dateObj = new Date();
    }

    var days = ['อาทิตย์', 'จันทร์', 'อังคาร', 'พุธ', 'พฤหัสบดี', 'ศุกร์', 'เสาร์'];
    var dayName = days[dateObj.getDay()];
    var normEmail = String(teacherEmail || '').trim().toLowerCase();
    Logger.log('🗓️ Looking for: Day=' + dayName);

    var result = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || row.length < 9) continue;

      var rowDay = String(row[1] || '').trim();
      var rowEmailNorm = String(row[5] || '').trim().toLowerCase();
      if (rowDay === dayName && rowEmailNorm === normEmail) {
        result.push({
          timetableId: String(row[0] || ''),
          period: row[2],
          subjectCode: String(row[3] || ''),
          subjectName: String(row[4] || ''),
          gradeLevel: String(row[6] || ''),
          room: String(row[7] || ''),
          classroomId: String(row[8] || '')
        });
      }
    }

    Logger.log('✅ Found ' + result.length + ' periods in ' + (Date.now() - startedAt) + 'ms total');
    return { status: 'success', data: result };
  } catch (error) {
    var message = error && error.message ? String(error.message) : 'Unknown server error';
    Logger.log('❌ getTeacherSubjectSchedule failed after ' + (Date.now() - startedAt) + 'ms: ' + message);
    return { status: 'error', message: message, data: [] };
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

// ========================================
// PHASE 4: DISCIPLINE WORKFLOW
// ========================================

const DISCIPLINE_SHEETS = { CASES: 'Discipline_Cases', LOGS: 'Discipline_Logs' };

const DISCIPLINE_HEADERS = {
  CASES: [
    'case_id','student_id','student_name','classroom_id','classroom_name',
    'stage','status','trigger',
    'warn_doc_no','warn_date','bond_doc_no','bond_date','bond_start','bond_end',
    'activity_doc_no','activity_date','parent_doc_no','parent_date',
    'parent_name','parent_relation','activity_data','activity_result',
    'meeting_date','meeting_time','meeting_location',
    'signed_director_at','signed_parent_at','signed_student_at',
    'signed_witness1_at','signed_witness2_at',
    'printed_warning_at','printed_bond_at','printed_activity_at','printed_parent_at',
    'notes','created_at','updated_at','trigger_data'
  ],
  LOGS: ['log_id','case_id','action','stage_from','stage_to','actor','note','created_at']
};

function getDisciplineSheet(kind) {
  const ss = getSpreadsheet();
  const name = kind === 'CASES' ? DISCIPLINE_SHEETS.CASES : DISCIPLINE_SHEETS.LOGS;
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    const headers = DISCIPLINE_HEADERS[kind];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#1e40af').setFontColor('#fff');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function rowToCaseObj(row) {
  const safeJson = (v) => { try { return v ? JSON.parse(v) : null; } catch(e) { return null; } };
  return {
    case_id: String(row[0] || ''), student_id: String(row[1] || ''),
    student_name: String(row[2] || ''), classroom_id: String(row[3] || ''),
    classroom_name: String(row[4] || ''), stage: String(row[5] || ''),
    status: String(row[6] || ''), trigger: String(row[7] || ''),
    warn_doc_no: String(row[8] || ''), warn_date: formatDateStr(row[9]),
    bond_doc_no: String(row[10] || ''), bond_date: formatDateStr(row[11]),
    bond_start: formatDateStr(row[12]), bond_end: formatDateStr(row[13]),
    activity_doc_no: String(row[14] || ''), activity_date: formatDateStr(row[15]),
    parent_doc_no: String(row[16] || ''), parent_date: formatDateStr(row[17]),
    parent_name: String(row[18] || ''), parent_relation: String(row[19] || ''),
    activity_data: safeJson(row[20]), activity_result: String(row[21] || ''),
    meeting_date: formatDateStr(row[22]), meeting_time: String(row[23] || ''),
    meeting_location: String(row[24] || ''),
    signed_director_at: String(row[25] || ''), signed_parent_at: String(row[26] || ''),
    signed_student_at: String(row[27] || ''),
    signed_witness1_at: String(row[28] || ''), signed_witness2_at: String(row[29] || ''),
    printed_warning_at: String(row[30] || ''), printed_bond_at: String(row[31] || ''),
    printed_activity_at: String(row[32] || ''), printed_parent_at: String(row[33] || ''),
    notes: String(row[34] || ''), created_at: String(row[35] || ''),
    updated_at: String(row[36] || ''),
    trigger_data: safeJson(row[37])
  };
}

const DISCIPLINE_GARUDA_URL = 'https://www.xn--12cr3ayd4cc5c1a6ccp8m.com/img/articles/105785016020260201_115039.png';

function getDisciplineConfig() {
  const sheet = getOrCreateSheet(SHEET_NAMES.CONFIG);
  const data = sheet.getDataRange().getValues();
  const cfg = {
    school_name: '',
    director_name: '',
    director_position: '',
    discipline_teacher: '',
    garuda_url: DISCIPLINE_GARUDA_URL,
    activity_volunteer_hours: null,
    activity_cleaning_count: null,
    activity_counseling_count: null,
    warnings: []
  };
  const keyMap = {
    schoolName: 'school_name',
    director_name: 'director_name',
    director_position: 'director_position',
    discipline_teacher: 'discipline_teacher',
    discipline_activity_volunteer_hours: 'activity_volunteer_hours',
    discipline_activity_cleaning_count: 'activity_cleaning_count',
    discipline_activity_counseling_count: 'activity_counseling_count',
    garuda_url: 'garuda_url'
  };
  data.slice(1).forEach(r => {
    const key = String(r[0] || '').trim();
    if (!keyMap[key]) return;
    const value = String(r[1] == null ? '' : r[1]).trim();
    const configKey = keyMap[key];
    if (configKey.indexOf('activity_') === 0) {
      cfg[configKey] = value === '' ? null : Number(value);
    } else if (value) {
      cfg[configKey] = value;
    }
  });

  if (!cfg.school_name) cfg.warnings.push('ยังไม่ได้ตั้งค่า schoolName ในชีต Config');
  if (!cfg.director_name) cfg.warnings.push('ยังไม่ได้ตั้งค่า director_name ในชีต Config');
  if (!cfg.director_position) cfg.warnings.push('ยังไม่ได้ตั้งค่า director_position ในชีต Config');
  if (!cfg.discipline_teacher) cfg.warnings.push('ยังไม่ได้ตั้งค่า discipline_teacher ในชีต Config');
  if (!Number.isFinite(cfg.activity_volunteer_hours) || cfg.activity_volunteer_hours <= 0) {
    cfg.warnings.push('ยังไม่ได้ตั้งค่า discipline_activity_volunteer_hours ในชีต Config');
  }
  if (!Number.isFinite(cfg.activity_cleaning_count) || cfg.activity_cleaning_count <= 0) {
    cfg.warnings.push('ยังไม่ได้ตั้งค่า discipline_activity_cleaning_count ในชีต Config');
  }
  if (!Number.isFinite(cfg.activity_counseling_count) || cfg.activity_counseling_count <= 0) {
    cfg.warnings.push('ยังไม่ได้ตั้งค่า discipline_activity_counseling_count ในชีต Config');
  }
  return cfg;
}

function getClassroomAdvisor(classroomId) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.TEACHERS);
    const data = sheet.getDataRange().getValues();
    const teacher = data.slice(1).find(row =>
      String(row[4] || '') === '1' &&
      String(row[5] || '').split(',').map(id => id.trim()).includes(String(classroomId))
    );
    if (!teacher) return { name: '', phone: '' };
    return {
      name: String(teacher[1] || '').trim(),
      phone: String(teacher[6] || '').trim()
    };
  } catch (e) {
    Logger.log('getClassroomAdvisor error: ' + e.message);
    return { name: '', phone: '' };
  }
}

/**
 * เพิ่มเฉพาะคีย์ Config ที่ยังไม่มีและหัวคอลัมน์ที่ขาด โดยไม่เติมชื่อ/เบอร์/ข้อมูลตัวอย่าง
 * ฟังก์ชันนี้เป็น migration สำหรับผู้ดูแลเท่านั้น และไม่ได้ถูกรันในงานนี้
 */
function setupDisciplineConfig() {
  try {
    const configSheet = getOrCreateSheet(SHEET_NAMES.CONFIG);
    const values = configSheet.getDataRange().getValues();
    const existing = new Set(values.slice(1).map(row => String(row[0] || '').trim()).filter(Boolean));
    const year = new Date().getFullYear() + 543;
    const yearKey = 'docCounter_' + year;
    const rows = [
      ['director_name', ''],
      ['director_position', ''],
      ['discipline_teacher', ''],
      ['discipline_activity_volunteer_hours', ''],
      ['discipline_activity_cleaning_count', ''],
      ['discipline_activity_counseling_count', ''],
      ['garuda_url', DISCIPLINE_GARUDA_URL],
      [yearKey, '0']
    ].filter(row => !existing.has(row[0]));

    // Validate both headers before appending config rows or modifying header cells.
    const casesSheet = getDisciplineSheet('CASES');
    const caseHeader = casesSheet.getRange(1, 1, 1, Math.max(38, casesSheet.getLastColumn())).getValues()[0];
    if (String(caseHeader[37] || '').trim() && String(caseHeader[37]).trim() !== 'trigger_data') {
      throw new Error('คอลัมน์ AL มีหัวข้ออื่นอยู่แล้ว จึงไม่แก้ทับ');
    }
    const teachersSheet = getOrCreateSheet(SHEET_NAMES.TEACHERS);
    const teacherHeader = teachersSheet.getRange(1, 1, 1, Math.max(7, teachersSheet.getLastColumn())).getValues()[0];
    if (String(teacherHeader[6] || '').trim() && String(teacherHeader[6]).trim() !== 'phone') {
      throw new Error('คอลัมน์ G ของ Teachers มีหัวข้ออื่นอยู่แล้ว จึงไม่แก้ทับ');
    }

    rows.forEach(row => configSheet.appendRow(row));
    if (String(caseHeader[37] || '').trim() !== 'trigger_data') casesSheet.getRange(1, 38).setValue('trigger_data');
    if (String(teacherHeader[6] || '').trim() !== 'phone') teachersSheet.getRange(1, 7).setValue('phone');

    return { status: 'success', addedConfigKeys: rows.map(row => row[0]), message: 'เพิ่มเฉพาะหัวข้อและคีย์ว่างที่จำเป็น โดยไม่เติมข้อมูลบุคคล' };
  } catch (err) {
    Logger.log('setupDisciplineConfig error: ' + err.message);
    return { status: 'error', message: err.toString() };
  }
}

/**
 * คำนวณกลุ่มประเมินของสัปดาห์ย้อนหลัง โดยใช้ช่วงวันเดียวกับการประเมินปัจจุบัน
 * สัปดาห์ที่ไม่มีข้อมูลจะไม่ถูกนับเป็น Red
 */
function getAssessmentHistoryForStudents(weeksBack, currentStartDate, currentEndDate) {
  const results = [];
  const addDays = (dateString, days) => {
    const date = new Date(String(dateString) + 'T00:00:00Z');
    if (isNaN(date.getTime())) throw new Error('ช่วงวันที่ประเมินไม่ถูกต้อง');
    date.setUTCDate(date.getUTCDate() + days);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  };
  const count = Math.max(0, Math.min(12, parseInt(weeksBack, 10) || 0));
  for (let i = 1; i <= count; i++) {
    const start = addDays(currentStartDate, -7 * i);
    const end = addDays(currentEndDate, -7 * i);
    try {
      const assessment = getStudentRiskAssessment(start, end, '', '');
      if (assessment.status !== 'success') {
        results.push({ weekStart: start, weekEnd: end, byStudent: {}, error: assessment.message || 'ประเมินไม่สำเร็จ' });
        continue;
      }
      const byStudent = {};
      (assessment.data.students || []).forEach(student => { byStudent[String(student.id)] = student.category; });
      results.push({ weekStart: start, weekEnd: end, byStudent });
    } catch (err) {
      results.push({ weekStart: start, weekEnd: end, byStudent: {}, error: err.message || String(err) });
    }
  }
  return results;
}

function getCurrentDisciplineWeekRange() {
  const today = new Date();
  const day = today.getDay();
  const monday = new Date(today);
  monday.setDate(today.getDate() + (day === 0 ? -6 : 1 - day));
  const friday = new Date(monday);
  friday.setDate(monday.getDate() + 4);
  return {
    startDate: Utilities.formatDate(monday, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    endDate: Utilities.formatDate(friday, Session.getScriptTimeZone(), 'yyyy-MM-dd')
  };
}

// ---------- CRUD ----------

function getDisciplineCases(filter) {
  try {
    filter = filter || {};
    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    let cases = data.slice(1).map(rowToCaseObj).filter(c => c.case_id);

    if (filter.stage && filter.stage !== 'all') cases = cases.filter(c => c.stage === String(filter.stage));
    if (filter.classroom) cases = cases.filter(c => c.classroom_id === filter.classroom);
    if (filter.search) {
      const q = String(filter.search).toLowerCase();
      cases = cases.filter(c => c.student_name.toLowerCase().includes(q) || c.student_id.includes(q) || c.case_id.toLowerCase().includes(q));
    }

    const order = { '1': 0, '2': 1, '3': 2, '4': 3, 'closed': 4 };
    cases.sort((a, b) => {
      if (a.stage !== b.stage) return (order[a.stage] ?? 9) - (order[b.stage] ?? 9);
      return new Date(b.created_at) - new Date(a.created_at);
    });

    const stats = { stage1: 0, stage2: 0, stage3: 0, stage4: 0, closed: 0 };
    data.slice(1).forEach(r => {
      const st = String(r[5] || '');
      if (st === '1') stats.stage1++;
      else if (st === '2') stats.stage2++;
      else if (st === '3') stats.stage3++;
      else if (st === '4') stats.stage4++;
      else if (st === 'closed') stats.closed++;
    });

    return { status: 'success', data: { cases, stats } };
  } catch (err) {
    return { status: 'error', message: err.toString(), data: { cases: [], stats: {} } };
  }
}

function getDisciplineCaseDetail(caseId) {
  try {
    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    const row = data.slice(1).find(r => String(r[0]) === String(caseId));
    if (!row) return { status: 'error', message: 'ไม่พบเคส' };

    const caseData = rowToCaseObj(row);
    const cfg = getDisciplineConfig();

    // Load logs
    const logSheet = getDisciplineSheet('LOGS');
    const logData = logSheet.getDataRange().getValues();
    const logs = logData.slice(1)
      .filter(l => String(l[1]) === String(caseId))
      .map(l => ({
        log_id: String(l[0] || ''), action: String(l[2] || ''),
        stage_from: String(l[3] || ''), stage_to: String(l[4] || ''),
        actor: String(l[5] || ''), note: String(l[6] || ''),
        created_at: String(l[7] || '')
      }))
      .sort((a, b) => new Date(b.created_at) - new Date(a.created_at));

    return {
      status: 'success',
      data: {
        case: caseData, logs,
        config: cfg,
        advisor: getClassroomAdvisor(caseData.classroom_id)
      }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function addDisciplineLog(caseId, action, stageFrom, stageTo, actor, note) {
  try {
    const logSheet = getDisciplineSheet('LOGS');
    const logId = 'L' + Date.now() + Math.floor(Math.random() * 1000);
    logSheet.appendRow([
      logId, caseId, action, stageFrom || '', stageTo || '',
      actor || '', note || '', new Date().toISOString()
    ]);
  } catch(e) { Logger.log('Log error: ' + e.message); }
}

// ---------- Auto-sync from assessment ----------

function syncDisciplineFromWeekly(startDate, endDate, actor) {
  try {
    const assessment = getStudentRiskAssessment(startDate, endDate, '', '');
    if (assessment.status !== 'success') throw new Error(assessment.message);
    const students = assessment.data.students || [];
    const existingResult = getDisciplineCases({});
    if (existingResult.status !== 'success') throw new Error(existingResult.message || 'อ่านเคสวินัยไม่สำเร็จ');
    const existing = existingResult.data.cases || [];
    const activeCases = existing.filter(c => c.stage !== 'closed' && c.status !== 'closed');
    const created = [], advanced = [], blocked = [];
    const year = new Date().getFullYear() + 543;
    let maxNum = 0;
    existing.forEach(c => {
      const match = String(c.case_id || '').match(/D-(\d+)-(\d+)/);
      if (match && Number(match[1]) === year) maxNum = Math.max(maxNum, Number(match[2]));
    });

    students.forEach(student => {
      const active = activeCases.find(c => c.student_id === String(student.id));
      const now = new Date().toISOString();
      if (!active) {
        if (student.category !== 'red' && student.category !== 'yellow') return;
        maxNum++;
        const caseId = `D-${year}-${String(maxNum).padStart(3, '0')}`;
        const trigger = student.category === 'red' ? 'red_1' : 'yellow';
        const triggerData = {
          weekStart: String(startDate),
          weekEnd: String(endDate),
          category: student.category,
          composite: student.composite,
          breakdown: student.breakdown || null
        };
        const row = new Array(38).fill('');
        row[0] = caseId; row[1] = student.id; row[2] = student.name;
        row[3] = student.classroom_id; row[4] = student.classroom_name;
        row[5] = '1'; row[6] = 'active'; row[7] = trigger;
        row[35] = now; row[36] = now; row[37] = JSON.stringify(triggerData);
        getDisciplineSheet('CASES').appendRow(row);
        addDisciplineLog(caseId, 'created', '', '1', actor || 'system',
          `สร้างเคสจากผลประเมิน ${student.category} (${startDate} ถึง ${endDate})`);
        created.push({ case_id: caseId, student_name: student.name, category: student.category });
        return;
      }

      if (active.stage === '1' && (student.category === 'red' || student.category === 'yellow')) {
        const result = advanceDisciplineStage_(active.case_id, '2',
          `ประเมินซ้ำหลังครบ 1 สัปดาห์ (${student.category}, ${student.composite}%)`,
          actor || 'system', startDate, endDate);
        if (result.status === 'success') advanced.push({ case_id: active.case_id, from: '1', to: '2' });
        else blocked.push({ case_id: active.case_id, stage: '1', reason: result.message });
        return;
      }

      if (active.stage === '2' && student.category === 'red') {
        const result = advanceDisciplineStage_(active.case_id, '3',
          `ผลประเมิน Red ต่อเนื่อง 3 สัปดาห์ (${startDate} ถึง ${endDate})`,
          actor || 'system', startDate, endDate);
        if (result.status === 'success') advanced.push({ case_id: active.case_id, from: '2', to: '3' });
        else if (result.code !== 'ASSESSMENT_STREAK_NOT_MET') {
          blocked.push({ case_id: active.case_id, stage: '2', reason: result.message });
        }
      }
    });

    return { status: 'success', data: { created, advanced, blocked, totalProcessed: students.length } };
  } catch (err) {
    Logger.log('syncDisciplineFromWeekly error: ' + err.message);
    return { status: 'error', message: err.toString(), data: { created: [], advanced: [], blocked: [] } };
  }
}

function advanceDisciplineStage(caseId, newStage, note, actor) {
  const range = getCurrentDisciplineWeekRange();
  return advanceDisciplineStage_(caseId, newStage, note, actor, range.startDate, range.endDate);
}

/** Internal transition gate. Trailing underscore prevents direct client invocation in Apps Script. */
function advanceDisciplineStage_(caseId, newStage, note, actor, assessmentStartDate, assessmentEndDate) {
  try {
    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    const idx = data.slice(1).findIndex(r => String(r[0]) === String(caseId)) + 2;
    if (idx <= 1) return { status: 'error', message: 'ไม่พบเคส' };

    const row = data[idx - 1];
    const oldStage = String(row[5] || '');
    newStage = String(newStage || '');
    const allowedNextStages = { '1': ['2', 'closed'], '2': ['3', 'closed'], '3': ['4', 'closed'], '4': ['closed'] };
    if (!(allowedNextStages[oldStage] || []).includes(newStage)) {
      return { status: 'error', message: `ไม่สามารถเปลี่ยนจาก Stage ${oldStage} เป็น ${newStage} ได้` };
    }
    const now = new Date().toISOString();

    if ((oldStage === '1' && newStage === '2') || (oldStage === '2' && newStage === '3')) {
      const start = String(assessmentStartDate || '').trim();
      const end = String(assessmentEndDate || '').trim();
      if (!/^\d{4}-\d{2}-\d{2}$/.test(start) || !/^\d{4}-\d{2}-\d{2}$/.test(end)) {
        return { status: 'error', code: 'INVALID_ASSESSMENT_PERIOD', message: 'ช่วงวันที่ประเมินไม่ถูกต้อง' };
      }
      const currentAssessment = getStudentRiskAssessment(start, end, '', '');
      if (currentAssessment.status !== 'success') {
        return { status: 'error', code: 'ASSESSMENT_UNAVAILABLE', message: currentAssessment.message || 'อ่านผลประเมินไม่สำเร็จ' };
      }
      const student = (currentAssessment.data.students || []).find(item => String(item.id) === String(row[1]));
      if (!student) return { status: 'error', code: 'ASSESSMENT_MISSING', message: 'ไม่พบผลประเมินของนักเรียนในช่วงวันที่ที่เลือก' };

      if (oldStage === '1' && newStage === '2') {
        if (student.category !== 'red' && student.category !== 'yellow') {
          return { status: 'error', code: 'ASSESSMENT_NOT_LOW', message: 'ยังไม่ยกระดับ: ผลประเมินล่าสุดไม่อยู่กลุ่ม Red หรือ Yellow' };
        }
        const createdAt = new Date(row[35]);
        const assessmentEnd = new Date(end + 'T23:59:59Z');
        const elapsedDays = Math.floor((assessmentEnd.getTime() - createdAt.getTime()) / 86400000);
        if (isNaN(createdAt.getTime()) || elapsedDays < 7) {
          return { status: 'error', code: 'TOO_EARLY', message: `ยังไม่ยกระดับ: ต้องมีผลประเมิน Red/Yellow หลังสร้างเคสอย่างน้อย 7 วัน (ช่วงนี้ห่าง ${isNaN(elapsedDays) ? 0 : elapsedDays} วัน)` };
        }
      }

      if (oldStage === '2' && newStage === '3') {
        const required = [
          { label: 'ผู้อำนวยการ', value: row[25] },
          { label: 'ผู้ปกครอง', value: row[26] },
          { label: 'นักเรียน', value: row[27] }
        ];
        const missing = required.filter(item => !String(item.value || '').trim()).map(item => item.label);
        if (missing.length) {
          return { status: 'error', code: 'SIGNATURES_INCOMPLETE', message: `ยังไม่ยกระดับ: ขาดลายเซ็น ${missing.join(', ')}` };
        }
        if (student.category !== 'red') {
          return { status: 'error', code: 'ASSESSMENT_STREAK_NOT_MET', message: 'ผลประเมินสัปดาห์นี้ไม่ใช่ Red' };
        }
        const history = getAssessmentHistoryForStudents(2, start, end);
        const redThreeWeeks = history.length === 2 && history.every(week => !week.error && week.byStudent[String(row[1])] === 'red');
        if (!redThreeWeeks) {
          return { status: 'error', code: 'ASSESSMENT_STREAK_NOT_MET', message: 'ยังไม่ยกระดับ: ต้องมีผลประเมิน Red ติดต่อกัน 3 สัปดาห์ โดยสัปดาห์ที่ไม่มีข้อมูลจะไม่ถูกนับ' };
        }
      }
    }

    if (oldStage === '3' && newStage === '4' && String(row[21] || '').trim() !== 'fail') {
      return { status: 'error', code: 'ACTIVITY_NOT_FAILED', message: 'ยังไม่ยกระดับ: ต้องบันทึกผลกิจกรรมเป็น “ไม่ผ่าน” ก่อน' };
    }
    if (oldStage === '3' && newStage === 'closed' && String(row[21] || '').trim() !== 'pass') {
      return { status: 'error', code: 'ACTIVITY_NOT_PASSED', message: 'ยังไม่ปิดเคส: ต้องบันทึกผลกิจกรรมเป็น “ผ่าน” ก่อน' };
    }

    let activityData = null;
    if (newStage === '3' && !String(row[20] || '').trim()) {
      const config = getDisciplineConfig();
      const targets = [config.activity_volunteer_hours, config.activity_cleaning_count, config.activity_counseling_count];
      if (targets.some(value => !Number.isFinite(value) || value <= 0)) {
        return { status: 'error', code: 'ACTIVITY_CONFIG_MISSING', message: 'ยังไม่ได้กำหนดเป้าหมายกิจกรรมจริงในชีต Config จึงยังเปิด Stage 3 ไม่ได้' };
      }
      const start = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const end = Utilities.formatDate(new Date(Date.now() + 14 * 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      activityData = {
        items: [
          { id: 'a1', label: 'บำเพ็ญประโยชน์', unit: 'ชั่วโมง', target: config.activity_volunteer_hours, progress: 0 },
          { id: 'a2', label: 'ทำความสะอาดเขตพื้นที่', unit: 'ครั้ง', target: config.activity_cleaning_count, progress: 0 },
          { id: 'a3', label: 'รับคำปรึกษา', unit: 'ครั้ง', target: config.activity_counseling_count, progress: 0 }
        ], start, end
      };
    }

    sheet.getRange(idx, 6).setValue(newStage);
    sheet.getRange(idx, 7).setValue(newStage === 'closed' ? 'closed' : 'active');
    sheet.getRange(idx, 37).setValue(now);
    if (activityData) {
      sheet.getRange(idx, 21).setValue(JSON.stringify(activityData));
      sheet.getRange(idx, 22).setValue('pending');
    }
    if (newStage === '2' && !String(row[12] || '').trim()) {
      const bondStart = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const bondEnd = Utilities.formatDate(new Date(Date.now() + 90 * 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      sheet.getRange(idx, 13).setValue(bondStart);
      sheet.getRange(idx, 14).setValue(bondEnd);
    }

    addDisciplineLog(caseId, 'stage_advance', oldStage, newStage, actor || '', note || '');
    return { status: 'success', message: `ยกระดับเป็น Stage ${newStage} สำเร็จ`, data: { oldStage, newStage } };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function recordSignature(caseId, party, meta) {
  try {
    meta = meta || {};
    const colMap = {
      director: 26, parent: 27, student: 28, witness1: 29, witness2: 30
    };
    if (!Object.prototype.hasOwnProperty.call(colMap, party)) return { status: 'error', message: 'party ไม่ถูกต้อง' };
    const col = colMap[party];

    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    const idx = data.slice(1).findIndex(r => String(r[0]) === caseId) + 2;
    if (idx <= 1) return { status: 'error', message: 'ไม่พบเคส' };
    if (String(data[idx - 1][5] || '') !== '2') return { status: 'error', message: 'การลงนามใช้ได้เฉพาะ Stage 2 เท่านั้น' };
    if (String(data[idx - 1][col - 1] || '').trim()) return { status: 'error', message: 'ฝ่ายนี้บันทึกการลงนามแล้ว' };
    if (party === 'parent' && (!String(meta.name || '').trim() || !String(meta.relation || '').trim())) {
      return { status: 'error', message: 'กรุณาระบุชื่อจริงและความสัมพันธ์ของผู้ปกครองก่อนบันทึกการลงนาม' };
    }

    const now = new Date().toISOString();
    sheet.getRange(idx, col).setValue(now);

    // ถ้าเป็น Stage 2 → เก็บชื่อผู้ปกครอง/ความสัมพันธ์
    if (party === 'parent') {
      if (meta.name) sheet.getRange(idx, 19).setValue(meta.name);
      if (meta.relation) sheet.getRange(idx, 20).setValue(meta.relation);
    }

    sheet.getRange(idx, 37).setValue(now);
    addDisciplineLog(caseId, 'signed', '', '', party, `${party} signed`);

    // เช็คว่า Stage 2 ครบ 3 ฝ่ายหรือยัง
    const row = sheet.getRange(idx, 1, 1, 37).getValues()[0];
    const stage = String(row[5]);
    if (stage === '2') {
      const s1 = String(row[25] || '').trim();
      const s2 = String(row[26] || '').trim();
      const s3 = String(row[27] || '').trim();
      if (s1 && s2 && s3) {
        sheet.getRange(idx, 7).setValue('active');
        addDisciplineLog(caseId, 'signatures_complete', '', '', 'system', 'ครบ 3 ฝ่าย → active');
      }
    }

    return { status: 'success', message: 'บันทึกการลงนามสำเร็จ' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function updateActivityProgress(caseId, itemId, delta) {
  try {
    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    const idx = data.slice(1).findIndex(r => String(r[0]) === caseId) + 2;
    if (idx <= 1) return { status: 'error', message: 'ไม่พบเคส' };
    if (String(data[idx - 1][5] || '') !== '3') return { status: 'error', message: 'การอัปเดตกิจกรรมใช้ได้เฉพาะ Stage 3 เท่านั้น' };

    let actData;
    try { actData = JSON.parse(data[idx - 1][20] || '{}'); } catch(e) { actData = { items: [] }; }

    const item = (actData.items || []).find(i => i.id === itemId);
    if (!item) return { status: 'error', message: 'ไม่พบกิจกรรม' };
    item.progress = Math.max(0, Math.min(item.target, (item.progress || 0) + delta));

    sheet.getRange(idx, 21).setValue(JSON.stringify(actData));
    sheet.getRange(idx, 37).setValue(new Date().toISOString());
    return { status: 'success', data: { activity: actData } };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function markActivityResult(caseId, result) {
  try {
    if (!['pass', 'fail'].includes(result)) return { status: 'error', message: 'result ไม่ถูกต้อง' };
    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    const idx = data.slice(1).findIndex(r => String(r[0]) === caseId) + 2;
    if (idx <= 1) return { status: 'error', message: 'ไม่พบเคส' };
    if (String(data[idx - 1][5] || '') !== '3') return { status: 'error', message: 'การสรุปผลกิจกรรมใช้ได้เฉพาะ Stage 3 เท่านั้น' };

    sheet.getRange(idx, 22).setValue(result);
    sheet.getRange(idx, 37).setValue(new Date().toISOString());
    addDisciplineLog(caseId, 'activity_' + result, '', '', 'system', `กิจกรรม ${result}`);

    return { status: 'success', message: result === 'pass' ? 'บันทึกว่าผ่านกิจกรรม' : 'บันทึกว่าไม่ผ่าน → ยกระดับเป็น Stage 4' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function markDocPrinted(caseId, docType) {
  try {
    const colMap = { warning: 31, bond: 32, activity: 33, parent: 34 };
    const col = colMap[docType];
    if (!col) return { status: 'error', message: 'docType ไม่ถูกต้อง' };
    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    const idx = data.slice(1).findIndex(r => String(r[0]) === String(caseId)) + 2;
    if (idx <= 1) return { status: 'error', message: 'ไม่พบเคส' };
    sheet.getRange(idx, col).setValue(new Date().toISOString());
    return { status: 'success' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function reserveDocNumber() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('จองเลขหนังสือไม่สำเร็จ กรุณาลองใหม่');
  try {
    const year = new Date().getFullYear() + 543;
    const yearShort = String(year).slice(-2);
    const configKey = `docCounter_${year}`;
    const configSheet = getOrCreateSheet(SHEET_NAMES.CONFIG);
    const configData = configSheet.getDataRange().getValues();
    let configRow = -1;
    let counter = 0;
    configData.slice(1).forEach((row, index) => {
      if (String(row[0] || '').trim() === configKey) {
        configRow = index + 2;
        counter = Math.max(counter, parseInt(row[1], 10) || 0);
      }
    });

    // Protect against a missing/stale counter by reading previously issued real document numbers.
    const cases = getDisciplineSheet('CASES').getDataRange().getValues();
    cases.slice(1).forEach(row => {
      [8, 10, 14, 16].forEach(index => {
        const match = String(row[index] || '').match(/\/(\d{2})(\d+)$/);
        if (match && match[1] === yearShort) counter = Math.max(counter, parseInt(match[2], 10) || 0);
      });
    });
    counter += 1;
    if (counter > 999999) throw new Error('เลขหนังสือประจำปีเกินขีดจำกัด');
    if (configRow > 0) configSheet.getRange(configRow, 2).setValue(counter);
    else configSheet.appendRow([configKey, counter]);
    return `ศธ 04201/${yearShort}${String(counter).padStart(4, '0')}`;
  } finally {
    lock.releaseLock();
  }
}

// ---------- Document HTML builders ----------

function buildDocStyle() {
  return `<style>
    @page { size: A4; margin: 2cm 2cm 2cm 2.5cm; }
    body { font-family: 'TH Sarabun New', 'TH SarabunPSK', 'Sarabun', sans-serif; font-size: 16pt; line-height: 1.4; margin: 0; color: #000; }
    .toolbar { position: fixed; top: 10px; right: 10px; background: #fff; padding: 8px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,.15); z-index: 999; }
    .toolbar button { padding: 6px 14px; border: 1px solid #ccc; border-radius: 6px; cursor: pointer; font-weight: bold; }
    @media print { .toolbar { display: none; } }
    .garuda-wrap { text-align: center; margin-bottom: 4px; }
    .garuda-wrap img { width: 60px; height: 60px; object-fit: contain; display: inline-block; }
    .header { text-align: center; font-size: 18pt; font-weight: bold; }
    .header-sub { text-align: center; font-size: 16pt; margin-bottom: 20px; }
    .meta { display: flex; justify-content: space-between; margin-bottom: 16px; }
    .indent { text-indent: 2.5cm; text-align: justify; margin: 8px 0; }
    .check-line { padding-left: 3cm; text-indent: -0.8cm; margin: 4px 0; }
    .signature { text-align: right; margin-top: 40px; }
    .sig-row { display: flex; justify-content: space-around; margin-top: 40px; }
    .sig-row > div { text-align: center; width: 45%; }
    .cut-line { border-top: 1px dashed #999; margin: 30px 0 20px; }
    .page-break { page-break-after: always; }
  </style>`;
}

function buildDocWrap(title, body, caseId, docType) {
  const metadata = JSON.stringify({ caseId: String(caseId || ''), docType: String(docType || '') }).replace(/</g, '\\u003c');
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${title}</title>
<link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap" rel="stylesheet">
${buildDocStyle()}</head><body>
<div class="toolbar"><button onclick="window.print()">🖨️ พิมพ์ / บันทึก PDF</button></div>
${body}
<script>
(function () {
  const printMetadata = ${metadata};
  window.addEventListener('afterprint', function () {
    if (window.opener && !window.opener.closed) {
      window.opener.postMessage(Object.assign({ type: 'nkw-discipline-print-complete' }, printMetadata), '*');
    }
  });
})();
</script>
</body></html>`;
}

function formatThaiDateLong(ds) {
  if (!ds) return '...............';
  const p = ds.split('-');
  if (p.length !== 3) return ds;
  const months = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];
  return `${parseInt(p[2])} ${months[parseInt(p[1]) - 1]} ${parseInt(p[0]) + 543}`;
}

function disciplineEscapeHtml(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function sanitizeDisciplineDocumentData(caseData, config, advisor) {
  const safeCase = JSON.parse(JSON.stringify(caseData || {}));
  [
    'case_id','student_id','student_name','classroom_id','classroom_name','stage','status','trigger',
    'warn_doc_no','bond_doc_no','activity_doc_no','parent_doc_no','parent_name','parent_relation',
    'activity_result','warn_date','bond_date','bond_start','bond_end','activity_date','parent_date',
    'meeting_date','meeting_time','meeting_location','notes','created_at','signed_director_at',
    'signed_parent_at','signed_student_at','signed_witness1_at','signed_witness2_at'
  ].forEach(key => { safeCase[key] = disciplineEscapeHtml(safeCase[key]); });
  if (safeCase.activity_data) {
    safeCase.activity_data.start = disciplineEscapeHtml(safeCase.activity_data.start);
    safeCase.activity_data.end = disciplineEscapeHtml(safeCase.activity_data.end);
    safeCase.activity_data.items = (safeCase.activity_data.items || []).map(item => Object.assign({}, item, {
      id: disciplineEscapeHtml(item.id), label: disciplineEscapeHtml(item.label),
      unit: disciplineEscapeHtml(item.unit), target: disciplineEscapeHtml(item.target),
      progress: disciplineEscapeHtml(item.progress)
    }));
  }
  const safeConfig = {};
  Object.keys(config || {}).forEach(key => {
    const value = config[key];
    safeConfig[key] = key === 'warnings' ? value : disciplineEscapeHtml(value);
  });
  const safeAdvisor = {
    name: disciplineEscapeHtml(advisor && advisor.name),
    phone: disciplineEscapeHtml(advisor && advisor.phone)
  };
  return { caseData: safeCase, config: safeConfig, advisor: safeAdvisor };
}

function getCategoryLabelFromTrigger(trigger, triggerData) {
  const value = String((triggerData && triggerData.category) || trigger || '').toLowerCase();
  if (value === 'yellow' || value.indexOf('yellow') === 0) return 'ห่วงใย';
  if (value === 'red' || value.indexOf('red') === 0 || value === 'fail_activity') return 'ใกล้ชิด';
  return '';
}

function getCategoryColorFromTrigger(trigger, triggerData) {
  return getCategoryLabelFromTrigger(trigger, triggerData) === 'ห่วงใย' ? '#eab308' : '#ef4444';
}

function formatSignatureDate(timestamp) {
  if (!timestamp) return '......... / ......... / .........';
  const date = new Date(timestamp);
  if (isNaN(date.getTime())) return '......... / ......... / .........';
  const timezone = Session.getScriptTimeZone();
  const day = Number(Utilities.formatDate(date, timezone, 'd'));
  const month = Number(Utilities.formatDate(date, timezone, 'M'));
  const year = Number(Utilities.formatDate(date, timezone, 'yyyy')) + 543;
  return `${day} / ${month} / ${year}`;
}

function getGarudaHtml(cfg) {
  const url = String((cfg && cfg.garuda_url) || DISCIPLINE_GARUDA_URL).trim();
  if (!/^https:\/\//i.test(url)) throw new Error('URL รูปครุฑต้องเป็น HTTPS');
  return `<img src="${url}" alt="ตราครุฑ" width="60" height="60" loading="eager" decoding="sync" onerror="this.style.display='none';">`;
}

function buildViolationsHtml(cData) {
  let breakdown = cData.trigger_data && cData.trigger_data.breakdown;
  // For pre-migration cases, read the real assessment from the week the case was created.
  if ((!breakdown || typeof breakdown !== 'object') && cData.created_at) {
    try {
      const timeZone = Session.getScriptTimeZone();
      const localDate = Utilities.formatDate(new Date(cData.created_at), timeZone, 'yyyy-MM-dd');
      const monday = new Date(localDate + 'T00:00:00Z');
      const weekday = monday.getUTCDay();
      monday.setUTCDate(monday.getUTCDate() - (weekday === 0 ? 6 : weekday - 1));
      const friday = new Date(monday);
      friday.setUTCDate(friday.getUTCDate() + 4);
      const weekStart = Utilities.formatDate(monday, timeZone, 'yyyy-MM-dd');
      const weekEnd = Utilities.formatDate(friday, timeZone, 'yyyy-MM-dd');
      const assessment = getStudentRiskAssessment(weekStart, weekEnd, '', '');
      if (assessment.status === 'success') {
        const student = (assessment.data.students || []).find(item => String(item.id) === String(cData.student_id));
        if (student) breakdown = student.breakdown;
      }
    } catch (error) {
      Logger.log('buildViolationsHtml historical assessment error: ' + error.message);
    }
  }
  if (!breakdown || typeof breakdown !== 'object') {
    throw new Error('ไม่พบ breakdown จริงของสัปดาห์ที่สร้างเคส จึงยังไม่สามารถออกเอกสารทัณฑ์บนได้');
  }
  const items = [];
  const metrics = [
    { key: 'attendance', label: 'เข้าแถวไม่ครบ' },
    { key: 'cleaning', label: 'ไม่ทำเวร/หนีเวร' },
    { key: 'subject', label: 'ขาดเรียน/เข้าเรียนไม่ครบ' }
  ];
  metrics.forEach(metric => {
    const data = breakdown[metric.key];
    if (!data || !Number.isFinite(Number(data.points)) || !Number.isFinite(Number(data.total))) return;
    const missed = Math.max(0, Math.round(Number(data.total) - Number(data.points)));
    const score = Number.isFinite(Number(data.score)) ? Number(data.score) : (Number(data.total) ? Math.round(Number(data.points) / Number(data.total) * 100) : 100);
    if (missed > 0 && score < 85) items.push(`${metric.label} <strong>${missed} ครั้ง</strong> (${score}%)`);
  });
  if (!items.length) return '<div class="check-line">ไม่พบรายการต่ำกว่าเกณฑ์จาก breakdown ที่บันทึกไว้</div>';
  return items.map(item => `<div class="check-line">☑ ${item}</div>`).join('');
}

function buildWarningDocHtml(cData, cfg, advisor) {
  const categoryLabel = getCategoryLabelFromTrigger(cData.trigger, cData.trigger_data);
  if (!categoryLabel) throw new Error('ไม่พบกลุ่มประเมินจริงสำหรับหนังสือตักเตือน');
  const categoryColor = getCategoryColorFromTrigger(cData.trigger, cData.trigger_data);
  const { docNo, docDate } = ensureDocNo(cData.case_id, 'warning');
  const body = `
    <div class="garuda-wrap">${getGarudaHtml(cfg)}</div>
    <div class="header">${cfg.school_name}</div>
    <div class="header-sub">${cfg.school_name}</div>
    <div class="meta"><div>ที่ ${docNo}</div><div>${formatThaiDateLong(docDate)}</div></div>
    <div><strong>เรื่อง</strong> ตักเตือนเรื่องความประพฤติ</div>
    <div><strong>เรียน</strong> ${cData.student_name} ชั้น ${cData.classroom_name} เลขประจำตัว ${cData.student_id}</div>
    <div class="indent">ด้วย${cfg.school_name} ได้ติดตามพฤติกรรมของนักเรียนอย่างต่อเนื่อง พบว่านักเรียนมีผลการประเมินความประพฤติอยู่ในกลุ่ม <strong style="color:${categoryColor};">“${categoryLabel}”</strong> ซึ่งต่ำกว่าเกณฑ์ที่โรงเรียนกำหนด</div>
    <div class="indent">โรงเรียนจึงขอว่ากล่าวตักเตือนนักเรียนให้ปรับปรุงพฤติกรรมด้านการเข้าเรียน การร่วมกิจกรรม และการทำเวรให้ดีขึ้น หากยังฝ่าฝืนอีก โรงเรียนจะดำเนินการตามระเบียบกระทรวงศึกษาธิการ ว่าด้วยการลงโทษนักเรียนและนักศึกษา พ.ศ. 2548 ต่อไป</div>
    <div class="indent">จึงเรียนมาเพื่อโปรดทราบและปฏิบัติตาม</div>
    <div class="signature"><div>ขอแสดงความนับถือ</div><div style="margin-top:40px;">ลงชื่อ .....................................................</div><div>(${advisor.name})</div><div>ครูที่ปรึกษา</div></div>`;
  return buildDocWrap('หนังสือตักเตือน', body, cData.case_id, 'warning');
}

function buildBondDocHtml(cData, cfg, advisor) {
  const bondStart = cData.bond_start || '';
  const bondEnd = cData.bond_end || '';
  const parentName = cData.parent_name || '.....................................................';
  const parentRel = cData.parent_relation || '................................';
  const categoryLabel = getCategoryLabelFromTrigger(cData.trigger, cData.trigger_data);
  if (!categoryLabel) throw new Error('ไม่พบกลุ่มประเมินจริงสำหรับเอกสารทัณฑ์บน');
  const violationsHtml = buildViolationsHtml(cData);
  const { docNo, docDate } = ensureDocNo(cData.case_id, 'bond');

  const page1 = `
    <div class="garuda-wrap">${getGarudaHtml(cfg)}</div>
    <div class="header" style="font-size:20pt;margin:10px 0;">บันทึกทัณฑ์บน</div>
    <div class="meta"><div>ที่ ${docNo}</div><div>${formatThaiDateLong(docDate)}</div></div>
    <div><strong>เรื่อง</strong> ทำทัณฑ์บนนักเรียน</div>
    <div><strong>เรียน</strong> ผู้อำนวยการ${cfg.school_name}</div>
    <div style="margin:12px 0;"><strong>1. ข้อมูลนักเรียน</strong><div style="padding-left:1cm;margin-top:4px;"><div>ชื่อ-สกุล: <strong>${cData.student_name}</strong></div><div>ชั้น: ${cData.classroom_name} &nbsp;&nbsp; เลขประจำตัว: ${cData.student_id}</div><div>ครูที่ปรึกษา: ${advisor.name}</div></div></div>
    <div style="margin:12px 0;"><strong>2. พฤติกรรมที่กระทำผิด</strong><div style="padding-left:1cm;margin-top:4px;">${violationsHtml}<div class="check-line">ผลการประเมินอยู่ในกลุ่ม “${categoryLabel}”</div></div></div>
    <div style="margin:12px 0;"><strong>3. เงื่อนไขการทำทัณฑ์บน</strong><div style="padding-left:1cm;margin-top:4px;">
      <div>3.1 ระยะเวลาทัณฑ์บน: ตั้งแต่ ${bondStart ? formatThaiDateLong(bondStart) : '................................'} ถึง ${bondEnd ? formatThaiDateLong(bondEnd) : '................................'}</div>
      <div>3.2 ระหว่างระยะเวลาทัณฑ์บน นักเรียนต้อง:</div>
      <div class="check-line">(1) ไม่กระทำความผิดซ้ำในลักษณะเดียวกัน</div><div class="check-line">(2) เข้าเรียนและร่วมกิจกรรมอย่างสม่ำเสมอ</div><div class="check-line">(3) รายงานตัวต่อครูที่ปรึกษาทุกสัปดาห์</div><div class="check-line">(4) ปฏิบัติตามคำสั่งของโรงเรียนอย่างเคร่งครัด</div>
      <div style="margin-top:6px;">3.3 หากฝ่าฝืนหรือไม่ปฏิบัติตามเงื่อนไข โรงเรียนจะพิจารณาลงโทษตามระเบียบกระทรวงศึกษาธิการ ว่าด้วยการลงโทษนักเรียนและนักศึกษา พ.ศ. 2548 ต่อไป</div>
    </div></div>
    <div class="indent">จึงเรียนมาเพื่อโปรดพิจารณาลงนามอนุมัติ</div>
    <div style="text-align:right;margin-top:30px;"><div>ลงชื่อ ..................................................... พยาน</div><div>(${advisor.name})</div><div>ครูที่ปรึกษา</div><div style="margin-top:6px;">วันที่ ${formatSignatureDate(cData.signed_witness1_at)}</div></div>
    <div class="page-break"></div>`;

  const page2 = `
    <div class="header" style="font-size:20pt;">บันทึกทัณฑ์บน (หน้า 2)</div>
    <div style="text-align:center;margin:30px 0;font-weight:bold;font-size:15pt;">คู่สัญญาทั้ง 3 ฝ่ายได้อ่านและเข้าใจข้อความในบันทึกทัณฑ์บนนี้โดยตลอดแล้ว<br>จึงได้ลงลายมือชื่อไว้เป็นหลักฐาน</div>
    <div style="margin:20px 0;"><div style="text-align:center;font-weight:bold;margin-bottom:12px;">─── ฝ่ายที่ 1: ผู้ทำทัณฑ์บน (นักเรียน) ───</div><div style="text-align:center;"><div>ข้าพเจ้าขอยอมรับว่ากระทำความผิดตามที่ระบุไว้จริง และขอสัญญาว่าจะไม่กระทำความผิดซ้ำอีก</div><div style="margin-top:25px;">ลงชื่อ ..................................................... ผู้ทำทัณฑ์บน</div><div>( ${cData.student_name} )</div><div>วันที่ ${formatSignatureDate(cData.signed_student_at)}</div></div></div>
    <div style="margin:20px 0;"><div style="text-align:center;font-weight:bold;margin-bottom:12px;">─── ฝ่ายที่ 2: ผู้รับรองการทำทัณฑ์บน (ผู้ปกครอง) ───</div><div style="text-align:center;"><div>ข้าพเจ้าได้รับทราบถึงความผิดของบุตรหลานแล้ว และขอรับรองว่าจะดูแลให้ปฏิบัติตามเงื่อนไขของทัณฑ์บน</div><div style="margin-top:25px;">ลงชื่อ ..................................................... ผู้ปกครอง</div><div>( ${parentName} )</div><div>เกี่ยวข้องเป็น ${parentRel}</div><div>วันที่ ${formatSignatureDate(cData.signed_parent_at)}</div></div></div>
    <div style="margin:20px 0;"><div style="text-align:center;font-weight:bold;margin-bottom:12px;">─── ฝ่ายที่ 3: ผู้ลงโทษ (ผู้บริหารสถานศึกษา) ───</div><div style="text-align:center;"><div>อาศัยอำนาจตามระเบียบกระทรวงศึกษาธิการ ว่าด้วยการลงโทษนักเรียนและนักศึกษา พ.ศ. 2548 จึงสั่งให้ทำทัณฑ์บนไว้เป็นหนังสือ</div><div style="margin-top:25px;">ลงชื่อ ..................................................... ผู้สั่งลงโทษ</div><div>( ${cfg.director_name} )</div><div>${cfg.director_position}</div><div>${cfg.school_name}</div><div>วันที่ ${formatSignatureDate(cData.signed_director_at)}</div></div></div>
    <div style="margin:20px 0;"><div style="text-align:center;font-weight:bold;margin-bottom:12px;">─── พยานในการรับทราบ ───</div><div class="sig-row"><div><div style="margin-top:25px;">ลงชื่อ ..................................................... พยาน</div><div>( ${advisor.name} )</div><div>ครูที่ปรึกษา</div><div>วันที่ ${formatSignatureDate(cData.signed_witness1_at)}</div></div><div><div style="margin-top:25px;">ลงชื่อ ..................................................... พยาน</div><div>( ${cfg.discipline_teacher} )</div><div>หัวหน้างานกิจการนักเรียน</div><div>วันที่ ${formatSignatureDate(cData.signed_witness2_at)}</div></div></div></div>`;
  return buildDocWrap('บันทึกทัณฑ์บน', page1 + page2, cData.case_id, 'bond');
}

function buildActivityDocHtml(cData, cfg, advisor) {
  const { docNo, docDate } = ensureDocNo(cData.case_id, 'activity');
  const act = cData.activity_data || { items: [], start: '', end: '' };
  const activitiesHtml = (act.items || []).map(item =>
    `<div class="check-line">☐ ${item.label} จำนวน ${item.target} ${item.unit}</div>`
  ).join('');
  const body = `
    <div class="garuda-wrap">${getGarudaHtml(cfg)}</div>
    <div class="header">คำสั่ง${cfg.school_name}</div><div class="header">ที่ ${docNo}</div><div class="header">เรื่อง มอบหมายกิจกรรมเพื่อปรับเปลี่ยนพฤติกรรม</div>
    <div style="text-align:center;margin:20px 0;">─────────────</div>
    <div class="indent">อาศัยอำนาจตามระเบียบกระทรวงศึกษาธิการ ว่าด้วยการลงโทษนักเรียนและนักศึกษา พ.ศ. 2548 และบันทึกทัณฑ์บนที่ ${cData.bond_doc_no || '................................'} ลงวันที่ ${cData.bond_date ? formatThaiDateLong(cData.bond_date) : '................................'} ประกอบกับนักเรียน <strong>${cData.student_name}</strong> ชั้น ${cData.classroom_name} เลขประจำตัว ${cData.student_id} ได้ฝ่าฝืนเงื่อนไขของทัณฑ์บน</div>
    <div class="indent">โรงเรียนจึงมีคำสั่งให้นักเรียนดำเนินกิจกรรมเพื่อปรับเปลี่ยนพฤติกรรม ดังต่อไปนี้:</div>
    <div style="margin:12px 0;padding-left:1cm;">${activitiesHtml || 'ไม่พบรายการกิจกรรมที่กำหนดไว้ในข้อมูลเคส'}</div>
    <div class="indent"><strong>ระยะเวลาดำเนินการ:</strong> ตั้งแต่วันที่ ${act.start ? formatThaiDateLong(act.start) : '................................'} ถึง ${act.end ? formatThaiDateLong(act.end) : '................................'}</div>
    <div class="indent"><strong>ผู้ติดตามผล:</strong> ${advisor.name} ครูที่ปรึกษา</div>
    <div class="indent">ขอให้นักเรียนและผู้ปกครองให้ความร่วมมือในการดำเนินกิจกรรมตามที่กำหนด หากนักเรียนไม่ดำเนินการหรือดำเนินการไม่ครบถ้วน โรงเรียนจะพิจารณาแจ้งผู้ปกครองและดำเนินการตามระเบียบวินัยต่อไป</div>
    <div style="text-align:right;margin-top:40px;"><div>สั่ง ณ วันที่ ${formatThaiDateLong(docDate)}</div><div style="margin-top:40px;">ลงชื่อ ..................................................... ผู้สั่ง</div><div>( ${cfg.director_name} )</div><div>${cfg.director_position}</div><div>${cfg.school_name}</div></div>
    <div class="sig-row"><div><div style="font-weight:bold;">รับทราบ</div><div style="margin-top:30px;">ลงชื่อ ..................................................... นักเรียน</div><div>( ${cData.student_name} )</div><div>วันที่ ${formatSignatureDate(cData.signed_student_at)}</div></div><div><div style="font-weight:bold;">รับทราบ</div><div style="margin-top:30px;">ลงชื่อ ..................................................... ผู้ปกครอง</div><div>( ${cData.parent_name || '.....................................................'} )</div><div>วันที่ ${formatSignatureDate(cData.signed_parent_at)}</div></div></div>`;
  return buildDocWrap('คำสั่งมอบหมายกิจกรรม', body, cData.case_id, 'activity');
}

function buildParentNoticeDocHtml(cData, cfg, advisor) {
  const { docNo, docDate } = ensureDocNo(cData.case_id, 'parent');
  const meetingDate = cData.meeting_date ? formatThaiDateLong(cData.meeting_date) : '...........................................';
  const meetingTime = cData.meeting_time || '................................';
  const meetingLoc = cData.meeting_location || '...........................................';
  const categoryLabel = getCategoryLabelFromTrigger(cData.trigger, cData.trigger_data);
  if (!categoryLabel) throw new Error('ไม่พบกลุ่มประเมินจริงสำหรับหนังสือเชิญผู้ปกครอง');
  const phoneText = advisor.phone ? ` โทร. ${advisor.phone}` : '';
  const body = `
    <div class="garuda-wrap">${getGarudaHtml(cfg)}</div><div class="header">${cfg.school_name}</div><div class="header-sub">${cfg.school_name}</div>
    <div class="meta"><div>ที่ ${docNo}</div><div>${formatThaiDateLong(docDate)}</div></div>
    <div><strong>เรื่อง</strong> ขอเชิญประชุมผู้ปกครองเพื่อหาแนวทางแก้ไขพฤติกรรมนักเรียน</div><div><strong>เรียน</strong> ผู้ปกครองของ ${cData.student_name}</div>
    <div class="indent">ด้วย${cfg.school_name} ได้ติดตามพฤติกรรมของ <strong>${cData.student_name}</strong> ชั้น ${cData.classroom_name} เลขประจำตัว ${cData.student_id} อย่างต่อเนื่อง พบว่านักเรียนมีผลการประเมินความประพฤติอยู่ในกลุ่ม <strong>“${categoryLabel}”</strong> และได้ผ่านกระบวนการตามระเบียบกระทรวงศึกษาธิการ ว่าด้วยการลงโทษนักเรียนและนักศึกษา พ.ศ. 2548 แล้ว คือ</div>
    <div style="margin:8px 0;padding-left:1cm;"><div>1. หนังสือตักเตือน เลขที่ ${cData.warn_doc_no || '................................'} ลงวันที่ ${cData.warn_date ? formatThaiDateLong(cData.warn_date) : '................................'}</div><div>2. บันทึกทัณฑ์บน เลขที่ ${cData.bond_doc_no || '................................'} ลงวันที่ ${cData.bond_date ? formatThaiDateLong(cData.bond_date) : '................................'}</div><div>3. คำสั่งมอบหมายกิจกรรม เลขที่ ${cData.activity_doc_no || '................................'} ลงวันที่ ${cData.activity_date ? formatThaiDateLong(cData.activity_date) : '................................'}</div></div>
    <div class="indent">แต่นักเรียนยังไม่สามารถดำเนินการตามเงื่อนไขได้สำเร็จ โรงเรียนจึงขอเชิญท่านเข้าร่วมประชุมเพื่อหาแนวทางแก้ไขร่วมกัน</div>
    <div style="margin:12px 0;padding-left:1cm;"><div><strong>วันประชุม:</strong> ${meetingDate}</div><div><strong>เวลา:</strong> ${meetingTime}</div><div><strong>สถานที่:</strong> ${meetingLoc}</div></div>
    <div class="indent">จึงเรียนมาเพื่อโปรดเข้าร่วมประชุมตามวันและเวลาดังกล่าว หากไม่สามารถมาได้ กรุณาติดต่อครูที่ปรึกษา ${advisor.name}${phoneText} เพื่อนัดหมายใหม่</div><div class="indent">ขอแสดงความนับถือ</div>
    <div style="text-align:right;margin-top:40px;"><div>ลงชื่อ .....................................................</div><div>( ${cfg.director_name} )</div><div>${cfg.director_position}</div><div>${cfg.school_name}</div></div>
    <div class="cut-line"></div><div style="text-align:center;font-weight:bold;font-size:18pt;margin:20px 0;">ใบตอบรับ</div>
    <div style="text-align:right;margin-bottom:10px;">เขียนที่ .........................................<br>วันที่ ......... / ......... / .........</div>
    <div style="margin:8px 0;"><strong>เรื่อง</strong> ตอบรับหนังสือเชิญประชุม</div><div style="margin:8px 0;"><strong>เรียน</strong> ผู้อำนวยการ${cfg.school_name}</div>
    <div class="indent">ตามที่โรงเรียนได้มีหนังสือที่ ${docNo} ลงวันที่ ${formatThaiDateLong(docDate)} ขอเชิญประชุมเรื่องพฤติกรรมของ ${cData.student_name} นั้น</div>
    <div class="indent">ข้าพเจ้า ..................................................... ผู้ปกครองของ ${cData.student_name} ได้รับหนังสือแล้ว และ</div>
    <div style="margin:8px 0;padding-left:1cm;"><div>☐ ยินดีเข้าร่วมประชุมตามวันและเวลาดังกล่าว</div><div>☐ ไม่สามารถเข้าร่วมได้ เพราะ ............................................ จะติดต่อนัดหมายใหม่</div></div>
    <div style="text-align:right;margin-top:30px;"><div>ลงชื่อ ..................................................... ผู้ปกครอง</div><div>( ..................................................... )</div><div>วันที่ ......... / ......... / .........</div></div>`;
  return buildDocWrap('หนังสือแจ้งผู้ปกครอง', body, cData.case_id, 'parent');
}

function ensureDocNo(caseId, docType) {
  const sheet = getDisciplineSheet('CASES');
  const data = sheet.getDataRange().getValues();
  const idx = data.slice(1).findIndex(row => String(row[0]) === String(caseId)) + 2;
  if (idx <= 1) return { docNo: '', docDate: '' };
  const row = data[idx - 1];
  const colMap = { warning: [8, 9], bond: [10, 11], activity: [14, 15], parent: [16, 17] };
  const positions = colMap[docType];
  if (!positions) throw new Error('docType ไม่ถูกต้อง');
  const [noIndex, dateIndex] = positions;
  let docNo = String(row[noIndex] || '').trim();
  let docDate = formatDateStr(row[dateIndex]);
  if (!docNo) {
    // Deliberately fail closed if the counter cannot be reserved; never generate a random official number.
    docNo = reserveDocNumber();
    docDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    sheet.getRange(idx, noIndex + 1).setValue(docNo);
    sheet.getRange(idx, dateIndex + 1).setValue(docDate);
  }
  return { docNo: disciplineEscapeHtml(docNo), docDate: disciplineEscapeHtml(docDate) };
}

// ---------- Master doc getter ----------

function getDisciplineDocHtml(caseId, docType) {
  try {
    const res = getDisciplineCaseDetail(caseId);
    if (res.status !== 'success') throw new Error(res.message);
    if (!['warning', 'bond', 'activity', 'parent'].includes(docType)) throw new Error('docType ไม่ถูกต้อง');
    const rawCase = res.data.case;
    const config = res.data.config || {};
    const advisor = res.data.advisor || {};
    if (docType === 'parent' && rawCase.stage !== '4') throw new Error('หนังสือผู้ปกครองใช้ได้เฉพาะ Stage 4 เท่านั้น');
    if (!config.school_name) throw new Error('ยังไม่ได้ตั้งชื่อโรงเรียนจริงในชีต Config');
    if (!advisor.name) throw new Error('ไม่พบชื่อครูที่ปรึกษาจริงของห้องนี้ในชีต Teachers');
    if (['bond', 'activity', 'parent'].includes(docType) && (!config.director_name || !config.director_position)) {
      throw new Error('กรุณากำหนดชื่อและตำแหน่งผู้อำนวยการจริงในชีต Config ก่อนพิมพ์');
    }
    if (docType === 'bond' && !config.discipline_teacher) throw new Error('กรุณากำหนดชื่อหัวหน้างานกิจการนักเรียนจริงในชีต Config ก่อนพิมพ์');
    const safeData = sanitizeDisciplineDocumentData(rawCase, config, advisor);
    const cData = safeData.caseData;
    const cfg = safeData.config;
    const safeAdvisor = safeData.advisor;
    let html;
    if (docType === 'warning') html = buildWarningDocHtml(cData, cfg, safeAdvisor);
    else if (docType === 'bond') html = buildBondDocHtml(cData, cfg, safeAdvisor);
    else if (docType === 'activity') html = buildActivityDocHtml(cData, cfg, safeAdvisor);
    else if (docType === 'parent') html = buildParentNoticeDocHtml(cData, cfg, safeAdvisor);
    else throw new Error('docType ไม่ถูกต้อง');
    // Printed status is recorded only after the document window signals that its print dialog completed.
    return { status: 'success', data: { html } };
  } catch (err) {
    Logger.log('getDisciplineDocHtml error: ' + err.message);
    return { status: 'error', message: err.toString() };
  }
}

/**
 * บันทึกข้อมูลนัดประชุมผู้ปกครอง (Stage 4)
 * @param {string} caseId - รหัสเคส
 * @param {string} dateStr - yyyy-MM-dd
 * @param {string} timeStr - เช่น "09:00"
 * @param {string} location - สถานที่ประชุม
 */
/** Save Stage 4 appointment metadata; no guardian contact is sent. */
function saveMeetingMeta(caseId, dateStr, timeStr, location) {
  try {
    if (!caseId) return { status: 'error', message: 'ไม่พบ caseId' };

    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    const idx = data.slice(1).findIndex(r => String(r[0]) === caseId) + 2;
    if (idx <= 1) return { status: 'error', message: 'ไม่พบเคสในระบบ' };

    const currentStage = String(data[idx - 1][5] || '');
    if (currentStage !== '4') {
      return { status: 'error', message: 'การนัดประชุมใช้ได้เฉพาะ Stage 4 (แจ้งผู้ปกครอง) เท่านั้น' };
    }

    const safeDate = String(dateStr || '').trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(safeDate)) {
      return { status: 'error', message: 'วันที่ประชุมไม่ถูกต้อง' };
    }
    const parsedDate = new Date(safeDate + 'T00:00:00Z');
    if (isNaN(parsedDate.getTime()) || parsedDate.toISOString().substring(0, 10) !== safeDate) {
      return { status: 'error', message: 'วันที่ประชุมไม่ถูกต้อง' };
    }
    const safeTime = String(timeStr || '').trim().substring(0, 20);
    const safeLoc = String(location || '').trim().substring(0, 200);
    if (!safeTime) return { status: 'error', message: 'กรุณาระบุเวลาประชุม' };
    if (!safeLoc) return { status: 'error', message: 'กรุณาระบุสถานที่ประชุม' };

    // อัปเดตคอลัมน์ 23 (meeting_date), 24 (meeting_time), 25 (meeting_location), 37 (updated_at)
    sheet.getRange(idx, 23).setValue(safeDate);
    sheet.getRange(idx, 24).setValue(safeTime);
    sheet.getRange(idx, 25).setValue(safeLoc);
    sheet.getRange(idx, 37).setValue(new Date().toISOString());

    addDisciplineLog(
      caseId, 'meeting_scheduled', '', '',
      'system',
      `นัดประชุม ${safeDate || '?'} ${safeTime || ''} @ ${safeLoc || '?'}`
    );

    return {
      status: 'success',
      message: 'บันทึกการนัดประชุมสำเร็จ',
      data: { meeting_date: safeDate, meeting_time: safeTime, meeting_location: safeLoc }
    };
  } catch (err) {
    Logger.log('saveMeetingMeta error: ' + err.message);
    return { status: 'error', message: err.toString() };
  }
}
