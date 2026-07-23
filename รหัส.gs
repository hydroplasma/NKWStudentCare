// --- ส่วนที่ 1: ปรับ Routing ใน doGet ---
function doGet(e) {
  try {
    const configRes = getConfig();
    const site_config = configRes.data ? configRes.data.site_config : {};
    const systemName = site_config['systemName'] || 'Student Attendance System';

    // หน้า stats เดิม
    if (e.parameter.page == 'stats') {
      return HtmlService.createTemplateFromFile('stats').evaluate()
        .setTitle(systemName + " - สถิติ")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
    }

    // หน้า stats2 ใหม่ (ที่คุณต้องการเพิ่ม)
    if (e.parameter.page == 'stats2') {
      return HtmlService.createTemplateFromFile('stats2').evaluate()
        .setTitle(systemName + " - วิเคราะห์รายชั้น")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // หน้าหลัก
    return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle(systemName)
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

// Headers for each sheet
const HEADERS = {
  [SHEET_NAMES.CONFIG]: ['key', 'value'],
  [SHEET_NAMES.STATUS_LIST]: ['label', 'color', 'hexColor', 'status'],
  [SHEET_NAMES.TEACHERS]: ['id', 'name', 'username', 'password', 'status', 'classroom_ids'],
  [SHEET_NAMES.PERIODS]: ['id', 'name', 'status'],
  [SHEET_NAMES.CLASSROOMS]: ['id', 'name', 'status'],
  [SHEET_NAMES.STUDENTS]: ['classroom_id', 'student_id', 'no', 'name', 'status'],
  [SHEET_NAMES.ATTENDANCE]: ['date', 'period', 'room_id', 'teacher', 'students', 'updated_at']
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

      if (!Array.isArray(students) || students.some(s => !s.student_id || !s.status)) {
        return createResponse('error', 'Invalid students data', null, { code: 'INVALID_STUDENTS' });
      }

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
      const teachersData = [
        [1, 'ครูสมชาย ใจดี', 'somchai', '1234', 1, 'm1-1,m2-1']
      ];
      teachersSheet.getRange(2, 1, teachersData.length, teachersData[0].length).setValues(teachersData);

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

const CLEANING_SHEET_NAMES = {
  CRITERIA: 'CleaningCriteria',
  ATTENDANCE: 'CleaningAttendance',
  SCORES: 'CleaningScores'
};

const CLEANING_HEADERS = {
  [CLEANING_SHEET_NAMES.CRITERIA]: ['id', 'name', 'max_score'],
  [CLEANING_SHEET_NAMES.ATTENDANCE]: ['date', 'classroom_id', 'checker_id', 'student_data', 'updated_at'],
  [CLEANING_SHEET_NAMES.SCORES]: ['date', 'classroom_id', 'checker_id', 'score_json', 'total_score', 'updated_at']
};

function getCleaningConfig(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const classroomsResponse = getClassrooms([], spreadsheetMode);
    if (classroomsResponse.status !== 'success') throw new Error('Failed to load classrooms');
    
    const studentsResponse = getAllStudents(spreadsheetMode);
    if (studentsResponse.status !== 'success') throw new Error('Failed to load students');
    
    const criteriaSheet = getOrCreateSheet(CLEANING_SHEET_NAMES.CRITERIA, spreadsheetMode);
    const criteriaData = criteriaSheet.getDataRange().getValues();
    const criteria = criteriaData.length <= 1 ? [] : criteriaData.slice(1).map(row => ({
      id: row[0],
      name: row[1],
      max_score: row[2]
    }));
    
    return createResponse('success', 'Config loaded successfully', {
      classrooms: classroomsResponse.data,
      students: studentsResponse.data,
      criteria: criteria
    });
  } catch (error) {
    return createResponse('error', 'Failed to load config: ' + error.message, null, error);
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
    const attendanceSheet = getOrCreateSheet(CLEANING_SHEET_NAMES.ATTENDANCE, spreadsheetMode);
    const attendanceData = attendanceSheet.getDataRange().getValues();
    const scoresSheet = getOrCreateSheet(CLEANING_SHEET_NAMES.SCORES, spreadsheetMode);
    const scoresData = scoresSheet.getDataRange().getValues();

    // กรองข้อมูลซ้ำ
    const cleanAttendance = getCleanedCleaningRows(attendanceData);
    const cleanScores = getCleanedCleaningRows(scoresData);

    const filteredAttendance = cleanAttendance.filter(row => {
      const rowDate = String(row[0]).trim();
      const rowClassroom = String(row[1]).trim();
      const dateMatch = rowDate >= startDate && rowDate <= endDate;
      const classroomMatch = !classroom_id || rowClassroom === classroom_id;
      return dateMatch && classroomMatch;
    });

    const filteredScores = cleanScores.filter(row => {
      const rowDate = String(row[0]).trim();
      const rowClassroom = String(row[1]).trim();
      const dateMatch = rowDate >= startDate && rowDate <= endDate;
      const classroomMatch = !classroom_id || rowClassroom === classroom_id;
      return dateMatch && classroomMatch;
    });

    const studentStats = {};
    filteredAttendance.forEach(row => {
      const students = JSON.parse(row[3] || '[]');
      students.forEach(student => {
        const sid = student.student_id;
        if (!studentStats[sid]) {
          studentStats[sid] = { student_id: sid, name: student.name || sid, total: 0, done: 0, not_done: 0, escaped: 0 };
        }
        studentStats[sid].total++;
        if (student.status === 'ทำ') studentStats[sid].done++;
        else if (student.status === 'ไม่ทำ') studentStats[sid].not_done++;
        else if (student.status === 'หนีเวร') studentStats[sid].escaped++;
      });
    });

    const topStudents = Object.values(studentStats).sort((a, b) => b.done - a.done).slice(0, 5);

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

    const scoreDetails = filteredScores.map(row => ({
      date: row[0],
      classroom_id: row[1],
      checker: row[2],
      scores: JSON.parse(row[3] || '{}'),
      total_score: row[4]
    }));

    return createResponse('success', 'Stats retrieved successfully', {
      studentStats: Object.values(studentStats), topStudents, avgScore, scoreDetails
    });
  } catch (error) {
    return createResponse('error', 'Failed to get stats: ' + error.message, null, error);
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
