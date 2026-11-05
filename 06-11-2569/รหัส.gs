// ═══════════════════════════════════════════════════════════════
// STUDENT ATTENDANCE SYSTEM — Code.gs
// Version: Sprint 1-4 Complete (2025-11-03)
// ═══════════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════════
// 1. ROUTING & UTILITIES
// ═══════════════════════════════════════════════════════════════

function doGet(e) {
  try {
    const configRes = getConfig();
    const site_config = configRes.data ? configRes.data.site_config : {};
    const systemName = site_config['systemName'] || 'Student Attendance System';

    if (e.parameter.page === 'stats') {
      return HtmlService.createTemplateFromFile('stats').evaluate()
        .setTitle(systemName + " - วิเคราะห์รายชั้น")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    else if (e.parameter.page === 'cleaning') {
      return HtmlService.createTemplateFromFile('cleaning').evaluate()
        .setTitle(systemName + " - เขตพื้นที่")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    else if (e.parameter.page === 'subject_attendance') {
      return HtmlService.createTemplateFromFile('subject_attendance').evaluate()
        .setTitle(systemName + " - เช็คชื่อรายวิชา")
        .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle(systemName)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    return HtmlService.createHtmlOutput('เกิดข้อผิดพลาด: ' + error.message);
  }
}

function getScriptUrl() {
  return ScriptApp.getService().getUrl();
}

function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (error) {
    Logger.log('Error in include ' + filename + ': ' + error.message);
    return '';
  }
}

// ═══════════════════════════════════════════════════════════════
// 2. CONSTANTS
// ═══════════════════════════════════════════════════════════════

const SPREADSHEET_ID = '1BiNWBtHfg7wnZmLwbOSTJS1bfVO5Mnfb6LMGdQcoOLk';
const API_VERSION = '1.0';
const DEFAULT_SPREADSHEET_MODE = 'byId';

const SHEET_NAMES = {
  CONFIG: 'Config',
  STATUS_LIST: 'StatusList',
  TEACHERS: 'Teachers',
  PERIODS: 'Periods',
  CLASSROOMS: 'Classrooms',
  STUDENTS: 'Students',
  ATTENDANCE: 'Attendance'
};

const CLEANING_SHEET_NAMES = {
  ATTENDANCE: 'Cleaning_Attendance',
  SCORES: 'Cleaning_Scores',
  CRITERIA: 'Cleaning_Criteria'
};

const HEADERS = {
  [SHEET_NAMES.CONFIG]: ['key', 'value'],
  [SHEET_NAMES.STATUS_LIST]: ['label', 'color', 'hexColor', 'status'],
  [SHEET_NAMES.TEACHERS]: ['id', 'name', 'username', 'password', 'status', 'classroom_ids', 'imageUrl', 'phone', 'role'],
  [SHEET_NAMES.PERIODS]: ['id', 'name', 'status'],
  [SHEET_NAMES.CLASSROOMS]: ['id', 'name', 'status'],
  [SHEET_NAMES.STUDENTS]: ['classroom_id', 'student_id', 'no', 'name', 'status', 'password', 'role', 'council_classroom_ids'],
  [SHEET_NAMES.ATTENDANCE]: ['date', 'period', 'room_id', 'teacher', 'students', 'updated_at'],
  [CLEANING_SHEET_NAMES.ATTENDANCE]: ['Date', 'Classroom ID', 'Checker ID', 'Student Data', 'Updated At'],
  [CLEANING_SHEET_NAMES.SCORES]: ['Date', 'Classroom ID', 'Checker ID', 'Scores', 'Total Score', 'Updated At'],
  [CLEANING_SHEET_NAMES.CRITERIA]: ['ID', 'Name', 'Max Score']
};

const ROLE_LABELS = {
  director: '👔 ผู้บริหาร',
  discipline: '📋 งานกิจการนักเรียน',
  council: '🎓 สภานักเรียน',
  teacher: '👨‍🏫 ครู',
  student: '🎓 นักเรียน'
};

const PERMISSIONS = {
  director:   ['attendance','cleaning','subject','stats-att','stats-cln','stats-subj','summary-weekly','summary-monthly','discipline','approval','settings'],
  discipline: ['attendance','cleaning','subject','stats-att','stats-cln','stats-subj','summary-weekly','summary-monthly','discipline','settings'],
  council:    ['attendance','cleaning','stats-att','stats-cln','student-self'],
  teacher:    ['attendance','cleaning','subject','stats-att','stats-cln','stats-subj','summary-weekly','summary-monthly','discipline'],
  student:    ['student-self']
};

// ═══════════════════════════════════════════════════════════════
// 3. SPREADSHEET ACCESS
// ═══════════════════════════════════════════════════════════════

function getSpreadsheet(mode = DEFAULT_SPREADSHEET_MODE) {
  try {
    if (mode === 'byId') {
      if (!SPREADSHEET_ID) throw new Error('Spreadsheet ID is not set.');
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    } else if (mode === 'active') {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      if (!ss) throw new Error('No active spreadsheet found.');
      return ss;
    }
    throw new Error('Invalid Spreadsheet mode');
  } catch (error) {
    Logger.log('Error in getSpreadsheet: ' + error.message);
    throw error;
  }
}

function getOrCreateSheet(sheetName, mode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const ss = getSpreadsheet(mode);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      const headers = HEADERS[sheetName];
      if (headers) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      if (sheetName === SHEET_NAMES.ATTENDANCE) sheet.getRange('A:A').setNumberFormat('@');
    } else {
      const existingHeaders = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
      if (existingHeaders.filter(h => h).length === 0) {
        const headers = HEADERS[sheetName];
        if (headers) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      }
      if (sheetName === SHEET_NAMES.ATTENDANCE) sheet.getRange('A:A').setNumberFormat('@');
    }
    return sheet;
  } catch (error) {
    Logger.log('Error in getOrCreateSheet: ' + error.message);
    throw error;
  }
}

function createResponse(status, message, data = null, error = null) {
  const response = {
    status, message, data,
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

// ═══════════════════════════════════════════════════════════════
// 4. CACHE HELPERS (Timetable V3)
// ═══════════════════════════════════════════════════════════════

function getCachedTimetableData() {
  const cache = CacheService.getScriptCache();
  const key = 'TIMETABLE_V3';

  try {
    const cached = cache.get(key);
    if (cached) {
      Logger.log('✅ [Cache HIT] Timetable V3');
      return JSON.parse(cached);
    }
  } catch (e) { /* ignore */ }

  Logger.log('🔄 [Cache MISS] Loading Timetable V3...');
  const t0 = Date.now();
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName('Timetable');
  if (!sheet) return [];

  const range = sheet.getDataRange();
  const rawData = range.getValues();
  const displayData = range.getDisplayValues();
  const locale = (ss.getSpreadsheetLocale && ss.getSpreadsheetLocale()) || 'th_TH';

  Logger.log('📋 Locale: ' + locale + ', rows: ' + rawData.length + ' in ' + (Date.now() - t0) + 'ms');

  const data = rawData.map(function(row, rowIdx) {
    if (rowIdx === 0) return row;
    return row.map(function(cell, colIdx) {
      if (colIdx === 2) return normalizePeriod(cell, displayData[rowIdx][colIdx], locale);
      if (cell instanceof Date) {
        return Utilities.formatDate(cell, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      return cell;
    });
  });

  try {
    const json = JSON.stringify(data);
    if (json.length < 90000) cache.put(key, json, 300);
  } catch (e) { /* ignore */ }

  return data;
}

function normalizePeriod(cellValue, displayValue, locale) {
  const disp = String(displayValue || '').trim();
  if (disp) {
    const isDateLike = /^\d{1,2}[\/\.]\d{1,2}[\/\.]\d{2,4}$/.test(disp) ||
                       /^[A-Za-z]{3,}\s/.test(disp) ||
                       /^\d{1,2}\s[A-Za-z]{3}/.test(disp);
    if (!isDateLike) return disp;
  }
  if (cellValue instanceof Date) {
    const day = cellValue.getDate();
    const month = cellValue.getMonth() + 1;
    const isUS = /^en[_\-](US|CA)/i.test(locale || '');
    return isUS ? (month + '-' + day) : (day + '-' + month);
  }
  if (cellValue === null || cellValue === undefined || cellValue === '') return '';
  return String(cellValue).trim();
}

function clearTimetableCache() {
  const cache = CacheService.getScriptCache();
  ['TIMETABLE_V1', 'TIMETABLE_V2', 'TIMETABLE_V3'].forEach(k => cache.remove(k));
  return { status: 'success', message: 'ล้าง cache ทุก version แล้ว' };
}

function refreshTimetableCache() {
  clearTimetableCache();
  Logger.log('✅ Timetable cache cleared');
}

// ═══════════════════════════════════════════════════════════════
// 5. PERMISSION LAYER
// ═══════════════════════════════════════════════════════════════

function getTeacherRole(username) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.TEACHERS);
    const data = sheet.getDataRange().getValues();
    const row = data.slice(1).find(r =>
      String(r[2] || '').trim() === String(username).trim() &&
      String(r[4] || '') === '1'
    );
    if (!row) return null;
    const role = String(row[8] || '').trim().toLowerCase();
    return ['director','discipline','council','teacher'].includes(role) ? role : 'teacher';
  } catch (e) {
    Logger.log('getTeacherRole error: ' + e.message);
    return null;
  }
}

function hasPermission(role, permission) {
  return (PERMISSIONS[role] || []).includes(permission);
}

function requirePermission(username, roleList) {
  const role = getTeacherRole(username);
  if (!role) throw new Error('AUTH_FAILED: ไม่พบครู');
  if (!roleList.includes(role)) {
    throw new Error('PERMISSION_DENIED: role "' + role + '" ไม่มีสิทธิ์');
  }
  return role;
}

function getStudentCouncilInfo(studentId) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.STUDENTS);
    const data = sheet.getDataRange().getValues();
    const row = data.slice(1).find(r =>
      String(r[1] || '').trim() === String(studentId).trim() &&
      String(r[4] || '') === '1'
    );
    if (!row) return null;
    const role = String(row[6] || '').trim().toLowerCase();
    if (role !== 'council') return null;
    return {
      student_id: String(row[1]).trim(),
      name: String(row[3]).trim(),
      classroom_id: String(row[0]).trim(),
      council_classroom_ids: String(row[7] || '').split(',').map(s => s.trim()).filter(Boolean)
    };
  } catch (e) {
    Logger.log('getStudentCouncilInfo error: ' + e.message);
    return null;
  }
}

function getActorInfo(username) {
  const teacherRole = getTeacherRole(username);
  if (teacherRole) {
    const tRow = getOrCreateSheet(SHEET_NAMES.TEACHERS).getDataRange().getValues()
      .slice(1).find(r => String(r[2]).trim() === String(username).trim());
    return {
      role: teacherRole, type: 'teacher',
      name: String(tRow[1] || '').trim(),
      classroom_ids: tRow[5] ? String(tRow[5]).split(',').map(s => s.trim()) : []
    };
  }
  const council = getStudentCouncilInfo(username);
  if (council) {
    return {
      role: 'council', type: 'student_council',
      name: council.name,
      classroom_ids: council.council_classroom_ids
    };
  }
  return null;
}

// ═══════════════════════════════════════════════════════════════
// 6. AUTHENTICATION & LOGIN
// ═══════════════════════════════════════════════════════════════

function getLoginData(username, password, mode) {
  mode = mode || 'teacher';
  try {
    if (mode === 'student') return loginStudent(username, password);
    return loginTeacher(username, password);
  } catch (err) {
    Logger.log('getLoginData error: ' + err.message);
    return createResponse('error', 'เกิดข้อผิดพลาด: ' + err.message, null, err);
  }
}

function loginTeacher(username, password) {
  const tSheet = getOrCreateSheet(SHEET_NAMES.TEACHERS);
  const tData = tSheet.getDataRange().getValues();

  const teacher = tData.slice(1).find(row =>
    String(row[2] || '').trim() === String(username).trim() &&
    String(row[3] || '').trim() === String(password).trim() &&
    String(row[4] || '') === '1'
  );

  if (!teacher) {
    return createResponse('error', 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง', null, { code: 'AUTH_FAILED' });
  }

  const role = ['director','discipline','council','teacher'].includes(String(teacher[8]||'').toLowerCase())
    ? String(teacher[8]).toLowerCase() : 'teacher';

  const teacherData = {
    type: 'teacher',
    id: String(teacher[0] || '').trim(),
    name: String(teacher[1] || '').trim(),
    username: String(teacher[2] || '').trim(),
    classroom_ids: teacher[5] ? String(teacher[5]).split(',').map(s => s.trim()) : [],
    imageUrl: String(teacher[6] || '').trim(),
    phone: String(teacher[7] || '').trim(),
    role, roleLabel: ROLE_LABELS[role],
    permissions: PERMISSIONS[role] || []
  };

  const cData = getOrCreateSheet(SHEET_NAMES.CLASSROOMS).getDataRange().getValues();
  const sData = getOrCreateSheet(SHEET_NAMES.STUDENTS).getDataRange().getValues();
  const classrooms = cData.slice(1)
    .filter(r => String(r[2]) === '1' && (teacherData.classroom_ids.length === 0 || teacherData.classroom_ids.includes(String(r[0]).trim())))
    .map(r => ({ id: String(r[0]).trim(), name: String(r[1]).trim(), status: r[2] }));

  const studentsByClassroom = {};
  teacherData.classroom_ids.forEach(id => {
    studentsByClassroom[id] = sData.slice(1)
      .filter(r => String(r[0]).trim() === id && String(r[4]) === '1')
      .map(r => ({
        classroom_id: String(r[0]).trim(), student_id: String(r[1]).trim(),
        no: String(r[2]).trim(), name: String(r[3]).trim(), status: r[4]
      }));
  });

  const studentCounts = [];
  Object.keys(studentsByClassroom).forEach(rid => {
    studentCounts.push({ room_id: rid, count: studentsByClassroom[rid].length });
  });

  return createResponse('success', 'Login successful', { teacher: teacherData, classrooms, studentCounts });
}

function loginStudent(studentId, password) {
  try {
    const sData = getOrCreateSheet(SHEET_NAMES.STUDENTS).getDataRange().getValues();
    const row = sData.slice(1).find(r =>
      String(r[1] || '').trim() === String(studentId).trim() &&
      String(r[4] || '') === '1'
    );
    if (!row) {
      return createResponse('error', 'ไม่พบเลขประจำตัวนักเรียน', null, { code: 'AUTH_FAILED' });
    }

    const storedPwd = String(row[5] || '').trim() || String(row[1]).trim();
    if (String(password || '').trim() !== storedPwd) {
      return createResponse('error', 'รหัสผ่านไม่ถูกต้อง', null, { code: 'AUTH_FAILED' });
    }

    const role = String(row[6] || '').trim().toLowerCase();
    const isCouncil = role === 'council';
    const councilRooms = String(row[7] || '').split(',').map(s => s.trim()).filter(Boolean);

    const cData = getOrCreateSheet(SHEET_NAMES.CLASSROOMS).getDataRange().getValues();
    const classroomRow = cData.slice(1).find(r => String(r[0]).trim() === String(row[0]).trim());

    const studentData = {
      type: isCouncil ? 'student_council' : 'student',
      student_id: String(row[1]).trim(),
      name: String(row[3]).trim(),
      no: String(row[2]).trim(),
      classroom_id: String(row[0]).trim(),
      classroom_name: classroomRow ? String(classroomRow[1]).trim() : String(row[0]).trim(),
      role: isCouncil ? 'council' : 'student',
      roleLabel: isCouncil ? '🎓 สภานักเรียน' : '🎓 นักเรียน',
      classroom_ids: isCouncil ? councilRooms : [],
      permissions: isCouncil
        ? ['attendance','cleaning','stats-att','stats-cln','student-self']
        : ['student-self']
    };

    if (isCouncil) {
      const classrooms = cData.slice(1)
        .filter(r => String(r[2]) === '1' && (councilRooms.length === 0 || councilRooms.includes(String(r[0]).trim())))
        .map(r => ({ id: String(r[0]).trim(), name: String(r[1]).trim(), status: r[2] }));

      const studentsByClassroom = {};
      councilRooms.forEach(id => {
        studentsByClassroom[id] = sData.slice(1)
          .filter(r => String(r[0]).trim() === id && String(r[4]) === '1')
          .map(r => ({
            classroom_id: String(r[0]).trim(), student_id: String(r[1]).trim(),
            no: String(r[2]).trim(), name: String(r[3]).trim(), status: r[4]
          }));
      });
      const studentCounts = [];
      Object.keys(studentsByClassroom).forEach(rid => {
        studentCounts.push({ room_id: rid, count: studentsByClassroom[rid].length });
      });

      return createResponse('success', 'Login successful', {
        student: studentData, classrooms, studentCounts
      });
    }

    return createResponse('success', 'Login successful', { student: studentData });
  } catch (err) {
    return createResponse('error', 'เกิดข้อผิดพลาด: ' + err.message, null, err);
  }
}

// ═══════════════════════════════════════════════════════════════
// 7. BASIC GETTERS
// ═══════════════════════════════════════════════════════════════

function getConfig(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.CONFIG, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const config = { site_config: {}, school_config: {} };
    data.slice(1).forEach(row => {
      const key = row[0], value = row[1];
      if (['logoUrl', 'systemName', 'systemAdmin'].includes(key)) config.site_config[key] = value;
      else if (['schoolName', 'semester', 'schoolYear'].includes(key)) config.school_config[key] = value;
    });
    return createResponse('success', 'Config retrieved', config);
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getStudentAnalysisUrl() {
  try {
    const sheet = getSpreadsheet().getSheetByName('Config');
    return sheet.getRange('B8').getValue();
  } catch (e) { return ''; }
}

function getStatusList(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.STATUS_LIST, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const statusList = data.length <= 1 ? [] : data.slice(1)
      .filter(row => row[3] === 1)
      .map(row => ({
        label: row[0], color: row[1], hexColor: row[2], status: row[3]
      }));
    return createResponse('success', 'Status list retrieved', statusList);
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getTeachers(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.TEACHERS, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const teachers = data.length <= 1 ? [] : data.slice(1).map(row => ({
      id: row[0], name: row[1], username: row[2], password: row[3],
      status: row[4],
      classroom_ids: row[5] ? String(row[5]).split(',').map(s => s.trim()) : [],
      imageUrl: String(row[6] || '').trim(),
      phone: String(row[7] || '').trim(),
      role: String(row[8] || 'teacher').trim().toLowerCase()
    }));
    return createResponse('success', 'Teachers retrieved', teachers);
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getPeriods(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.PERIODS, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const periods = data.length <= 1 ? [] : data.slice(1)
      .filter(row => row[2] === 1)
      .map(row => ({ id: row[0], name: row[1], status: row[2] }));
    return createResponse('success', 'Periods retrieved', periods);
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getClassrooms(teacherClassroomIds = [], spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    if (!teacherClassroomIds || !Array.isArray(teacherClassroomIds)) {
      return createResponse('error', 'Invalid classroom_ids', null, { code: 'INVALID_INPUT' });
    }
    const sheet = getOrCreateSheet(SHEET_NAMES.CLASSROOMS, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const classrooms = data.length <= 1 ? [] : data.slice(1)
      .filter(row => String(row[2]).trim() === '1' && (teacherClassroomIds.length === 0 || teacherClassroomIds.includes(String(row[0]).trim())))
      .map(row => ({ id: String(row[0]).trim(), name: String(row[1]).trim(), status: row[2] }));
    return createResponse('success', 'Classrooms retrieved', classrooms);
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getStudents(classroomIds, spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    if (!classroomIds || !Array.isArray(classroomIds) || !classroomIds.length) {
      return createResponse('error', 'Invalid classroom_ids', null, { code: 'INVALID_INPUT' });
    }
    const sheet = getOrCreateSheet(SHEET_NAMES.STUDENTS, spreadsheetMode);
    const data = sheet.getDataRange().getDisplayValues();
    const studentsByClassroom = {};
    classroomIds.forEach(id => {
      studentsByClassroom[id] = data.length <= 1 ? [] : data.slice(1)
        .filter(row => String(row[0]).trim() === id && String(row[4]).trim() === '1')
        .map(row => ({
          classroom_id: String(row[0]).trim(), student_id: String(row[1]).trim(),
          no: String(row[2]).trim(), name: String(row[3]).trim(), status: row[4]
        }));
    });
    return createResponse('success', 'Students retrieved', studentsByClassroom);
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getAllStudents(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.STUDENTS, spreadsheetMode);
    const data = sheet.getDataRange().getDisplayValues();
    const studentsByClassroom = {};
    if (data.length > 1) {
      data.slice(1).filter(row => String(row[4]).trim() === '1').forEach(row => {
        const cid = String(row[0]).trim();
        if (!studentsByClassroom[cid]) studentsByClassroom[cid] = [];
        studentsByClassroom[cid].push({
          classroom_id: cid, student_id: String(row[1]).trim(),
          no: String(row[2]).trim(), name: String(row[3]).trim(), status: row[4]
        });
      });
    }
    return createResponse('success', 'All students', studentsByClassroom);
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getClassroomStudents(classroomId) {
  try {
    if (!classroomId) return { students: [], error: 'ไม่มีรหัสห้องเรียน' };
    const key = String(classroomId).trim().toLowerCase();
    const sheet = getOrCreateSheet(SHEET_NAMES.STUDENTS);
    const data = sheet.getDataRange().getDisplayValues();
    const students = data.slice(1)
      .filter(row => String(row[0]).trim().toLowerCase() === key && String(row[4]).trim() === '1')
      .map(row => ({
        classroom_id: String(row[0]).trim(),
        student_id: String(row[1]).trim(),
        no: String(row[2]).trim(),
        name: String(row[3]).trim(),
        status: String(row[4]).trim()
      }));
    Logger.log("Found " + students.length + " students in classroom '" + key + "'");
    if (students.length === 0) {
      return { students: [], error: `ไม่พบนักเรียนในห้อง "${key}"` };
    }
    return { students };
  } catch (e) {
    Logger.log('Error in getClassroomStudents: ' + e.message);
    return { students: [], error: e.message };
  }
}

// ═══════════════════════════════════════════════════════════════
// 8. ATTENDANCE — CLEAN & GET
// ═══════════════════════════════════════════════════════════════

function formatDateStr(d) {
  if (!d) return '';
  if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(d).trim().substring(0, 10);
}

function getCleanedAttendanceRows(data) {
  if (!data || data.length <= 1) return [];
  const uniqueRecords = {};
  data.slice(1).forEach(row => {
    const rDate = formatDateStr(row[0]);
    const rPeriod = String(row[1] || '').trim();
    const rRoom = String(row[2] || '').trim();
    if (rDate && rPeriod && rRoom) uniqueRecords[`${rDate}_${rPeriod}_${rRoom}`] = row;
  });
  return Object.values(uniqueRecords);
}

function getCleanedCleaningRows(data) {
  if (!data || data.length <= 1) return [];
  const uniqueRecords = {};
  data.slice(1).forEach(row => {
    const rDate = formatDateStr(row[0]);
    const rRoom = String(row[1] || '').trim();
    if (rDate && rRoom) uniqueRecords[`${rDate}_${rRoom}`] = row;
  });
  return Object.values(uniqueRecords);
}

function getAttendance(date, period, roomId, spreadsheetMode) {
  if (spreadsheetMode === undefined) spreadsheetMode = DEFAULT_SPREADSHEET_MODE;
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.ATTENDANCE, spreadsheetMode);
    const cleanData = getCleanedAttendanceRows(sheet.getDataRange().getValues());

    const attendance = cleanData.filter(row => {
      const rowDate = formatDateStr(row[0]);
      const rowPeriod = String(row[1] || '').trim();
      const rowRoomId = String(row[2] || '').trim();
      if (!rowDate || !rowPeriod || !rowRoomId || !/^\d{4}-\d{2}-\d{2}$/.test(rowDate)) return false;
      return (!date || rowDate === date.trim()) &&
             (!period || rowPeriod === period.trim()) &&
             (!roomId || rowRoomId === roomId.trim());
    }).map(row => ({
      date: formatDateStr(row[0]),
      period: String(row[1] || '').trim(),
      room_id: String(row[2] || '').trim(),
      teacher: String(row[3] || '').trim(),
      students: JSON.parse(row[4] || '[]'),
      updated_at: String(row[5] || '').trim()
    }));

    return createResponse('success', 'Attendance retrieved', attendance);
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getStudentsAndAttendance(classroomId, date, period, spreadsheetMode) {
  if (spreadsheetMode === undefined) spreadsheetMode = DEFAULT_SPREADSHEET_MODE;
  try {
    const sData = getOrCreateSheet(SHEET_NAMES.STUDENTS, spreadsheetMode).getDataRange().getValues();
    const students = sData.slice(1)
      .filter(row => String(row[0]).trim() === classroomId && row[4] === 1)
      .map(row => ({
        classroom_id: String(row[0]).trim(), student_id: String(row[1]).trim(),
        no: String(row[2]).trim(), name: String(row[3]).trim(), status: row[4]
      }));

    const attData = getOrCreateSheet(SHEET_NAMES.ATTENDANCE, spreadsheetMode).getDataRange().getValues();
    const cleanData = getCleanedAttendanceRows(attData);
    const attendance = cleanData.filter(row => {
      return formatDateStr(row[0]) === date &&
             String(row[1] || '').trim() === period &&
             String(row[2] || '').trim() === classroomId;
    }).map(row => ({
      date: formatDateStr(row[0]),
      period: String(row[1] || '').trim(),
      room_id: String(row[2] || '').trim(),
      teacher: String(row[3] || '').trim(),
      students: JSON.parse(row[4] || '[]'),
      updated_at: String(row[5] || '').trim()
    }));

    return createResponse('success', 'Data', { students, attendance: attendance[0] || null });
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getAttendanceRange(startDate, endDate, period, spreadsheetMode) {
  if (spreadsheetMode === undefined) spreadsheetMode = DEFAULT_SPREADSHEET_MODE;
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.ATTENDANCE, spreadsheetMode);
    const cleanData = getCleanedAttendanceRows(sheet.getDataRange().getValues());

    const attendance = cleanData.filter(row => {
      const rowDate = formatDateStr(row[0]);
      const rowPeriod = String(row[1] || '').trim();
      if (!rowDate || !rowPeriod || !/^\d{4}-\d{2}-\d{2}$/.test(rowDate)) return false;
      return rowDate >= startDate.trim() && rowDate <= endDate.trim() &&
             (!period || rowPeriod === period.trim());
    }).map(row => ({
      date: formatDateStr(row[0]),
      period: String(row[1] || '').trim(),
      room_id: String(row[2] || '').trim(),
      teacher: String(row[3] || '').trim(),
      students: JSON.parse(row[4] || '[]'),
      updated_at: String(row[5] || '').trim()
    }));

    return createResponse('success', 'Range retrieved', attendance);
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getWeeklyAttendanceByRoom(weekDates, period, spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.ATTENDANCE, spreadsheetMode);
    const data = sheet.getDataRange().getValues();
    const statusMap = {};

    for (let i = 1; i < data.length; i++) {
      const rowDate = formatDateStr(data[i][0]);
      const rowPeriod = String(data[i][1]).trim();
      const rowRoom = String(data[i][2]).trim();
      if (weekDates.includes(rowDate) && rowPeriod === period) {
        if (!statusMap[rowRoom]) statusMap[rowRoom] = {};
        statusMap[rowRoom][rowDate] = true;
      }
    }
    return createResponse('success', 'Weekly status', statusMap);
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

// ═══════════════════════════════════════════════════════════════
// 9. SAVE ATTENDANCE (with permission)
// ═══════════════════════════════════════════════════════════════

function saveAttendance(jsonData, spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return createResponse('error', 'Server is busy', null, { code: 'LOCK_TIMEOUT' });
    }
    try {
      const entry = JSON.parse(jsonData);
      let { date, period, room_id, teacher, students, updated_at } = entry;

      if (!date || !period || !room_id || !teacher || !students || !updated_at) {
        return createResponse('error', 'Missing fields', null, { code: 'INVALID_INPUT' });
      }
      date = date.trim(); period = period.trim(); room_id = room_id.trim();
      teacher = teacher.trim(); updated_at = updated_at.trim();

      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
        return createResponse('error', 'Invalid date', null, { code: 'INVALID_DATE' });
      }
      if (!Array.isArray(students) || students.some(s => !s.student_id || !s.status)) {
        return createResponse('error', 'Invalid students', null, { code: 'INVALID_STUDENTS' });
      }

      // ✅ PERMISSION CHECK
      const actor = getActorInfo(teacher);
      if (!actor || !['teacher','discipline','council','director'].includes(actor.role)) {
        return createResponse('error', 'ไม่มีสิทธิ์เช็คแถว', null, { code: 'PERMISSION_DENIED' });
      }
      if (!actor.classroom_ids.includes(room_id)) {
        if (!(['director','discipline'].includes(actor.role) && actor.classroom_ids.length === 0)) {
          return createResponse('error', `ไม่มีสิทธิ์เช็คห้อง ${room_id}`, null, { code: 'PERMISSION_DENIED' });
        }
      }

      const sheet = getOrCreateSheet(SHEET_NAMES.ATTENDANCE, spreadsheetMode);
      const data = sheet.getDataRange().getValues();
      const existingIndex = data.slice(1).findIndex(row => {
        return formatDateStr(row[0]) === date &&
               String(row[1] || '').trim() === period &&
               String(row[2] || '').trim() === room_id;
      }) + 2;

      const rowData = [date, period, room_id, teacher, JSON.stringify(students), updated_at];
      sheet.getRange('A:A').setNumberFormat('@');

      if (existingIndex > 1) {
        sheet.getRange(existingIndex, 1, 1, rowData.length).setValues([rowData]);
        return createResponse('success', 'Updated attendance');
      } else {
        sheet.appendRow(rowData);
        return createResponse('success', 'Saved attendance');
      }
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    Logger.log('saveAttendance error: ' + error.message);
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

// ═══════════════════════════════════════════════════════════════
// 10. GET INITIAL DATA
// ═══════════════════════════════════════════════════════════════

function getInitialData(spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
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
    if (data.config.length > 1) {
      data.config.slice(1).forEach(row => {
        if (row.length < 2) return;
        const key = row[0]?.toString().trim();
        const value = row[1]?.toString().trim();
        if (['logoUrl','systemName','systemAdmin'].includes(key)) config.site_config[key] = value;
        else if (['schoolName','semester','schoolYear'].includes(key)) config.school_config[key] = value;
      });
    }

    const statusList = data.statusList.length <= 1 ? [] : data.statusList.slice(1)
      .filter(row => row.length >= 4 && row[3] === 1)
      .map(row => ({
        label: row[0]?.toString().trim(), color: row[1]?.toString().trim(),
        hexColor: row[2]?.toString().trim(), status: row[3]
      }));

    const teachers = data.teachers.length <= 1 ? [] : data.teachers.slice(1)
      .filter(row => row.length >= 6)
      .map(row => ({
        id: row[0]?.toString().trim(), name: row[1]?.toString().trim(),
        username: row[2]?.toString().trim(), password: row[3]?.toString().trim(),
        status: row[4],
        classroom_ids: row[5] ? row[5].toString().trim().split(',').map(s => s.trim()) : [],
        imageUrl: row[6]?.toString().trim() || '',
        phone: row[7]?.toString().trim() || '',
        role: row[8]?.toString().trim().toLowerCase() || 'teacher'
      }));

    const periods = data.periods.length <= 1 ? [] : data.periods.slice(1)
      .filter(row => row.length >= 3 && row[2] === 1)
      .map(row => ({
        id: row[0]?.toString().trim(), name: row[1]?.toString().trim(), status: row[2]
      }));

    return createResponse('success', 'Initial data', { config, statusList, teachers, periods });
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function testGetConfig() {
  const data = getConfig().data;
  Logger.log(data.site_config['systemName']);
}

// ═══════════════════════════════════════════════════════════════
// 11. CLEANING SYSTEM
// ═══════════════════════════════════════════════════════════════

function getCleaningConfig() {
  try {
    let classrooms = [];
    try {
      const crRes = getClassrooms([]);
      classrooms = crRes.data || [];
    } catch (e) { throw new Error("ไม่สามารถเรียก getClassrooms"); }

    let students = {};
    try {
      const roomIds = classrooms.map(c => c.id);
      const stRes = getStudents(roomIds);
      students = stRes.data || {};
    } catch (e) { throw new Error("ไม่สามารถเรียก getStudents"); }

    const criteria = [
      { name: "ความสะอาดพื้นและบริเวณโดยรอบ", max_score: 10 },
      { name: "การจัดการขยะและถังขยะ", max_score: 10 },
      { name: "ความเรียบร้อยของการจัดเก็บอุปกรณ์", max_score: 10 }
    ];

    return createResponse('success', 'Config loaded', { classrooms, students, criteria });
  } catch (error) {
    return createResponse('error', error.message, null, error);
  }
}

function saveCleaningData(jsonData, spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
      return createResponse('error', 'Server is busy', null, { code: 'LOCK_TIMEOUT' });
    }
    try {
      const data = JSON.parse(jsonData);
      const { date, classroom_id, checker_id, student_data, scores } = data;

      if (!date || !classroom_id || !checker_id || !student_data || !scores) {
        return createResponse('error', 'Missing fields', null, { code: 'INVALID_INPUT' });
      }
      if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
        return createResponse('error', 'Invalid date', null, { code: 'INVALID_DATE' });
      }

      // ✅ PERMISSION CHECK
      const actor = getActorInfo(checker_id);
      if (!actor || !['teacher','discipline','council','director'].includes(actor.role)) {
        return createResponse('error', 'ไม่มีสิทธิ์เช็คเขตพื้นที่', null, { code: 'PERMISSION_DENIED' });
      }
      if (!actor.classroom_ids.includes(classroom_id)) {
        if (!(['director','discipline'].includes(actor.role) && actor.classroom_ids.length === 0)) {
          return createResponse('error', `ไม่มีสิทธิ์เช็คห้อง ${classroom_id}`, null, { code: 'PERMISSION_DENIED' });
        }
      }

      const updated_at = new Date().toISOString();

      const attSheet = getOrCreateSheet(CLEANING_SHEET_NAMES.ATTENDANCE, spreadsheetMode);
      const attData = attSheet.getDataRange().getValues();
      const existingAtt = attData.slice(1).findIndex(row => {
        return String(row[0]).trim() === date && String(row[1]).trim() === classroom_id;
      }) + 2;

      const attRow = [date, classroom_id, checker_id, JSON.stringify(student_data), updated_at];
      if (existingAtt > 1) attSheet.getRange(existingAtt, 1, 1, attRow.length).setValues([attRow]);
      else attSheet.appendRow(attRow);

      let total = 0, maxTotal = 0;
      for (const k in scores) total += scores[k] || 0;
      const critData = getOrCreateSheet(CLEANING_SHEET_NAMES.CRITERIA, spreadsheetMode).getDataRange().getValues();
      critData.slice(1).forEach(r => { maxTotal += r[2] || 0; });
      const totalScore = `${total}/${maxTotal}`;

      const scoSheet = getOrCreateSheet(CLEANING_SHEET_NAMES.SCORES, spreadsheetMode);
      const scoData = scoSheet.getDataRange().getValues();
      const existingSco = scoData.slice(1).findIndex(row => {
        return String(row[0]).trim() === date && String(row[1]).trim() === classroom_id;
      }) + 2;

      const scoRow = [date, classroom_id, checker_id, JSON.stringify(scores), totalScore, updated_at];
      if (existingSco > 1) scoSheet.getRange(existingSco, 1, 1, scoRow.length).setValues([scoRow]);
      else scoSheet.appendRow(scoRow);

      return createResponse('success', existingAtt > 1 ? 'Updated' : 'Saved', { total_score: totalScore });
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getCleaningStats(filter, spreadsheetMode = DEFAULT_SPREADSHEET_MODE) {
  try {
    const { startDate, endDate, classroom_id } = filter;

    const attSheet = getOrCreateSheet(CLEANING_SHEET_NAMES.ATTENDANCE, spreadsheetMode);
    const scoSheet = getOrCreateSheet(CLEANING_SHEET_NAMES.SCORES, spreadsheetMode);
    const tData = getOrCreateSheet(SHEET_NAMES.TEACHERS, spreadsheetMode).getDataRange().getValues();

    const teacherMap = {};
    tData.slice(1).forEach(row => {
      const name = String(row[1]).trim();
      const username = String(row[2]).trim();
      if (username) teacherMap[username] = name;
    });

    const cleanAtt = getCleanedCleaningRows(attSheet.getDataRange().getValues());
    const cleanSco = getCleanedCleaningRows(scoSheet.getDataRange().getValues());

    const fmt = (d) => d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(d || '').trim();

    const fAtt = cleanAtt.filter(row => {
      if (row[0] === 'Date') return false;
      const d = fmt(row[0]);
      const c = String(row[1]).trim();
      return d >= startDate && d <= endDate && (!classroom_id || c === classroom_id);
    });

    const fSco = cleanSco.filter(row => {
      if (row[0] === 'Date') return false;
      const d = fmt(row[0]);
      const c = String(row[1]).trim();
      return d >= startDate && d <= endDate && (!classroom_id || c === classroom_id);
    });

    const studentStats = {};
    fAtt.forEach(row => {
      const roomId = String(row[1]).trim();
      const students = JSON.parse(row[3] || '[]');
      students.forEach(stu => {
        const sid = stu.student_id;
        if (!studentStats[sid]) {
          studentStats[sid] = { student_id: sid, name: stu.name || sid, classroom_id: roomId,
                                total: 0, done: 0, not_done: 0, leave: 0 };
        }
        studentStats[sid].total++;
        if (stu.status === 'ทำ') studentStats[sid].done++;
        else if (stu.status === 'ไม่ทำ') studentStats[sid].not_done++;
        else if (stu.status === 'ลา' || stu.status === 'หนีเวร') studentStats[sid].leave++;
      });
    });

    const allStudents = Object.values(studentStats).sort((a, b) => {
      if (a.classroom_id !== b.classroom_id) return a.classroom_id.localeCompare(b.classroom_id);
      return String(a.student_id).localeCompare(String(b.student_id));
    });

    const topStudents = Object.values(studentStats).sort((a, b) => b.done - a.done).slice(0, 5);

    let totalScore = 0, count = 0;
    fSco.forEach(row => {
      const s = String(row[4]);
      if (s.includes('/')) {
        const [sc, mx] = s.split('/').map(Number);
        if (!isNaN(sc) && !isNaN(mx) && mx > 0) { totalScore += (sc/mx)*100; count++; }
      }
    });
    const avgScore = count > 0 ? (totalScore/count).toFixed(2) : 0;

    const scoreDetails = fSco.map(row => ({
      date: fmt(row[0]),
      classroom_id: row[1],
      checker: teacherMap[String(row[2]).trim()] || String(row[2]).trim(),
      scores: JSON.parse(row[3] || '{}'),
      total_score: row[4]
    }));

    return createResponse('success', 'Stats', {
      studentStats: allStudents, allStudents, topStudents, avgScore, scoreDetails
    });
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getCleaningWeeklyStatus(weekDates) {
  try {
    const sheet = getOrCreateSheet(CLEANING_SHEET_NAMES.ATTENDANCE);
    const cleaned = getCleanedCleaningRows(sheet.getDataRange().getValues());
    const statusMap = {};

    cleaned.forEach(row => {
      const rowDate = formatDateStr(row[0]);
      const rowRoom = String(row[1]).trim();
      if (weekDates.includes(rowDate) && rowRoom) {
        if (!statusMap[rowRoom]) statusMap[rowRoom] = {};
        statusMap[rowRoom][rowDate] = true;
      }
    });
    return createResponse('success', 'Weekly status', statusMap);
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getCleaningDataForEdit(date, classroomId) {
  try {
    const ss = getSpreadsheet();
    let studentData = [], scores = {}, hasExisting = false;

    const attSheet = ss.getSheetByName(CLEANING_SHEET_NAMES.ATTENDANCE);
    if (attSheet) {
      const cleaned = getCleanedCleaningRows(attSheet.getDataRange().getValues());
      const row = cleaned.find(r => formatDateStr(r[0]) === date && String(r[1]).trim() === classroomId);
      if (row) { try { studentData = JSON.parse(row[3] || '[]'); } catch(e) {} hasExisting = true; }
    }

    const scoSheet = ss.getSheetByName(CLEANING_SHEET_NAMES.SCORES);
    if (scoSheet) {
      const cleaned = getCleanedCleaningRows(scoSheet.getDataRange().getValues());
      const row = cleaned.find(r => formatDateStr(r[0]) === date && String(r[1]).trim() === classroomId);
      if (row) { try { scores = JSON.parse(row[3] || '{}'); } catch(e) {} hasExisting = true; }
    }

    return createResponse('success', 'Data', { studentData, scores, hasExisting });
  } catch (error) {
    return createResponse('error', error.message, { studentData: [], scores: {}, hasExisting: false }, error);
  }
}

function getCleaningStudentDetail(studentId, startDate, endDate) {
  try {
    const sheet = getOrCreateSheet(CLEANING_SHEET_NAMES.ATTENDANCE);
    const cleaned = getCleanedCleaningRows(sheet.getDataRange().getValues());
    const result = [];
    const targetStu = String(studentId || '').trim();

    cleaned.forEach(row => {
      const rowDate = formatDateStr(row[0]);
      const rowRoom = String(row[1] || '').trim();
      if (startDate && rowDate < startDate) return;
      if (endDate && rowDate > endDate) return;
      let students = [];
      try { students = JSON.parse(row[3] || '[]'); } catch(e) {}
      const found = students.find(s => String(s.student_id || '').trim().replace(/^0+/, '') === targetStu.replace(/^0+/, ''));
      if (found) {
        result.push({
          date: rowDate, classroom_id: rowRoom,
          status: found.status, role: found.role || '',
          checker: String(row[2] || '').trim()
        });
      }
    });

    result.sort((a, b) => b.date.localeCompare(a.date));
    return { status: 'success', data: result };
  } catch (err) {
    Logger.log('getCleaningStudentDetail error: ' + err.message);
    return { status: 'error', message: err.toString(), data: [] };
  }
}

function setupCleaningSheets() {
  const ss = getSpreadsheet();
  const names = { ATTENDANCE: 'Cleaning_Attendance', SCORES: 'Cleaning_Scores', CRITERIA: 'Cleaning_Criteria' };

  const attSheet = ss.getSheetByName(names.ATTENDANCE) || ss.insertSheet(names.ATTENDANCE);
  attSheet.getRange(1, 1, 1, 5).setValues([['Date', 'Classroom ID', 'Checker ID', 'Student Data', 'Updated At']])
    .setFontWeight("bold").setBackground("#d9ead3");

  const scoSheet = ss.getSheetByName(names.SCORES) || ss.insertSheet(names.SCORES);
  scoSheet.getRange(1, 1, 1, 6).setValues([['Date', 'Classroom ID', 'Checker ID', 'Scores', 'Total Score', 'Updated At']])
    .setFontWeight("bold").setBackground("#c9daf8");

  let criSheet = ss.getSheetByName(names.CRITERIA);
  if (!criSheet) {
    criSheet = ss.insertSheet(names.CRITERIA);
    criSheet.getRange(1, 1, 4, 3).setValues([
      ['ID', 'Name', 'Max Score'],
      ['C01', 'กวาดพื้น/ถูพื้น', 10],
      ['C02', 'ทิ้งขยะ', 5],
      ['C03', 'ความเป็นระเบียบ', 5]
    ]).setFontWeight("bold");
  }
  Logger.log('✅ Cleaning sheets ready');
}

// ═══════════════════════════════════════════════════════════════
// 12. SUBJECT ATTENDANCE
// ═══════════════════════════════════════════════════════════════

function getTeacherSubjectSchedule(teacherEmail, targetDate) {
  const t0 = Date.now();
  try {
    Logger.log('📅 getTeacherSubjectSchedule: email=' + teacherEmail + ', date=' + targetDate);
    const data = getCachedTimetableData();
    Logger.log('⏱️ After cache: ' + (Date.now() - t0) + 'ms, rows: ' + data.length);

    if (!data || data.length <= 1) {
      return { status: 'success', data: [], message: 'ชีต Timetable ว่าง' };
    }

    let dateObj;
    if (targetDate) {
      const p = String(targetDate).split('-');
      dateObj = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
    } else {
      dateObj = new Date();
    }

    const days = ["อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"];
    const dayName = days[dateObj.getDay()];
    const normEmail = String(teacherEmail || '').trim().toLowerCase();
    Logger.log('🗓️ Day=' + dayName + ', email=' + normEmail);

    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row || row.length < 9) continue;
      const rowDay = String(row[1] || '').trim();
      const rowEmailNorm = String(row[5] || '').trim().toLowerCase();
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
    Logger.log('✅ Found ' + result.length + ' periods in ' + (Date.now() - t0) + 'ms');
    return { status: 'success', data: result };
  } catch (e) {
    Logger.log('❌ Error in ' + (Date.now() - t0) + 'ms: ' + e.message);
    return { status: 'error', message: e.message || 'Server error', data: [] };
  }
}

function saveSubjectAttendance(jsonPayload) {
  try {
    const payload = JSON.parse(jsonPayload);

    // ✅ Council check
    const actor = getActorInfo(payload.teacherEmail);
    if (actor && actor.type === 'student_council') {
      return { status: 'error', message: 'สภานักเรียนไม่มีสิทธิ์เช็คเข้าเรียน' };
    }

    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Attendance_Subjects');
    if (!sheet) {
      sheet = ss.insertSheet('Attendance_Subjects');
      sheet.appendRow(['Timestamp','Date','SubjectCode','SubjectName','TeacherEmail','GradeLevel','Room','StudentID','StudentName','Status','Note']);
    }

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { status: 'error', message: 'ระบบทำงานอยู่' };

    try {
      const targetDate = String(payload.date || '').trim();
      const targetCode = String(payload.subjectCode || '').trim();
      const targetName = String(payload.subjectName || '').trim();
      const targetGrade = String(payload.gradeLevel || '').trim();
      const targetRoom = String(payload.room || '').trim();

      const data = sheet.getDataRange().getValues();
      const toDelete = [];
      for (let i = data.length - 1; i >= 1; i--) {
        const row = data[i];
        const rowDate = formatDateStr(row[1]);
        const rowCode = String(row[2] || '').trim();
        const rowName = String(row[3] || '').trim();
        const rowGrade = String(row[5] || '').trim();
        const rowRoom = String(row[6] || '').trim();
        const codeMatch = targetCode !== '' ? (rowCode === targetCode) : (rowCode === '' && rowName === targetName);
        if (rowDate === targetDate && codeMatch && rowGrade === targetGrade && rowRoom === targetRoom) {
          toDelete.push(i + 1);
        }
      }
      toDelete.forEach(idx => sheet.deleteRow(idx));

      const timestamp = new Date();
      const toAdd = [];
      payload.students.forEach(stu => {
        toAdd.push([timestamp, targetDate, targetCode, targetName, payload.teacherEmail,
                    targetGrade, targetRoom, stu.student_id, stu.name, stu.status, stu.note || '']);
      });
      if (toAdd.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, toAdd.length, toAdd[0].length).setValues(toAdd);
      }

      return {
        status: 'success',
        message: toDelete.length > 0 ? 'อัปเดตผลการเข้าเรียนสำเร็จ' : 'บันทึกผลการเข้าเรียนสำเร็จ'
      };
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getSubjectAttendanceForEdit(date, subjectCode, subjectName, gradeLevel, room) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Attendance_Subjects');
    if (!sheet) return { status: 'success', data: [] };

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { status: 'success', data: [] };

    const targetDate = formatDateStr(date);
    const targetCode = String(subjectCode || '').trim();
    const targetName = String(subjectName || '').trim();
    const targetGrade = String(gradeLevel || '').trim();
    const targetRoom = String(room || '').trim();

    const result = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDate = formatDateStr(row[1]);
      const rowCode = String(row[2] || '').trim();
      const rowName = String(row[3] || '').trim();
      const rowGrade = String(row[5] || '').trim();
      const rowRoom = String(row[6] || '').trim();
      const codeMatch = targetCode !== '' ? (rowCode === targetCode) : (rowCode === '' && rowName === targetName);
      if (rowDate === targetDate && codeMatch && rowGrade === targetGrade && rowRoom === targetRoom) {
        result.push({
          student_id: String(row[7] || '').trim(),
          name: String(row[8] || '').trim(),
          status: String(row[9] || '').trim(),
          note: String(row[10] || '').trim()
        });
      }
    }
    return { status: 'success', data: result };
  } catch (err) {
    return { status: 'error', message: err.toString(), data: [] };
  }
}

function getSubjectAttendanceSummary(teacherIdentifier, subjectKey) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName('Attendance_Subjects');
    if (!sheet) return { status: 'error', message: 'ไม่มีข้อมูล' };

    const data = sheet.getDataRange().getValues();
    const summary = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowCode = String(row[2]).trim();
      const rowName = String(row[3]).trim();
      const rowTeacher = String(row[4]).trim();
      const matchSubj = rowCode !== '' ? (rowCode === subjectKey) : (rowName === subjectKey);
      if (matchSubj && rowTeacher.toLowerCase() === String(teacherIdentifier).toLowerCase()) {
        const sid = String(row[7]).trim();
        if (!summary[sid]) summary[sid] = { id: sid, name: String(row[8]).trim(), present: 0, absent: 0, leave: 0, late: 0, total: 0 };
        const st = String(row[9]).trim();
        if (st === 'มา') summary[sid].present++;
        else if (st === 'ขาด') summary[sid].absent++;
        else if (st === 'ลา') summary[sid].leave++;
        else if (st === 'สาย') summary[sid].late++;
        summary[sid].total++;
      }
    }
    const arr = Object.keys(summary).map(k => summary[k]).sort((a,b) => String(a.id).localeCompare(String(b.id)));
    return { status: 'success', data: arr };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getTeacherSubjectsList(teacherIdentifier) {
  try {
    const data = getCachedTimetableData();
    if (!data || data.length <= 1) return { status: 'success', data: [] };

    const subjectsMap = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowTeacher = String(row[5]).trim();
      if (rowTeacher.toLowerCase() === String(teacherIdentifier).toLowerCase()) {
        const subjCode = String(row[3]).trim();
        const subjName = String(row[4]).trim();
        const grade = String(row[6]).trim();
        const room = String(row[7]).trim();
        let gradeRoom = '';
        if (grade !== '' && room !== '') gradeRoom = grade + '/' + room;
        else if (grade !== '') gradeRoom = grade;
        else if (room !== '') gradeRoom = room;
        const key = subjCode !== '' ? subjCode : subjName;
        if (!subjectsMap[key]) {
          subjectsMap[key] = { code: subjCode, name: subjName, classes: gradeRoom ? [gradeRoom] : [] };
        } else if (gradeRoom && subjectsMap[key].classes.indexOf(gradeRoom) === -1) {
          subjectsMap[key].classes.push(gradeRoom);
        }
      }
    }

    const result = Object.keys(subjectsMap).map(key => {
      const subj = subjectsMap[key];
      let displayName = subj.name;
      if (subj.code !== '') displayName += ' (' + subj.code + ')';
      if (subj.classes.length > 0) displayName += ' - ' + subj.classes.join(', ');
      return { key, code: subj.code, name: subj.name, classes: subj.classes, displayName };
    });
    return { status: 'success', data: result };
  } catch (err) {
    return { status: 'error', message: err.toString(), data: [] };
  }
}

function setupSubjectDatabase() {
  const ss = getSpreadsheet();
  try {
    let ttSheet = ss.getSheetByName('Timetable');
    if (!ttSheet) {
      ttSheet = ss.insertSheet('Timetable');
      ttSheet.appendRow(['TimetableID','DayOfWeek','Period','SubjectCode','SubjectName','TeacherEmail','GradeLevel','Room','ClassroomId']);
      ttSheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#f9cb9c').setHorizontalAlignment('center');
      ttSheet.setFrozenRows(1);
      const exEmail = Session.getActiveUser().getEmail() || 'teacher@school.ac.th';
      ttSheet.appendRow(['T001','จันทร์',1,'ว23101','วิทยาศาสตร์',exEmail,'ม.3','1','m3-1']);
      ttSheet.autoResizeColumns(1, 9);
    }

    let attSheet = ss.getSheetByName('Attendance_Subjects');
    if (!attSheet) {
      attSheet = ss.insertSheet('Attendance_Subjects');
      attSheet.appendRow(['Timestamp','Date','SubjectCode','SubjectName','TeacherEmail','GradeLevel','Room','StudentID','StudentName','Status','Note']);
      attSheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#d9ead3').setHorizontalAlignment('center');
      attSheet.setFrozenRows(1);
      attSheet.autoResizeColumns(1, 11);
    }
    return { status: 'success', message: 'สร้างชีตสำเร็จ' };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════
// 13. ATTENDANCE SUMMARY V2 (Phase 3)
// ═══════════════════════════════════════════════════════════════

function getAttendanceSummaryV2(startDate, endDate, period, classroomFilter) {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.ATTENDANCE);
    
    // ✅ 1. ดึงรายชื่อจริงจากชีต Students มาทำ Map เพื่อความแม่นยำ
    const stdSheet = getSpreadsheet().getSheetByName(SHEET_NAMES.STUDENTS);
    const stdData = stdSheet ? stdSheet.getDataRange().getDisplayValues() : [];
    const nameMap = {};
    const idMap = {}; // ใช้เทียบรหัสแบบไม่มี 0
    
    stdData.slice(1).forEach(r => { 
      const id = String(r[1]).trim(); // student_id (คอลัมน์ B)
      const name = String(r[3]).trim(); // name (คอลัมน์ D)
      nameMap[id] = name; 
      idMap[id.replace(/^0+/, '')] = { id: id, name: name };
    });

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

        // ✅ 2. ค้นหาชื่อจริงจาก Students sheet (แก้ปัญหาชื่อกลายเป็นรหัส)
        const normId = sid.replace(/^0+/, '');
        let finalName = nameMap[sid] || (idMap[normId] ? idMap[normId].name : (s.name || sid));
        let finalId = idMap[normId] ? idMap[normId].id : sid;

        if (!summary[finalId]) {
          summary[finalId] = {
            id: finalId, name: finalName, classroom_id: rowRoom,
            present: 0, absent: 0, late: 0, leave: 0, sick: 0, total: 0
          };
        }
        const rec = summary[finalId];
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
      return String(a.id).localeCompare(String(b.id));
    });

    const totals = { present: 0, absent: 0, late: 0, leave: 0, sick: 0, total: 0 };
    students.forEach(s => {
      totals.present += s.present; totals.absent += s.absent;
      totals.late += s.late; totals.leave += s.leave;
      totals.sick += s.sick; totals.total += s.total;
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
    const targetStu = String(studentId || '').trim().replace(/^0+/, ''); // ✅ ตัด 0 ออกเพื่อเทียบ
    const targetPeriod = String(period || '').trim();

    cleanData.forEach(row => {
      const rowDate = formatDateStr(row[0]);
      const rowPeriod = String(row[1] || '').trim();
      if (startDate && rowDate < startDate) return;
      if (endDate && rowDate > endDate) return;
      if (targetPeriod && rowPeriod !== targetPeriod) return;

      let students = [];
      try { students = JSON.parse(row[4] || '[]'); } catch (e) {}
      
      // ✅ ใช้การ replace 0 เพื่อเทียบรหัส
      const found = students.find(s => String(s.student_id || '').trim().replace(/^0+/, '') === targetStu);
      
      if (found) {
        result.push({ 
          date: rowDate, 
          period: rowPeriod, 
          status: found.status || '', 
          reason: found.reason || '' 
        });
      }
    });

    result.sort((a, b) => b.date.localeCompare(a.date));
    return { status: 'success', data: result };
  } catch (err) {
    return { status: 'error', message: err.toString(), data: [] };
  }
}

function getWeeklyCheckReport(startDate, endDate, mode, options) {
  try {
    const opts = options || {};
    const weekDates = [];
    const [sy, sm, sd] = String(startDate).split('-').map(Number);
    for (let i = 0; i < 7; i++) {
      const d = new Date(sy, sm - 1, sd + i);
      const ds = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (ds >= startDate && ds <= endDate) weekDates.push(ds);
    }

    const rows = [];

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

      const classrooms = (getClassrooms([]).data || []).sort((a, b) => a.name.localeCompare(b.name, 'th'));
      const studentCounts = {};
      const allStu = getAllStudents();
      if (allStu.status === 'success') {
        Object.keys(allStu.data).forEach(rid => { studentCounts[rid] = allStu.data[rid].length; });
      }

      classrooms.forEach(c => {
        const dates = {};
        weekDates.forEach(d => { dates[d] = !!(checkMap[c.id] && checkMap[c.id][d]); });
        rows.push({ id: c.id, name: c.name, sublabel: `${studentCounts[c.id] || 0} คน`, dates, clickable: true, mode: 'attendance' });
      });
    }
    else if (mode === 'cleaning') {
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

      const classrooms = (getClassrooms([]).data || []).sort((a, b) => a.name.localeCompare(b.name, 'th'));
      const studentCounts = {};
      const allStu = getAllStudents();
      if (allStu.status === 'success') {
        Object.keys(allStu.data).forEach(rid => { studentCounts[rid] = allStu.data[rid].length; });
      }

      classrooms.forEach(c => {
        const dates = {};
        weekDates.forEach(d => { dates[d] = !!(checkMap[c.id] && checkMap[c.id][d]); });
        rows.push({ id: c.id, name: c.name, sublabel: `${studentCounts[c.id] || 0} คน`, dates, clickable: true, mode: 'cleaning' });
      });
    }
    else if (mode === 'subject') {
      const teacherEmail = String(opts.teacherEmail || '').trim().toLowerCase();
      const ttData = getCachedTimetableData();
      if (!ttData || ttData.length === 0) return { status: 'success', data: { weekDates, rows: [] } };

      const days = ["อาทิตย์", "จันทร์", "อังคาร", "พุธ", "พฤหัสบดี", "ศุกร์", "เสาร์"];
      const scheduleByDate = {};
      weekDates.forEach(d => {
        const dt = new Date(d);
        const dayName = days[dt.getDay()];
        scheduleByDate[d] = [];
        for (let i = 1; i < ttData.length; i++) {
          const r = ttData[i];
          if (String(r[1] || '').trim() === dayName && String(r[5] || '').trim().toLowerCase() === teacherEmail) {
            scheduleByDate[d].push({
              code: String(r[3] || '').trim(), name: String(r[4] || '').trim(),
              grade: String(r[6] || '').trim(), room: String(r[7] || '').trim(),
              classroomId: String(r[8] || '').trim()
            });
          }
        }
      });

      const checkMap = {};
      const attSheet = getSpreadsheet().getSheetByName('Attendance_Subjects');
      if (attSheet) {
        const attData = attSheet.getDataRange().getValues();
        for (let i = 1; i < attData.length; i++) {
          const r = attData[i];
          const rowDate = formatDateStr(r[1]);
          if (!weekDates.includes(rowDate)) continue;
          if (String(r[4] || '').trim().toLowerCase() !== teacherEmail) continue;
          const key = `${rowDate}_${String(r[2] || '').trim() || String(r[3] || '').trim()}_${String(r[5] || '').trim()}_${String(r[6] || '').trim()}`;
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
              id: key, name: sch.name + (sch.code ? ` (${sch.code})` : ''),
              sublabel: gradeRoom, dates: {},
              meta: { code: sch.code, name: sch.name, grade: sch.grade, room: sch.room, classroomId: sch.classroomId },
              clickable: true, mode: 'subject'
            };
          }
          rowMap[key].dates[d] = !!checkMap[`${d}_${sch.code || sch.name}_${sch.grade}_${sch.room}`];
        });
      });

      Object.values(rowMap).sort((a, b) => a.name.localeCompare(b.name, 'th')).forEach(r => rows.push(r));
    }

    return { status: 'success', data: { weekDates, rows } };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════
// 14. RISK ASSESSMENT
// ═══════════════════════════════════════════════════════════════

const RISK_WEIGHT_KEYS = {
  ATTENDANCE: 'riskWeight_attendance',
  CLEANING: 'riskWeight_cleaning',
  SUBJECT: 'riskWeight_subject'
};
const RISK_DEFAULT_WEIGHTS = { attendance: 40, cleaning: 30, subject: 30 };

function getRiskWeights() {
  try {
    const data = getOrCreateSheet(SHEET_NAMES.CONFIG).getDataRange().getValues();
    const weights = Object.assign({}, RISK_DEFAULT_WEIGHTS);
    data.slice(1).forEach(row => {
      const key = String(row[0] || '').trim();
      const val = Number(row[1]);
      if (isNaN(val) || val < 0) return;
      if (key === RISK_WEIGHT_KEYS.ATTENDANCE) weights.attendance = val;
      if (key === RISK_WEIGHT_KEYS.CLEANING) weights.cleaning = val;
      if (key === RISK_WEIGHT_KEYS.SUBJECT) weights.subject = val;
    });
    return { status: 'success', data: weights };
  } catch (err) {
    return { status: 'error', message: err.message, data: RISK_DEFAULT_WEIGHTS };
  }
}

function saveRiskWeights(weightsJson) {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return { status: 'error', message: 'Server busy' };
    try {
      const w = JSON.parse(weightsJson);
      const total = (w.attendance || 0) + (w.cleaning || 0) + (w.subject || 0);
      if (total <= 0) return { status: 'error', message: 'ผลรวมต้องมากกว่า 0' };

      const sheet = getOrCreateSheet(SHEET_NAMES.CONFIG);
      const data = sheet.getDataRange().getValues();
      const updateOrAdd = (key, val) => {
        for (let i = 1; i < data.length; i++) {
          if (String(data[i][0] || '').trim() === key) {
            sheet.getRange(i + 1, 2).setValue(val);
            return;
          }
        }
        sheet.appendRow([key, val]);
      };
      updateOrAdd(RISK_WEIGHT_KEYS.ATTENDANCE, Number(w.attendance) || 0);
      updateOrAdd(RISK_WEIGHT_KEYS.CLEANING, Number(w.cleaning) || 0);
      updateOrAdd(RISK_WEIGHT_KEYS.SUBJECT, Number(w.subject) || 0);
      return { status: 'success', message: 'บันทึกน้ำหนักสำเร็จ' };
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getStudentRiskAssessment(startDate, endDate, classroomFilter, weightsJson) {
  try {
    const weights = weightsJson ? JSON.parse(weightsJson) : getRiskWeights().data;
    const normId = id => String(id == null ? '' : id).trim().replace(/^0+(?=\d)/, '');
    const normRoom = id => String(id || '').trim().toLowerCase()
      .replace(/ม\./g, 'm').replace(/\s+/g, '').replace(/\//g, '-');

    const dateList = [];
    const [sy, sm, sd] = startDate.split('-').map(Number);
    const [ey, em, ed] = endDate.split('-').map(Number);
    const endD = new Date(ey, em - 1, ed);
    const cur = new Date(sy, sm - 1, sd);
    while (cur <= endD) {
      dateList.push(Utilities.formatDate(cur, Session.getScriptTimeZone(), 'yyyy-MM-dd'));
      cur.setDate(cur.getDate() + 1);
    }

    const allStuRes = getAllStudents();
    if (allStuRes.status !== 'success') throw new Error('Failed to load students');
    const classrooms = (getClassrooms([]).data || []);
    const roomNameMap = {};
    classrooms.forEach(c => { roomNameMap[c.id] = c.name; });

    const studentsMap = {};
    Object.keys(allStuRes.data).forEach(rid => {
      if (classroomFilter && normRoom(rid) !== normRoom(classroomFilter)) return;
      allStuRes.data[rid].forEach(s => {
        studentsMap[normId(s.student_id)] = {
          id: String(s.student_id), name: s.name,
          classroom_id: rid, classroom_name: roomNameMap[rid] || rid
        };
      });
    });

    // Attendance
    const attData = getCleanedAttendanceRows(getOrCreateSheet(SHEET_NAMES.ATTENDANCE).getDataRange().getValues());
    const attMap = {};
    attData.forEach(row => {
      const rDate = formatDateStr(row[0]);
      if (rDate < startDate || rDate > endDate) return;
      const rRoom = String(row[2] || '').trim();
      if (classroomFilter && normRoom(rRoom) !== normRoom(classroomFilter)) return;
      let students = [];
      try { students = JSON.parse(row[4] || '[]'); } catch (e) {}
      students.forEach(s => {
        const sid = normId(s.student_id);
        if (!studentsMap[sid]) return;
        const st = String(s.status || '');
        let pts = 0;
        if (st === 'มา') pts = 1;
        else if (st === 'สาย') pts = 0.75;
        else if (st === 'ลา' || st === 'ป่วย') pts = 0.5;
        if (!attMap[sid]) attMap[sid] = {};
        if (!attMap[sid][rDate]) attMap[sid][rDate] = { total: 0, points: 0 };
        attMap[sid][rDate].total++;
        attMap[sid][rDate].points += pts;
      });
    });

    // Cleaning
    const clnData = getCleanedCleaningRows(getOrCreateSheet(CLEANING_SHEET_NAMES.ATTENDANCE).getDataRange().getValues());
    const clnMap = {};
    clnData.forEach(row => {
      const rDate = formatDateStr(row[0]);
      if (rDate < startDate || rDate > endDate) return;
      const rRoom = String(row[1] || '').trim();
      if (classroomFilter && normRoom(rRoom) !== normRoom(classroomFilter)) return;
      let students = [];
      try { students = JSON.parse(row[3] || '[]'); } catch (e) {}
      students.forEach(s => {
        const sid = normId(s.student_id);
        if (!studentsMap[sid]) return;
        const st = String(s.status || '');
        let pts = 0;
        if (st === 'ทำ') pts = 1;
        else if (st === 'ลา' || st === 'หนีเวร') pts = 0.5;
        if (!clnMap[sid]) clnMap[sid] = {};
        if (!clnMap[sid][rDate]) clnMap[sid][rDate] = { total: 0, points: 0 };
        clnMap[sid][rDate].total++;
        clnMap[sid][rDate].points += pts;
      });
    });

    // Subject
    const subjSheet = getSpreadsheet().getSheetByName('Attendance_Subjects');
    const subjData = subjSheet ? subjSheet.getDataRange().getValues() : [];
    const subjMap = {};
    for (let i = 1; i < subjData.length; i++) {
      const row = subjData[i];
      const rDate = formatDateStr(row[1]);
      if (rDate < startDate || rDate > endDate) continue;
      const sid = normId(row[7]);
      if (!studentsMap[sid]) continue;
      if (classroomFilter) {
        const rowRoom = String(row[6] || '').trim();
        const rowGrade = String(row[5] || '').trim();
        const rowClassroomId = [normRoom(rowGrade), normRoom(rowRoom)].filter(Boolean).join('-');
        const targetRoom = normRoom(classroomFilter);
        if (rowClassroomId !== targetRoom && normRoom(rowRoom) !== targetRoom &&
            normRoom(rowGrade) !== targetRoom) continue;
      }
      const st = String(row[9] || '').trim();
      let pts = 0;
      if (st === 'มา') pts = 1;
      else if (st === 'สาย') pts = 0.75;
      else if (st === 'ลา' || st === 'ป่วย') pts = 0.5;
      if (!subjMap[sid]) subjMap[sid] = {};
      if (!subjMap[sid][rDate]) subjMap[sid][rDate] = { total: 0, points: 0 };
      subjMap[sid][rDate].total++;
      subjMap[sid][rDate].points += pts;
    }

    const results = [];
    Object.keys(studentsMap).forEach(normSid => {
      const stu = studentsMap[normSid];
      const attDay = attMap[normSid] || {};
      const clnDay = clnMap[normSid] || {};
      const subjDay = subjMap[normSid] || {};

      const sum = (dayMap) => {
        let t = 0, p = 0;
        Object.values(dayMap).forEach(d => { t += d.total; p += d.points; });
        return { total: t, points: p };
      };
      const att = sum(attDay), cln = sum(clnDay), subj = sum(subjDay);

      const attPct = att.total > 0 ? (att.points / att.total) * 100 : null;
      const clnPct = cln.total > 0 ? (cln.points / cln.total) * 100 : null;
      const subjPct = subj.total > 0 ? (subj.points / subj.total) * 100 : null;

      const avail = [];
      if (attPct !== null) avail.push({ pct: attPct, w: weights.attendance });
      if (clnPct !== null) avail.push({ pct: clnPct, w: weights.cleaning });
      if (subjPct !== null) avail.push({ pct: subjPct, w: weights.subject });

      const tw = avail.reduce((s, a) => s + a.w, 0);
      const composite = tw > 0 ? avail.reduce((s, a) => s + (a.pct * a.w / tw), 0) : 0;

      let maxStreak = 0, cur = 0;
      dateList.slice().sort().forEach(d => {
        const day = attDay[d];
        if (day && day.total > 0 && day.points === 0) { cur++; if (cur > maxStreak) maxStreak = cur; }
        else if (day && day.total > 0) { cur = 0; }
      });

      const dailyScores = {};
      dateList.forEach(d => {
        const item = {};
        if (attDay[d]) { item.attPts = attDay[d].points; item.attTotal = attDay[d].total; }
        if (clnDay[d]) { item.clnPts = clnDay[d].points; item.clnTotal = clnDay[d].total; }
        if (subjDay[d]) { item.subjPts = subjDay[d].points; item.subjTotal = subjDay[d].total; }

        const dayAvail = [];
        if (item.attTotal > 0) dayAvail.push({ pct: (item.attPts / item.attTotal) * 100, w: weights.attendance });
        if (item.clnTotal > 0) dayAvail.push({ pct: (item.clnPts / item.clnTotal) * 100, w: weights.cleaning });
        if (item.subjTotal > 0) dayAvail.push({ pct: (item.subjPts / item.subjTotal) * 100, w: weights.subject });

        if (dayAvail.length === 0) { dailyScores[d] = null; return; }

        const dtw = dayAvail.reduce((s, a) => s + a.w, 0);
        const daily = dayAvail.reduce((s, a) => s + (a.pct * a.w / dtw), 0);
        dailyScores[d] = {
          score: Math.round(daily),
          category: daily >= 85 ? 'green' : daily >= 60 ? 'yellow' : 'red',
          breakdown: {
            attendance: item.attTotal > 0 ? { points: +item.attPts.toFixed(1), total: item.attTotal } : null,
            cleaning: item.clnTotal > 0 ? { points: +item.clnPts.toFixed(1), total: item.clnTotal } : null,
            subject: item.subjTotal > 0 ? { points: +item.subjPts.toFixed(1), total: item.subjTotal } : null
          }
        };
      });

      let category = 'green';
      if (composite < 60 || maxStreak >= 3) category = 'red';
      else if (composite < 85 || maxStreak === 2) category = 'yellow';

      results.push({
        id: stu.id, name: stu.name,
        classroom_id: stu.classroom_id, classroom_name: stu.classroom_name,
        composite: Math.round(composite), category, streak: maxStreak,
        breakdown: {
          attendance: att.total > 0 ? { score: Math.round(attPct), total: att.total, points: +att.points.toFixed(1) } : null,
          cleaning: cln.total > 0 ? { score: Math.round(clnPct), total: cln.total, points: +cln.points.toFixed(1) } : null,
          subject: subj.total > 0 ? { score: Math.round(subjPct), total: subj.total, points: +subj.points.toFixed(1) } : null
        },
        dailyScores
      });
    });

    const order = { red: 0, yellow: 1, green: 2 };
    results.sort((a, b) => {
      if (a.category !== b.category) return order[a.category] - order[b.category];
      return a.composite - b.composite;
    });

    const summary = { red: 0, yellow: 0, green: 0, avgScore: 0, totalStudents: results.length };
    results.forEach(r => { summary[r.category]++; summary.avgScore += r.composite; });
    if (results.length > 0) summary.avgScore = Math.round(summary.avgScore / results.length);

    return {
      status: 'success',
      data: { period: { startDate, endDate, dates: dateList }, weights, summary, students: results }
    };
  } catch (err) {
    Logger.log('getStudentRiskAssessment error: ' + err.message);
    return { status: 'error', message: err.toString() };
  }
}

function getAssessmentHistoryForStudents(weeksBack, endDate) {
  const results = [];
  const weights = getRiskWeights().data;
  const endDateObj = new Date(endDate);

  for (let i = 1; i <= weeksBack; i++) {
    const weekEnd = new Date(endDateObj);
    weekEnd.setDate(weekEnd.getDate() - (i * 7));
    const weekStart = new Date(weekEnd);
    weekStart.setDate(weekEnd.getDate() - 6);

    const s = Utilities.formatDate(weekStart, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    const e = Utilities.formatDate(weekEnd, Session.getScriptTimeZone(), 'yyyy-MM-dd');

    try {
      const assess = getStudentRiskAssessment(s, e, '', JSON.stringify(weights));
      if (assess.status === 'success') {
        const byStudent = {};
        (assess.data.students || []).forEach(x => { byStudent[x.id] = x.category; });
        results.push({ weekStart: s, weekEnd: e, byStudent });
      } else {
        results.push({ weekStart: s, weekEnd: e, byStudent: {}, error: assess.message });
      }
    } catch (err) {
      results.push({ weekStart: s, weekEnd: e, byStudent: {}, error: err.message });
    }
  }
  return results;
}

const BEHAVIOR_ALERTS_SHEET = 'Behavior_Alerts';
const BEHAVIOR_RECORDS_SHEET = 'Behavior_Records';

function getBehaviorPeriod_() {
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.CONFIG);
  if (!sheet) throw new Error('ไม่พบชีต Config');

  const settings = {};
  sheet.getDataRange().getValues().slice(1).forEach(row => {
    const value = row[1] instanceof Date ? formatDateStr(row[1]) : String(row[1] || '').trim();
    settings[String(row[0] || '').trim()] = value;
  });
  const startDate = settings.behaviorStartDate;
  const endDate = settings.behaviorEndDate;
  const isValidDate = value => {
    const match = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value || '');
    if (!match) return false;
    const date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    return date.getFullYear() === Number(match[1]) &&
      date.getMonth() === Number(match[2]) - 1 &&
      date.getDate() === Number(match[3]);
  };

  if (!isValidDate(startDate) || !isValidDate(endDate) || startDate > endDate) {
    throw new Error('กรุณาตั้งค่า behaviorStartDate และ behaviorEndDate ในชีต Config เป็นวันที่ YYYY-MM-DD และให้วันเริ่มต้นไม่เกินวันสิ้นสุด');
  }
  return { startDate, endDate };
}

function normalizeBehaviorStudentId_(studentId) {
  return String(studentId == null ? '' : studentId).trim().replace(/^0+(?=\d)/, '');
}

function calculateBehaviorScores_(studentIds) {
  const period = getBehaviorPeriod_();
  const ids = [...new Set((studentIds || []).map(normalizeBehaviorStudentId_).filter(Boolean))];
  const scoreMap = {};
  ids.forEach(id => {
    scoreMap[id] = { score: 100, violations: [], attendedDates: {} };
  });

  const parseStudents = (raw, rowNumber, sourceName) => {
    try {
      const students = JSON.parse(raw || '[]');
      return Array.isArray(students) ? students : [];
    } catch (error) {
      Logger.log(`ข้อมูลนักเรียนใน ${sourceName} แถว ${rowNumber} ไม่ใช่ JSON ที่ถูกต้อง: ${error.message}`);
      return [];
    }
  };

  const attendanceRows = getCleanedAttendanceRows(
    getOrCreateSheet(SHEET_NAMES.ATTENDANCE).getDataRange().getValues()
  );
  attendanceRows.forEach((row, index) => {
    const date = formatDateStr(row[0]);
    if (!date || date < period.startDate || date > period.endDate) return;
    parseStudents(row[4], index + 2, 'Attendance').forEach(student => {
      const id = normalizeBehaviorStudentId_(student.student_id);
      const record = scoreMap[id];
      if (!record) return;
      const status = String(student.status || '').trim();
      if (status === 'มา') record.attendedDates[date] = true;
      if (status === 'ขาด' || status === 'สาย') {
        record.score -= 1;
        record.violations.push({
          date, type: status === 'ขาด' ? 'ไม่เข้าแถว' : 'เข้าแถวสาย', deduct: -1
        });
      }
    });
  });

  const cleaningRows = getCleanedCleaningRows(
    getOrCreateSheet(CLEANING_SHEET_NAMES.ATTENDANCE).getDataRange().getValues()
  );
  cleaningRows.forEach((row, index) => {
    const date = formatDateStr(row[0]);
    if (!date || date < period.startDate || date > period.endDate) return;
    parseStudents(row[3], index + 2, 'Cleaning_Attendance').forEach(student => {
      const id = normalizeBehaviorStudentId_(student.student_id);
      const record = scoreMap[id];
      if (!record || String(student.status || '').trim() !== 'ไม่ทำ') return;
      record.score -= 2;
      record.violations.push({ date, type: 'ไม่ทำเวรเขตพื้นที่', deduct: -2 });
    });
  });

  const subjectSheet = getSpreadsheet().getSheetByName('Attendance_Subjects');
  const subjectRows = subjectSheet ? subjectSheet.getDataRange().getValues() : [];
  for (let i = 1; i < subjectRows.length; i++) {
    const row = subjectRows[i];
    const date = formatDateStr(row[1]);
    if (!date || date < period.startDate || date > period.endDate) continue;
    const record = scoreMap[normalizeBehaviorStudentId_(row[7])];
    if (!record) continue;

    const status = String(row[9] || '').trim();
    const subjectName = String(row[3] || row[2] || '').trim();
    if (status === 'ขาด') {
      record.score -= 3;
      const skipClass = record.attendedDates[date] ? ' (โดดร่ม)' : '';
      record.violations.push({
        date, type: `ขาดเรียนวิชา ${subjectName}${skipClass}`, deduct: -3
      });
    } else if (status === 'สาย') {
      record.score -= 1;
      record.violations.push({ date, type: `เข้าเรียนสาย ${subjectName}`, deduct: -1 });
    }
  }

  Object.keys(scoreMap).forEach(id => {
    scoreMap[id].violations.sort((a, b) => b.date.localeCompare(a.date));
    delete scoreMap[id].attendedDates;
    scoreMap[id].period = period;
  });
  return scoreMap;
}

function calculateBehaviorScore_(studentId) {
  const id = normalizeBehaviorStudentId_(studentId);
  if (!id) throw new Error('ไม่พบรหัสนักเรียน');
  return calculateBehaviorScores_([id])[id];
}

function getStudentBehaviorScore(studentId, password) {
  try {
    const login = loginStudent(studentId, password);
    if (login.status !== 'success') throw new Error(login.message);
    const student = login.data.student;
    const result = calculateBehaviorScore_(student.student_id);
    return { status: 'success', ...result };
  } catch (error) {
    return { status: 'error', message: error.message || String(error) };
  }
}

function getBehaviorAlertsSheet_() {
  const spreadsheet = getSpreadsheet();
  let sheet = spreadsheet.getSheetByName(BEHAVIOR_ALERTS_SHEET);
  if (!sheet) sheet = spreadsheet.insertSheet(BEHAVIOR_ALERTS_SHEET);
  if (sheet.getLastRow() === 0 || !sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, 5).setValues([[
      'student_id', 'period_start', 'period_end', 'threshold', 'notified_at'
    ]]);
  }
  return sheet;
}

function sendLineBehaviorAlert_(message) {
  const properties = PropertiesService.getScriptProperties();
  const token = properties.getProperty('LINE_ACCESS_TOKEN');
  const recipientId = properties.getProperty('LINE_RECIPIENT_ID');
  if (!token || !recipientId) {
    throw new Error('กรุณาตั้งค่า LINE_ACCESS_TOKEN และ LINE_RECIPIENT_ID ใน Script Properties');
  }

  const response = UrlFetchApp.fetch('https://api.line.me/v2/bot/message/push', {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: `Bearer ${token}` },
    payload: JSON.stringify({
      to: recipientId,
      messages: [{ type: 'text', text: message }]
    }),
    muteHttpExceptions: true
  });
  const responseCode = response.getResponseCode();
  if (responseCode < 200 || responseCode >= 300) {
    throw new Error(`LINE Messaging API ส่งข้อความไม่สำเร็จ (HTTP ${responseCode}): ${response.getContentText()}`);
  }
}

function dailyBehaviorCheck() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('มีการตรวจคะแนนความประพฤติทำงานอยู่ กรุณาลองใหม่ภายหลัง');
  try {
    const studentsResult = getAllStudents();
    if (studentsResult.status !== 'success') {
      throw new Error(studentsResult.message || 'ไม่สามารถโหลดรายชื่อนักเรียนได้');
    }
    const students = Object.values(studentsResult.data).flat();
    const scoreMap = calculateBehaviorScores_(students.map(student => student.student_id));
    const period = getBehaviorPeriod_();
    const sheet = getBehaviorAlertsSheet_();
    const rows = sheet.getDataRange().getValues();
    const notified = new Set(rows.slice(1).map(row =>
      `${normalizeBehaviorStudentId_(row[0])}|${formatDateStr(row[1])}|${formatDateStr(row[2])}|${Number(row[3])}`
    ));
    const hasPendingAlert = students.some(student => {
      const id = normalizeBehaviorStudentId_(student.student_id);
      const score = scoreMap[id] && scoreMap[id].score;
      return score <= 80 && (!notified.has(`${id}|${period.startDate}|${period.endDate}|80`) ||
        (score <= 60 && !notified.has(`${id}|${period.startDate}|${period.endDate}|60`)));
    });
    if (hasPendingAlert) {
      const properties = PropertiesService.getScriptProperties();
      if (!properties.getProperty('LINE_ACCESS_TOKEN') || !properties.getProperty('LINE_RECIPIENT_ID')) {
        throw new Error('กรุณาตั้งค่า LINE_ACCESS_TOKEN และ LINE_RECIPIENT_ID ใน Script Properties');
      }
    }
    const errors = [];
    let sent = 0;

    students.forEach(student => {
      const id = normalizeBehaviorStudentId_(student.student_id);
      const result = scoreMap[id];
      if (!result || result.score > 80) return;

      const keyFor = threshold => `${id}|${period.startDate}|${period.endDate}|${threshold}`;
      const has80 = notified.has(keyFor(80));
      const has60 = notified.has(keyFor(60));
      if (has80 && (result.score > 60 || has60)) return;

      const threshold = result.score <= 60 ? 60 : 80;
      const message = [
        '🚨 แจ้งเตือนความประพฤติ',
        `นักเรียน: ${student.name} (${student.classroom_id})`,
        `คะแนนคงเหลือ: ${result.score}/100`,
        `สถานะ: ${result.score <= 60 ? 'วิกฤต (ต้องทำกิจกรรม)' : 'เฝ้าระวัง'}`,
        'กรุณาเข้าตรวจสอบในระบบ'
      ].join('\n');

      try {
        sendLineBehaviorAlert_(message);
        const now = new Date();
        const thresholds = result.score <= 60 && !has80 && !has60 ? [80, 60] : [threshold];
        thresholds.forEach(value => {
          sheet.appendRow([student.student_id, period.startDate, period.endDate, value, now]);
          notified.add(keyFor(value));
        });
        sent++;
      } catch (error) {
        const failure = `${student.student_id} (${student.name}): ${error.message}`;
        Logger.log('ส่ง LINE แจ้งเตือนไม่สำเร็จ: ' + failure);
        errors.push(failure);
      }
    });

    if (errors.length) throw new Error(`ส่ง LINE แจ้งเตือนไม่สำเร็จ ${errors.length} รายการ: ${errors.join('; ')}`);
    return { status: 'success', checked: students.length, sent };
  } finally {
    lock.releaseLock();
  }
}

function installDailyBehaviorTrigger() {
  const existing = ScriptApp.getProjectTriggers()
    .filter(trigger => trigger.getHandlerFunction() === 'dailyBehaviorCheck');
  existing.forEach(trigger => ScriptApp.deleteTrigger(trigger));
  ScriptApp.newTrigger('dailyBehaviorCheck').timeBased().everyDays(1).atHour(16).nearMinute(0).create();
  Logger.log('ติดตั้ง Trigger ตรวจคะแนนประจำวันประมาณ 16:00 น. แล้ว');
  return { status: 'success', message: 'ติดตั้ง Trigger ตรวจคะแนนประจำวันประมาณ 16:00 น. แล้ว' };
}

// ═══════════════════════════════════════════════════════════════
// 15. DISCIPLINE WORKFLOW
// ═══════════════════════════════════════════════════════════════

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
    'notes','created_at','updated_at','trigger_data',
    'approval_status','approved_at','approved_by','approval_note'
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
    trigger_data: safeJson(row[37]),
    approval_status: String(row[38] || ''),
    approved_at: String(row[39] || ''),
    approved_by: String(row[40] || ''),
    approval_note: String(row[41] || '')
  };
}

function getDisciplineConfig() {
  const sheet = getOrCreateSheet(SHEET_NAMES.CONFIG);
  const data = sheet.getDataRange().getValues();
  const cfg = {
    school_name: 'โรงเรียน',
    director_name: '', director_position: 'ผู้อำนวยการโรงเรียน',
    discipline_teacher: '', garuda_url: '', garuda_base64: '',
    warnings: []
  };

  data.slice(1).forEach(r => {
    const k = String(r[0] || '').trim(), v = String(r[1] || '').trim();
    if (k === 'schoolName') cfg.school_name = v || cfg.school_name;
    if (k === 'director_name') cfg.director_name = v;
    if (k === 'director_position') cfg.director_position = v || cfg.director_position;
    if (k === 'discipline_teacher') cfg.discipline_teacher = v;
    if (k === 'garuda_base64') cfg.garuda_base64 = v;
  });

  // ✅ ดึงลิงก์ตราครุฑจากเซลล์ B14 โดยเฉพาะ
  try {
    const garudaUrlB14 = String(sheet.getRange('B14').getValue() || '').trim();
    if (garudaUrlB14) {
      cfg.garuda_url = garudaUrlB14;
    }
  } catch (e) {
    Logger.log('อ่านค่า B14 ไม่สำเร็จ: ' + e.message);
  }

  if (!cfg.director_name || cfg.director_name.includes('...')) {
    cfg.warnings.push('⚠️ ยังไม่ได้ตั้งค่า director_name');
  }
  if (!cfg.discipline_teacher || cfg.discipline_teacher.includes('...')) {
    cfg.warnings.push('⚠️ ยังไม่ได้ตั้งค่า discipline_teacher');
  }

  if (!cfg.director_name) cfg.director_name = '.....................................................';
  if (!cfg.discipline_teacher) cfg.discipline_teacher = '.....................................................';

  return cfg;
}

function getClassroomAdvisor(classroomId) {
  try {
    const teachers = getTeachers().data || [];
    const t = teachers.find(t => (t.classroom_ids || []).map(String).includes(String(classroomId)));
    return t ? { name: t.name, phone: String(t.phone || '').trim() }
             : { name: '.....................................................', phone: '' };
  } catch (e) {
    return { name: '.....................................................', phone: '' };
  }
}

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
      if (a.stage !== b.stage) return (order[a.stage] || 9) - (order[b.stage] || 9);
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
      data: { case: caseData, logs, config: cfg, advisor: getClassroomAdvisor(caseData.classroom_id) }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function addDisciplineLog(caseId, action, stageFrom, stageTo, actor, note) {
  try {
    const logSheet = getDisciplineSheet('LOGS');
    const logId = 'L' + Date.now() + Math.floor(Math.random() * 1000);
    logSheet.appendRow([logId, caseId, action, stageFrom || '', stageTo || '', actor || '', note || '', new Date().toISOString()]);
  } catch(e) { Logger.log('Log error: ' + e.message); }
}

function syncDisciplineFromWeekly(startDate, endDate, actor) {
  try {
    const current = getStudentRiskAssessment(startDate, endDate, '', '');
    if (current.status !== 'success') throw new Error(current.message);
    const students = current.data.students || [];

    const existing = getDisciplineCases({}).data.cases;
    const activeCases = existing.filter(c => c.stage !== 'closed');

    const created = [], advanced = [];
    const year = new Date().getFullYear() + 543;
    let maxNum = 0;
    existing.forEach(c => {
      const m = c.case_id.match(/D-(\d+)-(\d+)/);
      if (m && +m[1] === year) maxNum = Math.max(maxNum, +m[2]);
    });

   // const hasStage2 = activeCases.some(c => c.stage === '2');
   // const history = hasStage2 ? getAssessmentHistoryForStudents(2, endDate) : [];

    students.forEach(s => {
      const active = activeCases.find(c => c.student_id === s.id);
      const now = new Date().toISOString();

      if (!active) {
        if (s.category === 'red' || s.category === 'yellow') {
          maxNum++;
          const caseId = `D-${year}-${String(maxNum).padStart(3, '0')}`;
          const trigger = s.category === 'red' ? 'red_1' : 'yellow';
          const triggerData = {
            weekStart: startDate, weekEnd: endDate,
            category: s.category, composite: s.composite, streak: s.streak || 0,
            breakdown: s.breakdown
          };

          const row = new Array(42).fill('');
          row[0] = caseId; row[1] = s.id; row[2] = s.name;
          row[3] = s.classroom_id; row[4] = s.classroom_name;
          row[5] = '1'; row[6] = 'active'; row[7] = trigger;
          row[35] = now; row[36] = now;
          row[37] = JSON.stringify(triggerData);

          getDisciplineSheet('CASES').appendRow(row);
          addDisciplineLog(caseId, 'created', '', '1', actor || 'system',
            `สร้างเคสอัตโนมัติ (${trigger}, ${s.composite}%)`);
          created.push({ case_id: caseId, student_name: s.name });
        }
        return;
      }

      // Stage 1 → 2
      if (active.stage === '1') {
        const createdDate = new Date(active.created_at);
        const daysSince = Math.floor((Date.now() - createdDate.getTime()) / (1000 * 60 * 60 * 24));
        if (daysSince >= 7 && (s.category === 'red' || s.category === 'yellow')) {
          advanceDisciplineStage(active.case_id, '2',
            `ฝ่าฝืนซ้ำหลังตักเตือน ${daysSince} วัน (${s.category}, ${s.composite}%)`,
            actor || 'system');
          advanced.push({ case_id: active.case_id, from: '1', to: '2' });
        }
        return;
      }
          // ===== STAGE 2 → 3 =====
      if (active.stage === '2') {
        if (s.category !== 'red') return; // สัปดาห์นี้ไม่ Red → ยังไม่ escalate

        // ✅ เปลี่ยนมาเช็คจากวันที่สร้างเคสแทน (เร็วกว่า ไม่ต้องคำนวณย้อนหลัง)
        const createdDate = new Date(active.created_at);
        const daysSince = Math.floor((Date.now() - createdDate.getTime()) / (1000 * 60 * 60 * 24));
        
        // ถ้าอยู่ใน Stage 2 มาแล้วเกิน 14 วัน (ประมาณ 2 สัปดาห์) และยังแดงอยู่ → ยกระดับ
        if (daysSince >= 14) {
          advanceDisciplineStage(active.case_id, '3',
            `อยู่ในกลุ่มใกล้ชิดต่อเนื่องเกิน 14 วัน (${s.composite}%)`,
            actor || 'system');
          advanced.push({
            case_id: active.case_id, from: '2', to: '3',
            reason: 'อยู่ในกลุ่มใกล้ชิดเกิน 2 สัปดาห์'
          });
        }
        return;
      }
    });

    return { status: 'success', data: { created, advanced, totalProcessed: students.length } };
  } catch (err) {
    Logger.log('syncDisciplineFromWeekly error: ' + err.message);
    return { status: 'error', message: err.toString(), data: { created: [], advanced: [] } };
  }
}

function advanceDisciplineStage(caseId, newStage, note, actor) {
  try {
    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    const idx = data.slice(1).findIndex(r => String(r[0]) === caseId) + 2;
    if (idx <= 1) return { status: 'error', message: 'ไม่พบเคส' };

    const row = data[idx - 1];
    const oldStage = String(row[5]);
    const now = new Date().toISOString();

    // ✅ LOCKS
    if (oldStage === '1' && newStage === '2') {
      const createdDate = row[35] ? new Date(row[35]) : new Date();
      const daysSince = Math.floor((Date.now() - createdDate.getTime()) / (1000 * 60 * 60 * 24));
      if (daysSince < 7) {
        return { status: 'error', code: 'TOO_EARLY',
          message: `ต้องรออย่างน้อย 7 วันหลังสร้างเคส (ผ่านมา ${daysSince} วัน)` };
      }
    }

    if (oldStage === '2' && newStage === '3') {
      const missing = [];
      if (!String(row[25] || '').trim()) missing.push('ผู้อำนวยการ');
      if (!String(row[26] || '').trim()) missing.push('ผู้ปกครอง');
      if (!String(row[27] || '').trim()) missing.push('นักเรียน');
      if (missing.length > 0) {
        return { status: 'error', code: 'SIGNATURES_INCOMPLETE',
          message: `ยังขาดการลงนาม: ${missing.join(', ')}` };
      }
    }

    if (oldStage === '3' && newStage === '4') {
      if (String(row[21] || '').trim() !== 'fail') {
        return { status: 'error', code: 'ACTIVITY_NOT_FAILED',
          message: 'ต้องบันทึกผลกิจกรรมเป็น "ไม่ผ่าน" ก่อน' };
      }
    }

    if (oldStage === '3' && newStage === 'closed') {
      if (String(row[21] || '').trim() !== 'pass') {
        return { status: 'error', code: 'ACTIVITY_NOT_PASSED',
          message: 'ต้องบันทึกผลกิจกรรมเป็น "ผ่าน" ก่อน' };
      }
    }

    sheet.getRange(idx, 6).setValue(newStage);
    sheet.getRange(idx, 7).setValue(newStage === 'closed' ? 'closed' : 'active');
    sheet.getRange(idx, 37).setValue(now);

    if (newStage === '3' && !String(row[20] || '').trim()) {
      const start = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const end = Utilities.formatDate(new Date(Date.now() + 14 * 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const actData = {
        items: [
          { id: 'a1', label: 'บำเพ็ญประโยชน์', unit: 'ชั่วโมง', target: 10, progress: 0 },
          { id: 'a2', label: 'ทำความสะอาดเขตพื้นที่', unit: 'ครั้ง', target: 3, progress: 0 },
          { id: 'a3', label: 'รับคำปรึกษา', unit: 'ครั้ง', target: 2, progress: 0 }
        ],
        start, end
      };
      sheet.getRange(idx, 21).setValue(JSON.stringify(actData));
      sheet.getRange(idx, 22).setValue('pending');
    }

    if (newStage === '2' && !String(row[12] || '').trim()) {
      const start = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const end = Utilities.formatDate(new Date(Date.now() + 90 * 86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
      sheet.getRange(idx, 13).setValue(start);
      sheet.getRange(idx, 14).setValue(end);
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
    const colMap = { director: 26, parent: 27, student: 28, witness1: 29, witness2: 30 };
    const col = colMap[party];
    if (!col) return { status: 'error', message: 'party ไม่ถูกต้อง' };

    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    const idx = data.slice(1).findIndex(r => String(r[0]) === caseId) + 2;
    if (idx <= 1) return { status: 'error', message: 'ไม่พบเคส' };

    const now = new Date().toISOString();
    sheet.getRange(idx, col).setValue(now);

    if (party === 'parent') {
      if (meta.name) sheet.getRange(idx, 19).setValue(meta.name);
      if (meta.relation) sheet.getRange(idx, 20).setValue(meta.relation);
    }

    sheet.getRange(idx, 37).setValue(now);
    addDisciplineLog(caseId, 'signed', '', '', party, `${party} signed`);

    const row = sheet.getRange(idx, 1, 1, 37).getValues()[0];
    if (String(row[5]) === '2') {
      const s1 = String(row[25] || '').trim();
      const s2 = String(row[26] || '').trim();
      const s3 = String(row[27] || '').trim();
      if (s1 && s2 && s3) {
        sheet.getRange(idx, 7).setValue('active');
        addDisciplineLog(caseId, 'signatures_complete', '', '', 'system', 'ครบ 3 ฝ่าย');
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

    sheet.getRange(idx, 22).setValue(result);
    sheet.getRange(idx, 37).setValue(new Date().toISOString());
    addDisciplineLog(caseId, 'activity_' + result, '', '', 'system', `กิจกรรม ${result}`);
    return { status: 'success', message: result === 'pass' ? 'ผ่านกิจกรรม' : 'ไม่ผ่านกิจกรรม' };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function markDocPrinted(caseId, docType) {
  try {
    const colMap = { warning: 31, bond: 32, activity: 33, parent: 34 };
    const col = colMap[docType];
    if (!col) return { status: 'error' };
    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    const idx = data.slice(1).findIndex(r => String(r[0]) === caseId) + 2;
    if (idx <= 1) return { status: 'error' };
    sheet.getRange(idx, col).setValue(new Date().toISOString());
    return { status: 'success' };
  } catch (err) { return { status: 'error', message: err.toString() }; }
}

function saveMeetingMeta(caseId, dateStr, timeStr, location) {
  try {
    if (!caseId) return { status: 'error', message: 'ไม่พบ caseId' };
    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    const idx = data.slice(1).findIndex(r => String(r[0]) === caseId) + 2;
    if (idx <= 1) return { status: 'error', message: 'ไม่พบเคส' };

    if (String(data[idx - 1][5]) !== '4') {
      return { status: 'error', message: 'ใช้ได้เฉพาะ Stage 4' };
    }

    const safeDate = /^\d{4}-\d{2}-\d{2}$/.test(dateStr || '') ? dateStr : '';
    const safeTime = String(timeStr || '').trim().substring(0, 20);
    const safeLoc = String(location || '').trim().substring(0, 200);

    sheet.getRange(idx, 23).setValue(safeDate);
    sheet.getRange(idx, 24).setValue(safeTime);
    sheet.getRange(idx, 25).setValue(safeLoc);
    sheet.getRange(idx, 37).setValue(new Date().toISOString());

    addDisciplineLog(caseId, 'meeting_scheduled', '', '', 'system',
      `นัดประชุม ${safeDate || '?'} ${safeTime || ''} @ ${safeLoc || '?'}`);

    return { status: 'success', message: 'บันทึกการนัดประชุมสำเร็จ',
      data: { meeting_date: safeDate, meeting_time: safeTime, meeting_location: safeLoc } };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════
// 16. APPROVAL WORKFLOW
// ═══════════════════════════════════════════════════════════════

function submitForApproval(caseId, docType) {
  try {
    if (!['bond','activity','parent'].includes(docType)) {
      return { status: 'success', message: 'ไม่ต้องอนุมัติ', data: { required: false } };
    }
    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    const idx = data.slice(1).findIndex(r => String(r[0]) === caseId) + 2;
    if (idx <= 1) return { status: 'error', message: 'ไม่พบเคส' };

    const row = data[idx - 1];
    if (String(row[38] || '').trim() === 'approved') {
      return { status: 'success', message: 'อนุมัติแล้ว', data: { required: true, status: 'approved' } };
    }

    sheet.getRange(idx, 39).setValue('pending');
    sheet.getRange(idx, 42).setValue(new Date().toISOString());
    addDisciplineLog(caseId, 'submitted_for_approval', '', '', 'system', docType);

    return { status: 'success', message: 'ส่งขออนุมัติแล้ว', data: { required: true, status: 'pending' } };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function approveDocument(caseId, decision, note, directorUsername) {
  try {
    requirePermission(directorUsername, ['director']);

    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();
    const idx = data.slice(1).findIndex(r => String(r[0]) === caseId) + 2;
    if (idx <= 1) return { status: 'error', message: 'ไม่พบเคส' };

    sheet.getRange(idx, 39).setValue(decision);
    sheet.getRange(idx, 40).setValue(new Date().toISOString());
    sheet.getRange(idx, 41).setValue(String(directorUsername || '').trim());
    sheet.getRange(idx, 42).setValue(String(note || '').trim());

    addDisciplineLog(caseId, 'approval_' + decision, '', '', directorUsername, note || '');

    return { status: 'success',
      message: decision === 'approved' ? '✅ อนุมัติแล้ว' : '❌ ปฏิเสธแล้ว',
      data: { decision, note } };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function getPendingApprovals(directorUsername) {
  try {
    requirePermission(directorUsername, ['director']);
    const sheet = getDisciplineSheet('CASES');
    const data = sheet.getDataRange().getValues();

    const pending = data.slice(1)
      .filter(r => String(r[38] || '').trim() === 'pending')
      .map(rowToCaseObj)
      .sort((a, b) => new Date(b.updated_at) - new Date(a.updated_at));

    return { status: 'success', data: pending };
  } catch (err) {
    return { status: 'error', message: err.toString(), data: [] };
  }
}

// ═══════════════════════════════════════════════════════════════
// 17. STUDENT SELF REPORT
// ═══════════════════════════════════════════════════════════════

function getStudentSelfReport(studentId, password, month) {
  try {
    const login = loginStudent(studentId, password);
    if (login.status !== 'success') throw new Error(login.message);

    const stu = login.data.student;
    const targetMonth = month || Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');

    const [y, m] = targetMonth.split('-').map(Number);
    const startDate = `${y}-${String(m).padStart(2,'0')}-01`;
    const lastDay = new Date(y, m, 0).getDate();
    const endDate = `${y}-${String(m).padStart(2,'0')}-${String(lastDay).padStart(2,'0')}`;

    // ✅ ฟังก์ชันสำหรับตัดเลข 0 ข้างหน้าออกก่อนเปรียบเทียบ (แก้ปัญหารหัส 01383 กับ 1383)
    const normId = (id) => String(id || '').replace(/^0+/, '');

    // 1. ดึงข้อมูลการเช็คแถว
    const attData = getCleanedAttendanceRows(getOrCreateSheet(SHEET_NAMES.ATTENDANCE).getDataRange().getValues());
    const attDaily = {};
    attData.forEach(row => {
      const d = formatDateStr(row[0]);
      if (d < startDate || d > endDate) return;
      if (String(row[2]).trim() !== stu.classroom_id) return;
      let students = [];
      try { students = JSON.parse(row[4] || '[]'); } catch(e) {}
      // ✅ ใช้ normId ในการเทียบรหัส
      const found = students.find(s => normId(s.student_id) === normId(stu.student_id));
      if (found) {
        if (!attDaily[d]) attDaily[d] = { attendance: null, cleaning: null, subject: null };
        attDaily[d].attendance = { status: found.status, reason: found.reason || '' };
      }
    });

    // 2. ดึงข้อมูลเขตพื้นที่
    const clnData = getCleanedCleaningRows(getOrCreateSheet(CLEANING_SHEET_NAMES.ATTENDANCE).getDataRange().getValues());
    clnData.forEach(row => {
      const d = formatDateStr(row[0]);
      if (d < startDate || d > endDate) return;
      if (String(row[1]).trim() !== stu.classroom_id) return;
      let students = [];
      try { students = JSON.parse(row[3] || '[]'); } catch(e) {}
      // ✅ ใช้ normId ในการเทียบรหัส
      const found = students.find(s => normId(s.student_id) === normId(stu.student_id));
      if (found) {
        if (!attDaily[d]) attDaily[d] = { attendance: null, cleaning: null, subject: null };
        attDaily[d].cleaning = { status: found.status, role: found.role || '' };
      }
    });

    // 3. ดึงข้อมูลรายวิชา
    const subjSheet = getSpreadsheet().getSheetByName('Attendance_Subjects');
    const subjSummary = {};
    if (subjSheet) {
      const subjData = subjSheet.getDataRange().getValues();
      for (let i = 1; i < subjData.length; i++) {
        const row = subjData[i];
        const d = formatDateStr(row[1]);
        if (d < startDate || d > endDate) continue;
        // ✅ ใช้ normId ในการเทียบรหัส
        if (normId(String(row[7]).trim()) !== normId(stu.student_id)) continue;

        const code = String(row[2] || '').trim();
        const name = String(row[3] || '').trim();
        const key = code || name;
        const status = String(row[9] || '').trim();

        if (!subjSummary[key]) {
          subjSummary[key] = { code, name, present:0, absent:0, late:0, leave:0, sick:0, total:0 };
        }
        if (status === 'มา') subjSummary[key].present++;
        else if (status === 'ขาด') subjSummary[key].absent++;
        else if (status === 'สาย') subjSummary[key].late++;
        else if (status === 'ลา') subjSummary[key].leave++;
        else if (status === 'ป่วย') subjSummary[key].sick++;
        subjSummary[key].total++;

        if (!attDaily[d]) attDaily[d] = { attendance: null, cleaning: null, subject: null };
        if (!attDaily[d].subject) attDaily[d].subject = [];
        attDaily[d].subject.push({ code, name, status });
      }
    }

    // 4. คำนวณคะแนนภาพรวม
    let presentCount = 0, totalCount = 0;
    Object.values(attDaily).forEach(d => {
      if (d.attendance) {
        const st = d.attendance.status;
        if (st === 'มา') presentCount += 1;
        else if (st === 'สาย') presentCount += 0.75;
        else if (st === 'ลา' || st === 'ป่วย') presentCount += 0.5;
        totalCount += 1;
      }
    });
    const composite = totalCount > 0 ? Math.round((presentCount / totalCount) * 100) : 0;

    // 5. ดึงข้อมูลวินัย
    let disciplineCases = [];
    try {
      const dData = getDisciplineSheet('CASES').getDataRange().getValues();
      disciplineCases = dData.slice(1)
        .filter(r => normId(r[1]) === normId(stu.student_id)) // ✅ ใช้ normId
        .map(r => ({
          case_id: String(r[0]), stage: String(r[5]), status: String(r[6]),
          created_at: String(r[35]), approval_status: String(r[38] || '')
        }));
    } catch(e) { /* ignore */ }

    return {
      status: 'success',
      data: {
        student: stu, month: targetMonth, composite,
        daily: attDaily,
        subjects: Object.values(subjSummary).sort((a,b) => (a.code||a.name).localeCompare(b.code||b.name)),
        discipline: disciplineCases
      }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════
// 18. MIGRATION
// ═══════════════════════════════════════════════════════════════

function setupDisciplineConfig() {
  try {
    const sheet = getOrCreateSheet(SHEET_NAMES.CONFIG);
    const data = sheet.getDataRange().getValues();
    const existingKeys = new Set();
    for (let i = 1; i < data.length; i++) {
      const k = String(data[i][0] || '').trim();
      if (k) existingKeys.add(k);
    }

    const defaults = [
      ['schoolName', 'โรงเรียนน้ำคำวิทยา'],
      ['director_name', 'นาย.....................'],
      ['director_position', 'ผู้อำนวยการโรงเรียน'],
      ['discipline_teacher', 'นาย.....................'],
      ['garuda_url', ''],
      ['garuda_base64', ''],
      ['behaviorStartDate', ''],
      ['behaviorEndDate', ''],
      ['docCounter_' + (new Date().getFullYear() + 543), 0]
    ];

    const toAdd = defaults.filter(([k]) => !existingKeys.has(k));
    toAdd.forEach(([k, v]) => sheet.appendRow([k, v]));

    const casesSheet = getDisciplineSheet('CASES');
    const maxCol = Math.max(42, casesSheet.getLastColumn());
    const headerRow = casesSheet.getRange(1, 1, 1, maxCol).getValues()[0];
    if (String(headerRow[37] || '').trim() !== 'trigger_data') {
      casesSheet.getRange(1, 38).setValue('trigger_data');
    }

    Logger.log(`✅ setupDisciplineConfig: +${toAdd.length} แถว`);
    return { status: 'success', message: `เพิ่ม ${toAdd.length} แถว`, added: toAdd.map(r => r[0]) };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

function migrateSprint1to4() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) return { status: 'error', message: 'Server busy' };

  try {
    const ss = getSpreadsheet();
    const log = [];

    // Teachers
    const tSheet = ss.getSheetByName('Teachers');
    if (tSheet) {
      const tData = tSheet.getDataRange().getValues();
      const oldHeaders = tData[0] || [];
      const hasPhoneAtG = String(oldHeaders[6] || '').toLowerCase() === 'phone';
      const hasImageAtH = String(oldHeaders[7] || '').toLowerCase() === 'imageurl';

      if (hasPhoneAtG && !hasImageAtH) {
        const phoneVals = tData.slice(1).map(r => [r[6] || '']);
        if (phoneVals.length > 0) {
          tSheet.getRange(2, 8, phoneVals.length, 1).setValues(phoneVals);
        }
        tSheet.getRange(2, 7, Math.max(phoneVals.length, 1), 1).clearContent();
        log.push('✅ ย้าย phone: G → H');
      }

      const newHeaders = ['id','name','username','password','status','classroom_ids','imageUrl','phone','role'];
      tSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders])
        .setFontWeight('bold').setBackground('#1e40af').setFontColor('#fff');

      const dataAfter = tSheet.getDataRange().getValues();
      for (let i = 1; i < dataAfter.length; i++) {
        if (!String(dataAfter[i][2] || '').trim()) continue;
        if (!String(dataAfter[i][8] || '').trim()) tSheet.getRange(i + 1, 9).setValue('teacher');
      }
      log.push(`✅ Teachers headers + role (${dataAfter.length - 1} แถว)`);
    }

    // Students
    const sSheet = ss.getSheetByName('Students');
    if (sSheet) {
      const newHeaders = ['classroom_id','student_id','no','name','status','password','role','council_classroom_ids'];
      sSheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders])
        .setFontWeight('bold').setBackground('#1e40af').setFontColor('#fff');

      const sData = sSheet.getDataRange().getValues();
      for (let i = 1; i < sData.length; i++) {
        if (!String(sData[i][1] || '').trim()) continue;
        if (!String(sData[i][5] || '').trim()) sSheet.getRange(i + 1, 6).setValue(String(sData[i][1]).trim());
        if (!String(sData[i][6] || '').trim()) sSheet.getRange(i + 1, 7).setValue('student');
      }
      log.push(`✅ Students password + role (${sData.length - 1} แถว)`);
    }

    // Discipline_Cases
    const dSheet = getDisciplineSheet('CASES');
    const dMaxCol = Math.max(42, dSheet.getLastColumn());
    const dHeaders = dSheet.getRange(1, 1, 1, dMaxCol).getValues()[0];

    const newCols = [
      { idx: 38, name: 'approval_status' },
      { idx: 39, name: 'approved_at' },
      { idx: 40, name: 'approved_by' },
      { idx: 41, name: 'approval_note' }
    ];
    newCols.forEach(col => {
      if (String(dHeaders[col.idx - 1] || '').trim() !== col.name) {
        dSheet.getRange(1, col.idx).setValue(col.name)
          .setFontWeight('bold').setBackground('#1e40af').setFontColor('#fff');
        log.push(`✅ Discipline: +${col.name}`);
      }
    });

    // Config flag
    const cfgSheet = getOrCreateSheet(SHEET_NAMES.CONFIG);
    const cfgData = cfgSheet.getDataRange().getValues();
    const cfgKeys = new Set();
    cfgData.slice(1).forEach(r => { if (r[0]) cfgKeys.add(String(r[0]).trim()); });
    if (!cfgKeys.has('sprint1to4_migrated')) {
      cfgSheet.appendRow(['sprint1to4_migrated', new Date().toISOString()]);
      log.push('✅ Config flag');
    }

    Logger.log(log.join('\n'));
    return { status: 'success', message: 'Migration สำเร็จ', log };
  } catch (err) {
    Logger.log('Migration error: ' + err.message);
    return { status: 'error', message: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// ═══════════════════════════════════════════════════════════════
// 19. DISCIPLINE DOCUMENT BUILDERS
// ═══════════════════════════════════════════════════════════════

function buildDocStyle() {
  return `<style>
    /* ตั้งค่ากระดาษ A4 */
    @page { size: A4; margin: 2cm 2cm 2cm 2.5cm; }
    
    /* ฟอนต์มาตรฐานราชการ */
    body { 
      font-family: 'TH Sarabun New', 'TH SarabunPSK', 'Sarabun', sans-serif; 
      font-size: 16pt; 
      line-height: 1.5; /* เพิ่มระยะบรรทัดให้อ่านง่าย */
      margin: 0; 
      color: #000; 
    }
    
    .toolbar { position: fixed; top: 10px; right: 10px; background: #fff; padding: 8px; border-radius: 8px; box-shadow: 0 2px 10px rgba(0,0,0,.15); z-index: 999; }
    .toolbar button { padding: 6px 14px; border: 1px solid #ccc; border-radius: 6px; cursor: pointer; font-weight: bold; }
    @media print { .toolbar { display: none; } }
    
    /* จัดตำแหน่งครุฑและหัวกระดาษ */
    .garuda-wrap { text-align: center; margin-bottom: 8px; min-height: 60px; }
    .header { text-align: center; font-size: 20pt; font-weight: bold; margin-bottom: 4px; }
    .header-sub { text-align: center; font-size: 16pt; margin-bottom: 24px; }
    
    /* การจัดวางข้อมูล */
    .meta { display: flex; justify-content: space-between; margin-bottom: 16px; }
    .indent { text-indent: 2.5cm; text-align: justify; margin: 6px 0; }
    .check-line { padding-left: 3cm; text-indent: -0.8cm; margin: 4px 0; }
    
    /* ลายเซ็น (จัดชิดขวาแบบเป็นทางการ) */
    .signature { 
      margin-top: 30px; 
      display: flex; 
      flex-direction: column; 
      align-items: center; 
      margin-left: auto; 
      margin-right: 1.5cm; 
      width: 250px; 
    }
    .sig-row { display: flex; justify-content: space-around; margin-top: 40px; }
    .sig-row > div { text-align: center; width: 45%; }
    
    /* อื่นๆ */
    .cut-line { border-top: 1px dashed #999; margin: 30px 0 20px; }
    .page-break { page-break-after: always; }
  </style>`;
}

function buildDocWrap(title, body) {
  return `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>${title}</title>
  <link href="https://fonts.googleapis.com/css2?family=Sarabun:wght@400;700&display=swap" rel="stylesheet">
  ${buildDocStyle()}</head><body>
  <div class="toolbar">
    <button onclick="window.print()" style="background:#1e40af;color:#fff;border:none;">🖨️ พิมพ์ / บันทึก PDF</button>
  </div>
  ${body}
  </body></html>`;
}

// ✅ เพิ่มฟังก์ชันสำหรับสร้างบล็อกอนุมัติ (ถ้าอนุมัติแล้ว)
function buildApprovalBlock(cData) {
  if (cData.approval_status !== 'approved') return '';
  
  const approvedDate = cData.approved_at ? new Date(cData.approved_at) : new Date();
  const thaiDate = Utilities.formatDate(approvedDate, Session.getScriptTimeZone(), "d/M/yyyy");
  const thaiYear = parseInt(thaiDate.split('/')[2]) + 543;
  const fullDate = thaiDate.replace(/\/\d{4}$/, '/' + thaiYear);

  return `
    <div style="margin-top: 40px; border-top: 1px dashed #999; padding-top: 20px;">
      <div style="text-align: center; font-size: 18pt; font-weight: bold; color: #000000;">
  ✅ ได้รับการอนุมัติแล้ว
</div>
      <div style="text-align: right; margin-top: 20px; font-size: 14pt;">
        <div>ลงชื่อ ..................................................... ผู้อนุมัติ</div>
        <div style="margin-top: 4px;">(${cData.approved_by || '........................................'})</div>
        <div style="margin-top: 4px;">อนุมัติเมื่อ: ${fullDate}</div>
      </div>
    </div>
  `;
}

function formatThaiDateLong(ds) {
  if (!ds) return '...............';
  const p = ds.split('-');
  if (p.length !== 3) return ds;
  const months = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];
  return `${parseInt(p[2])} ${months[parseInt(p[1]) - 1]} ${parseInt(p[0]) + 543}`;
}

function formatSignatureDate(timestamp) {
  if (!timestamp) return '......... / ......... / .........';
  try {
    const d = new Date(timestamp);
    if (isNaN(d.getTime())) return '......... / ......... / .........';
    return `${d.getDate()} / ${d.getMonth() + 1} / ${d.getFullYear() + 543}`;
  } catch (e) {
    return '......... / ......... / .........';
  }
}

function getGarudaHtml(cfg) {
  // ขนาดรูปให้พอดีกับกระดาษ A4 (สูง 60px)
  const style = 'width: auto; height: 60px; margin: 0 auto; display: block; object-fit: contain;';
  
  let url = cfg && cfg.garuda_url ? String(cfg.garuda_url).trim() : '';

  // ✅ ระบบแปลงลิงก์ Google Drive ให้เป็นลิงก์รูปภาพตรง (Direct Link)
  if (url.includes('drive.google.com')) {
    const match = url.match(/\/d\/([a-zA-Z0-9_-]+)/) || url.match(/id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      // ใช้ลิงก์แบบ lh3 ที่แสดงรูปตรงๆ ได้ดีที่สุด
      url = 'https://lh3.googleusercontent.com/d/' + match[1];
    }
  }

  if (url) {
    return `<img src="${url}" alt="ครุฑ" style="${style}" onerror="this.style.display='none';this.nextElementSibling.style.display='block';"><div style="text-align: center; font-size: 40pt; display: none;">🦁</div>`;
  }
  
  if (cfg && cfg.garuda_base64 && String(cfg.garuda_base64).trim()) {
    return `<img src="${cfg.garuda_base64}" alt="ครุฑ" style="${style}">`;
  }

  // Fallback: ถ้าระบบโหลดรูปไม่ได้เลย จะแสดงรูปจำลอง
  return `
    <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100" style="width: 60px; height: 60px; display: block; margin: 0 auto;">
      <circle cx="50" cy="50" r="45" fill="none" stroke="#8B0000" stroke-width="2"/>
      <circle cx="50" cy="50" r="40" fill="none" stroke="#8B0000" stroke-width="1" stroke-dasharray="2,2"/>
      <text x="50" y="46" text-anchor="middle" font-size="14" fill="#8B0000" font-weight="bold">ครุฑ</text>
      <text x="50" y="62" text-anchor="middle" font-size="8" fill="#8B0000">(รูปไม่ขึ้น)</text>
    </svg>
  `;
}

function getCategoryLabelFromTrigger(trigger) {
  const t = String(trigger || '').toLowerCase();
  if (t.startsWith('yellow')) return 'ห่วงใย';
  return 'ใกล้ชิด';
}

function getCategoryColorFromTrigger(trigger) {
  const t = String(trigger || '').toLowerCase();
  if (t.startsWith('yellow')) return '#eab308';
  return '#ef4444';
}

function buildViolationsHtml(cData) {
  const td = cData.trigger_data;
  let breakdown = td && td.breakdown;

  if (!breakdown) {
    try {
      const weights = getRiskWeights().data;
      const weekStart = td && td.weekStart ? td.weekStart : (function() {
        const d = new Date();
        const dow = d.getDay();
        d.setDate(d.getDate() - dow + (dow === 0 ? -6 : 1));
        return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      })();
      const weekEnd = td && td.weekEnd ? td.weekEnd : (function() {
        const d = new Date(weekStart);
        d.setDate(d.getDate() + 4);
        return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      })();

      const assess = getStudentRiskAssessment(weekStart, weekEnd, '', JSON.stringify(weights));
      if (assess.status === 'success') {
        const stu = (assess.data.students || []).find(x => x.id === cData.student_id);
        if (stu) breakdown = stu.breakdown;
      }
    } catch (e) { /* ignore */ }
  }

  if (!breakdown) return '<div class="check-line">☑ มีผลการประเมินความประพฤติอยู่ในกลุ่มที่ต้องติดตาม</div>';

  const items = [];
  const att = breakdown.attendance;
  if (att && att.score < 85) {
    items.push(`เข้าแถวไม่ครบ/ขาดแถว <strong>${Math.round(att.total - att.points)} ครั้ง</strong> (${att.score}%)`);
  }
  const cln = breakdown.cleaning;
  if (cln && cln.score < 85) {
    items.push(`ไม่ทำเวร/หนีเวร <strong>${Math.round(cln.total - cln.points)} ครั้ง</strong> (${cln.score}%)`);
  }
  const subj = breakdown.subject;
  if (subj && subj.score < 85) {
    items.push(`ขาดเรียน/เข้าเรียนไม่ครบ <strong>${Math.round(subj.total - subj.points)} ครั้ง</strong> (${subj.score}%)`);
  }

  if (items.length === 0) items.push('มีผลการประเมินความประเมินความประพฤติอยู่ในกลุ่มที่ต้องติดตาม');
  return items.map(i => `<div class="check-line">☑ ${i}</div>`).join('');
}

function reserveDocNumber() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) throw new Error('ไม่สามารถจองเลขหนังสือ');
  try {
    const year = new Date().getFullYear() + 543;
    const yearShort = String(year).substring(2);
    const configKey = `docCounter_${year}`;

    const sheet = getOrCreateSheet(SHEET_NAMES.CONFIG);
    const data = sheet.getDataRange().getValues();
    let counter = 0, rowIndex = -1;

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0] || '').trim() === configKey) {
        counter = parseInt(data[i][1]) || 0;
        rowIndex = i + 1;
        break;
      }
    }
    counter++;
    if (rowIndex > 0) sheet.getRange(rowIndex, 2).setValue(counter);
    else sheet.appendRow([configKey, counter]);

    return `ศธ 04201/${yearShort}${String(counter).padStart(4, '0')}`;
  } finally {
    lock.releaseLock();
  }
}

function ensureDocNo(caseId, docType) {
  const sheet = getDisciplineSheet('CASES');
  const data = sheet.getDataRange().getValues();
  const idx = data.slice(1).findIndex(r => String(r[0]) === caseId) + 2;
  if (idx <= 1) return { docNo: '', docDate: '' };

  const row = data[idx - 1];
  const colMap = { warning: [8, 9], bond: [10, 11], activity: [14, 15], parent: [16, 17] };
  const [noCol, dateCol] = colMap[docType] || [8, 9];

  let docNo = String(row[noCol] || '').trim();
  let docDate = formatDateStr(row[dateCol]);

  if (!docNo) {
    try {
      docNo = reserveDocNumber();
    } catch (err) {
      const year = new Date().getFullYear() + 543;
      docNo = `ศธ 04201/${String(year).substring(2)}${String(Math.floor(Math.random() * 9000) + 1000)}`;
    }
    docDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    sheet.getRange(idx, noCol + 1).setValue(docNo);
    sheet.getRange(idx, dateCol + 1).setValue(docDate);
  }
  return { docNo, docDate };
}

function buildWarningDocHtml(cData, cfg, advisor) {
  const { docNo, docDate } = ensureDocNo(cData.case_id, 'warning');
  const categoryLabel = getCategoryLabelFromTrigger(cData.trigger);
  const categoryColor = getCategoryColorFromTrigger(cData.trigger);

  const body = `
    <div class="garuda-wrap">${getGarudaHtml(cfg)}</div>
    <div class="header">${cfg.school_name}</div>
    <div class="header-sub">${cfg.school_name}</div>
    <div class="meta"><div>ที่ ${docNo}</div><div>${formatThaiDateLong(docDate)}</div></div>
    <div><strong>เรื่อง</strong> ตักเตือนเรื่องความประพฤติ</div>
    <div><strong>เรียน</strong> ${cData.student_name} ชั้น ${cData.classroom_name} เลขประจำตัว ${cData.student_id}</div>
    <div class="indent">
  ด้วย${cfg.school_name} ได้ติดตามพฤติกรรมของนักเรียนอย่างต่อเนื่อง พบว่านักเรียนมีผลการประเมินความประพฤติอยู่ในกลุ่ม <strong>"${categoryLabel}"</strong> ซึ่งต่ำกว่าเกณฑ์ที่โรงเรียนกำหนด
</div>
    <div class="indent">โรงเรียนจึงขอว่ากล่าวตักเตือนนักเรียนให้ปรับปรุงพฤติกรรม หากยังฝ่าฝืนอีก โรงเรียนจะดำเนินการตามระเบียบกระทรวงศึกษาธิการ ว่าด้วยการลงโทษนักเรียนและนักศึกษา พ.ศ. 2548 ต่อไป</div>
    <div class="indent">จึงเรียนมาเพื่อโปรดทราบและปฏิบัติตาม</div>
    <div class="signature">
      <div>ขอแสดงความนับถือ</div>
      <div style="margin-top: 40px;">ลงชื่อ .....................................................</div>
      <div>(${advisor.name})</div>
      <div>ครูที่ปรึกษา</div>
      <div style="margin-top:6px;">วันที่ ${formatSignatureDate(cData.printed_warning_at)}</div>
    </div>
  `;
  return buildDocWrap('หนังสือตักเตือน', body);
}

function buildBondDocHtml(cData, cfg, advisor) {
  const { docNo, docDate } = ensureDocNo(cData.case_id, 'bond');
  const bondStart = cData.bond_start || formatDateStr(new Date());
  const bondEnd = cData.bond_end || Utilities.formatDate(new Date(Date.now() + 90*86400000), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const parentName = cData.parent_name || '.....................................................';
  const parentRel = cData.parent_relation || 'บิดา/มารดา/ผู้ปกครอง';

  const page1 = `
    <div class="garuda-wrap">${getGarudaHtml(cfg)}</div>
    <div class="header" style="font-size: 20pt; margin: 10px 0;">บันทึกทัณฑ์บน</div>
    <div class="meta"><div>ที่ ${docNo}</div><div>${formatThaiDateLong(docDate)}</div></div>
    <div><strong>เรื่อง</strong> ทำทัณฑ์บนนักเรียน</div>
    <div><strong>เรียน</strong> ผู้อำนวยการ${cfg.school_name}</div>
    <div style="margin: 12px 0;"><strong>1. ข้อมูลนักเรียน</strong>
      <div style="padding-left: 1cm; margin-top: 4px;">
        <div>ชื่อ-สกุล: <strong>${cData.student_name}</strong></div>
        <div>ชั้น: ${cData.classroom_name} &nbsp;&nbsp; เลขประจำตัว: ${cData.student_id}</div>
        <div>ครูที่ปรึกษา: ${advisor.name}</div>
      </div>
    </div>
    <div style="margin: 12px 0;"><strong>2. พฤติกรรมที่กระทำผิด</strong>
      <div style="padding-left: 1cm; margin-top: 4px;">${buildViolationsHtml(cData)}</div>
    </div>
    <div style="margin: 12px 0;"><strong>3. เงื่อนไขการทำทัณฑ์บน</strong>
      <div style="padding-left: 1cm; margin-top: 4px;">
        <div>3.1 ระยะเวลาทัณฑ์บน: ตั้งแต่ ${formatThaiDateLong(bondStart)} ถึง ${formatThaiDateLong(bondEnd)}</div>
        <div>3.2 ระหว่างระยะเวลาทัณฑ์บน นักเรียนต้อง:</div>
        <div class="check-line">(1) ไม่กระทำความผิดซ้ำในลักษณะเดียวกัน</div>
        <div class="check-line">(2) เข้าเรียนและร่วมกิจกรรมอย่างสม่ำเสมอ</div>
        <div class="check-line">(3) รายงานตัวต่อครูที่ปรึกษาทุกสัปดาห์</div>
        <div class="check-line">(4) ปฏิบัติตามคำสั่งของโรงเรียนอย่างเคร่งครัด</div>
        <div style="margin-top: 6px;">3.3 หากฝ่าฝืนหรือไม่ปฏิบัติตามเงื่อนไข โรงเรียนจะพิจารณาลงโทษตามระเบียบกระทรวงศึกษาธิการ ว่าด้วยการลงโทษนักเรียนและนักศึกษา พ.ศ. 2548 ต่อไป</div>
      </div>
    </div>
    <div class="indent">จึงเรียนมาเพื่อโปรดพิจารณาลงนามอนุมัติ</div>
    <div style="text-align: right; margin-top: 30px;">
      <div>ลงชื่อ ..................................................... พยาน</div>
      <div>(${advisor.name})</div>
      <div>ครูที่ปรึกษา</div>
      <div style="margin-top: 6px;">วันที่ ${formatSignatureDate(cData.signed_witness1_at)}</div>
    </div>
    <div class="page-break"></div>
  `;

  const page2 = `
    <div class="header" style="font-size: 20pt;">บันทึกทัณฑ์บน (หน้า 2)</div>
    <div style="text-align: center; margin: 30px 0; font-weight: bold; font-size: 15pt;">
      คู่สัญญาทั้ง 3 ฝ่าย ได้อ่านและเข้าใจข้อความในบันทึกทัณฑ์บนนี้โดยตลอดแล้ว<br>จึงได้ลงลายมือชื่อไว้เป็นหลักฐาน
    </div>
    <div style="margin: 20px 0;">
      <div style="text-align: center; font-weight: bold; margin-bottom: 12px;">─── ฝ่ายที่ 1: ผู้ทำทัณฑ์บน (นักเรียน) ───</div>
      <div style="text-align: center;">
        <div>ข้าพเจ้าขอยอมรับว่ากระทำความผิดตามที่ระบุไว้จริง และขอสัญญาว่าจะไม่กระทำความผิดซ้ำอีก</div>
        <div style="margin-top: 25px;">ลงชื่อ ..................................................... ผู้ทำทัณฑ์บน</div>
        <div>( ${cData.student_name} )</div>
        <div>วันที่ ${formatSignatureDate(cData.signed_student_at)}</div>
      </div>
    </div>
    <div style="margin: 20px 0;">
      <div style="text-align: center; font-weight: bold; margin-bottom: 12px;">─── ฝ่ายที่ 2: ผู้รับรองการทำทัณฑ์บน (ผู้ปกครอง) ───</div>
      <div style="text-align: center;">
        <div>ข้าพเจ้าได้รับทราบถึงความผิดของบุตรหลานแล้ว และขอรับรองว่าจะดูแลให้ปฏิบัติตามเงื่อนไข</div>
        <div style="margin-top: 25px;">ลงชื่อ ..................................................... ผู้ปกครอง</div>
        <div>( ${parentName} )</div>
        <div>เกี่ยวข้องเป็น ${parentRel}</div>
        <div>วันที่ ${formatSignatureDate(cData.signed_parent_at)}</div>
      </div>
    </div>
    <div style="margin: 20px 0;">
      <div style="text-align: center; font-weight: bold; margin-bottom: 12px;">─── ฝ่ายที่ 3: ผู้ลงโทษ (ผู้บริหารสถานศึกษา) ───</div>
      <div style="text-align: center;">
        <div>อาศัยอำนาจตามระเบียบกระทรวงศึกษาธิการ ว่าด้วยการลงโทษนักเรียนและนักศึกษา พ.ศ. 2548 จึงสั่งให้ทำทัณฑ์บนไว้เป็นหนังสือ</div>
        <div style="margin-top: 25px;">ลงชื่อ ..................................................... ผู้สั่งลงโทษ</div>
        <div>( ${cfg.director_name} )</div>
        <div>${cfg.director_position}</div>
        <div>${cfg.school_name}</div>
        <div>วันที่ ${formatSignatureDate(cData.signed_director_at)}</div>
      </div>
    </div>
       <div style="margin: 20px 0;">
      <div style="text-align: center; font-weight: bold; margin-bottom: 12px;">─── พยานในการรับทราบ ───</div>
      <div class="sig-row">
        <div>
          <div style="margin-top: 25px;">ลงชื่อ ..................................................... พยาน</div>
          <div>( ${advisor.name} )</div>
          <div>ครูที่ปรึกษา</div>
          <div>วันที่ ${formatSignatureDate(cData.signed_witness1_at)}</div>
        </div>
        <div>
          <div style="margin-top: 25px;">ลงชื่อ ..................................................... พยาน</div>
          <div>( ${cfg.discipline_teacher} )</div>
          <div>หัวหน้างานกิจการนักเรียน</div>
          <div>วันที่ ${formatSignatureDate(cData.signed_witness2_at)}</div>
        </div>
      </div>
    </div>

    <!-- ✅ แทรกบล็อกอนุมัติตรงนี้ -->
    ${buildApprovalBlock(cData)}
  `;
  return buildDocWrap('บันทึกทัณฑ์บน', page1 + page2);
}

function buildActivityDocHtml(cData, cfg, advisor) {
  const { docNo, docDate } = ensureDocNo(cData.case_id, 'activity');
  const act = cData.activity_data || { items: [], start: '', end: '' };

  let activitiesHtml = '<div style="padding-left: 1cm;">';
  (act.items || []).forEach(it => {
    activitiesHtml += `<div class="check-line">☐ ${it.label} จำนวน ${it.target} ${it.unit}</div>`;
  });
  activitiesHtml += '</div>';

  const body = `
    <div class="garuda-wrap">${getGarudaHtml(cfg)}</div>
    <div class="header">คำสั่ง${cfg.school_name}</div>
    <div class="header">ที่ ${docNo}</div>
    <div class="header">เรื่อง มอบหมายกิจกรรมเพื่อปรับเปลี่ยนพฤติกรรม</div>
    <div style="text-align: center; margin: 20px 0;">─────────────</div>
    <div class="indent">อาศัยอำนาจตามระเบียบกระทรวงศึกษาธิการ ว่าด้วยการลงโทษนักเรียนและนักศึกษา พ.ศ. 2548 และบันทึกทัณฑ์บนที่ ${cData.bond_doc_no || '.........'} ลงวันที่ ${cData.bond_date ? formatThaiDateLong(cData.bond_date) : '.........'} ประกอบกับนักเรียน <strong>${cData.student_name}</strong> ชั้น ${cData.classroom_name} เลขประจำตัว ${cData.student_id} ได้ฝ่าฝืนเงื่อนไขของทัณฑ์บน</div>
    <div class="indent">โรงเรียนจึงมีคำสั่งให้นักเรียนดำเนินกิจกรรมเพื่อปรับเปลี่ยนพฤติกรรม ดังต่อไปนี้:</div>
    <div style="margin: 12px 0;">${activitiesHtml}</div>
    <div class="indent"><strong>ระยะเวลาดำเนินการ:</strong> ตั้งแต่วันที่ ${act.start ? formatThaiDateLong(act.start) : '.........'} ถึง ${act.end ? formatThaiDateLong(act.end) : '.........'}</div>
    <div class="indent"><strong>ผู้ติดตามผล:</strong> ${advisor.name} ครูที่ปรึกษา</div>
    <div class="indent">ขอให้นักเรียนและผู้ปกครองให้ความร่วมมือในการดำเนินกิจกรรมตามที่กำหนด หากไม่ดำเนินการ โรงเรียนจะพิจารณาแจ้งผู้ปกครองและดำเนินการตามระเบียบวินัยต่อไป</div>
    <div style="text-align: right; margin-top: 40px;">
      <div>สั่ง ณ วันที่ ${formatThaiDateLong(docDate)}</div>
      <div style="margin-top: 40px;">ลงชื่อ ..................................................... ผู้สั่ง</div>
      <div>( ${cfg.director_name} )</div>
      <div>${cfg.director_position}</div>
      <div>${cfg.school_name}</div>
    </div>
    <div class="sig-row">
      <div>
        <div style="font-weight: bold;">รับทราบ</div>
        <div style="margin-top: 30px;">ลงชื่อ ..................................................... นักเรียน</div>
        <div>( ${cData.student_name} )</div>
        <div>วันที่ ${formatSignatureDate(cData.signed_student_at)}</div>
      </div>
      <div>
        <div style="font-weight: bold;">รับทราบ</div>
        <div style="margin-top: 30px;">ลงชื่อ ..................................................... ผู้ปกครอง</div>
        <div>( ${cData.parent_name || '.............'} )</div>
        <div>วันที่ ${formatSignatureDate(cData.signed_parent_at)}</div>
      </div>
    </div>
  `;
  return buildDocWrap('คำสั่งมอบหมายกิจกรรม', body);
}

function buildParentNoticeDocHtml(cData, cfg, advisor) {
  const { docNo, docDate } = ensureDocNo(cData.case_id, 'parent');
  const meetingDate = cData.meeting_date ? formatThaiDateLong(cData.meeting_date) : '...........................................';
  const meetingTime = cData.meeting_time || '........................';
  const meetingLoc = cData.meeting_location || '...........................................';

  const body = `
    <div class="garuda-wrap">${getGarudaHtml(cfg)}</div>
    <div class="header">${cfg.school_name}</div>
    <div class="header-sub">${cfg.school_name}</div>
    <div class="meta"><div>ที่ ${docNo}</div><div>${formatThaiDateLong(docDate)}</div></div>
    <div><strong>เรื่อง</strong> ขอเชิญประชุมผู้ปกครองเพื่อหาแนวทางแก้ไขพฤติกรรมนักเรียน</div>
    <div><strong>เรียน</strong> ผู้ปกครองของ ${cData.student_name}</div>
    <div class="indent">ด้วย${cfg.school_name} ได้ติดตามพฤติกรรมของ <strong>${cData.student_name}</strong> ชั้น ${cData.classroom_name} เลขประจำตัว ${cData.student_id} พบว่านักเรียนมีผลการประเมินอยู่ในกลุ่ม <strong>"ใกล้ชิด"</strong> ต่อเนื่องกัน และได้ผ่านกระบวนการลงโทษตามระเบียบแล้ว</div>
    <div style="margin: 12px 0; padding-left: 1cm;">
      <div>1. หนังสือตักเตือน เลขที่ ${cData.warn_doc_no || '.........'} ลงวันที่ ${cData.warn_date ? formatThaiDateLong(cData.warn_date) : '.........'}</div>
      <div>2. บันทึกทัณฑ์บน เลขที่ ${cData.bond_doc_no || '.........'} ลงวันที่ ${cData.bond_date ? formatThaiDateLong(cData.bond_date) : '.........'}</div>
      <div>3. คำสั่งมอบหมายกิจกรรม เลขที่ ${cData.activity_doc_no || '.........'} ลงวันที่ ${cData.activity_date ? formatThaiDateLong(cData.activity_date) : '.........'}</div>
    </div>
    <div class="indent">แต่นักเรียนยังไม่สามารถดำเนินการตามเงื่อนไขได้สำเร็จ โรงเรียนจึงขอเชิญท่านเข้าร่วมประชุม</div>
    <div style="margin: 12px 0; padding-left: 1cm;">
      <div><strong>วันประชุม:</strong> ${meetingDate}</div>
      <div><strong>เวลา:</strong> ${meetingTime}</div>
      <div><strong>สถานที่:</strong> ${meetingLoc}</div>
    </div>
    <div class="indent">จึงเรียนมาเพื่อโปรดเข้าร่วมประชุมตามวันและเวลาดังกล่าว หากไม่สามารถมาได้ กรุณาติดต่อครูที่ปรึกษา ${advisor.name}${advisor.phone ? ' โทร. ' + advisor.phone : ''} เพื่อนัดหมายใหม่</div>
    <div class="indent">ขอแสดงความนับถือ</div>
    <div style="text-align: right; margin-top: 40px;">
      <div>ลงชื่อ .....................................................</div>
      <div>( ${cfg.director_name} )</div>
      <div>${cfg.director_position}</div>
      <div>${cfg.school_name}</div>
    </div>
    <div class="cut-line"></div>
    <div style="text-align: center; font-weight: bold; font-size: 18pt; margin: 20px 0;">✂️ ใบตอบรับ</div>
    <div style="text-align: right; margin-bottom: 10px;">เขียนที่ .........................................<br>วันที่ ......... / ......... / .........</div>
    <div style="margin: 8px 0;"><strong>เรื่อง</strong> ตอบรับหนังสือเชิญประชุม</div>
    <div style="margin: 8px 0;"><strong>เรียน</strong> ผู้อำนวยการ${cfg.school_name}</div>
    <div class="indent">ตามที่โรงเรียนได้มีหนังสือที่ ${docNo} ลงวันที่ ${formatThaiDateLong(docDate)} ขอเชิญประชุมเรื่องพฤติกรรมของ ${cData.student_name} นั้น</div>
    <div class="indent">ข้าพเจ้า ..................................................... ผู้ปกครองของ ${cData.student_name} ได้รับหนังสือแล้ว และ</div>
    <div style="margin: 8px 0; padding-left: 1cm;">
      <div>☐ ยินดีเข้าร่วมประชุมตามวันและเวลาดังกล่าว</div>
      <div>☐ ไม่สามารถเข้าร่วมได้ เพราะ ............................................ จะติดต่อนัดหมายใหม่</div>
    </div>
    <div style="text-align: right; margin-top: 30px;">
      <div>ลงชื่อ ..................................................... ผู้ปกครอง</div>
      <div>( ..................................................... )</div>
      <div>วันที่ ......... / ......... / .........</div>
    </div>
  `;
  return buildDocWrap('หนังสือแจ้งผู้ปกครอง', body);
}

function escapeBehaviorDocumentHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

function getBehaviorRecordsSheet_() {
  const spreadsheet = getSpreadsheet();
  let sheet = spreadsheet.getSheetByName(BEHAVIOR_RECORDS_SHEET);
  if (!sheet) sheet = spreadsheet.insertSheet(BEHAVIOR_RECORDS_SHEET);
  if (sheet.getLastRow() === 0 || !sheet.getRange(1, 1).getValue()) {
    sheet.getRange(1, 1, 1, 3).setValues([['case_id', 'doc_no', 'doc_date']]);
  }
  return sheet;
}

function ensureBehaviorDocNo_(caseId) {
  const sheet = getBehaviorRecordsSheet_();
  const records = sheet.getDataRange().getValues();
  const existing = records.slice(1).find(row => String(row[0]) === String(caseId));
  if (existing && String(existing[1] || '').trim()) {
    return { docNo: String(existing[1]).trim(), docDate: formatDateStr(existing[2]) };
  }

  const docNo = reserveDocNumber();
  const docDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (existing) {
    const rowIndex = records.findIndex(row => String(row[0]) === String(caseId)) + 1;
    sheet.getRange(rowIndex, 2, 1, 2).setValues([[docNo, docDate]]);
  } else {
    sheet.appendRow([caseId, docNo, docDate]);
  }
  return { docNo, docDate };
}

function buildBehaviorRecordDocHtml_(cData, cfg, advisor) {
  const { docNo, docDate } = ensureBehaviorDocNo_(cData.case_id);
  const result = calculateBehaviorScore_(cData.student_id);
  const violations = result.violations.slice(0, 3);
  const configResult = getConfig();
  if (configResult.status !== 'success') {
    throw new Error(configResult.message || 'ไม่สามารถอ่านการตั้งค่าโรงเรียนได้');
  }
  const schoolConfig = configResult.data.school_config || {};
  const termLabel = schoolConfig.semester && schoolConfig.schoolYear
    ? `ภาคเรียนที่ ${escapeBehaviorDocumentHtml_(schoolConfig.semester)} ปีการศึกษา ${escapeBehaviorDocumentHtml_(schoolConfig.schoolYear)}`
    : `ระหว่างวันที่ ${formatThaiDateLong(result.period.startDate)} ถึง ${formatThaiDateLong(result.period.endDate)}`;
  const violationHtml = violations.length
    ? violations.map((violation, index) => `
        <div style="margin-left: 1cm; display: flex; justify-content: space-between; width: 90%;">
          <div>${index + 1}. ${escapeBehaviorDocumentHtml_(violation.type)} (${formatThaiDateLong(violation.date)})</div>
          <div>${violation.deduct} คะแนน</div>
        </div>`).join('')
    : '<div style="margin-left: 1cm;">- ไม่มีประวัติการกระทำผิด -</div>';

  const body = `
    <div class="garuda-wrap">${getGarudaHtml(cfg)}</div>
    <div class="header">${escapeBehaviorDocumentHtml_(cfg.school_name)}</div>
    <div class="meta" style="margin-top: 20px;">
      <div>ที่ ${escapeBehaviorDocumentHtml_(docNo)}</div>
      <div style="margin-right: 2cm;">${formatThaiDateLong(docDate)}</div>
    </div>
    <div style="display: flex;">
      <div style="width: 2.5cm; text-align: right;"><strong>เรื่อง</strong></div>
      <div style="flex: 1; padding-left: 0.5cm;">แจ้งข้อมูลพฤติกรรมนักเรียน</div>
    </div>
    <div style="display: flex; margin-top: 6px;">
      <div style="width: 2.5cm; text-align: right;"><strong>เรียน</strong></div>
      <div style="flex: 1; padding-left: 0.5cm;">ผู้ปกครองของ ${escapeBehaviorDocumentHtml_(cData.student_name)}</div>
    </div>
    <div class="indent" style="margin-top: 20px;">
      ด้วย <strong>${escapeBehaviorDocumentHtml_(cData.student_name)}</strong>
      เลขประจำตัวนักเรียน ${escapeBehaviorDocumentHtml_(cData.student_id)}
      ชั้น ${escapeBehaviorDocumentHtml_(cData.classroom_name)} ซึ่งเป็นนักเรียนในปกครองของท่าน
      มีข้อมูลพฤติกรรมตามระบบดูแลช่วยเหลือนักเรียนใน${termLabel} ดังนี้
    </div>
    <div style="margin-top: 15px; font-weight: bold;">1. พฤติกรรมที่บันทึก (3 รายการล่าสุด)</div>
    <div style="margin-top: 10px; line-height: 1.8;">${violationHtml}</div>
    <div class="indent" style="margin-top: 20px;">
      จากคะแนนเต็ม 100 คะแนน ขณะนี้นักเรียนมีคะแนนความประพฤติคงเหลือ
      <strong>${result.score}</strong> คะแนน
    </div>
    <div class="indent" style="margin-top: 15px;">
      จึงเรียนมาเพื่อโปรดทราบ และร่วมกันดูแลช่วยเหลือนักเรียนดังกล่าวให้มีพฤติกรรมที่พึงประสงค์และประสบความสำเร็จในการศึกษาเล่าเรียน
    </div>
    <div class="signature" style="margin-top: 40px; text-align: center; width: 100%; margin-left: 0;">
      <div style="text-align: center; margin-bottom: 30px;">ขอแสดงความนับถือ</div>
      <div style="text-align: center;">ลงชื่อ .....................................................</div>
      <div style="text-align: center;">(${escapeBehaviorDocumentHtml_(advisor.name)})</div>
      <div style="text-align: center;">ครูที่ปรึกษา</div>
    </div>
  `;
  return buildDocWrap('แจ้งข้อมูลพฤติกรรม', body);
}

function getDisciplineDocHtml(caseId, docType) {
  try {
    const res = getDisciplineCaseDetail(caseId);
    if (res.status !== 'success') throw new Error(res.message);
    const { case: cData, config: cfg, advisor } = res.data;

    let html = '';
    if (docType === 'behavior_record') html = buildBehaviorRecordDocHtml_(cData, cfg, advisor);
    else if (docType === 'warning') html = buildWarningDocHtml(cData, cfg, advisor);
    else if (docType === 'bond') html = buildBondDocHtml(cData, cfg, advisor);
    else if (docType === 'activity') html = buildActivityDocHtml(cData, cfg, advisor);
    else if (docType === 'parent') html = buildParentNoticeDocHtml(cData, cfg, advisor);
    else throw new Error('docType ไม่ถูกต้อง');

    if (docType !== 'behavior_record') markDocPrinted(caseId, docType);
    return { status: 'success', data: { html } };
  } catch (err) {
    Logger.log('getDisciplineDocHtml error: ' + err.message);
    return { status: 'error', message: err.toString() };
  }
}

function getStudentRiskHistory(studentId, weeksBack) {
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

      const w = getRiskWeights().data;
      const res = getStudentRiskAssessment(start, end, '', JSON.stringify(w));
      if (res.status === 'success') {
        const stu = (res.data.students || []).find(s => s.id === String(studentId));
        if (stu) results.push({
          startDate: start, endDate: end,
          composite: stu.composite, category: stu.category,
          streak: stu.streak, breakdown: stu.breakdown
        });
      }
    }

    const redCount = results.filter(r => r.category === 'red').length;
    const yellowCount = results.filter(r => r.category === 'yellow').length;

    return {
      status: 'success',
      data: {
        weeks: results, redCount, yellowCount,
        recommendation: redCount >= 3 ? 'activity' : redCount >= 1 ? 'warning' : 'none'
      }
    };
  } catch (err) {
    return { status: 'error', message: err.toString() };
  }
}

// ═══════════════════════════════════════════════════════════════
// 20. STATS & REPORTS
// ═══════════════════════════════════════════════════════════════

function getTopAbsentStudents(startDate, endDate, period, limit, spreadsheetMode) {
  if (spreadsheetMode === undefined) spreadsheetMode = DEFAULT_SPREADSHEET_MODE;
  try {
    if (!limit || limit <= 0) limit = 20;
    const attendanceResponse = getAttendanceRange(startDate, endDate, period, spreadsheetMode);
    if (attendanceResponse.status !== 'success') throw new Error('Failed');
    const attendanceData = attendanceResponse.data;
    if (attendanceData.length === 0) return createResponse('success', 'No data', []);

    const classroomResponse = getClassrooms([], spreadsheetMode);
    const allClassrooms = classroomResponse.data;
    const studentResponse = getStudents(allClassrooms.map(c => c.id), spreadsheetMode);
    const studentsByClassroom = studentResponse.data;

    const latestByRoomAndDate = {};
    attendanceData.forEach(entry => {
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
            const sid = String(student.student_id);
            if (!absentCounts[sid]) absentCounts[sid] = { student_id: sid, room_id: entry.room_id, count: 0 };
            absentCounts[sid].count++;
          }
        });
      }
    });

    const absentList = [];
    for (const sid in absentCounts) {
      const item = absentCounts[sid];
      let studentInfo = null, foundRoomId = null;
      for (const rid in studentsByClassroom) {
        const found = studentsByClassroom[rid].find(s => String(s.student_id) === sid);
        if (found) { studentInfo = found; foundRoomId = rid; break; }
      }
      if (studentInfo) {
        const classroom = allClassrooms.find(c => c.id === foundRoomId);
        absentList.push({
          student_id: sid, no: studentInfo.no || '', name: studentInfo.name || '',
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

    return createResponse('success', 'Top absent', absentList.slice(0, limit));
  } catch (error) {
    return createResponse('error', 'Failed: ' + error.message, null, error);
  }
}

function getDashboardStatsSummary(startDate, endDate, selectedPeriods, selectedStatuses, filterLevel) {
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
      if(row[1]) stdMap[row[1].toString().trim()] = { classroom_id: row[0].toString().trim(), name: row[3] };
    });

    const studentStats = {}, roomStats = {};
    const uniqueDates = new Set();
    const targetStatuses = Array.isArray(selectedStatuses) ? selectedStatuses : [selectedStatuses];
    const targetPeriods = Array.isArray(selectedPeriods) ? selectedPeriods.filter(p => p !== '') : [];

    const cleanData = getCleanedAttendanceRows(attData);

    cleanData.forEach(row => {
      const rDate = formatDateStr(row[0]);
      const rPeriod = String(row[1] || '').trim();
      const rRoom = String(row[2] || '').trim();
      const rJson = row[4];

      let isLevelMatch = false;
      if (filterLevel === 'all' || !filterLevel || filterLevel === '') isLevelMatch = true;
      else if (filterLevel === 'm1-3') isLevelMatch = ['m1','m2','m3'].some(l => rRoom.toLowerCase().startsWith(l));
      else if (filterLevel === 'm4-6') isLevelMatch = ['m4','m5','m6'].some(l => rRoom.toLowerCase().startsWith(l));
      else isLevelMatch = rRoom.toLowerCase().startsWith(filterLevel.toLowerCase());

      if (rDate >= startDate && rDate <= endDate && isLevelMatch) {
        uniqueDates.add(rDate);
        let students = [];
        try { students = JSON.parse(rJson); } catch (e) {}

        if (Array.isArray(students)) {
          if (!roomStats[rRoom]) roomStats[rRoom] = { classroom: roomNameMap[rRoom] || rRoom, periods: {}, total: 0 };

          if (targetPeriods.length === 0 || targetPeriods.includes(rPeriod)) {
            const pk = rPeriod || 'ไม่ระบุเวลา';
            if (!roomStats[rRoom].periods[pk]) roomStats[rRoom].periods[pk] = 0;
            roomStats[rRoom].periods[pk]++;
            roomStats[rRoom].total++;
          }

          students.forEach(s => {
            const sid = String(s.student_id).trim();
            if (!studentStats[sid]) {
              const info = stdMap[sid] || { name: 'ไม่พบชื่อ ('+sid+')', classroom_id: rRoom };
              studentStats[sid] = {
                student_id: sid, name: info.name,
                classroom_name: roomNameMap[info.classroom_id] || info.classroom_id,
                total_checks: 0, filtered_count: 0
              };
            }
            if (targetPeriods.length === 0 || targetPeriods.includes(rPeriod)) {
              studentStats[sid].total_checks++;
              if (targetStatuses.includes(s.status)) studentStats[sid].filtered_count++;
            }
          });
        }
      }
    });

    const sortedPeriodSummary = Object.values(roomStats).filter(r => r.total > 0).sort((a, b) => a.classroom.localeCompare(b.classroom, 'th'));

    return {
      status: 'success',
      data: {
        startDate, endDate, schoolDays: uniqueDates.size,
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

function getStudentReportData(studentId, startDate, endDate, statusFilters, classroomId, periodFilter) {
  try {
    const ss = getSpreadsheet();
    const attData = ss.getSheetByName(SHEET_NAMES.ATTENDANCE).getDataRange().getValues();
    const cleanData = getCleanedAttendanceRows(attData);

    const specificDates = [];
    const uniqueDateStrings = new Set();
    const targetStatuses = Array.isArray(statusFilters) ? statusFilters : [statusFilters];

    cleanData.forEach(row => {
      const rDateStr = formatDateStr(row[0]);
      const rPeriod = String(row[1] || '').trim();
      let dateMatch = (rDateStr >= startDate && rDateStr <= endDate);
      let periodMatch = (!periodFilter || periodFilter === '' || rPeriod === periodFilter);

      if (dateMatch && periodMatch) {
        let students = JSON.parse(row[4] || '[]');
        let rec = students.find(s => String(s.student_id) === String(studentId));
        if (rec && targetStatuses.includes(rec.status)) {
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

// ═══════════════════════════════════════════════════════════════
// 21. DIAGNOSTICS
// ═══════════════════════════════════════════════════════════════

function diagnoseSubjectAttendance() {
  const ss = getSpreadsheet();
  const out = [];

  const tSheet = ss.getSheetByName('Timetable');
  if (!tSheet) {
    out.push('❌ ไม่พบชีต Timetable');
  } else {
    const tData = tSheet.getDataRange().getValues();
    out.push('✅ Timetable: ' + tData.length + ' rows');
    if (tData.length > 0) out.push('   Headers: ' + JSON.stringify(tData[0]));

    const emails = {}, days = {};
    for (let i = 1; i < tData.length; i++) {
      const e = String(tData[i][5] || '').trim();
      const d = String(tData[i][1] || '').trim();
      if (e) emails[e] = (emails[e] || 0) + 1;
      if (d) days[d] = (days[d] || 0) + 1;
    }
    out.push('   TeacherEmails: ' + JSON.stringify(emails));
    out.push('   Days: ' + JSON.stringify(days));
  }

  const teachSheet = ss.getSheetByName('Teachers');
  if (teachSheet) {
    const tData2 = teachSheet.getDataRange().getValues();
    const users = tData2.slice(1).map(r => String(r[2] || '').trim());
    out.push('✅ Teachers usernames: ' + JSON.stringify(users));
  }

  Logger.log(out.join('\n'));
  return out.join('\n');
}

function debugSubjectSchedule() {
  const teacher = 'teacher602';
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const monday = (function() {
    const d = new Date();
    d.setDate(d.getDate() - d.getDay() + (d.getDay() === 0 ? -6 : 1));
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  })();

  Logger.log('═══ DEBUG getTeacherSubjectSchedule ═══');
  Logger.log('Teacher: ' + teacher);
  Logger.log('Today:   ' + today);
  Logger.log('Monday:  ' + monday);

  Logger.log('\n📅 TEST 1: TODAY');
  try {
    const t0 = Date.now();
    const r1 = getTeacherSubjectSchedule(teacher, today);
    Logger.log(`⏱️ ${Date.now() - t0}ms`);
    Logger.log(JSON.stringify(r1, null, 2));
  } catch (e) {
    Logger.log('❌ ' + e.message);
  }

  Logger.log('\n📅 TEST 2: MONDAY');
  try {
    const t0 = Date.now();
    const r2 = getTeacherSubjectSchedule(teacher, monday);
    Logger.log(`⏱️ ${Date.now() - t0}ms`);
    Logger.log(JSON.stringify(r2, null, 2));
  } catch (e) {
    Logger.log('❌ ' + e.message);
  }

  Logger.log('\n✅ DEBUG COMPLETE');
}

// ═══════════════════════════════════════════════════════════════
// END OF FILE
// ═══════════════════════════════════════════════════════════════
/**
 * ✅ DEBUG: ตรวจสอบ Cleaning จากฝั่ง Server
 * ⚠️ รันใน Apps Script Editor ได้ (ไม่ต้องใช้ browser)
 * 
 * วิธีใช้: แก้ค่า USERNAME_TO_DEBUG → Run → ดู Execution log
 */
function debugCleaningServer() {
  const USERNAME_TO_DEBUG = 'teacher602';   // ⬅️ เปลี่ยนเป็น username ที่ต้องการ Debug
  
  Logger.log('═══════════════ DEBUG CLEANING (Server) ═══════════════');
  
  try {
    // 1. Teacher Info
    const tSheet = getOrCreateSheet(SHEET_NAMES.TEACHERS);
    const tData = tSheet.getDataRange().getValues();
    const teacherRow = tData.slice(1).find(r => String(r[2] || '').trim() === USERNAME_TO_DEBUG);
    
    if (!teacherRow) {
      Logger.log(`❌ ไม่พบ username "${USERNAME_TO_DEBUG}" ในชีต Teachers`);
      return;
    }
    
    Logger.log('1️⃣ Teacher Info:');
    Logger.log('   id: ' + teacherRow[0]);
    Logger.log('   name: ' + teacherRow[1]);
    Logger.log('   username: ' + teacherRow[2]);
    Logger.log('   status: ' + teacherRow[4] + ' (ต้องเป็น 1)');
    Logger.log('   classroom_ids: "' + (teacherRow[5] || '') + '"');
    Logger.log('   role: ' + (teacherRow[8] || '(ว่าง → default=teacher)'));
    
    const teacherRooms = String(teacherRow[5] || '')
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);
    Logger.log('   → parsed rooms: ' + JSON.stringify(teacherRooms));
    
    // 2. Classrooms in system
    Logger.log('');
    Logger.log('2️⃣ Classrooms ในระบบ:');
    const cSheet = getOrCreateSheet(SHEET_NAMES.CLASSROOMS);
    const cData = cSheet.getDataRange().getValues();
    const activeRooms = cData.slice(1)
      .filter(r => String(r[2]) === '1')
      .map(r => ({ id: String(r[0]).trim(), name: String(r[1]).trim() }));
    
    Logger.log('   ทั้งหมด ' + activeRooms.length + ' ห้อง:');
    activeRooms.forEach(c => Logger.log('   - ' + c.id + ' = "' + c.name + '"'));
    
    // 3. Compare
    Logger.log('');
    Logger.log('3️⃣ เปรียบเทียบ:');
    const allRoomIds = activeRooms.map(c => c.id);
    
    teacherRooms.forEach(r => {
      const match = allRoomIds.includes(r);
      Logger.log(`   "${r}" → ${match ? '✅ ตรงกับ Classrooms' : '❌ ไม่มีใน Classrooms'}`);
    });
    
    // 4. Prediction
    Logger.log('');
    Logger.log('4️⃣ ผลลัพธ์ที่คาดหวัง:');
    const role = String(teacherRow[8] || 'teacher').trim().toLowerCase();
    const isFullAccess = ['director', 'discipline'].includes(role) && teacherRooms.length === 0;
    
    if (isFullAccess) {
      Logger.log('   ✅ ควรเห็นปุ่มเช็คทุกห้อง (' + activeRooms.length + ' ปุ่ม)');
    } else if (teacherRooms.length === 0) {
      Logger.log('   ❌ ไม่ควรเห็นปุ่มเลย (ไม่มี classroom_ids + ไม่ใช่ director/discipline)');
    } else {
      const validRooms = teacherRooms.filter(r => allRoomIds.includes(r));
      Logger.log('   ✅ ควรเห็น ' + validRooms.length + ' ปุ่ม: ' + JSON.stringify(validRooms));
      
      if (validRooms.length === 0) {
        Logger.log('   ⚠️ ปัญหา: ไม่มีห้องที่ตรงกัน → ตรวจสอบการพิมพ์ผิด (ขีดกลาง vs ขีดล่าง)');
      }
    }
    
  } catch (err) {
    Logger.log('❌ Error: ' + err.message);
    Logger.log('Stack: ' + err.stack);
  }
  
  Logger.log('═══════════════ END ═══════════════');
}
function testGetData() {
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  Logger.log('=== วันที่ทดสอบ: ' + today + ' ===');
  
  // 1. ทดสอบดึงการเช็คแถว (สมมติห้อง m6-1 คาบเช้า)
  var att = getAttendance(today, 'morning', 'm6-1');
  Logger.log('ข้อมูลเช็คแถว: ' + JSON.stringify(att));
  
  // 2. ทดสอบดึงข้อมูลเขตพื้นที่
  var cln = getCleaningWeeklyStatus([today]);
  Logger.log('ข้อมูลเขตพื้นที่: ' + JSON.stringify(cln));
}
// ========================================
// SUBJECT SUMMARY V2 (สำหรับหน้าสถิติรายวิชา)
// ========================================

function getSubjectAttendanceSummaryV2(teacherIdentifier, subjectKey, startDate, endDate, classroomFilter) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("Attendance_Subjects");
    if (!sheet) return { status: "success", data: { students: [], totals: {}, dailyTrend: {} } };

    // ✅ ดึงรายชื่อจริงจากชีต Students มาทำ Map เพื่อความแม่นยำ
    const stdSheet = ss.getSheetByName('Students');
    const stdData = stdSheet ? stdSheet.getDataRange().getDisplayValues() : [];
    const nameMap = {};
    const idMap = {}; // ใช้เทียบรหัสแบบไม่มี 0
    stdData.slice(1).forEach(r => { 
      const id = String(r[1]).trim();
      const name = String(r[3]).trim();
      nameMap[id] = name; 
      idMap[id.replace(/^0+/, '')] = { id: id, name: name };
    });

    const data = sheet.getDataRange().getValues();
    const summary = {};
    const dailyTrend = {};
    const norm = String(teacherIdentifier || "").toLowerCase().trim();
    const subjKey = String(subjectKey || "").trim();

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDate = formatDateStr(row[1]);
      const rowCode = String(row[2] || "").trim();
      const rowName = String(row[3] || "").trim();
      const rowTeacher = String(row[4] || "").trim().toLowerCase();
      const rowGrade = String(row[5] || "").trim();
      const rowRoom = String(row[6] || "").trim();
      let stuId = String(row[7] || "").trim();
      let stuName = String(row[8] || "").trim();
      const status = String(row[9] || "").trim();

      if (!rowDate || !stuId) continue;
      if (startDate && rowDate < startDate) continue;
      if (endDate && rowDate > endDate) continue;
      if (rowTeacher !== norm) continue;

      const codeMatch = subjKey !== "" ? (rowCode === subjKey) : (rowName === subjKey);
      if (!codeMatch) continue;

      if (classroomFilter) {
        const gradeRoom = rowGrade + "/" + rowRoom;
        if (gradeRoom !== classroomFilter && rowRoom !== classroomFilter && rowGrade !== classroomFilter) continue;
      }

      // ✅ ค้นหาชื่อจริงจาก Students sheet (แก้ปัญหาเบิ้ลรหัส)
      const normId = stuId.replace(/^0+/, '');
      let finalName = nameMap[stuId] || (idMap[normId] ? idMap[normId].name : stuName);
      let finalId = idMap[normId] ? idMap[normId].id : stuId;

      if (!summary[finalId]) {
        summary[finalId] = { id: finalId, name: finalName, gradeLevel: rowGrade, room: rowRoom, present: 0, absent: 0, leave: 0, late: 0, sick: 0, total: 0 };
      }
      const rec = summary[finalId];
      if (status === "มา") rec.present++;
      else if (status === "ขาด") rec.absent++;
      else if (status === "ลา") rec.leave++;
      else if (status === "สาย") rec.late++;
      else if (status === "ป่วย") rec.sick++;
      rec.total++;

      if (!dailyTrend[rowDate]) dailyTrend[rowDate] = { present: 0, absent: 0, leave: 0, late: 0, sick: 0 };
      if (status === "มา") dailyTrend[rowDate].present++;
      else if (status === "ขาด") dailyTrend[rowDate].absent++;
      else if (status === "ลา") dailyTrend[rowDate].leave++;
      else if (status === "สาย") dailyTrend[rowDate].late++;
      else if (status === "ป่วย") dailyTrend[rowDate].sick++;
    }

    const students = Object.values(summary);
    students.sort((a, b) => {
      const ga = a.gradeLevel + "/" + a.room;
      const gb = b.gradeLevel + "/" + b.room;
      if (ga !== gb) return ga.localeCompare(gb, 'th');
      return String(a.id).localeCompare(String(b.id));
    });

    const totals = { present: 0, absent: 0, leave: 0, late: 0, sick: 0, total: 0 };
    students.forEach(s => {
      totals.present += s.present; totals.absent += s.absent;
      totals.leave += s.leave; totals.late += s.late;
      totals.sick += s.sick; totals.total += s.total;
    });

    return { status: "success", data: { students, totals, dailyTrend } };
  } catch (err) {
    Logger.log('Error in getSubjectAttendanceSummaryV2: ' + err.message);
    return { status: "error", message: err.toString() };
  }
}

function getSubjectStudentDetail(teacherIdentifier, subjectKey, studentId, startDate, endDate) {
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName("Attendance_Subjects");
    if (!sheet) return { status: "success", data: [] };

    const data = sheet.getDataRange().getValues();
    const result = [];
    const norm = String(teacherIdentifier || "").toLowerCase().trim();
    const subjKey = String(subjectKey || "").trim();
    const targetStu = String(studentId || "").trim().replace(/^0+/, ''); // ✅ ตัด 0 ออกเพื่อเทียบ

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowDate = formatDateStr(row[1]);
      const rowCode = String(row[2] || "").trim();
      const rowName = String(row[3] || "").trim();
      const rowTeacher = String(row[4] || "").trim().toLowerCase();
      // ✅ ตัด 0 ออกเพื่อเทียบ
      const rowStuId = String(row[7] || "").trim().replace(/^0+/, ''); 

      if (rowStuId !== targetStu) continue;
      if (startDate && rowDate < startDate) continue;
      if (endDate && rowDate > endDate) continue;
      if (rowTeacher !== norm) continue;

      const codeMatch = subjKey !== "" ? (rowCode === subjKey) : (rowName === subjKey);
      if (!codeMatch) continue;

      result.push({ 
        date: rowDate, 
        status: String(row[9] || "").trim(), 
        note: String(row[10] || "").trim() 
      });
    }

    result.sort((a, b) => b.date.localeCompare(a.date));
    return { status: "success", data: result };
  } catch (err) {
    Logger.log('Error in getSubjectStudentDetail: ' + err.message);
    return { status: "error", message: err.toString(), data: [] };
  }
}
