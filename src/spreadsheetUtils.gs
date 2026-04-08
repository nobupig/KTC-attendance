function getMasterSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEETS.MASTER);
}

function getOperationSpreadsheet() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEETS.OPERATION);
}

function getMasterSheet(name) {
  return getMasterSpreadsheet().getSheetByName(name);
}

function getOperationSheet(name) {
  return getOperationSpreadsheet().getSheetByName(name);
}

function isSafeScriptCacheKey_(key) {
  const normalizedKey = String(key || '');
  return normalizedKey.length > 0 && normalizedKey.length <= 240;
}

function getScriptCacheJson_(key) {
  if (!isSafeScriptCacheKey_(key)) {
    Logger.log('Cache get skip (invalid key): length=' + String(key || '').length);
    return null;
  }

  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get(key);
    return cached ? JSON.parse(cached) : null;
  } catch (e) {
    Logger.log('Cache get failed: ' + e + ' key=' + key);
    return null;
  }
}

function putScriptCacheJson_(key, value, ttlSeconds) {
  if (!isSafeScriptCacheKey_(key)) {
    Logger.log('Cache skip (key too large): length=' + String(key || '').length);
    return;
  }

  const cache = CacheService.getScriptCache();
  const serialized = safeJsonStringifyForCache_(value);

  // Script Cache の1キーあたり上限は約100KB
  // 余裕を見て 90KB を超えるものはキャッシュしない
  if (!serialized) {
    return;
  }

  if (serialized.length > 90000) {
    Logger.log('Cache skip (too large): ' + key + ' size=' + serialized.length);
    return;
  }

  try {
    cache.put(key, serialized, ttlSeconds || 300);
  } catch (e) {
    Logger.log('Cache put failed: ' + e + ' key=' + key);
  }
}

function safeJsonStringifyForCache_(value) {
  try {
    return JSON.stringify(value);
  } catch (e) {
    Logger.log('Cache stringify failed: ' + e);
    return '';
  }
}

function removeScriptCacheKeys_(keys) {
  const cache = CacheService.getScriptCache();
  (keys || []).forEach(function(key) {
    if (!isSafeScriptCacheKey_(key)) {
      return;
    }
    try {
      cache.remove(key);
    } catch (e) {
      Logger.log('Cache remove failed: ' + e + ' key=' + key);
    }
  });
}

function getSheetDataCached_(spreadsheetType, sheetName, ttlSeconds) {
  const cacheKey = 'sheetData__' + spreadsheetType + '__' + sheetName;
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const ss = spreadsheetType === 'MASTER'
    ? getMasterSpreadsheet()
    : getOperationSpreadsheet();

  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(sheetName + ' シートが見つかりません');
  }

  const values = sheet.getDataRange().getValues();
  const data = {
    headers: values.length > 0 ? values[0] : [],
    rows: values.length > 1 ? values.slice(1) : []
  };

  putScriptCacheJson_(cacheKey, data, ttlSeconds || 300);
  return data;
}

function removeScriptCacheByPrefix_(prefix) {
  const cache = CacheService.getScriptCache();
  const keys = cache.getAll([]); // ダミー呼び出しでは一覧取得不可のため未使用
  // ScriptCache は prefix 一括削除ができないため、
  // この関数は将来拡張用に残す。
  // 現時点では個別キー削除を使う。
  return prefix;
}

function buildAttendanceSessionCacheKey_(classId, date, period) {
  return 'attendanceMap__' +
    String(classId || '').trim() + '__' +
    formatDateToYmd(date) + '__' +
    String(period || '').trim();
}

function getAttendanceSheetCacheKey_() {
  return 'sheetData__OPERATION__' + CONFIG.SHEETS.ATTENDANCE;
}

function getAttendanceSessionsSheetCacheKey_() {
  return 'sheetData__OPERATION__' + CONFIG.SHEETS.ATTENDANCE_SESSIONS;
}

function buildHomeroomSummaryCacheKey_(grade, unit, termFilter) {
  return 'homeroomSummary__' +
    String(grade || '').trim() + '__' +
    String(unit || '').trim() + '__' +
    String(termFilter || 'all').trim();
}

function buildHomeroomDetailCacheKey_(studentId, termFilter) {
  return 'homeroomDetail__' +
    String(studentId || '').trim() + '__' +
    String(termFilter || 'all').trim();
}

function buildHomeroomInitialCacheKey_(email, termFilter) {
  return 'homeroomInitial__' +
    String(email || '').trim().toLowerCase() + '__' +
    String(termFilter || 'all').trim();
}

function clearHomeroomCachesForTest() {
  const cache = CacheService.getScriptCache();
  const user = getCurrentUserContext();
  const keys = [
    buildHomeroomInitialCacheKey_(user.email, 'all'),
    buildHomeroomInitialCacheKey_(user.email, '前期'),
    buildHomeroomInitialCacheKey_(user.email, '後期'),
    buildHomeroomInitialCacheKey_(user.email, '通年')
  ];
  cache.removeAll(keys);
}

function clearCoreCachesForTest() {
  const cache = CacheService.getScriptCache();

  const keys = [
    'sheetData__MASTER__' + CONFIG.SHEETS.CLASSES,
    'sheetData__MASTER__' + CONFIG.SHEETS.STUDENTS,
    'sheetData__MASTER__' + CONFIG.SHEETS.SUBJECTS,
    'sheetData__OPERATION__' + CONFIG.SHEETS.TIMETABLE,
    'sheetData__OPERATION__' + CONFIG.SHEETS.CLASS_TEACHER_TEAMS,
    'sheetData__OPERATION__' + CONFIG.SHEETS.CLASS_SESSIONS,
    'sheetData__OPERATION__' + CONFIG.SHEETS.TEACHERS,
    'sheetData__OPERATION__' + CONFIG.SHEETS.HOMEROOM_ASSIGNMENTS,

    // teacherService.gs 側の script cache
    'classTeacherTeamRows__all'
  ];

  cache.removeAll(keys);
  Logger.log('Core caches cleared');
}