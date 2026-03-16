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

function getScriptCacheJson_(key) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);
  return cached ? JSON.parse(cached) : null;
}

function putScriptCacheJson_(key, value, ttlSeconds) {
  const cache = CacheService.getScriptCache();
  cache.put(key, JSON.stringify(value), ttlSeconds || 300);
}

function removeScriptCacheKeys_(keys) {
  const cache = CacheService.getScriptCache();
  (keys || []).forEach(function(key) {
    cache.remove(key);
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