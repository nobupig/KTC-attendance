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