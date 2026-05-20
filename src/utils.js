function getSpreadsheetByName(name) {
  const files = DriveApp.getFilesByName(name);
  if (!files.hasNext()) {
    throw new Error(`スプレッドシートが見つかりません: ${name}`);
  }

  const file = files.next();
  return SpreadsheetApp.openById(file.getId());
}

function getOrCreateSheet(spreadsheet, sheetName) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

function setHeader(sheet, headers) {
  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}


function formatDateToYmd(date) {
  return Utilities.formatDate(new Date(date), 'Asia/Tokyo', 'yyyy-MM-dd');
}

function fastYmdForCompare_(value) {
  return formatDateToYmd(value);
}

function fastYmdFromCell_(value) {
  return formatDateToYmd(value);
}

function normalizeYmdDisplayText_(value) {
  const s = String(value || '').trim();
  if (!s) return '';

  const m1 = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m1) {
    return m1[1] + '-' + String(m1[2]).padStart(2, '0') + '-' + String(m1[3]).padStart(2, '0');
  }

  const m2 = s.match(/^(\d{4})\/(\d{1,2})\/(\d{1,2})/);
  if (m2) {
    return m2[1] + '-' + String(m2[2]).padStart(2, '0') + '-' + String(m2[3]).padStart(2, '0');
  }

  return '';
}

function logPerf_(label, startedAtMs, extra) {
  const elapsed = Date.now() - startedAtMs;
  if (extra) {
    console.log('[PERF] ' + label + ': ' + elapsed + 'ms | ' + extra);
  } else {
    console.log('[PERF] ' + label + ': ' + elapsed + 'ms');
  }
}

function perfNow_() {
  return Date.now();
}