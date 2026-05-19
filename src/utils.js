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