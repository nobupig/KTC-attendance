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