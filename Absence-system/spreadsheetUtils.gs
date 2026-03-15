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