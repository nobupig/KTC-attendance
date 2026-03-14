function getCurrentUserEmail() {
  return Session.getActiveUser().getEmail();
}

function testGetCurrentUserEmail() {
  Logger.log(getCurrentUserEmail());
}

function doGet(e) {

  const page = (e && e.parameter && e.parameter.page)
    ? e.parameter.page
    : 'index';

  if (page === 'teacher') {

    const template = HtmlService.createTemplateFromFile('html/teacher');

    template.classId = e.parameter.classId || "";
    template.date = e.parameter.date || "";
    template.period = e.parameter.period || "";

    return template
      .evaluate()
      .setTitle('科目担当画面');

  }

  if (page === 'homeroom') {
    return HtmlService.createHtmlOutputFromFile('html/homeroom')
      .setTitle('担任画面');
  }

  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('html/admin')
      .setTitle('教務画面');
  }

  return HtmlService.createHtmlOutputFromFile('html/index')
    .setTitle('出席管理システム');

}

function getCurrentUserContext() {
  const email = Session.getActiveUser().getEmail().toLowerCase().trim();

  const ss = getOperationSpreadsheet();

  const teachersSheet = ss.getSheetByName(CONFIG.SHEETS.TEACHERS);
  const homeroomSheet = ss.getSheetByName(CONFIG.SHEETS.HOMEROOM_ASSIGNMENTS);

  if (!teachersSheet) {
    throw new Error('teachers シートが見つかりません');
  }
  if (!homeroomSheet) {
    throw new Error('homeroomAssignments シートが見つかりません');
  }

  const teachersData = teachersSheet.getDataRange().getValues();
  const homeroomData = homeroomSheet.getDataRange().getValues();

  const headers = teachersData.shift();
  const roleIndex = headers.indexOf("role");
  const emailIndex = headers.indexOf("email");
  const idIndex = headers.indexOf("teacherId");
  const nameIndex = headers.indexOf("name");

  let teacherId = null;
  let name = null;
  const roles = new Set();

  teachersData.forEach(row => {
    if (String(row[emailIndex]).toLowerCase().trim() === email) {
      teacherId = row[idIndex];
      name = row[nameIndex];
      roles.add(row[roleIndex]);
    }
  });

  const homeroomClasses = [];

  const hrHeaders = homeroomData.shift();
  const hrEmailIndex = hrHeaders.indexOf("email");
  const gradeIndex = hrHeaders.indexOf("grade");
  const unitIndex = hrHeaders.indexOf("unit");

  homeroomData.forEach(row => {
    if (String(row[hrEmailIndex]).toLowerCase().trim() === email) {
      homeroomClasses.push({
        grade: row[gradeIndex],
        unit: row[unitIndex]
      });
    }
  });

  return {
    email,
    teacherId,
    name,
    roles: Array.from(roles),
    homeroomClasses
  };
}

function getUserRoleInfo() {
  const user = getCurrentUserContext();

  return {
    roles: user.roles,
    homeroomClasses: user.homeroomClasses,
    name: user.name
  };
}