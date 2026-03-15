function getTeacherByClassId(classId) {

  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TIMETABLE);

  const values = sheet.getDataRange().getValues();
  const headers = values.shift();

  const classIndex = headers.indexOf("classId");
  const teacherIndex = headers.indexOf("teacherEmail");

  const row = values.find(r => r[classIndex] === classId);

  return row ? row[teacherIndex] : null;
}