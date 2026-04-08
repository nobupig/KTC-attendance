function initializeMasterSheets_legacy() {
  throw new Error('この初期化関数は旧構造用のため、現在は使用禁止です。');
  
}

function getTeacherNameByEmail(email) {

  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.TEACHERS);

  if (!sheet) {
    return email;
  }

  const values = sheet.getDataRange().getValues();
  const headers = values.shift();

  const emailIndex = headers.indexOf("email");
  const nameIndex = headers.indexOf("name");

  for (const row of values) {
    const rowEmail = String(row[emailIndex]).trim().toLowerCase();
    if (rowEmail === String(email).trim().toLowerCase()) {
      return row[nameIndex];
    }
  }

  return email;
}

function getClassDisplayName_legacy(classId) {

  const ss = getMasterSpreadsheet();

  const classesSheet = ss.getSheetByName(CONFIG.SHEETS.CLASSES);
  const subjectsSheet = ss.getSheetByName(CONFIG.SHEETS.SUBJECTS);

  const classes = classesSheet.getDataRange().getValues().slice(1);
  const subjects = subjectsSheet.getDataRange().getValues().slice(1);

  const targetClassId = String(classId).trim();

  for (const row of classes) {
    const rowClassId = String(row[0]).trim();   // ClassID
    const subjectId = String(row[1]).trim();    // SubjectID
    const grade = row[2];                       // 学年
    const unit = row[3];                        // 対象区分

    if (rowClassId === targetClassId) {
      const subjectRow = subjects.find(s => String(s[0]).trim() === subjectId);
      const subjectName = subjectRow ? String(subjectRow[1]).trim() : "";
      return `${grade}年${unit} ${subjectName}`.trim();
    }
  }

  return classId;
}