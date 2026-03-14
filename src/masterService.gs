function initializeMasterSheets() {
  const ss = getMasterSpreadsheet();

  const definitions = [
    {
      name: CONFIG.SHEETS.STUDENTS,
      headers: ['studentId', 'grade', 'unit', 'attendanceNumber', 'name', 'status']
    },
    {
      name: CONFIG.SHEETS.SUBJECTS,
      headers: ['subjectId', 'subjectName', 'isActive']
    },
    {
      name: CONFIG.SHEETS.CLASSES,
      headers: ['classId', 'subjectId', 'grade', 'unit', 'group']
    },
    {
      name: CONFIG.SHEETS.TEACHERS,
      headers: ['teacherId', 'name', 'email', 'role']
    },
    {
      name: CONFIG.SHEETS.TIMETABLE,
      headers: ['classId', 'weekday', 'period', 'teacherEmail']
    },
    {
      name: CONFIG.SHEETS.CALENDAR,
      headers: ['date', 'weekday', 'isClassDay']
    },
    {
      name: CONFIG.SHEETS.STUDENT_GROUPS,
      headers: ['studentId', 'subjectId', 'group']
    },
    {
      name: CONFIG.SHEETS.ATTENDANCE_STATUS,
      headers: ['code', 'label']
    },
    {
      name: CONFIG.SHEETS.HOMEROOM_TEACHERS,
      headers: ['grade', 'unit', 'teacherEmail']
    },
    {
      name: CONFIG.SHEETS.ATTENDANCE_SESSIONS,
      headers: ['classId', 'date', 'period', 'teacherEmail', 'accessedAt']
    },
    {
      name: CONFIG.SHEETS.ATTENDANCE,
      headers: ['classId', 'date', 'period', 'studentId', 'statusCode', 'recordedAt']
    },
    {
      name: CONFIG.SHEETS.CLASS_SESSIONS,
      headers: ['classId', 'date', 'period', 'sessionNumber']
    }
  ];

  definitions.forEach(def => {
    const sheet = getOrCreateSheet(ss, def.name);
    setHeader(sheet, def.headers);
  });
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

function getClassDisplayName(classId) {

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