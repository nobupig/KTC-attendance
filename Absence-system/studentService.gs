function getStudentsByClassId(classId) {
  const targetClassId = String(classId || "").trim();
  if (!targetClassId) {
    return [];
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = "studentsByClassId__" + targetClassId;
  const cached = cache.get(cacheKey);

  if (cached) {
    return JSON.parse(cached);
  }

  const ss = getMasterSpreadsheet();
  const classesSheet = ss.getSheetByName(CONFIG.SHEETS.CLASSES);
  const studentsSheet = ss.getSheetByName(CONFIG.SHEETS.STUDENTS);

  const classes = classesSheet.getDataRange().getValues();
  const students = studentsSheet.getDataRange().getValues();

  if (classes.length <= 1 || students.length <= 1) {
    return [];
  }

  const classRows = classes.slice(1);
  const studentRows = students.slice(1);

  const targetClass = classRows.find(row => String(row[0]).trim() === targetClassId);
  if (!targetClass) {
    return [];
  }

  const grade = String(targetClass[2]).trim();
  const unit = String(targetClass[3]).trim();

  const result = studentRows
    .filter(row => {
      return (
        String(row[1]).trim() === grade &&
        String(row[2]).trim() === unit &&
        String(row[5]).trim() === "active"
      );
    })
    .map(row => ({
      studentId: String(row[0]).trim(),
      grade: String(row[1]).trim(),
      unit: String(row[2]).trim(),
      attendanceNumber: row[3],
      name: String(row[4]).trim(),
      status: String(row[5]).trim()
    }))
    .sort((a, b) => Number(a.attendanceNumber) - Number(b.attendanceNumber));

  cache.put(cacheKey, JSON.stringify(result), 300);

  return result;
}

function getTeacherSessionDetail(classId, date, period) {
  return {
    students: getStudentsByClassId(classId),
    attendanceMap: getAttendanceMap(classId, date, period)
  };
}

function testGetStudentsByClassId() {
  const result = getStudentsByClassId('G1_MATH1A_1');
  Logger.log(JSON.stringify(result));
}