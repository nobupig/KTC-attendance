function getStudentsByClassId(classId) {
  const targetClassId = String(classId || "").trim();
  if (!targetClassId) {
    return [];
  }

  const cacheKey = "studentsByClassId__" + targetClassId;
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const classesData = getSheetDataCached_('MASTER', CONFIG.SHEETS.CLASSES, 300);
  const studentsData = getSheetDataCached_('MASTER', CONFIG.SHEETS.STUDENTS, 300);

  const classRows = classesData.rows;
  const studentRows = studentsData.rows;

  const targetClass = classRows.find(function(row) {
    return String(row[0]).trim() === targetClassId;
  });

  if (!targetClass) {
    return [];
  }

  const grade = String(targetClass[2]).trim();
  const unit = String(targetClass[3]).trim();

  const result = studentRows
    .filter(function(row) {
      return (
        String(row[1]).trim() === grade &&
        String(row[2]).trim() === unit &&
        String(row[5]).trim() === "active"
      );
    })
    .map(function(row) {
      return {
        studentId: String(row[0]).trim(),
        grade: String(row[1]).trim(),
        unit: String(row[2]).trim(),
        attendanceNumber: row[3],
        name: String(row[4]).trim(),
        status: String(row[5]).trim()
      };
    })
    .sort(function(a, b) {
      return Number(a.attendanceNumber) - Number(b.attendanceNumber);
    });

  putScriptCacheJson_(cacheKey, result, 300);
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