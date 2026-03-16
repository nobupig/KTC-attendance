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
  const students = getStudentsByClassId(classId);
  const attendanceMap = getAttendanceMap(classId, date, period);
  const riskMap = getStudentRiskMapForClass(classId, students);

  return {
    students: students,
    attendanceMap: attendanceMap,
    riskMap: riskMap
  };
}

function getStudentRiskMapForClass(classId, students) {
  const targetClassId = String(classId || '').trim();
  const riskMap = {};

  if (!targetClassId || !students || !students.length) {
    return riskMap;
  }

  const classesData = getSheetDataCached_('MASTER', CONFIG.SHEETS.CLASSES, 300);
  const subjectsData = getSheetDataCached_('MASTER', CONFIG.SHEETS.SUBJECTS, 300);

  const classRows = classesData.rows;
  const subjectRows = subjectsData.rows;

  const targetClass = classRows.find(function(row) {
    return String(row[0] || '').trim() === targetClassId;
  });

  if (!targetClass) {
    return riskMap;
  }

  const targetSubjectId = String(targetClass[1] || '').trim();
  if (!targetSubjectId) {
    return riskMap;
  }

  const targetSubjectRow = subjectRows.find(function(row) {
    return String(row[0] || '').trim() === targetSubjectId;
  });

  const subjectName = targetSubjectRow ? String(targetSubjectRow[1] || '').trim() : '';
  const defaultLimit = targetSubjectRow ? Number(targetSubjectRow[3] || 0) : 0;

  const bundle = buildAbsenceCalculationBundle_('all');

  students.forEach(function(student) {
    const studentId = String(student.studentId || '').trim();
    if (!studentId) return;

    try {
      const result = calculateStudentAbsenceRiskFromBundle_(studentId, 'all', bundle);
      const targetSubject = (result.subjects || []).find(function(subject) {
        return String(subject.subjectId || '').trim() === targetSubjectId;
      });

      if (targetSubject) {
        riskMap[studentId] = {
          subjectId: targetSubject.subjectId,
          subjectName: targetSubject.subjectName,
          normalAbsence: Number(targetSubject.normalAbsence || 0),
          officialAbsence: Number(targetSubject.officialAbsence || 0),
          late: Number(targetSubject.late || 0),
          early: Number(targetSubject.early || 0),
          limit: Number(targetSubject.limit || 0),
          remaining: Number(targetSubject.remaining || 0),
          riskLevel: String(targetSubject.riskLevel || 'normal'),
          riskLabel: String(targetSubject.riskLabel || '正常')
        };
      } else {
        riskMap[studentId] = {
          subjectId: targetSubjectId,
          subjectName: subjectName,
          normalAbsence: 0,
          officialAbsence: 0,
          late: 0,
          early: 0,
          limit: defaultLimit,
          remaining: defaultLimit,
          riskLevel: 'normal',
          riskLabel: '正常'
        };
      }
    } catch (e) {
      riskMap[studentId] = {
        subjectId: targetSubjectId,
        subjectName: subjectName,
        normalAbsence: 0,
        officialAbsence: 0,
        late: 0,
        early: 0,
        limit: defaultLimit,
        remaining: defaultLimit,
        riskLevel: 'normal',
        riskLabel: '正常'
      };
    }
  });

  return riskMap;
}

function testGetStudentsByClassId() {
  const result = getStudentsByClassId('G1_MATH1A_1');
  Logger.log(JSON.stringify(result));
}