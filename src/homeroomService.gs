function getMyHomeroomClasses() {
  const user = getCurrentUserContext();

  if (!user) {
    throw new Error('ユーザー情報を取得できませんでした');
  }

  const classes = Array.isArray(user.homeroomClasses)
    ? user.homeroomClasses
    : [];

  return sortHomeroomClasses_(
    classes.map(function(item) {
      return {
        grade: String(item.grade || '').trim(),
        unit: String(item.unit || '').trim()
      };
    })
  );
}

function getHomeroomInitialData(termFilter) {
  const user = getCurrentUserContext();
  const homeroomClasses = getMyHomeroomClasses();
  const normalizedTermFilter = normalizeTermFilter_(termFilter);

  const cacheKey = buildHomeroomInitialCacheKey_(user.email, normalizedTermFilter);
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  let firstSummary = null;
  if (homeroomClasses.length > 0) {
    const first = homeroomClasses[0];
    firstSummary = getHomeroomRiskSummary(
      first.grade,
      first.unit,
      normalizedTermFilter
    );
  }

  const result = {
    user: {
      email: user.email,
      name: user.name,
      roles: user.roles || []
    },
    homeroomClasses: homeroomClasses,
    selectedTermFilter: normalizedTermFilter,
    firstSummary: firstSummary
  };

  putScriptCacheJson_(cacheKey, result, 120);
  return result;
}

function getHomeroomRiskSummary(grade, unit, termFilter) {
  const targetGrade = String(grade || '').trim();
  const targetUnit = String(unit || '').trim();
  const normalizedTermFilter = normalizeTermFilter_(termFilter);

  if (!targetGrade || !targetUnit) {
    throw new Error('学年または組が指定されていません');
  }

  ensureHomeroomAccess_(targetGrade, targetUnit);

  const cacheKey = buildHomeroomSummaryCacheKey_(
    targetGrade,
    targetUnit,
    normalizedTermFilter
  );
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const result = calculateHomeroomRiskSummary(
    targetGrade,
    targetUnit,
    normalizedTermFilter
  );

  const response = {
    grade: result.grade,
    unit: result.unit,
    classLabel: buildHomeroomClassLabel_(result.grade, result.unit),
    termFilter: normalizedTermFilter,
    students: result.students || [],
    summary: result.summary || {
      totalStudents: 0,
      warningStudents: 0,
      officialOverStudents: 0,
      normalOverStudents: 0
    }
  };

  putScriptCacheJson_(cacheKey, response, 120);
  return response;
}

function getStudentAbsenceRiskDetail(studentId, termFilter) {
  const targetStudentId = String(studentId || '').trim();
  const normalizedTermFilter = normalizeTermFilter_(termFilter);

  if (!targetStudentId) {
    throw new Error('studentId が指定されていません');
  }

  const cacheKey = buildHomeroomDetailCacheKey_(
    targetStudentId,
    normalizedTermFilter
  );
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const student = getStudentBasicInfoById_(targetStudentId);
  if (!student) {
    throw new Error('対象学生が見つかりません: ' + targetStudentId);
  }

  ensureHomeroomAccess_(student.grade, student.unit);

  const result = calculateStudentAbsenceRisk(
    targetStudentId,
    normalizedTermFilter
  );

  const filteredSubjects = (result.subjects || []).filter(subject =>
    (subject.normalAbsence || 0) > 0 ||
    (subject.officialAbsence || 0) > 0 ||
    (subject.late || 0) > 0 ||
    (subject.normalOver || 0) > 0 ||
    (subject.officialOver || 0) > 0
  );

  const response = {
    student: {
      studentId: result.student.studentId,
      grade: result.student.grade,
      unit: result.student.unit,
      attendanceNumber: result.student.attendanceNumber,
      name: result.student.name,
      status: result.student.status || ''
    },
    classLabel: buildHomeroomClassLabel_(result.student.grade, result.student.unit),
    termFilter: normalizedTermFilter,
    subjects: filteredSubjects,
    summary: {
      subjectCount: filteredSubjects.length,
      warningSubjects: filteredSubjects.filter(s => s.riskLevel === 'warning').length,
      officialOverSubjects: filteredSubjects.filter(s => s.riskLevel === 'official_over').length,
      normalOverSubjects: filteredSubjects.filter(s => s.riskLevel === 'normal_over').length,
      highestRiskLevel: filteredSubjects.some(s => s.riskLevel === 'normal_over')
        ? 'normal_over'
        : filteredSubjects.some(s => s.riskLevel === 'official_over')
          ? 'official_over'
          : filteredSubjects.some(s => s.riskLevel === 'warning')
            ? 'warning'
            : 'normal',
      highestRiskLabel: filteredSubjects.some(s => s.riskLevel === 'normal_over')
        ? '留級対象'
        : filteredSubjects.some(s => s.riskLevel === 'official_over')
          ? '公欠超過'
          : filteredSubjects.some(s => s.riskLevel === 'warning')
            ? '注意'
            : '正常'
    }
  };

  putScriptCacheJson_(cacheKey, response, 120);
  return response;
}

function getHomeroomClassDetail(grade, unit, termFilter) {
  const summary = getHomeroomRiskSummary(grade, unit, termFilter);

  let firstStudentDetail = null;
  if (summary.students && summary.students.length > 0) {
    firstStudentDetail = getStudentAbsenceRiskDetail(
      summary.students[0].studentId,
      termFilter
    );
  }

  return {
    classInfo: {
      grade: summary.grade,
      unit: summary.unit,
      classLabel: summary.classLabel
    },
    termFilter: summary.termFilter,
    students: summary.students,
    summary: summary.summary,
    firstStudentDetail: firstStudentDetail
  };
}

function testGetMyHomeroomClasses() {
  const result = getMyHomeroomClasses();
  Logger.log(JSON.stringify(result, null, 2));
}

function testGetHomeroomRiskSummary() {
  const result = getHomeroomRiskSummary('1', '1', 'all');
  Logger.log(JSON.stringify(result, null, 2));
}

function testGetStudentAbsenceRiskDetail() {
  const result = getStudentAbsenceRiskDetail('S001', 'all');
  Logger.log(JSON.stringify(result, null, 2));
}

/* =========================
 * 内部ヘルパー
 * ========================= */

function ensureHomeroomAccess_(grade, unit) {
  const user = getCurrentUserContext();
  const targetGrade = String(grade || '').trim();
  const targetUnit = String(unit || '').trim();

  const homeroomClasses = Array.isArray(user.homeroomClasses)
    ? user.homeroomClasses
    : [];

  const allowed = homeroomClasses.some(function(item) {
    const assignedGrade = String(item.grade || '').trim();
    const assignedUnit = String(item.unit || '').trim();

    if (assignedGrade !== targetGrade) {
      return false;
    }

    return doesHomeroomUnitCoverTargetUnit_(assignedUnit, targetUnit);
  });

  if (!allowed) {
    throw new Error('このクラスの担任画面を閲覧する権限がありません');
  }

  return true;
}

function doesHomeroomUnitCoverTargetUnit_(assignedUnit, targetUnit) {
  const assigned = String(assignedUnit || '').trim().toUpperCase();
  const target = String(targetUnit || '').trim().toUpperCase();

  if (assigned === target) {
    return true;
  }

  if (assigned === 'CA' && (target === 'C' || target === 'A' || target === 'CA')) {
    return true;
  }

  return false;
}

function getAllHomeroomClasses_() {
  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.HOMEROOM_ASSIGNMENTS);

  if (!sheet) {
    return [];
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return [];
  }

  const headers = values[0];
  const rows = values.slice(1);

  const gradeIndex = headers.indexOf('grade');
  const unitIndex = headers.indexOf('unit');

  if (gradeIndex === -1 || unitIndex === -1) {
    throw new Error('homeroomAssignments シートに grade または unit 列がありません');
  }

  const map = {};

  rows.forEach(row => {
    const grade = String(row[gradeIndex] || '').trim();
    const unit = String(row[unitIndex] || '').trim();
    if (!grade || !unit) return;

    const key = grade + '__' + unit;
    if (!map[key]) {
      map[key] = { grade: grade, unit: unit };
    }
  });

  return Object.values(map);
}

function sortHomeroomClasses_(classes) {
  return (classes || [])
    .filter(item => String(item.grade || '').trim() && String(item.unit || '').trim())
    .sort((a, b) => {
      const gradeDiff = Number(a.grade) - Number(b.grade);
      if (gradeDiff !== 0) return gradeDiff;
      return String(a.unit).localeCompare(String(b.unit), 'ja');
    })
    .map(item => ({
      grade: String(item.grade).trim(),
      unit: String(item.unit).trim(),
      classLabel: buildHomeroomClassLabel_(item.grade, item.unit)
    }));
}

function buildHomeroomClassLabel_(grade, unit) {
  return String(grade || '').trim() + '年' + String(unit || '').trim();
}

function normalizeTermFilter_(termFilter) {
  const value = String(termFilter || '').trim();
  if (!value) return 'all';

  if (value === '前期' || value === '後期' || value === '通年' || value === 'all') {
    return value;
  }

  return 'all';
}

function getStudentBasicInfoById_(studentId) {
  const targetStudentId = String(studentId || '').trim();
  if (!targetStudentId) {
    return null;
  }

  const studentsData = getSheetDataCached_('MASTER', CONFIG.SHEETS.STUDENTS, 300);
  const headers = studentsData.headers;
  const rows = studentsData.rows;

  const col = {
    studentId: findColumnIndexForHomeroom_(headers, ['studentId', 'StudentID']),
    grade: findColumnIndexForHomeroom_(headers, ['grade', '学年']),
    unit: findColumnIndexForHomeroom_(headers, ['unit', '組', '対象区分', '組・コース']),
    attendanceNumber: findColumnIndexForHomeroom_(headers, ['attendanceNumber', 'number', '出席番号']),
    name: findColumnIndexForHomeroom_(headers, ['name', '氏名']),
    status: findColumnIndexForHomeroom_(headers, ['status', '在籍状態'])
  };

  if (
    col.studentId === -1 ||
    col.grade === -1 ||
    col.unit === -1 ||
    col.attendanceNumber === -1 ||
    col.name === -1
  ) {
    throw new Error('students シートに必要な列がありません');
  }

  const row = rows.find(r => String(r[col.studentId] || '').trim() === targetStudentId);
  if (!row) {
    return null;
  }

  return {
    studentId: String(row[col.studentId] || '').trim(),
    grade: String(row[col.grade] || '').trim(),
    unit: String(row[col.unit] || '').trim(),
    attendanceNumber: String(row[col.attendanceNumber] || '').trim(),
    name: String(row[col.name] || '').trim(),
    status: col.status === -1 ? '' : String(row[col.status] || '').trim()
  };
}

function findColumnIndexForHomeroom_(headers, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const idx = headers.indexOf(candidates[i]);
    if (idx !== -1) {
      return idx;
    }
  }
  return -1;
}