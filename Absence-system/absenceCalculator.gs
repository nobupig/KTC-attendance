const ABSENCE_CALCULATOR_CONFIG = {
  WARNING_REMAINING: 1 // 残り1コマ以下で「注意」
};

/**
 * 生徒1人の科目別欠席状況を計算
 * @param {string} studentId
 * @param {string=} termFilter 例: '前期', '後期', '通年', 'all'
 * @returns {{
 *   student: Object,
 *   subjects: Array,
 *   summary: Object
 * }}
 */
function calculateStudentAbsenceRisk(studentId, termFilter) {
  const bundle = buildAbsenceCalculationBundle_(termFilter);
  return calculateStudentAbsenceRiskFromBundle_(studentId, termFilter, bundle);
}

function calculateStudentAbsenceRiskFromBundle_(studentId, termFilter, bundle) {
  const targetStudentId = String(studentId || '').trim();

  if (!targetStudentId) {
    throw new Error('studentId が必要です');
  }

  const student = bundle.studentMap[targetStudentId];
  if (!student) {
    throw new Error('対象学生が見つかりません: ' + targetStudentId);
  }

  const targetKeys = getTargetSubjectKeysForStudent_(student, bundle.classMap, bundle.subjectMap, termFilter);
  const subjectTotals = initializeSubjectTotals_(targetKeys, bundle.subjectMap);

  bundle.attendanceRows.forEach(row => {
    const rowStudentId = String(row[bundle.attendanceCol.studentId] || '').trim();
    if (rowStudentId !== targetStudentId) return;

    const classId = String(row[bundle.attendanceCol.classId] || '').trim();
    const statusCode = normalizeStatusCode_(row[bundle.attendanceCol.statusCode]);

    const cls = bundle.classMap[classId];
    if (!cls) return;

    const subject = bundle.subjectMap[cls.subjectId];
    if (!subject) return;

    if (!isSubjectIncludedByTerm_(subject.term, termFilter)) return;

    const subjectKey = cls.subjectId;
    if (!subjectTotals[subjectKey]) {
      // 念のため、対象科目に含まれていないが attendance に存在する場合も拾う
      subjectTotals[subjectKey] = createEmptySubjectTotal_(subject);
    }

    const rule = bundle.statusRuleMap[statusCode] || createDefaultStatusRule_(statusCode);

subjectTotals[subjectKey].normalAbsence += rule.normalAbsence;
subjectTotals[subjectKey].officialAbsence += rule.officialAbsence;

if (statusCode === 'L') {
  subjectTotals[subjectKey].late++;
}

if (statusCode === 'E') {
  subjectTotals[subjectKey].early++;
}

subjectTotals[subjectKey].records.push({
  date: formatDateToYmd(row[bundle.attendanceCol.date]),
  period: row[bundle.attendanceCol.period],
  status: statusCode
});
  });

  const subjects = Object.keys(subjectTotals)
    .map(subjectId => finalizeSubjectRisk_(subjectTotals[subjectId]))
    .sort((a, b) => {
      if (a.riskOrder !== b.riskOrder) return b.riskOrder - a.riskOrder;
      return a.subjectName.localeCompare(b.subjectName, 'ja');
    });

  const summary = summarizeSubjectRisks_(subjects);

  return {
    student: student,
    subjects: subjects,
    summary: summary
  };
}

/**
 * 担任クラス全体の欠席リスクサマリーを計算
 * @param {string|number} grade
 * @param {string|number} unit
 * @param {string=} termFilter 例: '前期', '後期', '通年', 'all'
 * @returns {{
 *   grade: string,
 *   unit: string,
 *   students: Array,
 *   summary: Object
 * }}
 */
function calculateHomeroomRiskSummary(grade, unit, termFilter) {
  const bundle = buildAbsenceCalculationBundle_(termFilter);

  const targetGrade = String(grade || '').trim();
  const targetUnit = String(unit || '').trim();

  const allStudents = Object.values(bundle.studentMap)
    .filter(student =>
      String(student.grade || '').trim() === targetGrade &&
      String(student.unit || '').trim() === targetUnit &&
      isStudentActive_(student.status)
    );

  const evaluatedStudents = allStudents.map(student => {
    const result = calculateStudentAbsenceRiskFromBundle_(student.studentId, termFilter, bundle);
    const subjects = result.subjects || [];

    const totalNormalAbsence = subjects.reduce((sum, s) => sum + Number(s.normalAbsence || 0), 0);
    const totalOfficialAbsence = subjects.reduce((sum, s) => sum + Number(s.officialAbsence || 0), 0);
    const totalLate = subjects.reduce((sum, s) => sum + Number(s.late || 0), 0);

    const hasAttendanceIssue =
      totalNormalAbsence > 0 ||
      totalOfficialAbsence > 0 ||
      totalLate > 0;

    return {
      studentId: student.studentId,
      attendanceNumber: student.attendanceNumber,
      name: student.name,
      grade: student.grade,
      unit: student.unit,
      totalNormalAbsence: totalNormalAbsence,
      totalOfficialAbsence: totalOfficialAbsence,
      totalLate: totalLate,
      hasAttendanceIssue: hasAttendanceIssue,
      warningSubjects: result.summary.warningSubjects,
      officialOverSubjects: result.summary.officialOverSubjects,
      normalOverSubjects: result.summary.normalOverSubjects,
      highestRiskLevel: result.summary.highestRiskLevel,
      highestRiskLabel: result.summary.highestRiskLabel
    };
  });

  const visibleStudents = evaluatedStudents
    .filter(student => student.hasAttendanceIssue)
    .sort((a, b) => {
      const riskOrder = {
        normal_over: 3,
        official_over: 2,
        warning: 1,
        normal: 0
      };

      const riskDiff = (riskOrder[b.highestRiskLevel] || 0) - (riskOrder[a.highestRiskLevel] || 0);
      if (riskDiff !== 0) return riskDiff;

      const issueA = (a.totalNormalAbsence || 0) + (a.totalOfficialAbsence || 0) + (a.totalLate || 0);
      const issueB = (b.totalNormalAbsence || 0) + (b.totalOfficialAbsence || 0) + (b.totalLate || 0);
      if (issueB !== issueA) return issueB - issueA;

      return Number(a.attendanceNumber || 9999) - Number(b.attendanceNumber || 9999);
    });

  const summary = {
    totalStudents: allStudents.length,
    visibleStudents: visibleStudents.length,
    warningStudents: evaluatedStudents.filter(s => s.highestRiskLevel === 'warning').length,
    officialOverStudents: evaluatedStudents.filter(s => s.highestRiskLevel === 'official_over').length,
    normalOverStudents: evaluatedStudents.filter(s => s.highestRiskLevel === 'normal_over').length
  };

  return {
    grade: targetGrade,
    unit: targetUnit,
    students: visibleStudents,
    summary: summary
  };
}

/**
 * テスト用: 生徒1人の科目別結果をログ出力
 */
function testCalculateStudentAbsenceRisk() {
  const result = calculateStudentAbsenceRisk('S001', 'all');
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * テスト用: 担任クラス全体のサマリーをログ出力
 */
function testCalculateHomeroomRiskSummary() {
  const result = calculateHomeroomRiskSummary('1', '1', 'all');
  Logger.log(JSON.stringify(result, null, 2));
}

/* =========================
 * 以下、内部ヘルパー
 * ========================= */

function buildAbsenceCalculationBundle_(termFilter) {
  const masterSs = getMasterSpreadsheet();
  const operationSs = getOperationSpreadsheet();

  const studentsSheet = masterSs.getSheetByName(CONFIG.SHEETS.STUDENTS);
  const subjectsSheet = masterSs.getSheetByName(CONFIG.SHEETS.SUBJECTS);
  const classesSheet = masterSs.getSheetByName(CONFIG.SHEETS.CLASSES);
  const attendanceSheet = operationSs.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  const attendanceStatusSheet = operationSs.getSheetByName(CONFIG.SHEETS.ATTENDANCE_STATUS);

  if (!studentsSheet) throw new Error('students シートがありません');
  if (!subjectsSheet) throw new Error('Subjects シートがありません');
  if (!classesSheet) throw new Error('classes シートがありません');
  if (!attendanceSheet) throw new Error('attendance シートがありません');

  const studentsHeaders = getSheetHeadersForAbsence_(studentsSheet);
  const studentsRows = getSheetBodyValuesForAbsence_(studentsSheet);

  const subjectsHeaders = getSheetHeadersForAbsence_(subjectsSheet);
  const subjectsRows = getSheetBodyValuesForAbsence_(subjectsSheet);

  const classesHeaders = getSheetHeadersForAbsence_(classesSheet);
  const classesRows = getSheetBodyValuesForAbsence_(classesSheet);

  const attendanceHeaders = getSheetHeadersForAbsence_(attendanceSheet);
  const attendanceRows = getSheetBodyValuesForAbsence_(attendanceSheet);

  const studentCol = {
    studentId: findColumnIndexForAbsence_(studentsHeaders, ['studentId', 'StudentID']),
    grade: findColumnIndexForAbsence_(studentsHeaders, ['grade', '学年']),
    unit: findColumnIndexForAbsence_(studentsHeaders, ['unit', '組', '対象区分', '組・コース']),
    attendanceNumber: findColumnIndexForAbsence_(studentsHeaders, ['attendanceNumber', 'number', '出席番号']),
    name: findColumnIndexForAbsence_(studentsHeaders, ['name', '氏名']),
    status: findColumnIndexForAbsence_(studentsHeaders, ['status', '在籍状態'])
  };
  validateColumnsForAbsence_('students', studentCol, ['studentId', 'grade', 'unit', 'attendanceNumber', 'name']);

  const subjectCol = {
    subjectId: findColumnIndexForAbsence_(subjectsHeaders, ['SubjectID', 'subjectId']),
    subjectName: findColumnIndexForAbsence_(subjectsHeaders, ['科目名', 'subjectName']),
    term: findColumnIndexForAbsence_(subjectsHeaders, ['開設期', 'term']),
    limit: findColumnIndexForAbsence_(subjectsHeaders, ['欠席可能コマ数', 'limit'])
  };
  validateColumnsForAbsence_('Subjects', subjectCol, ['subjectId', 'subjectName', 'term', 'limit']);

  const classCol = {
    classId: findColumnIndexForAbsence_(classesHeaders, ['ClassID', 'classId']),
    subjectId: findColumnIndexForAbsence_(classesHeaders, ['SubjectID', 'subjectId']),
    grade: findColumnIndexForAbsence_(classesHeaders, ['学年', 'grade']),
    unit: findColumnIndexForAbsence_(classesHeaders, ['対象区分', 'unit', '組・コース'])
  };
  validateColumnsForAbsence_('classes', classCol, ['classId', 'subjectId', 'grade', 'unit']);

  const attendanceCol = {
    classId: findColumnIndexForAbsence_(attendanceHeaders, ['classId', 'ClassID']),
    date: findColumnIndexForAbsence_(attendanceHeaders, ['date', '日付']),
    period: findColumnIndexForAbsence_(attendanceHeaders, ['period', '時限']),
    studentId: findColumnIndexForAbsence_(attendanceHeaders, ['studentId', 'StudentID']),
    statusCode: findColumnIndexForAbsence_(attendanceHeaders, ['statusCode', 'code', '出欠コード'])
  };
  validateColumnsForAbsence_('attendance', attendanceCol, ['classId', 'studentId', 'statusCode']);

  const studentMap = {};
  studentsRows.forEach(row => {
    const studentId = String(row[studentCol.studentId] || '').trim();
    if (!studentId) return;

    studentMap[studentId] = {
      studentId: studentId,
      grade: String(row[studentCol.grade] || '').trim(),
      unit: String(row[studentCol.unit] || '').trim(),
      attendanceNumber: String(row[studentCol.attendanceNumber] || '').trim(),
      name: String(row[studentCol.name] || '').trim(),
      status: studentCol.status === -1 ? '' : String(row[studentCol.status] || '').trim()
    };
  });

  const subjectMap = {};
  subjectsRows.forEach(row => {
    const subjectId = String(row[subjectCol.subjectId] || '').trim();
    if (!subjectId) return;

    subjectMap[subjectId] = {
      subjectId: subjectId,
      subjectName: String(row[subjectCol.subjectName] || '').trim(),
      term: String(row[subjectCol.term] || '').trim(),
      limit: toNumberForAbsence_(row[subjectCol.limit])
    };
  });

  const classMap = {};
  classesRows.forEach(row => {
    const classId = String(row[classCol.classId] || '').trim();
    if (!classId) return;

    classMap[classId] = {
      classId: classId,
      subjectId: String(row[classCol.subjectId] || '').trim(),
      grade: String(row[classCol.grade] || '').trim(),
      unit: String(row[classCol.unit] || '').trim()
    };
  });

  const statusRuleMap = buildAttendanceStatusRuleMap_(attendanceStatusSheet);

  return {
    studentMap: studentMap,
    subjectMap: subjectMap,
    classMap: classMap,
    attendanceRows: attendanceRows,
    attendanceCol: attendanceCol,
    statusRuleMap: statusRuleMap
  };
}

function buildAttendanceStatusRuleMap_(attendanceStatusSheet) {
  const defaultMap = {
    P: { normalAbsence: 0, officialAbsence: 0 },
    A: { normalAbsence: 1, officialAbsence: 0 },
    O: { normalAbsence: 0, officialAbsence: 1 },
    L: { normalAbsence: 0, officialAbsence: 0 },
    E: { normalAbsence: 0, officialAbsence: 0 },

    // 旧コード対策
    ABS: { normalAbsence: 1, officialAbsence: 0 },
    OFF: { normalAbsence: 0, officialAbsence: 1 },
    LAT: { normalAbsence: 0, officialAbsence: 0 }
  };

  if (!attendanceStatusSheet) {
    return defaultMap;
  }

  const headers = getSheetHeadersForAbsence_(attendanceStatusSheet);
  const rows = getSheetBodyValuesForAbsence_(attendanceStatusSheet);

  const col = {
    code: findColumnIndexForAbsence_(headers, ['code', 'statusCode']),
    normalAbsence: findColumnIndexForAbsence_(headers, ['normalAbsence', '通常欠席換算']),
    officialAbsence: findColumnIndexForAbsence_(headers, ['officialAbsence', '公欠換算'])
  };

  if (col.code === -1 || col.normalAbsence === -1 || col.officialAbsence === -1) {
    return defaultMap;
  }

  const map = Object.assign({}, defaultMap);

  rows.forEach(row => {
    const code = normalizeStatusCode_(row[col.code]);
    if (!code) return;

    map[code] = {
      normalAbsence: toNumberForAbsence_(row[col.normalAbsence]),
      officialAbsence: toNumberForAbsence_(row[col.officialAbsence])
    };
  });

  return map;
}

function getTargetSubjectKeysForStudent_(student, classMap, subjectMap, termFilter) {
  const keys = [];

  Object.values(classMap).forEach(cls => {
    if (String(cls.grade) !== String(student.grade)) return;
    if (String(cls.unit) !== String(student.unit)) return;

    const subject = subjectMap[cls.subjectId];
    if (!subject) return;
    if (!isSubjectIncludedByTerm_(subject.term, termFilter)) return;

    if (!keys.includes(cls.subjectId)) {
      keys.push(cls.subjectId);
    }
  });

  return keys;
}

function initializeSubjectTotals_(subjectIds, subjectMap) {
  const result = {};

  subjectIds.forEach(subjectId => {
    const subject = subjectMap[subjectId];
    if (!subject) return;
    result[subjectId] = createEmptySubjectTotal_(subject);
  });

  return result;
}

function createEmptySubjectTotal_(subject) {
  return {
    subjectId: subject.subjectId,
    subjectName: subject.subjectName,
    term: subject.term,
    limit: toNumberForAbsence_(subject.limit),
    normalAbsence: 0,
    officialAbsence: 0,
    late: 0,
    early: 0,
    records: []
  };
}

function finalizeSubjectRisk_(item) {
  const limit = toNumberForAbsence_(item.limit);
  const normalAbsence = toNumberForAbsence_(item.normalAbsence);
  const officialAbsence = toNumberForAbsence_(item.officialAbsence);
  const totalAbsence = normalAbsence + officialAbsence;

  const normalOver = Math.max(normalAbsence - limit, 0);
  const officialOver = normalAbsence <= limit
    ? Math.max(totalAbsence - limit, 0)
    : 0;

  const remaining = limit - totalAbsence;

  let riskLevel = 'normal';
  let riskLabel = '正常';
  let riskOrder = 0;

  if (normalOver > 0) {
    riskLevel = 'normal_over';
    riskLabel = '留級対象';
    riskOrder = 3;
  } else if (officialOver > 0) {
    riskLevel = 'official_over';
    riskLabel = '公欠超過';
    riskOrder = 2;
  } else if (remaining <= ABSENCE_CALCULATOR_CONFIG.WARNING_REMAINING) {
    riskLevel = 'warning';
    riskLabel = '注意';
    riskOrder = 1;
  }

return {
  subjectId: item.subjectId,
  subjectName: item.subjectName,
  term: item.term,
  limit: limit,
  normalAbsence: normalAbsence,
  officialAbsence: officialAbsence,
  late: item.late,
  early: item.early,
  totalAbsence: totalAbsence,
  remaining: remaining,
  normalOver: normalOver,
  officialOver: officialOver,
  riskLevel: riskLevel,
  riskLabel: riskLabel,
  riskOrder: riskOrder,
  records: item.records
};
}

function summarizeSubjectRisks_(subjects) {
  const normalOverSubjects = subjects.filter(s => s.riskLevel === 'normal_over').length;
  const officialOverSubjects = subjects.filter(s => s.riskLevel === 'official_over').length;
  const warningSubjects = subjects.filter(s => s.riskLevel === 'warning').length;

  let highestRiskLevel = 'normal';
  let highestRiskLabel = '正常';

  if (normalOverSubjects > 0) {
    highestRiskLevel = 'normal_over';
    highestRiskLabel = '留級対象';
  } else if (officialOverSubjects > 0) {
    highestRiskLevel = 'official_over';
    highestRiskLabel = '公欠超過';
  } else if (warningSubjects > 0) {
    highestRiskLevel = 'warning';
    highestRiskLabel = '注意';
  }

  return {
    subjectCount: subjects.length,
    warningSubjects: warningSubjects,
    officialOverSubjects: officialOverSubjects,
    normalOverSubjects: normalOverSubjects,
    highestRiskLevel: highestRiskLevel,
    highestRiskLabel: highestRiskLabel
  };
}

function isSubjectIncludedByTerm_(subjectTerm, termFilter) {
  const term = String(subjectTerm || '').trim();
  const filter = String(termFilter || 'all').trim();

  if (!filter || filter === 'all') {
    return true;
  }

  if (term === '通年') {
    return true;
  }

  return term === filter;
}

function isStudentActive_(status) {
  const s = String(status || '').trim();
  if (!s) return true;
  return s !== 'inactive' && s !== '卒業' && s !== '退学';
}

function normalizeStatusCode_(value) {
  return String(value || '').trim().toUpperCase();
}

function createDefaultStatusRule_(statusCode) {
  const code = normalizeStatusCode_(statusCode);

  switch (code) {
    case 'A':
    case 'ABS':
      return { normalAbsence: 1, officialAbsence: 0 };
    case 'O':
    case 'OFF':
      return { normalAbsence: 0, officialAbsence: 1 };
    case 'P':
    case 'L':
    case 'E':
    case 'LAT':
    default:
      return { normalAbsence: 0, officialAbsence: 0 };
  }
}

function toNumberForAbsence_(value) {
  if (typeof value === 'number') {
    return value;
  }

  const normalized = String(value || '')
    .replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 65248))
    .replace(/[^\d.-]/g, '');

  const num = Number(normalized);
  return Number.isNaN(num) ? 0 : num;
}

function getSheetHeadersForAbsence_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return [];
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
}

function getSheetBodyValuesForAbsence_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn === 0) {
    return [];
  }

  return sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
}

function findColumnIndexForAbsence_(headers, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const idx = headers.indexOf(candidates[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

function validateColumnsForAbsence_(sheetName, colMap, requiredKeys) {
  requiredKeys.forEach(key => {
    if (colMap[key] === -1) {
      throw new Error(sheetName + ' シートに必要な列がありません: ' + key);
    }
  });
}