function getStudentsByClassId(classId) {
  const targetClassId = normalizeString_(classId);
  if (!targetClassId) {
    return [];
  }

  const cacheKey = 'studentsByClassId__' + targetClassId;
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const classInfo = getClassRecordById_(targetClassId);
  if (!classInfo) {
    return [];
  }

  const targetGrade = normalizeString_(classInfo.grade);
  const targetUnits = expandStudentUnitsForClassUnit_(classInfo.unit);

  const studentsData = getSheetDataCached_('MASTER', CONFIG.SHEETS.STUDENTS, 300);
  const headers = studentsData.headers;
  const rows = studentsData.rows;

  const col = {
    studentId: findColumnIndex_(headers, ['studentId', 'StudentID']),
    grade: findColumnIndex_(headers, ['grade', '学年']),
    unit: findColumnIndex_(headers, ['unit', '組・コース', '対象区分']),
    attendanceNumber: findColumnIndex_(headers, ['attendanceNumber', '出席番号']),
    name: findColumnIndex_(headers, ['name', '氏名']),
    status: findColumnIndex_(headers, ['status', '在籍状態'])
  };

  validateRequiredColumnsForStudentService_('students', col, [
    'studentId',
    'grade',
    'unit',
    'attendanceNumber',
    'name'
  ]);

  const result = rows
    .filter(function(row) {
      const rowGrade = normalizeString_(row[col.grade]);
      const rowUnit = normalizeString_(row[col.unit]);
      const rowStatus = col.status !== -1 ? normalizeString_(row[col.status]).toLowerCase() : 'active';

      if (rowGrade !== targetGrade) return false;
      if (!targetUnits.includes(rowUnit)) return false;
      if (!isActiveStudentStatus_(rowStatus)) return false;

      return true;
    })
    .map(function(row) {
      return {
        studentId: normalizeString_(row[col.studentId]),
        grade: normalizeString_(row[col.grade]),
        unit: normalizeString_(row[col.unit]),
        attendanceNumber: row[col.attendanceNumber],
        name: normalizeString_(row[col.name]),
        status: col.status !== -1 ? normalizeString_(row[col.status]) : 'active'
      };
    })
    .sort(compareStudentsByAttendanceNumber_);

  putScriptCacheJson_(cacheKey, result, 300);
  return result;
}

function getTeacherSessionDetailLight(classId, date, period, group) {
  const totalStartedAt = perfNow_();

  const targetClassId = normalizeString_(classId);
  const targetGroup = normalizeString_(group);

  const rosterStartedAt = perfNow_();
  let students = [];
  const rosterSource = getRosterSourceByClassId_(targetClassId);
  logPerf_('getTeacherSessionDetailLight getRosterSourceByClassId_', rosterStartedAt, 'classId=' + targetClassId + ' rosterSource=' + rosterSource);

  const studentsStartedAt = perfNow_();
  if (rosterSource === 'studentGroups') {
    if (isExperimentGroupTargetClass_(targetClassId)) {
      students = getStudentsByClassIdAndGroup(targetClassId, targetGroup);
    } else {
      students = getStudentsByStudentGroupsClassId_(targetClassId);
    }
  } else {
    students = getStudentsByClassId(targetClassId);
  }
  logPerf_('getTeacherSessionDetailLight load students', studentsStartedAt, 'count=' + students.length + ' group=' + targetGroup);

  const attendanceStartedAt = perfNow_();
  const attendanceMap = getAttendanceMap(targetClassId, date, period);
  logPerf_('getTeacherSessionDetailLight getAttendanceMap', attendanceStartedAt, 'attendanceCount=' + Object.keys(attendanceMap || {}).length);

  const savedInfoStartedAt = perfNow_();
  const lastSavedInfo = getLatestAttendanceSessionInfo_(
    targetClassId,
    date,
    period,
    ['normal', 'past-edit']
  );
  logPerf_('getTeacherSessionDetailLight getLatestAttendanceSessionInfo_', savedInfoStartedAt, 'hasSaved=' + (!!lastSavedInfo));

  const groupsStartedAt = perfNow_();
  const availableGroups = (rosterSource === 'studentGroups' && isExperimentGroupTargetClass_(targetClassId))
    ? getGroupsByClassId(targetClassId)
    : [];
  logPerf_('getTeacherSessionDetailLight get availableGroups', groupsStartedAt, 'count=' + availableGroups.length);

  const safeLastSavedInfo = lastSavedInfo ? {
    teacherEmail: lastSavedInfo.teacherEmail || '',
    savedAtText: lastSavedInfo.savedAtText || '',
    actionType: lastSavedInfo.actionType || '',
    targetSessionKey: lastSavedInfo.targetSessionKey || '',
    savedModeLabel: lastSavedInfo.savedModeLabel || '',
    savedByCurrentUser: !!lastSavedInfo.savedByCurrentUser
  } : null;

  const result = {
    students: students,
    attendanceMap: attendanceMap,
    hasSavedSession: !!safeLastSavedInfo,
    lastSavedInfo: safeLastSavedInfo,
    group: targetGroup,
    availableGroups: availableGroups
  };

  logPerf_(
    'getTeacherSessionDetailLight total',
    totalStartedAt,
    'classId=' + targetClassId + ' period=' + period + ' students=' + students.length
  );

  return result;
}

function getTeacherSessionDetail(classId, date, period, group) {
  const light = getTeacherSessionDetailLight(classId, date, period, group);
  light.riskMap = getStudentRiskMapForClass(classId, light.students);
  return light;
}

function getStudentRiskMapForTeacherSession(classId, studentIds) {
  const targetClassId = normalizeString_(classId);
  const ids = Array.isArray(studentIds) ? studentIds.map(normalizeString_).filter(Boolean) : [];
  const cacheKey = 'riskMapForTeacherSession__' + targetClassId + '__' + buildRiskCacheSignature_(ids);
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const classInfo = getClassRecordById_(targetClassId);
  if (!classInfo) {
    return {};
  }

  let students = [];

  if (ids.length > 0) {
    const allowMap = {};
    ids.forEach(function(id) {
      allowMap[id] = true;
    });

    if (isExperimentGroupTargetClass_(targetClassId)) {
      const studentData = getSheetDataCached_('MASTER', CONFIG.SHEETS.STUDENTS, 300);
      const studentHeaders = studentData.headers;
      const studentRows = studentData.rows;

      const studentCol = {
        studentId: findColumnIndex_(studentHeaders, ['studentId', 'StudentID']),
        grade: findColumnIndex_(studentHeaders, ['grade', '学年']),
        unit: findColumnIndex_(studentHeaders, ['unit', '組・コース', '対象区分']),
        attendanceNumber: findColumnIndex_(studentHeaders, ['attendanceNumber', '出席番号']),
        name: findColumnIndex_(studentHeaders, ['name', '氏名'])
      };

      ['studentId', 'grade', 'unit', 'attendanceNumber', 'name'].forEach(function(key) {
        if (studentCol[key] === -1) {
          throw new Error('students シートに必要な列がありません: ' + key);
        }
      });

      const targetGrade = normalizeString_(classInfo.grade);

      students = studentRows
        .map(function(row) {
          return {
            studentId: normalizeString_(row[studentCol.studentId]),
            grade: normalizeString_(row[studentCol.grade]),
            unit: normalizeString_(row[studentCol.unit]),
            attendanceNumber: normalizeString_(row[studentCol.attendanceNumber]),
            name: normalizeString_(row[studentCol.name])
          };
        })
        .filter(function(student) {
          return !!student.studentId &&
                 student.grade === targetGrade &&
                 !!allowMap[student.studentId];
        });
    } else {
      const classStudents = getStudentsByClassId(targetClassId);
      students = classStudents.filter(function(student) {
        return !!allowMap[normalizeString_(student.studentId)];
      });
    }
  }

  const result = getStudentRiskMapForClass(targetClassId, students);
  putScriptCacheJson_(cacheKey, result, 120);
  return result;
}

function getStudentRiskMapForClass(classId, students) {
  const targetClassId = normalizeString_(classId);
  const riskMap = {};

  if (!targetClassId || !students || !students.length) {
    return riskMap;
  }

  const studentIds = students
    .map(function(student) { return normalizeString_(student.studentId); })
    .filter(Boolean);
  const cacheKey = 'riskMapForClass__' + targetClassId + '__' + buildRiskCacheSignature_(studentIds);
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const classInfo = getClassRecordById_(targetClassId);
  if (!classInfo) {
    return riskMap;
  }

  const targetSubjectId = normalizeString_(classInfo.subjectId);
  const subjectName = normalizeString_(classInfo.subjectName);
  const defaultLimit = Number(classInfo.allowedAbsences || 0);

  if (!targetSubjectId) {
    return riskMap;
  }

  const bundle = buildAbsenceCalculationBundle_('all');

  students.forEach(function(student) {
    const studentId = normalizeString_(student.studentId);
    if (!studentId) return;

    try {
      const result = calculateStudentAbsenceRiskFromBundle_(studentId, 'all', bundle);
      const subjects = Array.isArray(result.subjects) ? result.subjects : [];

      const targetSubject = subjects.find(function(subject) {
        return normalizeString_(subject.subjectId) === targetSubjectId;
      });

      if (targetSubject) {
        riskMap[studentId] = {
          subjectId: normalizeString_(targetSubject.subjectId),
          subjectName: normalizeString_(targetSubject.subjectName) || subjectName,
          normalAbsence: Number(targetSubject.normalAbsence || 0),
          officialAbsence: Number(targetSubject.officialAbsence || 0),
          late: Number(targetSubject.late || 0),
          early: Number(targetSubject.early || 0),
          limit: Number(targetSubject.limit || defaultLimit || 0),
          remaining: Number(targetSubject.remaining || 0),
          riskLevel: normalizeString_(targetSubject.riskLevel) || 'normal',
          riskLabel: normalizeString_(targetSubject.riskLabel) || '正常'
        };
      } else {
        riskMap[studentId] = buildDefaultRiskRecord_(targetSubjectId, subjectName, defaultLimit);
      }
    } catch (e) {
      riskMap[studentId] = buildDefaultRiskRecord_(targetSubjectId, subjectName, defaultLimit);
    }
  });

  putScriptCacheJson_(cacheKey, riskMap, 120);
  return riskMap;
}

function expandStudentUnitsForClassUnit_(classUnit) {
  const unit = normalizeString_(classUnit).toUpperCase();

  // 出席管理上 CA は、学生所属上の C / A を両方対象にする
  if (unit === 'CA') {
    return ['C', 'A', 'CA'];
  }

  return [normalizeString_(classUnit)];
}

function isActiveStudentStatus_(status) {
  const s = normalizeString_(status).toLowerCase();

  // status 列が空なら現役扱い
  if (!s) return true;

  return ['active', '在籍', '有効'].includes(s);
}

function compareStudentsByUnitAndAttendanceNumber_(a, b) {
  const unitA = normalizeString_(a && a.unit);
  const unitB = normalizeString_(b && b.unit);

  const numA = parseInt(String(unitA || '').replace(/[^\d]/g, ''), 10);
  const numB = parseInt(String(unitB || '').replace(/[^\d]/g, ''), 10);

  const hasNumA = !Number.isNaN(numA);
  const hasNumB = !Number.isNaN(numB);

  if (hasNumA && hasNumB && numA !== numB) {
    return numA - numB;
  }

  if (unitA !== unitB) {
    return unitA.localeCompare(unitB, 'ja');
  }

  return compareStudentsByAttendanceNumber_(a, b);
}


function compareStudentsByAttendanceNumber_(a, b) {
  const aNum = Number(a.attendanceNumber);
  const bNum = Number(b.attendanceNumber);

  const aIsNum = !Number.isNaN(aNum);
  const bIsNum = !Number.isNaN(bNum);

  if (aIsNum && bIsNum) {
    return aNum - bNum;
  }

  return String(a.attendanceNumber).localeCompare(String(b.attendanceNumber), 'ja');
}

function buildRiskCacheSignature_(ids) {
  const normalizedIds = (Array.isArray(ids) ? ids : [])
    .map(function(id) { return normalizeString_(id); })
    .filter(Boolean)
    .sort();

  if (normalizedIds.length === 0) {
    return 'empty';
  }

  const raw = normalizedIds.join('|');
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    raw,
    Utilities.Charset.UTF_8
  );

  let hex = '';
  for (var i = 0; i < digest.length; i++) {
    var value = digest[i];
    if (value < 0) {
      value += 256;
    }
    const piece = value.toString(16);
    hex += piece.length === 1 ? '0' + piece : piece;
  }

  return normalizedIds.length + '__' + hex;
}

function buildDefaultRiskRecord_(subjectId, subjectName, defaultLimit) {
  return {
    subjectId: normalizeString_(subjectId),
    subjectName: normalizeString_(subjectName),
    normalAbsence: 0,
    officialAbsence: 0,
    late: 0,
    early: 0,
    limit: Number(defaultLimit || 0),
    remaining: Number(defaultLimit || 0),
    riskLevel: 'normal',
    riskLabel: '正常'
  };
}

function validateRequiredColumnsForStudentService_(sheetName, colMap, requiredKeys) {
  requiredKeys.forEach(function(key) {
    if (colMap[key] === -1) {
      throw new Error(sheetName + ' シートに必要な列がありません: ' + key);
    }
  });
}

function testGetStudentsByClassId() {
  const result = getStudentsByClassId('G1_1_英語1A_FA');
  Logger.log(JSON.stringify(result, null, 2));
}

function isExperimentGroupTargetClass_(classId) {
  const targetClassId = normalizeString_(classId);
  if (!targetClassId) return false;

  const classInfo = getClassRecordById_(targetClassId);
  if (!classInfo) return false;

  const subjectId = normalizeString_(classInfo.subjectId);
  return subjectId === 'G1_G_工学実験実習1_FY' || subjectId === 'G2_G_工学実験実習2_FY';
}

function getStudentGroupKeyByClassId_(classId) {
  const targetClassId = normalizeString_(classId);
  if (!targetClassId) return '';

  const classInfo = getClassRecordById_(targetClassId);
  if (!classInfo) return targetClassId;

  const subjectId = normalizeString_(classInfo.subjectId);

  if (subjectId === 'G1_G_工学実験実習1_FY' || subjectId === 'G2_G_工学実験実習2_FY') {
    return subjectId;
  }

  return targetClassId;
}

function getGroupsByClassId(classId) {
  const targetKey = getStudentGroupKeyByClassId_(classId);
  if (!targetKey) return [];

  if (!isExperimentGroupTargetClass_(classId)) {
    return [];
  }

  const data = getSheetDataCached_('OPERATION', CONFIG.SHEETS.STUDENT_GROUPS, 300);
  const headers = data.headers;
  const rows = data.rows;

  const col = {
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    group: findColumnIndex_(headers, ['group', '班'])
  };

  if (col.classId === -1 || col.group === -1) {
    throw new Error('studentGroups シートに classId / group 列がありません');
  }

  const groups = [];
  rows.forEach(function(row) {
    const rowClassId = normalizeString_(row[col.classId]);
    const rowGroup = normalizeString_(row[col.group]);
    if (rowClassId !== targetKey || !rowGroup) return;
    if (!groups.includes(rowGroup)) groups.push(rowGroup);
  });

  return groups.sort(compareGroupLabels_);
}

function compareGroupLabels_(a, b) {
  const aText = String(a || '');
  const bText = String(b || '');

  const aNum = parseInt(aText.replace(/[^\d]/g, ''), 10);
  const bNum = parseInt(bText.replace(/[^\d]/g, ''), 10);

  const aIsNum = !Number.isNaN(aNum);
  const bIsNum = !Number.isNaN(bNum);

  if (aIsNum && bIsNum && aNum !== bNum) {
    return aNum - bNum;
  }

  return aText.localeCompare(bText, 'ja');
}

function getStudentsByClassIdAndGroup(classId, group) {
  const targetClassId = normalizeString_(classId);
  const targetGroup = normalizeString_(group);
  const targetKey = getStudentGroupKeyByClassId_(targetClassId);

  if (!targetClassId) return [];

  if (!isExperimentGroupTargetClass_(targetClassId)) {
    return getStudentsByClassId(targetClassId);
  }

  if (!targetGroup) return [];

  const classInfo = getClassRecordById_(targetClassId);
  if (!classInfo) return [];

  const targetGrade = normalizeString_(classInfo.grade);

  // 学生マスタから「対象学年の全学生」を土台にする
  const studentData = getSheetDataCached_('MASTER', CONFIG.SHEETS.STUDENTS, 300);
  const studentHeaders = studentData.headers;
  const studentRows = studentData.rows;

  const studentCol = {
    studentId: findColumnIndex_(studentHeaders, ['studentId', 'StudentID']),
    grade: findColumnIndex_(studentHeaders, ['grade', '学年']),
    unit: findColumnIndex_(studentHeaders, ['unit', '組・コース', '対象区分']),
    attendanceNumber: findColumnIndex_(studentHeaders, ['attendanceNumber', '出席番号']),
    name: findColumnIndex_(studentHeaders, ['name', '氏名'])
  };

  ['studentId', 'grade', 'unit', 'attendanceNumber', 'name'].forEach(function(key) {
    if (studentCol[key] === -1) {
      throw new Error('students シートに必要な列がありません: ' + key);
    }
  });

  const activeStudentMap = {};
  studentRows.forEach(function(row) {
    const studentId = normalizeString_(row[studentCol.studentId]);
    const studentGrade = normalizeString_(row[studentCol.grade]);

    if (!studentId) return;
    if (studentGrade !== targetGrade) return;

    activeStudentMap[studentId] = {
      studentId: studentId,
      grade: normalizeString_(row[studentCol.grade]),
      unit: normalizeString_(row[studentCol.unit]),
      attendanceNumber: normalizeString_(row[studentCol.attendanceNumber]),
      name: normalizeString_(row[studentCol.name])
    };
  });

  const data = getSheetDataCached_('OPERATION', CONFIG.SHEETS.STUDENT_GROUPS, 300);
  const headers = data.headers;
  const rows = data.rows;

  const col = {
    studentId: findColumnIndex_(headers, ['studentId', 'StudentID']),
    grade: findColumnIndex_(headers, ['grade', '学年']),
    unit: findColumnIndex_(headers, ['unit', '組・コース', '対象区分']),
    attendanceNumber: findColumnIndex_(headers, ['attendanceNumber', '出席番号']),
    name: findColumnIndex_(headers, ['name', '氏名']),
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    group: findColumnIndex_(headers, ['group', '班'])
  };

  ['studentId', 'grade', 'unit', 'attendanceNumber', 'name', 'classId', 'group'].forEach(function(key) {
    if (col[key] === -1) {
      throw new Error('studentGroups シートに必要な列がありません: ' + key);
    }
  });

  return rows
    .filter(function(row) {
      return normalizeString_(row[col.classId]) === targetKey &&
             normalizeString_(row[col.group]) === targetGroup;
    })
    .map(function(row) {
      const studentId = normalizeString_(row[col.studentId]);
      return activeStudentMap[studentId] || null;
    })
    .filter(Boolean)
    .sort(compareStudentsByUnitAndAttendanceNumber_);
}

function hasStudentGroupRowsForClassId_(classId) {
  const targetClassId = normalizeString_(classId);
  if (!targetClassId) return false;

  const data = getSheetDataCached_('OPERATION', CONFIG.SHEETS.STUDENT_GROUPS, 300);
  const headers = data.headers;
  const rows = data.rows;

  const col = {
    classId: findColumnIndex_(headers, ['classId', 'ClassID'])
  };

  if (col.classId === -1) {
    throw new Error('studentGroups シートに classId 列がありません');
  }

  return rows.some(function(row) {
    return normalizeString_(row[col.classId]) === targetClassId;
  });
}

function getStudentsByStudentGroupsClassId_(classId) {
  const targetClassId = normalizeString_(classId);
  if (!targetClassId) return [];

  const baseStudents = getStudentsByClassId(targetClassId);
  const activeStudentMap = {};
  baseStudents.forEach(function(student) {
    activeStudentMap[normalizeString_(student.studentId)] = student;
  });

  const data = getSheetDataCached_('OPERATION', CONFIG.SHEETS.STUDENT_GROUPS, 300);
  const headers = data.headers;
  const rows = data.rows;

  const col = {
    studentId: findColumnIndex_(headers, ['studentId', 'StudentID']),
    classId: findColumnIndex_(headers, ['classId', 'ClassID'])
  };

  if (col.studentId === -1 || col.classId === -1) {
    throw new Error('studentGroups シートに studentId / classId 列がありません');
  }

  return rows
    .filter(function(row) {
      return normalizeString_(row[col.classId]) === targetClassId;
    })
    .map(function(row) {
      const studentId = normalizeString_(row[col.studentId]);
      return activeStudentMap[studentId] || null;
    })
    .filter(Boolean)
    .sort(compareStudentsByAttendanceNumber_);
}

function getRosterSourceByClassId_(classId) {
  const targetClassId = normalizeString_(classId);
  if (!targetClassId) return 'students';

  const classesData = getSheetDataCached_('MASTER', CONFIG.SHEETS.CLASSES, 300);
  const headers = classesData.headers || [];
  const rows = classesData.rows || [];

  const col = {
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    rosterSource: findColumnIndex_(headers, ['rosterSource', '名簿取得元'])
  };

  if (col.classId === -1) {
    throw new Error('classes シートに classId 列がありません');
  }

  const row = rows.find(function(r) {
    return normalizeString_(r[col.classId]) === targetClassId;
  });

  if (!row) return 'students';
  if (col.rosterSource === -1) return 'students';

  const value = normalizeString_(row[col.rosterSource]).toLowerCase();
  return value === 'studentgroups' ? 'studentGroups' : 'students';
}