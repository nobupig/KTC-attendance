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
  const isExperimentGroup = rosterSource === 'studentGroups' && isExperimentGroupTargetClass_(targetClassId);
  const relatedClassIds = isExperimentGroup
    ? getExperimentRelatedClassIdsByClassId_(targetClassId)
    : [targetClassId];

  logPerf_(
    'getTeacherSessionDetailLight getRosterSourceByClassId_',
    rosterStartedAt,
    'classId=' + targetClassId + ' rosterSource=' + rosterSource + ' relatedClassIds=' + relatedClassIds.join(',')
  );

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
  const targetStudentIds = getStudentIdsForAttendanceFilter_(students);

  const attendanceMap = isExperimentGroup
    ? getAttendanceMapForClassIds_(relatedClassIds, date, period, targetStudentIds)
    : getAttendanceMap(targetClassId, date, period);

  logPerf_(
    'getTeacherSessionDetailLight getAttendanceMap',
    attendanceStartedAt,
    'attendanceCount=' + Object.keys(attendanceMap || {}).length + ' experiment=' + isExperimentGroup
  );

  const savedInfoStartedAt = perfNow_();
  const lastSavedInfo = isExperimentGroup
    ? getLatestAttendanceSessionInfoForClassIds_(relatedClassIds, date, period, ['normal', 'past-edit'])
    : getLatestAttendanceSessionInfo_(targetClassId, date, period, ['normal', 'past-edit']);

  logPerf_('getTeacherSessionDetailLight getLatestAttendanceSessionInfo_', savedInfoStartedAt, 'hasSaved=' + (!!lastSavedInfo));

  const groupsStartedAt = perfNow_();
  const availableGroups = isExperimentGroup
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
    availableGroups: availableGroups,
    relatedClassIds: relatedClassIds
  };

  logPerf_(
    'getTeacherSessionDetailLight total',
    totalStartedAt,
    'classId=' + targetClassId + ' period=' + period + ' students=' + students.length
  );

  return result;
}

function normalizePreviousSessionCopyValue_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function getPreviousSessionCopyNumericValue_(value) {
  const normalized = normalizePreviousSessionCopyValue_(value);
  if (!normalized) return null;

  const numericValue = Number(normalized);
  return Number.isFinite(numericValue) ? numericValue : null;
}

function getPreviousSessionCopyTeacherSignature_(cls) {
  const teachers = Array.isArray(cls && cls.teachers) && cls.teachers.length > 0
    ? cls.teachers
    : [{
      teacherEmail: cls && cls.teacherEmail,
      teacherName: cls && cls.teacherName,
      roleType: cls && cls.roleType
    }];

  return teachers.map(function(teacher) {
    return JSON.stringify([
      normalizePreviousSessionCopyValue_(teacher && (teacher.teacherEmail || teacher.email)),
      normalizePreviousSessionCopyValue_(teacher && (teacher.teacherName || teacher.name)),
      normalizePreviousSessionCopyValue_(teacher && teacher.roleType)
    ]);
  }).filter(function(signature) {
    return signature !== '["","",""]';
  }).sort().join('|');
}

function isPreviousSessionCopyUnsupportedClass_(cls) {
  if (!cls) return true;

  const classId = normalizePreviousSessionCopyValue_(cls.classId);
  const subjectId = normalizePreviousSessionCopyValue_(cls.subjectId);

  return (
    classId.indexOf('\u5de5\u5b66\u5b9f\u9a13\u5b9f\u7fd21') !== -1 ||
    classId.indexOf('\u5de5\u5b66\u5b9f\u9a13\u5b9f\u7fd22') !== -1 ||
    subjectId === 'G1_G_\u5de5\u5b66\u5b9f\u9a13\u5b9f\u7fd21_FY' ||
    subjectId === 'G2_G_\u5de5\u5b66\u5b9f\u9a13\u5b9f\u7fd22_FY'
  );
}

function canCopyPreviousTeacherSession_(previousClass, nextClass) {
  if (!previousClass || !nextClass) return false;

  if (
    isPreviousSessionCopyUnsupportedClass_(previousClass) ||
    isPreviousSessionCopyUnsupportedClass_(nextClass)
  ) {
    return false;
  }

  const previousDate = normalizePreviousSessionCopyValue_(previousClass.date);
  const nextDate = normalizePreviousSessionCopyValue_(nextClass.date);
  const previousClassId = normalizePreviousSessionCopyValue_(previousClass.classId);
  const nextClassId = normalizePreviousSessionCopyValue_(nextClass.classId);
  const previousSubjectId = normalizePreviousSessionCopyValue_(previousClass.subjectId);
  const nextSubjectId = normalizePreviousSessionCopyValue_(nextClass.subjectId);

  if (!previousDate || previousDate !== nextDate) return false;
  if (!previousClassId || previousClassId !== nextClassId) return false;
  if (!previousSubjectId || previousSubjectId !== nextSubjectId) return false;

  const optionalIdentityFields = ['grade', 'unit', 'curriculumUnit'];
  const optionalIdentityMatches = optionalIdentityFields.every(function(fieldName) {
    return normalizePreviousSessionCopyValue_(previousClass[fieldName]) ===
      normalizePreviousSessionCopyValue_(nextClass[fieldName]);
  });
  if (!optionalIdentityMatches) return false;

  if (
    getPreviousSessionCopyTeacherSignature_(previousClass) !==
    getPreviousSessionCopyTeacherSignature_(nextClass)
  ) {
    return false;
  }

  const previousPeriod = getPreviousSessionCopyNumericValue_(previousClass.period);
  const nextPeriod = getPreviousSessionCopyNumericValue_(nextClass.period);
  if (previousPeriod === null || nextPeriod !== previousPeriod + 1) return false;

  const previousSessionNumber = getPreviousSessionCopyNumericValue_(previousClass.sessionNumber);
  const nextSessionNumber = getPreviousSessionCopyNumericValue_(nextClass.sessionNumber);
  if (previousSessionNumber === null || nextSessionNumber !== previousSessionNumber + 1) return false;

  return true;
}

function buildPreviousSessionCopyStudentIds_(students) {
  const result = [];
  const seen = {};

  (Array.isArray(students) ? students : []).forEach(function(student) {
    const studentId = normalizePreviousSessionCopyValue_(student && student.studentId);
    if (!studentId || seen[studentId]) return;

    seen[studentId] = true;
    result.push(studentId);
  });

  result.sort();
  return result;
}

function previousSessionCopyFail_(code, message) {
  return {
    ok: false,
    code: String(code || 'COPY_PREVIEW_FAILED'),
    message: String(message || 'Previous-session attendance copy is unavailable.')
  };
}

function getPreviousTeacherSessionAttendanceCopyPreview(destination) {
  const input = destination && typeof destination === 'object' ? destination : {};

  const targetClassId = normalizePreviousSessionCopyValue_(input.classId);
  const targetDate = formatDateToYmd(input.date);
  const targetPeriod = normalizePreviousSessionCopyValue_(input.period);
  const targetSessionNumber = normalizePreviousSessionCopyValue_(input.sessionNumber);

  if (!targetClassId || !targetDate || !targetPeriod || !targetSessionNumber) {
    return previousSessionCopyFail_(
      'INVALID_DESTINATION',
      'Destination session information is incomplete.'
    );
  }

  const teacherClasses = getClassesForCurrentUserByDate(targetDate);
  const normalizedTeacherClasses = Array.isArray(teacherClasses) ? teacherClasses : [];

  const destinationMatches = normalizedTeacherClasses.filter(function(cls) {
    return (
      normalizePreviousSessionCopyValue_(cls.classId) === targetClassId &&
      formatDateToYmd(cls.date) === targetDate &&
      normalizePreviousSessionCopyValue_(cls.period) === targetPeriod &&
      normalizePreviousSessionCopyValue_(cls.sessionNumber) === targetSessionNumber
    );
  });

  if (destinationMatches.length !== 1) {
    return previousSessionCopyFail_(
      'DESTINATION_NOT_FOUND',
      'Destination session could not be verified as a current teacher assignment.'
    );
  }

  const destinationClass = destinationMatches[0];

  if (isPreviousSessionCopyUnsupportedClass_(destinationClass)) {
    return previousSessionCopyFail_(
      'UNSUPPORTED_CLASS',
      'Previous-session copy is not supported for this class.'
    );
  }

  const sourceMatches = normalizedTeacherClasses.filter(function(cls) {
    return canCopyPreviousTeacherSession_(cls, destinationClass);
  });

  if (sourceMatches.length !== 1) {
    return previousSessionCopyFail_(
      'PREVIOUS_SESSION_NOT_FOUND',
      'An eligible immediate previous session could not be verified.'
    );
  }

  const sourceClass = sourceMatches[0];

  const sourceDetail = getTeacherSessionDetailLight(
    sourceClass.classId,
    sourceClass.date,
    sourceClass.period,
    ''
  );

  if (!sourceDetail || !sourceDetail.hasSavedSession) {
    return previousSessionCopyFail_(
      'SOURCE_NOT_SAVED',
      'The immediate previous session is not saved.'
    );
  }

  const destinationDetail = getTeacherSessionDetailLight(
    destinationClass.classId,
    destinationClass.date,
    destinationClass.period,
    ''
  );

  if (destinationDetail && destinationDetail.hasSavedSession) {
    return previousSessionCopyFail_(
      'DESTINATION_ALREADY_SAVED',
      'The destination session is already saved.'
    );
  }

  const sourceStudentIds = buildPreviousSessionCopyStudentIds_(
    sourceDetail && sourceDetail.students
  );
  const destinationStudentIds = buildPreviousSessionCopyStudentIds_(
    destinationDetail && destinationDetail.students
  );

  if (
    sourceStudentIds.length === 0 ||
    sourceStudentIds.length !== destinationStudentIds.length ||
    sourceStudentIds.some(function(studentId, index) {
      return studentId !== destinationStudentIds[index];
    })
  ) {
    return previousSessionCopyFail_(
      'ROSTER_MISMATCH',
      'Source and destination rosters do not match.'
    );
  }

  const sourceStudentIdSet = {};
  sourceStudentIds.forEach(function(studentId) {
    sourceStudentIdSet[studentId] = true;
  });

  const rawAttendanceMap = sourceDetail && sourceDetail.attendanceMap
    ? sourceDetail.attendanceMap
    : {};
  const safeAttendanceMap = {};
  const allowedCodes = {
    A: true,
    L: true,
    O: true
  };

  const attendanceKeys = Object.keys(rawAttendanceMap);
  for (let i = 0; i < attendanceKeys.length; i++) {
    const rawStudentId = attendanceKeys[i];
    const studentId = normalizePreviousSessionCopyValue_(rawStudentId);
    const statusCode = normalizePreviousSessionCopyValue_(rawAttendanceMap[rawStudentId]);

    if (!studentId || !sourceStudentIdSet[studentId] || !allowedCodes[statusCode]) {
      return previousSessionCopyFail_(
        'INVALID_SOURCE_ATTENDANCE',
        'Saved source attendance contains an unsupported value.'
      );
    }

    safeAttendanceMap[studentId] = statusCode;
  }

  return {
    ok: true,
    source: {
      classId: normalizePreviousSessionCopyValue_(sourceClass.classId),
      date: formatDateToYmd(sourceClass.date),
      period: normalizePreviousSessionCopyValue_(sourceClass.period),
      sessionNumber: normalizePreviousSessionCopyValue_(sourceClass.sessionNumber)
    },
    destination: {
      classId: normalizePreviousSessionCopyValue_(destinationClass.classId),
      date: formatDateToYmd(destinationClass.date),
      period: normalizePreviousSessionCopyValue_(destinationClass.period),
      sessionNumber: normalizePreviousSessionCopyValue_(destinationClass.sessionNumber)
    },
    studentIds: sourceStudentIds,
    attendanceMap: safeAttendanceMap
  };
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

function getExperimentRelatedClassIdsByClassId_(classId) {
  const targetClassId = normalizeString_(classId);
  if (!targetClassId) return [];

  const classInfo = getClassRecordById_(targetClassId);
  if (!classInfo) return [targetClassId];

  const subjectId = normalizeString_(classInfo.subjectId);
  if (!subjectId) return [targetClassId];

  const isExperimentSubject =
    subjectId === 'G1_G_工学実験実習1_FY' ||
    subjectId === 'G2_G_工学実験実習2_FY';

  if (!isExperimentSubject) {
    return [targetClassId];
  }

  const classesData = getSheetDataCached_('MASTER', CONFIG.SHEETS.CLASSES, 300);
  const headers = classesData.headers || [];
  const rows = classesData.rows || [];

  const col = {
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    subjectId: findColumnIndex_(headers, ['subjectId', 'SubjectID'])
  };

  if (col.classId === -1 || col.subjectId === -1) {
    return [targetClassId];
  }

  const seen = {};
  const result = [];

  rows.forEach(function(row) {
    const rowClassId = normalizeString_(row[col.classId]);
    const rowSubjectId = normalizeString_(row[col.subjectId]);

    if (!rowClassId || rowSubjectId !== subjectId) return;
    if (seen[rowClassId]) return;

    seen[rowClassId] = true;
    result.push(rowClassId);
  });

  if (!seen[targetClassId]) {
    result.push(targetClassId);
  }

  return result.sort(function(a, b) {
    return String(a).localeCompare(String(b), 'ja');
  });
}

function getStudentIdsForAttendanceFilter_(students) {
  return (Array.isArray(students) ? students : [])
    .map(function(student) {
      return normalizeString_(student && student.studentId);
    })
    .filter(Boolean);
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