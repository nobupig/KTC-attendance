function buildTeacherMasterCacheKey_() {
  return 'teacherMasterBundle__all';
}

function buildTeacherRecordFromRow_(row) {
  if (!row) return null;

  return {
    teacherId: normalizeString_(row.teacherId),
    name: normalizeString_(row.name),
    email: normalizeString_(row.email).toLowerCase(),
    roles: parseRoles_(row.roles)
  };
}

function getTeacherMasterBundle_() {
  const cacheKey = buildTeacherMasterCacheKey_();
  const cached = getScriptCacheJson_(cacheKey);
  if (cached !== null) {
    return cached;
  }

  const rows = getTeachersSheetObjectsCached_(300);
  const byEmail = {};
  const byId = {};
  const byName = {};

  rows.forEach(function(row) {
    const record = buildTeacherRecordFromRow_(row);
    if (!record) return;

    if (record.email && !byEmail[record.email]) {
      byEmail[record.email] = record;
    }
    if (record.teacherId && !byId[record.teacherId]) {
      byId[record.teacherId] = record;
    }
    if (record.name && !byName[record.name]) {
      byName[record.name] = record;
    }
  });

  const bundle = {
    byEmail: byEmail,
    byId: byId,
    byName: byName
  };

  putScriptCacheJson_(cacheKey, bundle, 300);
  return bundle;
}

function getTeacherRecordByEmail_(email) {
  const targetEmail = normalizeString_(email).toLowerCase();
  if (!targetEmail) return null;

  const bundle = getTeacherMasterBundle_();
  return bundle.byEmail[targetEmail] || null;
}

function getTeacherRecordById_(teacherId) {
  const targetTeacherId = normalizeString_(teacherId);
  if (!targetTeacherId) return null;

  const bundle = getTeacherMasterBundle_();
  return bundle.byId[targetTeacherId] || null;
}

function getTeacherRecordByName_(name) {
  const targetName = normalizeString_(name);
  if (!targetName) return null;

  const bundle = getTeacherMasterBundle_();
  return bundle.byName[targetName] || null;
}

function getClassTeacherTeamRows_() {
  const cacheKey = 'classTeacherTeamRows__all';
  const cached = getScriptCacheJson_(cacheKey);
  if (cached !== null) {
    return cached;
  }

  const ss = openOperationSpreadsheet_();
  let rows = [];
  try {
    rows = readSheetAsObjects_(ss, CONFIG.SHEETS.CLASS_TEACHER_TEAMS);
  } catch (e) {
    // シート未作成時の互換
    rows = [];
  }

  putScriptCacheJson_(cacheKey, rows, 300);
  return rows;
}

function buildTeacherTeamMember_(teacherId, teacherName, roleType) {
  let record = null;

  if (teacherId) {
    record = getTeacherRecordById_(teacherId);
  } else if (teacherName) {
    record = getTeacherRecordByName_(teacherName);
  }

  return {
    teacherId: record ? record.teacherId : normalizeString_(teacherId),
    teacherName: record ? record.name : normalizeString_(teacherName),
    teacherEmail: record ? record.email : '',
    roles: record ? record.roles : [],
    roleType: normalizeString_(roleType || 'support').toLowerCase() || 'support'
  };
}

/**
 * Resolves teacher members from one teacher-master snapshot for one teaching
 * assignment index build. This prevents one CacheService read/JSON parse per
 * timetable or team row while preserving buildTeacherTeamMember_ semantics.
 */
function createTeacherTeamMemberResolver_() {
  const startedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const bundle = getTeacherMasterBundle_() || {};
  const byId = bundle.byId || {};
  const byName = bundle.byName || {};
  const teacherCount = Object.keys(byId).length;

  if (typeof logPerf_ === 'function') {
    logPerf_('teaching assignment teacher bundle', startedAt, 'teacherCount=' + teacherCount);
  }

  const resolver = function(teacherId, teacherName, roleType) {
    let record = null;

    // Keep the original branch order: a supplied teacherId is authoritative;
    // only an absent teacherId permits teacherName fallback.
    if (teacherId) {
      record = byId[normalizeString_(teacherId)] || null;
    } else if (teacherName) {
      record = byName[normalizeString_(teacherName)] || null;
    }

    return {
      teacherId: record ? record.teacherId : normalizeString_(teacherId),
      teacherName: record ? record.name : normalizeString_(teacherName),
      teacherEmail: record ? record.email : '',
      roles: record ? record.roles : [],
      roleType: normalizeString_(roleType || 'support').toLowerCase() || 'support'
    };
  };
  resolver.teacherBundleLoadCount = 1;
  return resolver;
}

function buildTeachingAssignmentKey_(classId, term, weekday, period) {
  const normalizedClassId = normalizeString_(classId);
  const normalizedTerm = normalizeAcademicTerm_(term);
  const normalizedWeekday = normalizeWeekday_(weekday);
  const normalizedPeriod = normalizeTeachingAssignmentPeriod_(period);
  if (!normalizedClassId || !normalizedWeekday || !normalizedPeriod || (term && !normalizedTerm)) {
    return '';
  }
  return [normalizedClassId, normalizedTerm, normalizedWeekday, normalizedPeriod].join('__');
}

/**
 * Canonical period contract for teaching-assignment identity. Keep this local
 * to the assignment domain so file load order does not depend on planner code.
 */
function normalizeTeachingAssignmentPeriod_(value) {
  const text = normalizeString_(value);
  if (!/^\d+$/.test(text)) return '';
  const number = Number(text);
  return Number.isSafeInteger(number) && number > 0 ? String(number) : '';
}

function validateTeachingAssignmentColumns_(sheetName, colMap, requiredKeys) {
  requiredKeys.forEach(function(key) {
    if (colMap[key] === -1) {
      throw new Error(sheetName + ' シートに必要な列がありません: ' + key);
    }
  });
}

/**
 * Build the annual timetable/team index once. A termless timetable is the
 * explicit legacy mode; a present-but-empty term is never legacy data.
 */
function buildTeachingAssignmentIndex_(timetableData, teamData, resolveMember) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const ttHeaders = timetableData.headers || [];
  const teamHeaders = teamData.headers || [];
  const ttCol = {
    classId: findColumnIndex_(ttHeaders, ['classId', 'ClassID']),
    weekday: findColumnIndex_(ttHeaders, ['weekday', '曜日']),
    period: findColumnIndex_(ttHeaders, ['period', '時限']),
    teacherId: findColumnIndex_(ttHeaders, ['teacherId', 'TeacherID']),
    teacherName: findColumnIndex_(ttHeaders, ['teacherName', '担当者名', 'name']),
    term: findColumnIndex_(ttHeaders, ['term', '学期'])
  };
  validateTeachingAssignmentColumns_('timetable', ttCol, ['classId', 'weekday', 'period']);

  const teamCol = {
    classId: findColumnIndex_(teamHeaders, ['classId', 'ClassID']),
    weekday: findColumnIndex_(teamHeaders, ['weekday', '曜日']),
    period: findColumnIndex_(teamHeaders, ['period', '時限']),
    teacherId: findColumnIndex_(teamHeaders, ['teacherId', 'TeacherID']),
    teacherName: findColumnIndex_(teamHeaders, ['teacherName', '担当者名', 'name']),
    roleType: findColumnIndex_(teamHeaders, ['roleType', '役割']),
    term: findColumnIndex_(teamHeaders, ['term', '学期'])
  };

  const legacyTimetable = ttCol.term === -1;
  const legacyTeam = teamCol.term === -1;
  const byKey = {};
  const invalidKeys = {};
  const duplicateKeys = {};
  const warnings = [];
  const makeMember = typeof resolveMember === 'function'
    ? resolveMember
    : function(teacherId, teacherName, roleType) {
        return {
          teacherId: normalizeString_(teacherId),
          teacherName: normalizeString_(teacherName),
          teacherEmail: '',
          roles: [],
          roleType: normalizeString_(roleType || 'support').toLowerCase() || 'support'
        };
      };

  const timetableStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  let timetableValidRows = 0;
  let timetableDuplicateCount = 0;
  let timetableMemberResolveCount = 0;

  (timetableData.rows || []).forEach(function(row, index) {
    const classId = normalizeString_(row[ttCol.classId]);
    const weekday = normalizeWeekday_(row[ttCol.weekday]);
    const period = normalizeTeachingAssignmentPeriod_(row[ttCol.period]);
    const rawTerm = legacyTimetable ? '' : normalizeString_(row[ttCol.term]);
    const term = legacyTimetable ? '' : normalizeAcademicTerm_(rawTerm);
    if (!classId || !weekday || !period) return;
    if (!legacyTimetable && !term) {
      warnings.push('timetable row ' + (index + 2) + ': termが空欄または不正です');
      return;
    }

    const key = buildTeachingAssignmentKey_(classId, term, weekday, period);
    if (!key) return;
    timetableValidRows++;
    if (invalidKeys[key]) {
      duplicateKeys[key].rowNumbers.push(index + 2);
      timetableDuplicateCount++;
      warnings.push('timetable row ' + (index + 2) + ': duplicate assignment key (invalid)');
      return;
    }
    if (byKey[key]) {
      duplicateKeys[key] = {
        key: key,
        rowNumbers: [byKey[key].sourceRowNumber, index + 2]
      };
      invalidKeys[key] = true;
      delete byKey[key];
      timetableDuplicateCount++;
      warnings.push('timetable row ' + (index + 2) + ': duplicate assignment key (invalid)');
      return;
    }
    const teacherId = ttCol.teacherId === -1 ? '' : normalizeString_(row[ttCol.teacherId]);
    const teacherName = ttCol.teacherName === -1 ? '' : normalizeString_(row[ttCol.teacherName]);
    timetableMemberResolveCount++;
    const member = makeMember(teacherId, teacherName, 'main');
    byKey[key] = {
      classId: classId,
      term: term,
      weekday: weekday,
      period: period,
      teacherId: member && member.teacherId ? member.teacherId : '',
      teacherName: member && member.teacherName ? member.teacherName : teacherName,
      teacherIds: member && member.teacherId ? [member.teacherId] : [],
      teachers: member && member.teacherId ? [member] : [],
      sourceRowNumber: index + 2
    };
  });

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'buildTeachingAssignmentIndex timetable',
      timetableStartedAt,
      'totalRows=' + (timetableData.rows || []).length +
        ' validRows=' + timetableValidRows +
        ' duplicateCount=' + timetableDuplicateCount +
        ' memberResolveCount=' + timetableMemberResolveCount
    );
  }

  // A termful timetable cannot safely consume a termless team row. In legacy
  // mode both sheets are matched by the original classId/weekday/period key.
  const teamsStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  let teamMergedRows = 0;
  let teamOrphanCount = 0;
  let teamRejectedCount = 0;
  let teamMemberResolveCount = 0;
  (teamData.rows || []).forEach(function(row, index) {
    const classId = teamCol.classId === -1 ? '' : normalizeString_(row[teamCol.classId]);
    const weekday = teamCol.weekday === -1 ? '' : normalizeWeekday_(row[teamCol.weekday]);
    const period = teamCol.period === -1 ? '' : normalizeTeachingAssignmentPeriod_(row[teamCol.period]);
    const rawTerm = legacyTeam ? '' : normalizeString_(row[teamCol.term]);
    const term = legacyTeam ? '' : normalizeAcademicTerm_(rawTerm);
    if (!classId || !weekday || !period) {
      teamRejectedCount++;
      return;
    }
    if ((!legacyTeam && !term) || (legacyTimetable !== legacyTeam)) {
      teamRejectedCount++;
      warnings.push('classTeacherTeams row ' + (index + 2) + ': term契約がtimetableと一致しません');
      return;
    }

    const key = buildTeachingAssignmentKey_(classId, term, weekday, period);
    const bucket = byKey[key];
    if (!bucket) {
      teamOrphanCount++;
      warnings.push('classTeacherTeams row ' + (index + 2) + ': ' +
        (invalidKeys[key] ? '対応するtimetable行がduplicateで無効です' : '対応するtimetable行がありません'));
      return;
    }

    teamMemberResolveCount++;
    const member = makeMember(
      teamCol.teacherId === -1 ? '' : row[teamCol.teacherId],
      teamCol.teacherName === -1 ? '' : row[teamCol.teacherName],
      teamCol.roleType === -1 ? 'support' : row[teamCol.roleType]
    );
    if (!member || !member.teacherId || bucket.teacherIds.indexOf(member.teacherId) !== -1) return;
    bucket.teacherIds.push(member.teacherId);
    bucket.teachers.push(member);
    teamMergedRows++;
  });

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'buildTeachingAssignmentIndex teams',
      teamsStartedAt,
      'totalRows=' + (teamData.rows || []).length +
        ' mergedRows=' + teamMergedRows +
        ' orphanCount=' + teamOrphanCount +
        ' rejectedCount=' + teamRejectedCount +
        ' memberResolveCount=' + teamMemberResolveCount
    );
    logPerf_(
      'buildTeachingAssignmentIndex total',
      totalStartedAt,
      'bucketCount=' + Object.keys(byKey).length +
        ' warningCount=' + warnings.length +
        ' teacherBundleLoadCount=' + (Number(makeMember.teacherBundleLoadCount) || 0)
    );
  }

  return {
    byKey: byKey,
    legacyTimetable: legacyTimetable,
    legacyTeam: legacyTeam,
    warnings: warnings,
    invalidKeys: invalidKeys,
    duplicateKeys: duplicateKeys
  };
}

function getTeachingAssignmentForSessionFromIndex_(index, classId, date, period, context) {
  const classDayContext = context || getEffectiveClassDayContext_(date);
  if (!classDayContext.isClassDay || !classDayContext.effectiveWeekday) return null;
  const targetClassId = normalizeString_(classId);
  const targetPeriod = normalizeTeachingAssignmentPeriod_(period);
  if (!targetClassId || !targetPeriod) return null;

  const keys = index.legacyTimetable
    ? [buildTeachingAssignmentKey_(targetClassId, '', classDayContext.effectiveWeekday, targetPeriod)]
    : [
        buildTeachingAssignmentKey_(targetClassId, classDayContext.term, classDayContext.effectiveWeekday, targetPeriod),
        buildTeachingAssignmentKey_(targetClassId, 'FY', classDayContext.effectiveWeekday, targetPeriod)
      ];
  const teachers = [];
  const teacherIds = [];
  let primary = null;
  keys.forEach(function(key) {
    const bucket = index.byKey[key];
    if (!bucket) return;
    if (!primary) primary = bucket;
    (bucket.teachers || []).forEach(function(member) {
      if (!member.teacherId || teacherIds.indexOf(member.teacherId) !== -1) return;
      teacherIds.push(member.teacherId);
      teachers.push(member);
    });
  });
  if (!primary) return null;
  const main = teachers.find(function(member) { return member.roleType === 'main'; }) || teachers[0] || null;
  return {
    classId: targetClassId,
    date: classDayContext.ymd,
    term: classDayContext.term,
    weekday: classDayContext.effectiveWeekday,
    period: targetPeriod,
    teacherId: main ? main.teacherId : '',
    teacherName: main ? main.teacherName : '',
    teacherEmail: main ? main.teacherEmail : '',
    roles: main ? main.roles : [],
    teacherIds: teacherIds,
    teachers: teachers
  };
}

function getTeacherAssignmentsForClassSession_(classId, date, period) {
  const timetableData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.TIMETABLE, 300);
  const teamData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_TEACHER_TEAMS, 300);
  const index = buildTeachingAssignmentIndex_(
    timetableData,
    teamData,
    createTeacherTeamMemberResolver_()
  );
  const assignment = getTeachingAssignmentForSessionFromIndex_(index, classId, date, period);
  return assignment ? assignment.teachers : [];
}

function getTeacherAssignmentForClassSession_(classId, date, period) {
  const teachers = getTeacherAssignmentsForClassSession_(classId, date, period);
  if (!teachers.length) return null;
  const main = teachers.find(function(member) { return member.roleType === 'main'; }) || teachers[0];
  return Object.assign({}, main, { classId: normalizeString_(classId), date: formatDateToYmd(date), period: normalizeTeachingAssignmentPeriod_(period), teachers: teachers });
}

function getTeacherAssignmentsByClassId_(classId) {
  const targetClassId = normalizeString_(classId);
  if (!targetClassId) return [];

  const cacheKey = 'teacherAssignmentsByClassId__' + targetClassId +
    '__assignmentRevision__' + getTeachingAssignmentRevision_();
  const cached = getScriptCacheJson_(cacheKey);
  if (cached !== null) {
    return cached;
  }

  const ss = openOperationSpreadsheet_();
  const timetable = readSheetAsObjects_(ss, CONFIG.SHEETS.TIMETABLE);
  const teamRows = getClassTeacherTeamRows_();

  const resultMap = {};

  timetable
    .filter(item => normalizeString_(item.classId) === targetClassId)
    .forEach(function(row) {
      const teacherId = normalizeString_(row.teacherId);
      const teacherName = normalizeString_(row.teacherName);
      const member = buildTeacherTeamMember_(teacherId, teacherName, 'main');
      if (member.teacherId) {
        resultMap[member.teacherId] = member;
      }
    });

  teamRows
    .filter(item => normalizeString_(item.classId) === targetClassId)
    .forEach(function(row) {
      const teacherId = normalizeString_(row.teacherId);
      const teacherName = normalizeString_(row.teacherName);
      const roleType = normalizeString_(row.roleType || 'support').toLowerCase();
      const member = buildTeacherTeamMember_(teacherId, teacherName, roleType);
      if (member.teacherId) {
        resultMap[member.teacherId] = member;
      }
    });

  const result = Object.keys(resultMap).map(function(key) {
    return resultMap[key];
  });

  putScriptCacheJson_(cacheKey, result, 300);
  return result;
}

function getTeacherAssignmentsByClassPeriod_(classId, weekday, period) {
  const targetClassId = normalizeString_(classId);
  const targetWeekday = normalizeWeekday_(weekday);
  const targetPeriod = normalizeTeachingAssignmentPeriod_(period);

  if (!targetClassId || !targetPeriod) return [];

  const cacheKey = 'teacherAssignmentsByClassPeriod__' + [targetClassId, targetWeekday, targetPeriod].join('__') +
    '__assignmentRevision__' + getTeachingAssignmentRevision_();
  const cached = getScriptCacheJson_(cacheKey);
  if (cached !== null) {
    return cached;
  }

  const ss = openOperationSpreadsheet_();
  const timetable = readSheetAsObjects_(ss, CONFIG.SHEETS.TIMETABLE);
  const teamRows = getClassTeacherTeamRows_();

  const resultMap = {};

  timetable
    .filter(item =>
      normalizeString_(item.classId) === targetClassId &&
      normalizeTeachingAssignmentPeriod_(item.period) === targetPeriod &&
      normalizeWeekday_(item.weekday) === targetWeekday
    )
    .forEach(function(row) {
      const teacherId = normalizeString_(row.teacherId);
      const teacherName = normalizeString_(row.teacherName);
      const member = buildTeacherTeamMember_(teacherId, teacherName, 'main');
      if (member.teacherId) {
        resultMap[member.teacherId] = member;
      }
    });

  teamRows
    .filter(item =>
      normalizeString_(item.classId) === targetClassId &&
      normalizeTeachingAssignmentPeriod_(item.period) === targetPeriod &&
      normalizeWeekday_(item.weekday) === targetWeekday
    )
    .forEach(function(row) {
      const teacherId = normalizeString_(row.teacherId);
      const teacherName = normalizeString_(row.teacherName);
      const roleType = normalizeString_(row.roleType || 'support').toLowerCase();
      const member = buildTeacherTeamMember_(teacherId, teacherName, roleType);
      if (member.teacherId) {
        resultMap[member.teacherId] = member;
      }
    });

  const result = Object.keys(resultMap).map(function(key) {
    return resultMap[key];
  });

  putScriptCacheJson_(cacheKey, result, 300);
  return result;
}

function getTeacherAssignmentByClassId_(classId) {
  const targetClassId = normalizeString_(classId);
  if (!targetClassId) return null;

  const teachers = getTeacherAssignmentsByClassId_(targetClassId);
  if (!teachers || teachers.length === 0) return null;

  const mainTeacher = teachers.find(t => t.roleType === 'main') || teachers[0];

  return {
    classId: targetClassId,
    teacherId: mainTeacher.teacherId,
    teacherName: mainTeacher.teacherName,
    teacherEmail: mainTeacher.teacherEmail,
    roles: mainTeacher.roles,
    teachers: teachers
  };
}

function getTeacherAssignmentByClassPeriod_(classId, weekday, period) {
  const targetClassId = normalizeString_(classId);
  const targetWeekday = normalizeWeekday_(weekday);
  const targetPeriod = normalizeTeachingAssignmentPeriod_(period);

  if (!targetClassId || !targetPeriod) return null;

  const teachers = getTeacherAssignmentsByClassPeriod_(targetClassId, targetWeekday, targetPeriod);
  if (!teachers || teachers.length === 0) return null;

  const mainTeacher = teachers.find(t => t.roleType === 'main') || teachers[0];

  return {
    classId: targetClassId,
    weekday: targetWeekday,
    period: targetPeriod,
    teacherId: mainTeacher.teacherId,
    teacherName: mainTeacher.teacherName,
    teacherEmail: mainTeacher.teacherEmail,
    roles: mainTeacher.roles,
    teachers: teachers
  };
}

/**
 * 旧互換
 * 既存コードが email だけ欲しい場合のために一時的に残す
 */
function getTeacherByClassId(classId) {
  const assignment = getTeacherAssignmentByClassId_(classId);
  return assignment ? assignment.teacherEmail : '';
}

/**
 * 動作確認用
 */
function debugTeacherAssignmentsByClassId(classId) {
  const result = getTeacherAssignmentsByClassId_(classId);
  Logger.log(JSON.stringify(result, null, 2));
}

function debugTeacherAssignmentByClassPeriod(classId, weekday, period) {
  const result = getTeacherAssignmentByClassPeriod_(classId, weekday, period);
  Logger.log(JSON.stringify(result, null, 2));
}

function debugCurrentTeacherAssignment() {
  const user = getCurrentUserContext();
  Logger.log('current user=' + JSON.stringify(user, null, 2));

  if (user && user.email) {
    const byEmail = getTeacherRecordByEmail_(user.email);
    Logger.log('teacher by email=' + JSON.stringify(byEmail, null, 2));
  }
  if (user && user.teacherId) {
    const byId = getTeacherRecordById_(user.teacherId);
    Logger.log('teacher by id=' + JSON.stringify(byId, null, 2));
  }
}
