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

function getTeacherAssignmentsByClassId_(classId) {
  const targetClassId = normalizeString_(classId);
  if (!targetClassId) return [];

  const cacheKey = 'teacherAssignmentsByClassId__' + targetClassId;
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
  const targetPeriod = normalizeString_(period);

  if (!targetClassId || !targetPeriod) return [];

  const cacheKey = 'teacherAssignmentsByClassPeriod__' + [targetClassId, targetWeekday, targetPeriod].join('__');
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
      normalizeString_(item.period) === targetPeriod &&
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
      normalizeString_(item.period) === targetPeriod &&
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
  const targetPeriod = normalizeString_(period);

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
