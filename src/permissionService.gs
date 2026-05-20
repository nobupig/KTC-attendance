function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function hasRole_(user, role) {
  if (!user || !Array.isArray(user.roles)) {
    return false;
  }
  return user.roles.includes(role);
}

function canAccessTeacherPage() {
  const user = getCurrentUserContext();
  if (!user) return false;

  if (hasRole_(user, 'admin')) {
    return true;
  }

  if (hasRole_(user, 'teacher')) {
    return true;
  }

  return hasAnyTeachingAssignmentByTeacherId_(user.teacherId);
}

function canAccessHomeroomPage() {
  const user = getCurrentUserContext();
  if (!user) return false;

  return Array.isArray(user.homeroomClasses) && user.homeroomClasses.length > 0;
}

function canAccessAdminPage() {
  const user = getCurrentUserContext();
  if (!user) return false;

  return hasRole_(user, 'admin');
}

function canEditAttendance(session) {
  const user = getCurrentUserContext();
  if (!user || !user.teacherId) {
    return false;
  }

  if (hasRole_(user, 'admin')) {
    return true;
  }

  const canUseTeacher = hasRole_(user, 'teacher') || hasAnyTeachingAssignmentByTeacherId_(user.teacherId);
  if (!canUseTeacher) {
    return false;
  }

  if (!session || !session.classId || !session.period || !session.date) {
    return false;
  }

  const weekday = getWeekdayFromYmdJst_(formatDateToYmd(session.date));
  const assignment = getTeacherAssignmentByClassPeriod_(session.classId, weekday, session.period);

  if (!assignment || !Array.isArray(assignment.teachers)) {
    return false;
  }

  return assignment.teachers.some(function(t) {
    return normalizeString_(t.teacherId) === normalizeString_(user.teacherId);
  });
}

/**
 * 動作確認用
 */
function debugPermissionService() {
  const user = getCurrentUserContext();
  Logger.log('user=' + JSON.stringify(user, null, 2));

  Logger.log('canAccessTeacherPage=' + canAccessTeacherPage());
  Logger.log('canAccessHomeroomPage=' + canAccessHomeroomPage());
  Logger.log('canAccessAdminPage=' + canAccessAdminPage());

  const ss = getOperationSpreadsheet();
  const timetableSheet = ss.getSheetByName(CONFIG.SHEETS.TIMETABLE);
  if (!timetableSheet) {
    Logger.log('timetable シートが見つかりません');
    return;
  }

  const values = timetableSheet.getDataRange().getValues();
  if (values.length < 2) {
    Logger.log('timetable にデータがありません');
    return;
  }

  const headers = values[0];
  const rows = values.slice(1);
  const classIdCol = findColumnIndex_(headers, ['classId', 'ClassID']);
  const periodCol = findColumnIndex_(headers, ['period', '時限']);
  const weekdayCol = findColumnIndex_(headers, ['weekday', '曜日']);

  if (classIdCol === -1) {
    Logger.log('timetable に classId 列がありません');
    return;
  }

  const firstRow = rows.find(function(r) {
    return normalizeString_(r[classIdCol]);
  });

  if (!firstRow) {
    Logger.log('timetable に classId が入っていません');
    return;
  }

  const testWeekday = weekdayCol !== -1 ? normalizeWeekday_(firstRow[weekdayCol]) : 'Mon';
  const session = {
    classId: normalizeString_(firstRow[classIdCol]),
    date: getSampleDateByWeekday_(testWeekday),
    period: periodCol !== -1 ? normalizeString_(firstRow[periodCol]) : '1'
  };

  Logger.log('testSession=' + JSON.stringify(session, null, 2));
  Logger.log('canEditAttendance=' + canEditAttendance(session));
}

function getSampleDateByWeekday_(weekday) {
  const map = {
    Mon: '2026-04-06',
    Tue: '2026-04-07',
    Wed: '2026-04-08',
    Thu: '2026-04-09',
    Fri: '2026-04-10',
    Sat: '2026-04-11',
    Sun: '2026-04-12'
  };
  return map[normalizeWeekday_(weekday)] || '2026-04-06';
}

function canManageHomeroomAttendance(grade, unit) {
  try {
    ensureHomeroomAccess_(grade, unit);
    return true;
  } catch (e) {
    return false;
  }
}