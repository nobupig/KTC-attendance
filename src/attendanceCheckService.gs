function getUnsubmittedClasses(targetDate) {
  const ss = getOperationSpreadsheet();

  const sessionsSheet = ss.getSheetByName(CONFIG.SHEETS.CLASS_SESSIONS);
  const attendanceSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  const attendanceSessionsSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_SESSIONS);

  if (!sessionsSheet) {
    throw new Error('classSessions シートが見つかりません');
  }
  if (!attendanceSheet) {
    throw new Error('attendance シートが見つかりません');
  }
  if (!attendanceSessionsSheet) {
    throw new Error('attendanceSessions シートが見つかりません');
  }

  const sessionsValues = sessionsSheet.getDataRange().getValues();
  const attendanceValues = attendanceSheet.getDataRange().getValues();
  const attendanceSessionsValues = attendanceSessionsSheet.getDataRange().getValues();

  if (sessionsValues.length < 2) {
    return [];
  }

  const sessionHeaders = sessionsValues[0];
  const sessionRows = sessionsValues.slice(1);

  const attendanceHeaders = attendanceValues.length > 0 ? attendanceValues[0] : [];
  const attendanceRows = attendanceValues.length > 1 ? attendanceValues.slice(1) : [];

  const attendanceSessionHeaders = attendanceSessionsValues.length > 0 ? attendanceSessionsValues[0] : [];
  const attendanceSessionRows = attendanceSessionsValues.length > 1 ? attendanceSessionsValues.slice(1) : [];

  const colSession = {
    classId: findColumnIndex_(sessionHeaders, ['classId', 'ClassID']),
    date: findColumnIndex_(sessionHeaders, ['date', '日付']),
    period: findColumnIndex_(sessionHeaders, ['period', '時限'])
  };

  const colAttendance = {
    classId: findColumnIndex_(attendanceHeaders, ['classId', 'ClassID']),
    date: findColumnIndex_(attendanceHeaders, ['date', '日付']),
    period: findColumnIndex_(attendanceHeaders, ['period', '時限'])
  };

  const colAttendanceSession = {
    classId: findColumnIndex_(attendanceSessionHeaders, ['classId', 'ClassID']),
    date: findColumnIndex_(attendanceSessionHeaders, ['date', '日付']),
    period: findColumnIndex_(attendanceSessionHeaders, ['period', '時限'])
  };

  validateRequiredColumnsForAttendanceCheck_('classSessions', colSession, ['classId', 'date', 'period']);
  validateRequiredColumnsForAttendanceCheck_('attendance', colAttendance, ['classId', 'date', 'period']);
  validateRequiredColumnsForAttendanceCheck_('attendanceSessions', colAttendanceSession, ['classId', 'date', 'period']);

  const target = formatDateToYmd(targetDate);

  const targetSessionRows = sessionRows.filter(function(row) {
    const classId = normalizeString_(row[colSession.classId]);
    if (!classId) return false;
    return formatDateToYmd(row[colSession.date]) === target;
  });

  const attendanceSet = new Set();
  attendanceRows.forEach(function(row) {
    const classId = normalizeString_(row[colAttendance.classId]);
    const date = formatDateToYmd(row[colAttendance.date]);
    const period = normalizeString_(row[colAttendance.period]);

    if (!classId || !date || !period) return;
    if (date !== target) return;

    attendanceSet.add(buildAttendanceCheckKey_(classId, date, period));
  });

  const attendanceSessionSet = new Set();
  attendanceSessionRows.forEach(function(row) {
    const classId = normalizeString_(row[colAttendanceSession.classId]);
    const date = formatDateToYmd(row[colAttendanceSession.date]);
    const period = normalizeString_(row[colAttendanceSession.period]);

    if (!classId || !date || !period) return;
    if (date !== target) return;

    attendanceSessionSet.add(buildAttendanceCheckKey_(classId, date, period));
  });

  const unsubmitted = [];
  const seen = new Set();

  targetSessionRows.forEach(function(row) {
    const classId = normalizeString_(row[colSession.classId]);
    const date = formatDateToYmd(row[colSession.date]);
    const period = normalizeString_(row[colSession.period]);

    const key = buildAttendanceCheckKey_(classId, date, period);
    if (seen.has(key)) return;
    seen.add(key);

const hasAttendanceSession = attendanceSessionSet.has(key);

if (!hasAttendanceSession) {
  unsubmitted.push({
    classId: classId,
    date: date,
    period: period
  });
}
  });

  return unsubmitted;
}

function getUnsubmittedClassesFast_(targetDate) {
  const target = formatDateToYmd(targetDate);

  const sessionData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_SESSIONS, 30);
  const attendanceData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.ATTENDANCE, 15);
  const attendanceSessionData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.ATTENDANCE_SESSIONS, 15);

  const colSession = {
    classId: findColumnIndex_(sessionData.headers, ['classId', 'ClassID']),
    date: findColumnIndex_(sessionData.headers, ['date', '日付']),
    period: findColumnIndex_(sessionData.headers, ['period', '時限'])
  };

  const colAttendance = {
    classId: findColumnIndex_(attendanceData.headers, ['classId', 'ClassID']),
    date: findColumnIndex_(attendanceData.headers, ['date', '日付']),
    period: findColumnIndex_(attendanceData.headers, ['period', '時限'])
  };

  const colAttendanceSession = {
    classId: findColumnIndex_(attendanceSessionData.headers, ['classId', 'ClassID']),
    date: findColumnIndex_(attendanceSessionData.headers, ['date', '日付']),
    period: findColumnIndex_(attendanceSessionData.headers, ['period', '時限'])
  };

  validateRequiredColumnsForAttendanceCheck_('classSessions', colSession, ['classId', 'date', 'period']);
  validateRequiredColumnsForAttendanceCheck_('attendance', colAttendance, ['classId', 'date', 'period']);
  validateRequiredColumnsForAttendanceCheck_('attendanceSessions', colAttendanceSession, ['classId', 'date', 'period']);

  const attendanceSet = new Set();
  attendanceData.rows.forEach(function(row) {
    const classId = normalizeString_(row[colAttendance.classId]);
    const date = formatDateToYmd(row[colAttendance.date]);
    const period = normalizeString_(row[colAttendance.period]);
    if (!classId || !date || !period) return;
    if (date !== target) return;
    attendanceSet.add(buildAttendanceCheckKey_(classId, date, period));
  });

  const attendanceSessionSet = new Set();
  attendanceSessionData.rows.forEach(function(row) {
    const classId = normalizeString_(row[colAttendanceSession.classId]);
    const date = formatDateToYmd(row[colAttendanceSession.date]);
    const period = normalizeString_(row[colAttendanceSession.period]);
    if (!classId || !date || !period) return;
    if (date !== target) return;
    attendanceSessionSet.add(buildAttendanceCheckKey_(classId, date, period));
  });

  const result = [];
  const seen = new Set();

  sessionData.rows.forEach(function(row) {
    const classId = normalizeString_(row[colSession.classId]);
    const date = formatDateToYmd(row[colSession.date]);
    const period = normalizeString_(row[colSession.period]);

    if (!classId || !date || !period) return;
    if (date !== target) return;

    const key = buildAttendanceCheckKey_(classId, date, period);
    if (seen.has(key)) return;
    seen.add(key);

const hasAttendanceSession = attendanceSessionSet.has(key);

if (!hasAttendanceSession) {
  result.push({
    classId: classId,
    date: date,
    period: period
  });
}
  });

  return result;
}

function checkYesterdayAttendance() {
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);

  const unsubmitted = getUnsubmittedClasses(yesterday);

  if (unsubmitted.length === 0) {
    return;
  }

  const grouped = {};
  const logData = [];

  unsubmitted.forEach(function(item) {
    const weekday = getWeekdayFromYmdJst_(item.date);
    const assignment = getTeacherAssignmentByClassPeriod_(item.classId, weekday, item.period);

    if (!assignment || !Array.isArray(assignment.teachers) || assignment.teachers.length === 0) {
      return;
    }

    assignment.teachers.forEach(function(teacher) {
      const teacherEmail = normalizeString_(teacher.teacherEmail).toLowerCase();
      if (!teacherEmail) return;

      if (!grouped[teacherEmail]) {
        grouped[teacherEmail] = {
          teacherName: teacher.teacherName || '',
          items: []
        };
      }

      grouped[teacherEmail].items.push({
        classId: item.classId,
        date: item.date,
        period: item.period,
        roleType: teacher.roleType || ''
      });
    });
  });

  Object.keys(grouped).forEach(function(email) {
    const group = grouped[email];
    const classes = group.items;
    const teacherName = group.teacherName || email;

    let message = '【出席未入力通知】\n\n';
    message += '担当教員: ' + teacherName + '\n\n';

    classes.forEach(function(c) {
      const className = getClassDisplayName(c.classId);
      const url =
        CONFIG.APP.BASE_URL +
        '?page=teacher' +
        '&classId=' + encodeURIComponent(c.classId) +
        '&date=' + encodeURIComponent(c.date) +
        '&period=' + encodeURIComponent(c.period);

      message += '授業: ' + className + '\n';
      message += '日付: ' + c.date + '\n';
      message += '時限: ' + c.period + '\n';
      if (c.roleType) {
        message += '担当区分: ' + c.roleType + '\n';
      }
      message += '\n▶ 出席入力はこちら\n' + url + '\n\n';

      logData.push({
        classId: c.classId,
        className: className,
        date: c.date,
        period: c.period,
        teacherEmail: email,
        teacherName: teacherName
      });
    });

    sendSlackMessage(message);
  });

  recordMissingAttendanceLog(logData);
}

function recordMissingAttendanceLog(data) {
  if (!data || data.length === 0) {
    return;
  }

  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_MISSING_LOG);

  if (!sheet) {
    throw new Error('attendanceMissingLog シートが見つかりません');
  }

  const now = new Date();

  const rows = data.map(function(item) {
    return [
      now,
      item.classId,
      item.className,
      item.date,
      item.period,
      item.teacherEmail,
      item.teacherName
    ];
  });

  sheet.getRange(
    sheet.getLastRow() + 1,
    1,
    rows.length,
    rows[0].length
  ).setValues(rows);
}

function buildAttendanceCheckKey_(classId, date, period) {
  return [
    normalizeString_(classId),
    normalizeString_(date),
    normalizeString_(period)
  ].join('__');
}

function validateRequiredColumnsForAttendanceCheck_(sheetName, colMap, requiredKeys) {
  requiredKeys.forEach(function(key) {
    if (colMap[key] === -1) {
      throw new Error(sheetName + ' シートに必要な列がありません: ' + key);
    }
  });
}

/**
 * 動作確認用
 */
function debugAttendanceCheckService(targetDate) {
  const date = targetDate || new Date();
  const result = getUnsubmittedClasses(date);
  Logger.log(JSON.stringify(result, null, 2));
}