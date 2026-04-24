function getAdminDashboardData(targetDate, gradeFilter) {
  if (!canAccessAdminPage()) {
    throw new Error('教務画面の権限がありません。');
  }

  const ymd = formatDateToYmd(targetDate || new Date());
  const allRows = getAdminTargetSessions_(ymd);

  // 高速版を優先利用
  const missingList = typeof getUnsubmittedClassesFast_ === 'function'
    ? getUnsubmittedClassesFast_(ymd)
    : getUnsubmittedClasses(ymd);

  const missingSet = new Set(
    missingList.map(function(item) {
      return [item.classId, item.date, item.period].join('__');
    })
  );

  let rows = allRows
    .map(function(row) {
      const key = [row.classId, row.date, row.period].join('__');
      return Object.assign({}, row, {
        status: missingSet.has(key) ? '未入力' : '入力済'
      });
    })
    .filter(function(row) {
      return row.status === '未入力';
    });

  if (gradeFilter) {
    rows = rows.filter(function(row) {
      return String(row.grade) === String(gradeFilter);
    });
  }

  rows.sort(function(a, b) {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    if (Number(a.period) !== Number(b.period)) return Number(a.period) - Number(b.period);
    if (Number(a.grade) !== Number(b.grade)) return Number(a.grade) - Number(b.grade);
    return String(a.unit).localeCompare(String(b.unit), 'ja');
  });

  const totalCount = gradeFilter
    ? allRows.filter(function(row) {
        return String(row.grade) === String(gradeFilter);
      }).length
    : allRows.length;

  return {
    targetDate: ymd,
    totalCount: totalCount,
    missingCount: rows.length,
    updatedAt: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone() || 'Asia/Tokyo',
      'yyyy-MM-dd HH:mm:ss'
    ),
    rows: rows
  };
}

function getAdminMissingSummaryRange(startDate, endDate) {
  if (!canAccessAdminPage()) {
    throw new Error('教務画面の権限がありません。');
  }

  const startYmd = formatDateToYmd(startDate || '2026-04-07');
  const endYmd = formatDateToYmd(endDate);

  if (!endYmd) {
    throw new Error('終了日が指定されていません。');
  }

  if (endYmd < startYmd) {
    throw new Error('終了日は開始日以降を指定してください。');
  }

  const grouped = {};
  const dates = buildAdminDateRange_(startYmd, endYmd);

  dates.forEach(function(ymd) {
    const data = getAdminDashboardData(ymd, '');

    (data.rows || []).forEach(function(row) {
      const key = [
        row.date,
        row.classId,
        row.subjectName,
        row.teacherName
      ].join('__');

      if (!grouped[key]) {
        grouped[key] = {
          date: row.date,
          classId: row.classId,
          subjectName: row.subjectName || '(科目名未設定)',
          teacherName: row.teacherName || '(担当者未設定)',
          grade: row.grade || '',
          unit: row.unit || '',
          periods: []
        };
      }

      grouped[key].periods.push(Number(row.period));
    });
  });

  const rows = Object.keys(grouped).map(function(key) {
    const item = grouped[key];

    item.periods = item.periods
      .filter(function(p) {
        return !isNaN(p);
      })
      .sort(function(a, b) {
        return a - b;
      });

    return {
      date: item.date,
      dateText: formatAdminDateForDisplay_(item.date),
      classId: item.classId,
      subjectName: item.subjectName,
      teacherName: item.teacherName,
      grade: item.grade,
      unit: item.unit,
      periodText: formatAdminPeriodRange_(item.periods)
    };
  });

  rows.sort(function(a, b) {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    if (a.teacherName !== b.teacherName) return a.teacherName.localeCompare(b.teacherName, 'ja');
    if (a.subjectName !== b.subjectName) return a.subjectName.localeCompare(b.subjectName, 'ja');
    return a.classId.localeCompare(b.classId);
  });

  const copyText = buildAdminMissingSummaryCopyText_(startYmd, endYmd, rows);

  return {
    startDate: startYmd,
    endDate: endYmd,
    count: rows.length,
    rows: rows,
    copyText: copyText
  };
}

function buildAdminDateRange_(startYmd, endYmd) {
  const dates = [];
  let current = parseAdminYmd_(startYmd);
  const end = parseAdminYmd_(endYmd);

  while (current.getTime() <= end.getTime()) {
    dates.push(Utilities.formatDate(current, 'Asia/Tokyo', 'yyyy-MM-dd'));
    current.setDate(current.getDate() + 1);
  }

  return dates;
}

function parseAdminYmd_(ymd) {
  const parts = String(ymd).split('-');
  return new Date(
    Number(parts[0]),
    Number(parts[1]) - 1,
    Number(parts[2])
  );
}

function formatAdminDateForDisplay_(ymd) {
  const date = parseAdminYmd_(ymd);
  const weekdays = ['日', '月', '火', '水', '木', '金', '土'];

  return Utilities.formatDate(date, 'Asia/Tokyo', 'M月d日') +
    '(' + weekdays[date.getDay()] + ')';
}

function formatAdminPeriodRange_(periods) {
  if (!Array.isArray(periods) || periods.length === 0) {
    return '';
  }

  const unique = Array.from(new Set(periods)).sort(function(a, b) {
    return a - b;
  });

  const parts = [];
  let start = unique[0];
  let prev = unique[0];

  for (let i = 1; i < unique.length; i++) {
    const current = unique[i];

    if (current === prev + 1) {
      prev = current;
      continue;
    }

    parts.push(formatAdminPeriodPart_(start, prev));
    start = current;
    prev = current;
  }

  parts.push(formatAdminPeriodPart_(start, prev));

  return parts.join('、') + '限';
}

function formatAdminPeriodPart_(start, end) {
  if (start === end) {
    return String(start);
  }

  return start + '～' + end;
}

function buildAdminMissingSummaryCopyText_(startYmd, endYmd, rows) {
  const header = '【' +
    formatAdminDateForDisplay_(startYmd) +
    '～' +
    formatAdminDateForDisplay_(endYmd) +
    ' 未保存授業一覧】';

  if (!rows || rows.length === 0) {
    return header + '\n\n未保存の授業はありません。';
  }

  const lines = rows.map(function(row) {
    return [
      row.dateText,
      row.subjectName,
      row.periodText,
      row.teacherName
    ].join(' ｜ ');
  });

  return header + '\n\n' + lines.join('\n');
}

function getAdminTargetSessions_(targetDate) {
  const ymd = formatDateToYmd(targetDate);
  const weekday = getWeekdayFromYmdJst_(ymd);

  const classSessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_SESSIONS, 60);
  const classesData = getSheetDataCached_('MASTER', CONFIG.SHEETS.CLASSES, 300);
  const timetableData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.TIMETABLE, 300);
  const teamData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_TEACHER_TEAMS, 300);
  const teachersData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.TEACHERS, 300);

  const csCol = {
    classId: findColumnIndex_(classSessionsData.headers, ['classId', 'ClassID']),
    date: findColumnIndex_(classSessionsData.headers, ['date', '日付']),
    period: findColumnIndex_(classSessionsData.headers, ['period', '時限']),
    sessionNumber: findColumnIndex_(classSessionsData.headers, ['sessionNumber', '回', '回数'])
  };
  validateRequiredColumnsForAdmin_('classSessions', csCol, ['classId', 'date', 'period']);

  const classMap = buildAdminClassMap_(classesData);
  const teacherLookup = buildAdminTeacherLookup_(teachersData);
  const assignmentMap = buildAdminAssignmentMap_(timetableData, teamData, teacherLookup);

  return classSessionsData.rows
    .filter(function(row) {
      return formatDateToYmd(row[csCol.date]) === ymd;
    })
    .map(function(row) {
      const classId = normalizeString_(row[csCol.classId]);
      const period = normalizeString_(row[csCol.period]);
      const sessionNumber = csCol.sessionNumber !== -1 ? row[csCol.sessionNumber] : '';
      const classInfo = classMap[classId] || {};
      const assignKey = buildAdminAssignKey_(classId, weekday, period);
      const teacherAssign = assignmentMap[assignKey] || null;

      const mainTeacher = teacherAssign
        ? (teacherAssign.teachers.find(function(t) { return t.roleType === 'main'; }) || teacherAssign.teachers[0] || null)
        : null;

      return {
        classId: classId,
        date: ymd,
        period: period,
        sessionNumber: sessionNumber,
        subjectId: classInfo.subjectId || '',
        subjectName: classInfo.subjectName || '',
        grade: classInfo.grade || '',
        unit: classInfo.unit || '',
        term: classInfo.term || '',
        teacherId: mainTeacher ? mainTeacher.teacherId : '',
        teacherName: mainTeacher ? mainTeacher.teacherName : '',
        teachers: teacherAssign ? teacherAssign.teachers : []
      };
    });
}

function getAdminProxySessionForEdit(classId, date, period) {
  if (!canAccessAdminPage()) {
    throw new Error('教務画面の権限がありません。');
  }

  const ymd = formatDateToYmd(date);
  const rows = getAdminTargetSessions_(ymd);

  const hit = rows.find(function(row) {
    return row.classId === String(classId) &&
      row.date === ymd &&
      String(row.period) === String(period);
  });

  if (!hit) {
    throw new Error('対象授業が見つかりません。');
  }

  return hit;
}

function saveProxyAttendanceLog(log) {
  if (!canAccessAdminPage()) {
    throw new Error('教務画面の権限がありません。');
  }

  const sheet = getOperationSheet(CONFIG.SHEETS.PROXY_ATTENDANCE_LOG);
  if (!sheet) {
    throw new Error('proxyAttendanceLog シートが見つかりません。');
  }

  sheet.appendRow([
    new Date(),
    log.adminEmail || '',
    log.adminName || '',
    log.classId || '',
    log.subjectName || '',
    log.date || '',
    log.period || '',
    log.proxyForTeacherId || '',
    log.proxyForTeacherName || '',
    log.proxyReasonType || '',
    log.proxyReasonNote || '',
    log.savedCount || 0
  ]);
}

function buildAdminClassMap_(classesData) {
  const col = {
    classId: findColumnIndex_(classesData.headers, ['classId', 'ClassID']),
    subjectId: findColumnIndex_(classesData.headers, ['subjectId', 'SubjectID']),
    subjectName: findColumnIndex_(classesData.headers, ['subjectName', '科目名', 'className']),
    grade: findColumnIndex_(classesData.headers, ['grade', '学年']),
    unit: findColumnIndex_(classesData.headers, ['unit', '対象区分', '組・コース']),
    term: findColumnIndex_(classesData.headers, ['term', '開設期'])
  };
  validateRequiredColumnsForAdmin_('classes', col, ['classId', 'grade', 'unit']);

  const map = {};
  classesData.rows.forEach(function(row) {
    const classId = normalizeString_(row[col.classId]);
    if (!classId) return;

    map[classId] = {
      classId: classId,
      subjectId: col.subjectId !== -1 ? normalizeString_(row[col.subjectId]) : '',
      subjectName: col.subjectName !== -1 ? normalizeString_(row[col.subjectName]) : '',
      grade: col.grade !== -1 ? normalizeString_(row[col.grade]) : '',
      unit: col.unit !== -1 ? normalizeString_(row[col.unit]) : '',
      term: col.term !== -1 ? normalizeString_(row[col.term]) : ''
    };
  });

  return map;
}

function buildAdminTeacherLookup_(teachersData) {
  const col = {
    teacherId: findColumnIndex_(teachersData.headers, ['teacherId', 'TeacherID']),
    name: findColumnIndex_(teachersData.headers, ['name', '氏名', 'teacherName']),
    email: findColumnIndex_(teachersData.headers, ['email', 'メールアドレス']),
    roles: findColumnIndex_(teachersData.headers, ['roles', 'role'])
  };
  validateRequiredColumnsForAdmin_('teachers', col, ['teacherId', 'name', 'email']);

  const byId = {};
  const byName = {};

  teachersData.rows.forEach(function(row) {
    const teacherId = normalizeString_(row[col.teacherId]);
    const name = normalizeString_(row[col.name]);
    const email = normalizeString_(row[col.email]).toLowerCase();
    const roles = col.roles !== -1 ? parseRoles_(row[col.roles]) : [];

    const item = {
      teacherId: teacherId,
      teacherName: name,
      teacherEmail: email,
      roles: roles
    };

    if (teacherId) byId[teacherId] = item;
    if (name) byName[name] = item;
  });

  return {
    byId: byId,
    byName: byName
  };
}

function buildAdminAssignmentMap_(timetableData, teamData, teacherLookup) {
  const ttCol = {
    classId: findColumnIndex_(timetableData.headers, ['classId', 'ClassID']),
    weekday: findColumnIndex_(timetableData.headers, ['weekday', '曜日']),
    period: findColumnIndex_(timetableData.headers, ['period', '時限']),
    teacherId: findColumnIndex_(timetableData.headers, ['teacherId', 'TeacherID']),
    teacherName: findColumnIndex_(timetableData.headers, ['teacherName', '担当者名', 'name'])
  };
  validateRequiredColumnsForAdmin_('timetable', ttCol, ['classId', 'weekday', 'period']);

  const teamCol = {
    classId: findColumnIndex_(teamData.headers, ['classId', 'ClassID']),
    weekday: findColumnIndex_(teamData.headers, ['weekday', '曜日']),
    period: findColumnIndex_(teamData.headers, ['period', '時限']),
    teacherId: findColumnIndex_(teamData.headers, ['teacherId', 'TeacherID']),
    teacherName: findColumnIndex_(teamData.headers, ['teacherName', '担当者名', 'name']),
    roleType: findColumnIndex_(teamData.headers, ['roleType', '役割'])
  };

  const map = {};

  timetableData.rows.forEach(function(row) {
    const classId = normalizeString_(row[ttCol.classId]);
    const weekday = normalizeWeekday_(row[ttCol.weekday]);
    const period = normalizeString_(row[ttCol.period]);
    const teacherId = ttCol.teacherId !== -1 ? normalizeString_(row[ttCol.teacherId]) : '';
    const teacherName = ttCol.teacherName !== -1 ? normalizeString_(row[ttCol.teacherName]) : '';

    if (!classId || !weekday || !period) return;

    const key = buildAdminAssignKey_(classId, weekday, period);
    if (!map[key]) {
      map[key] = {
        teacherIds: [],
        teachers: []
      };
    }

    addAdminAssignmentMember_(map[key], teacherLookup, teacherId, teacherName, 'main');
  });

  teamData.rows.forEach(function(row) {
    const classId = teamCol.classId !== -1 ? normalizeString_(row[teamCol.classId]) : '';
    const weekday = teamCol.weekday !== -1 ? normalizeWeekday_(row[teamCol.weekday]) : '';
    const period = teamCol.period !== -1 ? normalizeString_(row[teamCol.period]) : '';
    const teacherId = teamCol.teacherId !== -1 ? normalizeString_(row[teamCol.teacherId]) : '';
    const teacherName = teamCol.teacherName !== -1 ? normalizeString_(row[teamCol.teacherName]) : '';
    const roleType = teamCol.roleType !== -1
      ? normalizeString_(row[teamCol.roleType]).toLowerCase()
      : 'support';

    if (!classId || !weekday || !period) return;

    const key = buildAdminAssignKey_(classId, weekday, period);
    if (!map[key]) {
      map[key] = {
        teacherIds: [],
        teachers: []
      };
    }

    addAdminAssignmentMember_(map[key], teacherLookup, teacherId, teacherName, roleType || 'support');
  });

  return map;
}

function addAdminAssignmentMember_(bucket, teacherLookup, teacherId, teacherName, roleType) {
  let record = null;

  if (teacherId && teacherLookup.byId[teacherId]) {
    record = teacherLookup.byId[teacherId];
  } else if (teacherName && teacherLookup.byName[teacherName]) {
    record = teacherLookup.byName[teacherName];
  }

  const resolvedTeacherId = record ? record.teacherId : normalizeString_(teacherId);
  const resolvedTeacherName = record ? record.teacherName : normalizeString_(teacherName);
  const resolvedTeacherEmail = record ? record.teacherEmail : '';
  const resolvedRoles = record ? record.roles : [];
  const dedupeKey = resolvedTeacherId || resolvedTeacherName;

  if (!dedupeKey) return;
  if (bucket.teacherIds.indexOf(dedupeKey) !== -1) return;

  bucket.teacherIds.push(dedupeKey);
  bucket.teachers.push({
    teacherId: resolvedTeacherId,
    teacherName: resolvedTeacherName,
    teacherEmail: resolvedTeacherEmail,
    roles: resolvedRoles,
    roleType: normalizeString_(roleType || 'support').toLowerCase() || 'support'
  });
}

function buildAdminAssignKey_(classId, weekday, period) {
  return [
    normalizeString_(classId),
    normalizeWeekday_(weekday),
    normalizeString_(period)
  ].join('__');
}

function validateRequiredColumnsForAdmin_(sheetName, colMap, requiredKeys) {
  requiredKeys.forEach(function(key) {
    if (colMap[key] === -1) {
      throw new Error(sheetName + ' シートに必要な列がありません: ' + key);
    }
  });
}