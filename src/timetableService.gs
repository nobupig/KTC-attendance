function getClassesForCurrentUserByDate(targetDate) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const userStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const user = getCurrentUserContext();
  if (typeof logPerf_ === 'function') {
    logPerf_('getClassesForCurrentUserByDate getCurrentUserContext', userStartedAt);
  }

  if (!user || !user.teacherId) {
    throw new Error('ログインユーザー情報を取得できませんでした。');
  }

  const currentTeacherId = normalizeString_(user.teacherId);
  const currentUserEmail = normalizeString_(user.email).toLowerCase();
  const ymd = targetDate
  ? formatDateToYmd(targetDate)
  : formatDateToYmd(new Date());

  const loadSheetsStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const timetableData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.TIMETABLE, 300);
  const classesData = getSheetDataCached_('MASTER', CONFIG.SHEETS.CLASSES, 300);
  const teamData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_TEACHER_TEAMS, 300);
  if (typeof logPerf_ === 'function') {
    logPerf_('getClassesForCurrentUserByDate load base sheet data', loadSheetsStartedAt);
  }

  const timetable = timetableData.rows;
  const classes = classesData.rows;
  const teamRows = teamData.rows;

  const timetableHeaders = timetableData.headers;
  const ttCol = {
    classId: findColumnIndex_(timetableHeaders, ['classId', 'ClassID']),
    weekday: findColumnIndex_(timetableHeaders, ['weekday', '曜日']),
    period: findColumnIndex_(timetableHeaders, ['period', '時限']),
    teacherName: findColumnIndex_(timetableHeaders, ['teacherName', '担当者名', 'name']),
    teacherId: findColumnIndex_(timetableHeaders, ['teacherId', 'TeacherID'])
  };
  validateRequiredColumnsForTimetable_('timetable', ttCol, ['classId', 'weekday', 'period']);

  const teamHeaders = teamData.headers;
  const teamCol = {
    classId: findColumnIndex_(teamHeaders, ['classId', 'ClassID']),
    weekday: findColumnIndex_(teamHeaders, ['weekday', '曜日']),
    period: findColumnIndex_(teamHeaders, ['period', '時限']),
    teacherName: findColumnIndex_(teamHeaders, ['teacherName', '担当者名', 'name']),
    teacherId: findColumnIndex_(teamHeaders, ['teacherId', 'TeacherID']),
    roleType: findColumnIndex_(teamHeaders, ['roleType', '役割'])
  };

  const timetableMapStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const timetableMap = {};
  timetable.forEach(function(row) {
    const classId = normalizeString_(row[ttCol.classId]);
    const period = normalizeString_(row[ttCol.period]);
    const weekday = ttCol.weekday !== -1 ? normalizeWeekday_(row[ttCol.weekday]) : '';
    const teacherName = ttCol.teacherName !== -1 ? normalizeString_(row[ttCol.teacherName]) : '';
    let teacherId = ttCol.teacherId !== -1 ? normalizeString_(row[ttCol.teacherId]) : '';

    if (!teacherId && teacherName) {
      const teacher = getTeacherRecordByName_(teacherName);
      teacherId = teacher ? teacher.teacherId : '';
    }

    if (!classId || !period || !weekday) return;

    const key = classId + '__' + weekday + '__' + period;
    timetableMap[key] = {
      classId: classId,
      weekday: weekday,
      period: period,
      teacherId: teacherId,
      teacherName: teacherName,
      teacherIds: teacherId ? [teacherId] : [],
      teachers: teacherId ? [{
        teacherId: teacherId,
        teacherName: teacherName,
        roleType: 'main'
      }] : []
    };
  });
  if (typeof logPerf_ === 'function') {
    logPerf_('getClassesForCurrentUserByDate build timetableMap', timetableMapStartedAt, 'rows=' + timetable.length);
  }

  const teamMergeStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  teamRows.forEach(function(row) {
    const classId = teamCol.classId !== -1 ? normalizeString_(row[teamCol.classId]) : '';
    const period = teamCol.period !== -1 ? normalizeString_(row[teamCol.period]) : '';
    const weekday = teamCol.weekday !== -1 ? normalizeWeekday_(row[teamCol.weekday]) : '';
    const teacherName = teamCol.teacherName !== -1 ? normalizeString_(row[teamCol.teacherName]) : '';
    let teacherId = teamCol.teacherId !== -1 ? normalizeString_(row[teamCol.teacherId]) : '';
    const roleType = teamCol.roleType !== -1
      ? normalizeString_(row[teamCol.roleType]).toLowerCase()
      : 'support';

    if (!teacherId && teacherName) {
      const teacher = getTeacherRecordByName_(teacherName);
      teacherId = teacher ? teacher.teacherId : '';
    }

    if (!classId || !period || !weekday || !teacherId) return;

    const key = classId + '__' + weekday + '__' + period;
    if (!timetableMap[key]) {
      timetableMap[key] = {
        classId: classId,
        weekday: weekday,
        period: period,
        teacherId: '',
        teacherName: '',
        teacherIds: [],
        teachers: []
      };
    }

    if (!timetableMap[key].teacherIds.includes(teacherId)) {
      timetableMap[key].teacherIds.push(teacherId);
      timetableMap[key].teachers.push({
        teacherId: teacherId,
        teacherName: teacherName,
        roleType: roleType || 'support'
      });
    }
  });
  if (typeof logPerf_ === 'function') {
    logPerf_('getClassesForCurrentUserByDate merge classTeacherTeams', teamMergeStartedAt, 'rows=' + teamRows.length);
  }

  // 先に「この教員が担当する授業キー」だけを抽出
  const teacherKeyStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const teacherSessionKeyMap = {};
  Object.keys(timetableMap).forEach(function(key) {
    const tt = timetableMap[key];
    if (tt && Array.isArray(tt.teacherIds) && tt.teacherIds.includes(currentTeacherId)) {
      teacherSessionKeyMap[key] = true;
    }
  });
  if (typeof logPerf_ === 'function') {
    logPerf_('getClassesForCurrentUserByDate build teacherSessionKeyMap', teacherKeyStartedAt, 'keys=' + Object.keys(teacherSessionKeyMap).length);
  }

  const savedSessionStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const savedSessionMapBase = getSavedSessionMapByDateCached_(ymd);
  const savedSessionMap = {};
  Object.keys(savedSessionMapBase).forEach(function(key) {
    const raw = savedSessionMapBase[key];
    savedSessionMap[key] = {
      teacherEmail: raw.teacherEmail || '',
      savedAtText: raw.savedAtText || '',
      actionType: raw.actionType || '',
      targetSessionKey: raw.targetSessionKey || '',
      savedModeLabel: raw.savedModeLabel || '',
      savedByCurrentUser: !!raw.teacherEmail && raw.teacherEmail === currentUserEmail
    };
  });
  if (typeof logPerf_ === 'function') {
    logPerf_('getClassesForCurrentUserByDate build savedSessionMap', savedSessionStartedAt, 'rows=' + Object.keys(savedSessionMap).length);
  }

  const classMapStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const classHeaders = classesData.headers;
  const clsCol = {
    classId: findColumnIndex_(classHeaders, ['classId', 'ClassID']),
    subjectId: findColumnIndex_(classHeaders, ['subjectId', 'SubjectID']),
    subjectName: findColumnIndex_(classHeaders, ['subjectName', '科目名']),
    grade: findColumnIndex_(classHeaders, ['grade', '学年']),
    unit: findColumnIndex_(classHeaders, ['unit', '対象区分', '組・コース']),
    term: findColumnIndex_(classHeaders, ['term', '開設期']),
    curriculumUnit: findColumnIndex_(classHeaders, ['curriculumUnit', '組・コース']),
    allowedAbsences: findColumnIndex_(classHeaders, ['allowedAbsences', '欠席可能コマ数'])
  };
  validateRequiredColumnsForTimetable_('classes', clsCol, ['classId', 'subjectId', 'subjectName', 'grade', 'unit']);

  const classMap = {};
  classes.forEach(function(row) {
    const classId = normalizeString_(row[clsCol.classId]);
    if (!classId) return;

    classMap[classId] = {
      classId: classId,
      subjectId: clsCol.subjectId !== -1 ? normalizeString_(row[clsCol.subjectId]) : '',
      subjectName: clsCol.subjectName !== -1 ? normalizeString_(row[clsCol.subjectName]) : '',
      grade: clsCol.grade !== -1 ? normalizeString_(row[clsCol.grade]) : '',
      unit: clsCol.unit !== -1 ? normalizeString_(row[clsCol.unit]) : '',
      term: clsCol.term !== -1 ? normalizeString_(row[clsCol.term]) : '',
      curriculumUnit: clsCol.curriculumUnit !== -1 ? normalizeString_(row[clsCol.curriculumUnit]) : '',
      allowedAbsences: clsCol.allowedAbsences !== -1 ? row[clsCol.allowedAbsences] : ''
    };
  });
  if (typeof logPerf_ === 'function') {
    logPerf_('getClassesForCurrentUserByDate build classMap', classMapStartedAt, 'rows=' + classes.length);
  }

  // classSessions は日付単位で小さくキャッシュしたものを使う
  const daySessionsStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const daySessions = getClassSessionsByDateCached_(ymd);
  if (typeof logPerf_ === 'function') {
    logPerf_('getClassesForCurrentUserByDate getClassSessionsByDateCached_', daySessionsStartedAt, 'rows=' + daySessions.length);
  }

  const resultBuildStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const result = [];

  daySessions.forEach(function(session) {
    const classId = session.classId;
    const period = session.period;
    const sessionNumber = session.sessionNumber;
    const sessionYmd = session.date;
    const weekday = session.weekday;

    const teacherKey = classId + '__' + weekday + '__' + period;
    if (!teacherSessionKeyMap[teacherKey]) {
      return;
    }

    const tt = timetableMap[teacherKey];
    if (!tt) {
      return;
    }

    const cls = classMap[classId];
    const saveInfo = savedSessionMap[[classId, sessionYmd, period].join('__')] || null;

    result.push({
      classId: classId,
      date: sessionYmd,
      period: period,
      sessionNumber: sessionNumber,
      subjectId: cls ? cls.subjectId : '',
      subjectName: cls ? cls.subjectName : '',
      grade: cls ? cls.grade : '',
      unit: cls ? cls.unit : '',
      term: cls ? cls.term : '',
      curriculumUnit: cls ? cls.curriculumUnit : '',
      allowedAbsences: cls ? cls.allowedAbsences : '',
      teacherId: tt.teacherId,
      teacherName: tt.teacherName,
      teacherIds: tt.teacherIds,
      teachers: tt.teachers,
      weekday: tt.weekday,
      isSaved: !!saveInfo,
      lastSavedInfo: saveInfo
    });
  });

  result.sort(function(a, b) {
    return Number(a.period) - Number(b.period);
  });

  if (typeof logPerf_ === 'function') {
    logPerf_('getClassesForCurrentUserByDate build result', resultBuildStartedAt, 'result=' + result.length);
    logPerf_('getClassesForCurrentUserByDate total', totalStartedAt, 'date=' + ymd);
  }

  return result;
}

/**
 * classSessions を日付単位で小さくキャッシュする
 */
function getClassSessionsByDateCached_(ymd) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const cacheKey = 'classSessionsByDate__v2__' + ymd;
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    if (typeof logPerf_ === 'function') {
      logPerf_('getClassSessionsByDateCached_ total', totalStartedAt, 'cache=hit rows=' + cached.length + ' ymd=' + ymd);
    }
    return cached;
  }

  const indexStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const byDateMap = getClassSessionsByDateIndexCached_();
  if (typeof logPerf_ === 'function') {
    logPerf_('getClassSessionsByDateCached_ getClassSessionsByDateIndexCached_', indexStartedAt, 'dates=' + Object.keys(byDateMap).length);
  }

  const result = byDateMap[ymd] || [];
  putScriptCacheJson_(cacheKey, result, 300);

  if (typeof logPerf_ === 'function') {
    logPerf_('getClassSessionsByDateCached_ total', totalStartedAt, 'cache=miss rows=' + result.length + ' ymd=' + ymd);
  }

  return result;
}

function getClassSessionsByDateIndexCached_() {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const cacheKey = 'classSessionsByDateIndex__v2';
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    if (typeof logPerf_ === 'function') {
      logPerf_('getClassSessionsByDateIndexCached_ total', totalStartedAt, 'cache=hit dates=' + Object.keys(cached).length);
    }
    return cached;
  }

  const loadStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const classSessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_SESSIONS, 300);
  if (typeof logPerf_ === 'function') {
    logPerf_('getClassSessionsByDateIndexCached_ load classSessionsData', loadStartedAt, 'rows=' + classSessionsData.rows.length);
  }

  const headerStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const classSessionHeaders = classSessionsData.headers;
  const csCol = {
    classId: findColumnIndex_(classSessionHeaders, ['classId', 'ClassID']),
    date: findColumnIndex_(classSessionHeaders, ['date', '日付']),
    period: findColumnIndex_(classSessionHeaders, ['period', '時限']),
    sessionNumber: findColumnIndex_(classSessionHeaders, ['sessionNumber', '回', '回数'])
  };
  validateRequiredColumnsForTimetable_('classSessions', csCol, ['classId', 'date', 'period']);
  if (typeof logPerf_ === 'function') {
    logPerf_('getClassSessionsByDateIndexCached_ resolve headers', headerStartedAt);
  }

  const buildStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const byDateMap = {};

  classSessionsData.rows.forEach(function(row) {
    const rowYmd = formatDateToYmd(row[csCol.date]);
    if (!rowYmd) return;

    if (!byDateMap[rowYmd]) {
      const weekday = getWeekdayFromYmdJst_(rowYmd);
      byDateMap[rowYmd] = {
        _weekday: weekday,
        _rows: []
      };
    }

    byDateMap[rowYmd]._rows.push({
      classId: normalizeString_(row[csCol.classId]),
      date: rowYmd,
      period: normalizeString_(row[csCol.period]),
      sessionNumber: csCol.sessionNumber !== -1 ? row[csCol.sessionNumber] : '',
      weekday: byDateMap[rowYmd]._weekday
    });
  });

  const normalizedMap = {};
  Object.keys(byDateMap).forEach(function(dateKey) {
    normalizedMap[dateKey] = byDateMap[dateKey]._rows;
  });

  if (typeof logPerf_ === 'function') {
    logPerf_('getClassSessionsByDateIndexCached_ build index', buildStartedAt, 'dates=' + Object.keys(normalizedMap).length);
  }

  putScriptCacheJson_(cacheKey, normalizedMap, 300);

  if (typeof logPerf_ === 'function') {
    logPerf_('getClassSessionsByDateIndexCached_ total', totalStartedAt, 'cache=miss dates=' + Object.keys(normalizedMap).length);
  }

  return normalizedMap;
}

/**
 * attendanceSessions を日付単位で小さくキャッシュする
 * savedByCurrentUser は呼び出し側で付与する
 */
function getSavedSessionMapByDateCached_(ymd) {
  const cacheKey = 'savedSessionMapByDate__' + ymd;
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const attendanceSessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.ATTENDANCE_SESSIONS, 60);
  const headers = attendanceSessionsData.headers;
  const rows = attendanceSessionsData.rows;

  const asCol = {
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    date: findColumnIndex_(headers, ['date', '日付']),
    period: findColumnIndex_(headers, ['period', '時限']),
    teacherEmail: findColumnIndex_(headers, ['teacherEmail', 'email']),
    accessedAt: findColumnIndex_(headers, ['accessedAt', 'savedAt']),
    actionType: findColumnIndex_(headers, ['actionType']),
    targetSessionKey: findColumnIndex_(headers, ['targetSessionKey']),
    savedModeLabel: findColumnIndex_(headers, ['savedModeLabel'])
  };
  validateRequiredColumnsForTimetable_('attendanceSessions', asCol, ['classId', 'date', 'period']);

  const map = {};

  rows.forEach(function(row) {
    const rowDate = formatDateToYmd(row[asCol.date]);
    if (rowDate !== ymd) return;

    const rowClassId = normalizeString_(row[asCol.classId]);
    const rowPeriod = normalizeString_(row[asCol.period]);
    if (!rowClassId || !rowPeriod) return;

    const key = [rowClassId, rowDate, rowPeriod].join('__');
    const teacherEmail = asCol.teacherEmail !== -1
      ? normalizeString_(row[asCol.teacherEmail]).toLowerCase()
      : '';

    const accessedAtRaw = asCol.accessedAt !== -1 ? row[asCol.accessedAt] : '';
    const accessedAt = accessedAtRaw instanceof Date ? accessedAtRaw : new Date(accessedAtRaw);
    const accessedAtMs = isNaN(accessedAt.getTime()) ? 0 : accessedAt.getTime();

    if (!map[key] || accessedAtMs >= map[key]._ms) {
      map[key] = {
        teacherEmail: teacherEmail,
        savedAtText: formatDateTimeJst_(accessedAtRaw),
        actionType: asCol.actionType !== -1 ? normalizeString_(row[asCol.actionType]) : '',
        targetSessionKey: asCol.targetSessionKey !== -1 ? normalizeString_(row[asCol.targetSessionKey]) : '',
        savedModeLabel: asCol.savedModeLabel !== -1 ? normalizeString_(row[asCol.savedModeLabel]) : '',
        _ms: accessedAtMs
      };
    }
  });

  putScriptCacheJson_(cacheKey, map, 60);
  return map;
}

function getSavedSessionKeySetByRangeCached_(startYmd, endYmd) {
  const cacheKey = 'savedSessionKeySetByRange__v2__' + String(startYmd || '') + '__' + String(endYmd || '');
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const attendanceSessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.ATTENDANCE_SESSIONS, 60);
  const headers = attendanceSessionsData.headers;
  const rows = attendanceSessionsData.rows;

  const asCol = {
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    date: findColumnIndex_(headers, ['date', '日付']),
    period: findColumnIndex_(headers, ['period', '時限'])
  };
  validateRequiredColumnsForTimetable_('attendanceSessions', asCol, ['classId', 'date', 'period']);

  const keySet = {};

  rows.forEach(function(row) {
    const rowDate = formatDateToYmd(row[asCol.date]);
    if (!rowDate || rowDate < startYmd || rowDate > endYmd) return;

    const classId = normalizeString_(row[asCol.classId]);
    const period = normalizeString_(row[asCol.period]);
    if (!classId || !period) return;

    const key = [classId, rowDate, period].join('__');
    keySet[key] = true;
  });

  putScriptCacheJson_(cacheKey, keySet, 60);
  return keySet;
}

function getWeekdayFromDate_(value) {
  const date = value instanceof Date ? value : new Date(value);
  return Utilities.formatDate(date, 'Asia/Tokyo', 'EEE'); // Sun, Mon, Tue...
}

function getTodayClassesForCurrentUser(targetDate) {
  const totalStartedAt = perfNow_();
  const result = getClassesForCurrentUserByDate(targetDate);
  logPerf_('getTodayClassesForCurrentUser total', totalStartedAt, 'result=' + (Array.isArray(result) ? result.length : 0));
  return result;
}

function getTeacherUnsavedSummary() {
  const user = getCurrentUserContext();
  if (!user || !user.teacherId) {
    throw new Error('ログインユーザー情報を取得できませんでした。');
  }

  const teacherId = normalizeString_(user.teacherId);

  // JST基準の今日を先に確定する
  const todayYmd = formatDateToYmd(new Date());
  const today = new Date(todayYmd + 'T12:00:00+09:00');

  const endDate = new Date(today);
  endDate.setDate(endDate.getDate() - 1);

  const startDate = getTeacherUnsavedStartDate_(today);
  const startYmd = formatDateToYmd(startDate);
  const endYmd = formatDateToYmd(endDate);

  const cacheKey = buildTeacherUnsavedSummaryCacheKey_(teacherId, endYmd);
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  if (!startYmd || !endYmd || endYmd < startYmd) {
    const emptyResult = {
      ok: true,
      count: 0,
      checkedRange: {
        start: startYmd || '',
        end: endYmd || ''
      }
    };
    putScriptCacheJson_(cacheKey, emptyResult, 60);
    return emptyResult;
  }

  const count = getTeacherUnsavedCount_(teacherId, startYmd, endYmd);

  const result = {
    ok: true,
    count: count,
    checkedRange: {
      start: startYmd,
      end: endYmd
    }
  };

  putScriptCacheJson_(cacheKey, result, 60);
  return result;
}

function getTeacherUnsavedDetails() {
  const user = getCurrentUserContext();
  if (!user || !user.teacherId) {
    throw new Error('ログインユーザー情報を取得できませんでした。');
  }

  const teacherId = normalizeString_(user.teacherId);

  // JST基準の今日を先に確定する
  const todayYmd = formatDateToYmd(new Date());
  const today = new Date(todayYmd + 'T12:00:00+09:00');

  const endDate = new Date(today);
  endDate.setDate(endDate.getDate() - 1);

  const startDate = getTeacherUnsavedStartDate_(today);
  const startYmd = formatDateToYmd(startDate);
  const endYmd = formatDateToYmd(endDate);

  const cacheKey = buildTeacherUnsavedDetailsCacheKey_(teacherId, endYmd);
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  if (!startYmd || !endYmd || endYmd < startYmd) {
    const emptyResult = {
      ok: true,
      items: [],
      checkedRange: {
        start: startYmd || '',
        end: endYmd || ''
      }
    };
    putScriptCacheJson_(cacheKey, emptyResult, 60);
    return emptyResult;
  }

  const items = getTeacherUnsavedSessionItems_(teacherId, startYmd, endYmd);

  const result = {
    ok: true,
    items: items,
    checkedRange: {
      start: startYmd,
      end: endYmd
    }
  };

  putScriptCacheJson_(cacheKey, result, 60);
  return result;
}

function getTeacherUnsavedCount_(teacherId, startYmd, endYmd) {
  if (!teacherId || !startYmd || !endYmd || endYmd < startYmd) {
    return 0;
  }

  const context = getTeacherUnsavedContext_(teacherId);
  const byDateMap = getClassSessionsByDateIndexCached_();
  const savedKeySet = getSavedSessionKeySetByRangeCached_(startYmd, endYmd);
  const dateKeys = Object.keys(byDateMap).sort();

  let count = 0;
  const seenSessionKeys = {};

  dateKeys.forEach(function(ymd) {
    if (ymd < startYmd || ymd > endYmd) return;

    const daySessions = byDateMap[ymd] || [];
    daySessions.forEach(function(session) {
      const classId = normalizeString_(session.classId);
      const period = normalizeString_(session.period);
      const weekday = normalizeWeekday_(session.weekday);
      const teacherKey = [classId, weekday, period].join('__');

      if (!context.teacherSessionKeyMap[teacherKey]) return;

      const saveKey = [classId, ymd, period].join('__');
      if (savedKeySet[saveKey]) return;

      const cls = context.classMap[classId] || {};
      const displayKey = buildTeacherUnsavedDisplayKey_(cls, classId, ymd, period);

      if (seenSessionKeys[displayKey]) return;
      seenSessionKeys[displayKey] = true;

      count += 1;
    });
  });

  return count;
}

function getTeacherUnsavedSessionItems_(teacherId, startYmd, endYmd) {
  if (!teacherId || !startYmd || !endYmd || endYmd < startYmd) {
    return [];
  }

  const context = getTeacherUnsavedContext_(teacherId);
  const byDateMap = getClassSessionsByDateIndexCached_();
  const savedKeySet = getSavedSessionKeySetByRangeCached_(startYmd, endYmd);
  const dateKeys = Object.keys(byDateMap).sort();
  const result = [];
  const seenSessionKeys = {};

  dateKeys.forEach(function(ymd) {
    if (ymd < startYmd || ymd > endYmd) return;

    const daySessions = byDateMap[ymd] || [];

    daySessions.forEach(function(session) {
      const classId = normalizeString_(session.classId);
      const period = normalizeString_(session.period);
      const weekday = normalizeWeekday_(session.weekday);
      const teacherKey = [classId, weekday, period].join('__');

      if (!context.teacherSessionKeyMap[teacherKey]) return;

      const saveKey = [classId, ymd, period].join('__');
      if (savedKeySet[saveKey]) return;

      const cls = context.classMap[classId] || {};
      const displayKey = buildTeacherUnsavedDisplayKey_(cls, classId, ymd, period);

      if (seenSessionKeys[displayKey]) return;
      seenSessionKeys[displayKey] = true;

      const isExperiment = isTeacherUnsavedExperimentClass_(classId);

      let targetLabel = '';
      if (isExperiment) {
        targetLabel = '班選択して記録';
      } else {
        const gradeText = cls.grade ? String(cls.grade) + '年' : '';
        const unitText = cls.unit ? String(cls.unit) + '組' : '';
        targetLabel = (gradeText + ' ' + unitText).trim();
      }

      result.push({
        classId: classId,
        date: ymd,
        period: period,
        sessionNumber: session.sessionNumber || '',
        subjectName: cls.subjectName || '',
        targetLabel: targetLabel
      });
    });
  });

  result.sort(function(a, b) {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    return Number(a.period) - Number(b.period);
  });

  return result;
}

function getTeacherUnsavedContext_(teacherId) {
  const cacheKey = buildTeacherUnsavedContextCacheKey_(teacherId);
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const timetableData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.TIMETABLE, 300);
  const classesData = getSheetDataCached_('MASTER', CONFIG.SHEETS.CLASSES, 300);
  const teamData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_TEACHER_TEAMS, 300);

  const timetable = timetableData.rows;
  const classes = classesData.rows;
  const teamRows = teamData.rows;

  const timetableHeaders = timetableData.headers;
  const ttCol = {
    classId: findColumnIndex_(timetableHeaders, ['classId', 'ClassID']),
    weekday: findColumnIndex_(timetableHeaders, ['weekday', '曜日']),
    period: findColumnIndex_(timetableHeaders, ['period', '時限']),
    teacherName: findColumnIndex_(timetableHeaders, ['teacherName', '担当者名', 'name']),
    teacherId: findColumnIndex_(timetableHeaders, ['teacherId', 'TeacherID'])
  };
  validateRequiredColumnsForTimetable_('timetable', ttCol, ['classId', 'weekday', 'period']);

  const teamHeaders = teamData.headers;
  const teamCol = {
    classId: findColumnIndex_(teamHeaders, ['classId', 'ClassID']),
    weekday: findColumnIndex_(teamHeaders, ['weekday', '曜日']),
    period: findColumnIndex_(teamHeaders, ['period', '時限']),
    teacherName: findColumnIndex_(teamHeaders, ['teacherName', '担当者名', 'name']),
    teacherId: findColumnIndex_(teamHeaders, ['teacherId', 'TeacherID']),
    roleType: findColumnIndex_(teamHeaders, ['roleType', '役割'])
  };

  const timetableMap = {};

  timetable.forEach(function(row) {
    const classId = normalizeString_(row[ttCol.classId]);
    const period = normalizeString_(row[ttCol.period]);
    const weekday = ttCol.weekday !== -1 ? normalizeWeekday_(row[ttCol.weekday]) : '';
    const teacherName = ttCol.teacherName !== -1 ? normalizeString_(row[ttCol.teacherName]) : '';
    let teacherIdInRow = ttCol.teacherId !== -1 ? normalizeString_(row[ttCol.teacherId]) : '';

    if (!teacherIdInRow && teacherName) {
      const teacher = getTeacherRecordByName_(teacherName);
      teacherIdInRow = teacher ? teacher.teacherId : '';
    }

    if (!classId || !period || !weekday) return;

    const key = classId + '__' + weekday + '__' + period;
    timetableMap[key] = {
      classId: classId,
      weekday: weekday,
      period: period,
      teacherId: teacherIdInRow,
      teacherName: teacherName,
      teacherIds: teacherIdInRow ? [teacherIdInRow] : []
    };
  });

  teamRows.forEach(function(row) {
    const classId = teamCol.classId !== -1 ? normalizeString_(row[teamCol.classId]) : '';
    const period = teamCol.period !== -1 ? normalizeString_(row[teamCol.period]) : '';
    const weekday = teamCol.weekday !== -1 ? normalizeWeekday_(row[teamCol.weekday]) : '';
    const teacherName = teamCol.teacherName !== -1 ? normalizeString_(row[teamCol.teacherName]) : '';
    let teacherIdInRow = teamCol.teacherId !== -1 ? normalizeString_(row[teamCol.teacherId]) : '';

    if (!teacherIdInRow && teacherName) {
      const teacher = getTeacherRecordByName_(teacherName);
      teacherIdInRow = teacher ? teacher.teacherId : '';
    }

    if (!classId || !period || !weekday || !teacherIdInRow) return;

    const key = classId + '__' + weekday + '__' + period;
    if (!timetableMap[key]) {
      timetableMap[key] = {
        classId: classId,
        weekday: weekday,
        period: period,
        teacherId: '',
        teacherName: '',
        teacherIds: []
      };
    }

    if (!timetableMap[key].teacherIds.includes(teacherIdInRow)) {
      timetableMap[key].teacherIds.push(teacherIdInRow);
    }
  });

  const teacherSessionKeyMap = {};
  Object.keys(timetableMap).forEach(function(key) {
    const tt = timetableMap[key];
    if (tt && Array.isArray(tt.teacherIds) && tt.teacherIds.includes(teacherId)) {
      teacherSessionKeyMap[key] = true;
    }
  });

  const classHeaders = classesData.headers;
  const clsCol = {
    classId: findColumnIndex_(classHeaders, ['classId', 'ClassID']),
    subjectId: findColumnIndex_(classHeaders, ['subjectId', 'SubjectID']),
    subjectName: findColumnIndex_(classHeaders, ['subjectName', '科目名']),
    grade: findColumnIndex_(classHeaders, ['grade', '学年']),
    unit: findColumnIndex_(classHeaders, ['unit', '対象区分', '組・コース'])
  };
  validateRequiredColumnsForTimetable_('classes', clsCol, ['classId', 'subjectId', 'subjectName']);

  const classMap = {};
  classes.forEach(function(row) {
    const classId = normalizeString_(row[clsCol.classId]);
    if (!classId) return;

  classMap[classId] = {
    classId: classId,
    subjectId: clsCol.subjectId !== -1 ? normalizeString_(row[clsCol.subjectId]) : '',
    subjectName: clsCol.subjectName !== -1 ? normalizeString_(row[clsCol.subjectName]) : '',
    grade: clsCol.grade !== -1 ? normalizeString_(row[clsCol.grade]) : '',
    unit: clsCol.unit !== -1 ? normalizeString_(row[clsCol.unit]) : ''
  };
  });

  const result = {
    timetableMap: timetableMap,
    teacherSessionKeyMap: teacherSessionKeyMap,
    classMap: classMap
  };

  putScriptCacheJson_(cacheKey, result, 300);
  return result;
}

function buildTeacherUnsavedSummaryCacheKey_(teacherId, endYmd) {
  return 'teacherUnsavedSummary__v2__' + String(teacherId || '') + '__' + String(endYmd || '');
}

function buildTeacherUnsavedDetailsCacheKey_(teacherId, endYmd) {
  return 'teacherUnsavedDetails__v2__' + String(teacherId || '') + '__' + String(endYmd || '');
}

function buildTeacherUnsavedContextCacheKey_(teacherId) {
  return 'teacherUnsavedContext__v2__' + String(teacherId || '');
}

function getTeacherUnsavedStartDate_(baseDate) {
  const d = new Date(baseDate);
  d.setHours(0, 0, 0, 0);

  const schoolYear = d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1;
  const start = new Date(schoolYear, 3, 7);
  start.setHours(0, 0, 0, 0);
  return start;
}

function isTeacherUnsavedExperimentClass_(classId) {
  const value = String(classId || '').trim();
  return (
    value.indexOf('工学実験実習1') !== -1 ||
    value.indexOf('工学実験実習2') !== -1
  );
}

function buildTeacherUnsavedDisplayKey_(cls, classId, ymd, period) {
  const normalizedClassId = normalizeString_(classId);
  const normalizedYmd = normalizeString_(ymd);
  const normalizedPeriod = normalizeString_(period);
  const subjectId = cls && cls.subjectId ? normalizeString_(cls.subjectId) : '';

  if (isTeacherUnsavedExperimentClass_(normalizedClassId) && subjectId) {
    return [subjectId, normalizedYmd, normalizedPeriod].join('__');
  }

  return [normalizedClassId, normalizedYmd, normalizedPeriod].join('__');
}

function testGetTodayClassesForCurrentUser() {
  const result = getTodayClassesForCurrentUser('2026-04-06');
  Logger.log(JSON.stringify(result, null, 2));
}

function debugTodayClassMatching(targetDate) {
  const user = getCurrentUserContext();
  Logger.log('currentUser=' + JSON.stringify(user, null, 2));

  const timezone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const today = targetDate
    ? formatDateToYmd(targetDate)
    : Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');

  const classSessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_SESSIONS, 60);
  const timetableData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.TIMETABLE, 60);
  const classesData = getSheetDataCached_('MASTER', CONFIG.SHEETS.CLASSES, 60);

  Logger.log('today=' + today);
  Logger.log('classSessionsHeaders=' + JSON.stringify(classSessionsData.headers));
  Logger.log('timetableHeaders=' + JSON.stringify(timetableData.headers));
  Logger.log('classesHeaders=' + JSON.stringify(classesData.headers));

  const classSessionHeaders = classSessionsData.headers;
  const csDateCol = findColumnIndex_(classSessionHeaders, ['date', '日付']);
  if (csDateCol === -1) {
    throw new Error('classSessions の date 列が見つかりません');
  }

  const filtered = classSessionsData.rows.filter(function(row) {
    return formatDateToYmd(row[csDateCol]) === today;
  });

  Logger.log('filteredClassSessions=' + JSON.stringify(filtered));
}

function validateRequiredColumnsForTimetable_(sheetName, colMap, requiredKeys) {
  requiredKeys.forEach(function(key) {
    if (colMap[key] === -1) {
      throw new Error(sheetName + ' シートに必要な列がありません: ' + key);
    }
  });
}

function debugTeacherClassesByDate() {
  targetDate = '2026-04-09';
  const user = getCurrentUserContext();
  Logger.log('user=' + JSON.stringify(user, null, 2));

  const currentTeacherId = normalizeString_(user.teacherId);
  const ymd = formatDateToYmd(targetDate);

  const classSessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_SESSIONS, 60);
  const timetableData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.TIMETABLE, 60);
  const teamData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_TEACHER_TEAMS, 60);

  const classSessions = classSessionsData.rows;
  const timetable = timetableData.rows;
  const teamRows = teamData.rows;

  const ttCol = {
    classId: findColumnIndex_(timetableData.headers, ['classId', 'ClassID']),
    weekday: findColumnIndex_(timetableData.headers, ['weekday', '曜日']),
    period: findColumnIndex_(timetableData.headers, ['period', '時限']),
    teacherId: findColumnIndex_(timetableData.headers, ['teacherId', 'TeacherID']),
    teacherName: findColumnIndex_(timetableData.headers, ['teacherName', '担当者名', 'name'])
  };

  const teamCol = {
    classId: findColumnIndex_(teamData.headers, ['classId', 'ClassID']),
    weekday: findColumnIndex_(teamData.headers, ['weekday', '曜日']),
    period: findColumnIndex_(teamData.headers, ['period', '時限']),
    teacherId: findColumnIndex_(teamData.headers, ['teacherId', 'TeacherID']),
    teacherName: findColumnIndex_(teamData.headers, ['teacherName', '担当者名', 'name']),
    roleType: findColumnIndex_(teamData.headers, ['roleType', '役割'])
  };

  const csCol = {
    classId: findColumnIndex_(classSessionsData.headers, ['classId', 'ClassID']),
    date: findColumnIndex_(classSessionsData.headers, ['date', '日付']),
    period: findColumnIndex_(classSessionsData.headers, ['period', '時限'])
  };

  const weekday = getWeekdayFromYmdJst_(ymd);
  Logger.log('targetDate=' + ymd + ', weekday=' + weekday + ', teacherId=' + currentTeacherId);

  const targetSessions = classSessions.filter(row => formatDateToYmd(row[csCol.date]) === ymd);
  Logger.log('targetSessions=' + JSON.stringify(targetSessions, null, 2));

  targetSessions.forEach(function(row) {
    const classId = normalizeString_(row[csCol.classId]);
    const period = normalizeString_(row[csCol.period]);

    const timetableHit = timetable.filter(t =>
      normalizeString_(t[ttCol.classId]) === classId &&
      normalizeWeekday_(t[ttCol.weekday]) === weekday &&
      normalizeString_(t[ttCol.period]) === period
    );

    const teamHit = teamRows.filter(t =>
      normalizeString_(t[teamCol.classId]) === classId &&
      normalizeWeekday_(t[teamCol.weekday]) === weekday &&
      normalizeString_(t[teamCol.period]) === period
    );

    Logger.log('---');
    Logger.log('session classId=' + classId + ', period=' + period);
    Logger.log('timetableHit=' + JSON.stringify(timetableHit, null, 2));
    Logger.log('teamHit=' + JSON.stringify(teamHit, null, 2));
  });
}

function getWeekdayFromYmdJst_(ymd) {
  const m = String(ymd || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';

  // JSTの正午を明示して、実行環境のタイムゾーン影響を受けないようにする
  const date = new Date(ymd + 'T12:00:00+09:00');
  const weekday = Utilities.formatDate(date, 'Asia/Tokyo', 'E'); // Mon, Tue, Wed...

  return normalizeWeekday_(weekday);
}