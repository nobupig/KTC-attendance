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

  const classes = classesData.rows;
  const classDayContext = getEffectiveClassDayContext_(ymd);
  if (!classDayContext.isClassDay || !classDayContext.effectiveWeekday) {
    return [];
  }
  // 年間累積 timetable では date の有効termで解決する。index はこのリクエスト中で共有する。
  const assignmentIndex = buildTeachingAssignmentIndex_(
    timetableData,
    teamData,
    buildTeacherTeamMember_
  );

  const savedSessionStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  // 初期表示高速化：授業一覧では保存済み判定を読まない。
  // 保存状態は「欠席者を入力」で開いた時、または保存後に反映する。
  const savedSessionMap = {};

  if (typeof logPerf_ === 'function') {
    logPerf_('getClassesForCurrentUserByDate build savedSessionMap', savedSessionStartedAt, 'skip-initial-list');
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
    const tt = getTeachingAssignmentForSessionFromIndex_(
      assignmentIndex,
      classId,
      sessionYmd,
      period,
      classDayContext
    );
    if (!tt || !tt.teacherIds.includes(currentTeacherId)) {
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
      lastSavedInfo: saveInfo,
      saveStatusNotLoaded: true
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
 * classSessions を指定日だけ読み込む軽量版
 * 本日の授業一覧では、全日付インデックスを作らず対象日のみ抽出する
 */
function getClassSessionsByDateCached_(ymd) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const targetYmd = formatDateToYmd(ymd);
  const classDayInfo = getEffectiveClassDayInfo_(targetYmd);

  if (!classDayInfo.isClassDay || !classDayInfo.weekday) {
    if (typeof logPerf_ === 'function') {
      logPerf_(
        'getClassSessionsByDateCached_ total',
        totalStartedAt,
        'non-class-day ymd=' + targetYmd
      );
    }
    return [];
  }

  const cacheKey = buildClassSessionsByDateCacheKey_(targetYmd);

  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    if (typeof logPerf_ === 'function') {
      logPerf_(
        'getClassSessionsByDateCached_ total',
        totalStartedAt,
        'cache=hit rows=' + cached.length + ' ymd=' + targetYmd
      );
    }
    return cached.map(function(session) {
      return Object.assign({}, session, { weekday: classDayInfo.weekday });
    });
  }

  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CLASS_SESSIONS);
  if (!sheet) {
    throw new Error('classSessions シートが見つかりません');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1 || lastCol <= 0) {
    putScriptCacheJson_(cacheKey, [], 300);
    return [];
  }

  const headerStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const csCol = {
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    date: findColumnIndex_(headers, ['date', '日付']),
    period: findColumnIndex_(headers, ['period', '時限']),
    sessionNumber: findColumnIndex_(headers, ['sessionNumber', '回', '回数'])
  };

  validateRequiredColumnsForTimetable_('classSessions', csCol, ['classId', 'date', 'period']);

  if (typeof logPerf_ === 'function') {
    logPerf_('getClassSessionsByDateCached_ resolve headers', headerStartedAt);
  }

  const loadStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const numRows = lastRow - 1;

  const values = sheet.getRange(2, 1, numRows, lastCol).getValues();

  // 日付列だけは表示値で読む。これが今回の安全ポイント。
  const dateDisplayValues = sheet
    .getRange(2, csCol.date + 1, numRows, 1)
    .getDisplayValues();

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getClassSessionsByDateCached_ load sheet direct',
      loadStartedAt,
      'rows=' + numRows
    );
  }

  const buildStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const weekday = classDayInfo.weekday;
  const result = [];

  values.forEach(function(row, index) {
    const rowYmd = normalizeYmdDisplayText_(dateDisplayValues[index][0]);
    if (rowYmd !== targetYmd) return;

    result.push({
      classId: normalizeString_(row[csCol.classId]),
      date: rowYmd,
      period: normalizeString_(row[csCol.period]),
      sessionNumber: csCol.sessionNumber !== -1 ? row[csCol.sessionNumber] : '',
      weekday: weekday
    });
  });

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getClassSessionsByDateCached_ build result',
      buildStartedAt,
      'rows=' + result.length + ' ymd=' + targetYmd
    );
  }

  putScriptCacheJson_(cacheKey, result, 300);

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getClassSessionsByDateCached_ total',
      totalStartedAt,
      'cache=miss rows=' + result.length + ' ymd=' + targetYmd
    );
  }

  return result;
}

function getClassSessionsByDateIndexCached_() {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const cacheKey = getClassSessionsByDateIndexCacheKey_();
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    if (typeof logPerf_ === 'function') {
      logPerf_(
        'getClassSessionsByDateIndexCached_ total',
        totalStartedAt,
        'cache=hit dates=' + Object.keys(cached).length
      );
    }
    return applyEffectiveClassDayInfoToSessionIndex_(cached, getEffectiveClassDayIndex_());
  }

  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CLASS_SESSIONS);
  if (!sheet) {
    throw new Error('classSessions シートが見つかりません');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1 || lastCol <= 0) {
    return {};
  }

  const headerStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const csCol = {
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    date: findColumnIndex_(headers, ['date', '日付']),
    period: findColumnIndex_(headers, ['period', '時限']),
    sessionNumber: findColumnIndex_(headers, ['sessionNumber', '回', '回数'])
  };

  validateRequiredColumnsForTimetable_('classSessions', csCol, ['classId', 'date', 'period']);

  if (typeof logPerf_ === 'function') {
    logPerf_('getClassSessionsByDateIndexCached_ resolve headers', headerStartedAt);
  }

  const loadStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const numRows = lastRow - 1;

  const values = sheet.getRange(2, 1, numRows, lastCol).getValues();

  // 日付列だけ表示値で読む。Utilities.formatDate の大量実行を避ける。
  const dateDisplayValues = sheet
    .getRange(2, csCol.date + 1, numRows, 1)
    .getDisplayValues();

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getClassSessionsByDateIndexCached_ load sheet direct',
      loadStartedAt,
      'rows=' + numRows
    );
  }

  const buildStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const byDateMap = {};

  values.forEach(function(row, index) {
    const rowYmd = normalizeYmdDisplayText_(dateDisplayValues[index][0]);
    if (!rowYmd) return;

    if (!byDateMap[rowYmd]) {
      byDateMap[rowYmd] = [];
    }

    byDateMap[rowYmd].push({
      classId: normalizeString_(row[csCol.classId]),
      date: rowYmd,
      period: normalizeString_(row[csCol.period]),
      sessionNumber: csCol.sessionNumber !== -1 ? row[csCol.sessionNumber] : '',
      weekday: ''
    });
  });

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getClassSessionsByDateIndexCached_ build index',
      buildStartedAt,
      'dates=' + Object.keys(byDateMap).length
    );
  }

  // サイズが大きい場合は内部で skip されるが、入ればラッキー程度。
  putScriptCacheJson_(cacheKey, byDateMap, 300);

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getClassSessionsByDateIndexCached_ total',
      totalStartedAt,
      'cache=miss dates=' + Object.keys(byDateMap).length
    );
  }

  return applyEffectiveClassDayInfoToSessionIndex_(byDateMap, getEffectiveClassDayIndex_());
}

function applyEffectiveClassDayInfoToSessionIndex_(byDateMap, calendarIndex) {
  const result = byDateMap || {};

  Object.keys(result).forEach(function(ymd) {
    const classDayInfo = getEffectiveClassDayInfo_(ymd, calendarIndex);

    if (!classDayInfo.isClassDay || !classDayInfo.weekday) {
      delete result[ymd];
      return;
    }

    (result[ymd] || []).forEach(function(session) {
      session.weekday = classDayInfo.weekday;
    });
  });

  return result;
}


function getTimetableExperimentSubjectIdByClassId_(classId) {
  const targetClassId = normalizeString_(classId);
  if (!targetClassId) return '';

  // 高速化のため、保存済み判定・未保存判定では classes 参照を行わない。
  // この関数は attendanceSessions の全行スキャン中に大量実行されるため、
  // getClassRecordById_ を呼ぶと極端に遅くなる。
  if (targetClassId.indexOf('工学実験実習1') !== -1) {
    return 'G1_G_工学実験実習1_FY';
  }

  if (targetClassId.indexOf('工学実験実習2') !== -1) {
    return 'G2_G_工学実験実習2_FY';
  }

  return '';
}

function buildExperimentSessionKeyForTimetable_(classId, date, period) {
  const subjectId = getTimetableExperimentSubjectIdByClassId_(classId);
  const ymd = formatDateToYmd(date);
  const targetPeriod = normalizeString_(period);

  if (!subjectId || !ymd || !targetPeriod) return '';
  return [subjectId, ymd, targetPeriod].join('__');
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
    const rowDate = fastYmdFromCell_(row[asCol.date]);
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
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const normalizedStartYmd = formatDateToYmd(startYmd);
  const normalizedEndYmd = formatDateToYmd(endYmd);

  const cacheKey =
    'savedSessionKeySetByRange__v4__' +
    String(normalizedStartYmd || '') +
    '__' +
    String(normalizedEndYmd || '');

  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    if (typeof logPerf_ === 'function') {
      logPerf_(
        'getSavedSessionKeySetByRangeCached_ total',
        totalStartedAt,
        'cache=hit keys=' + Object.keys(cached).length
      );
    }
    return cached;
  }

  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_SESSIONS);
  if (!sheet) {
    throw new Error('attendanceSessions シートが見つかりません');
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow <= 1 || lastCol <= 0) {
    return {};
  }

  const headerStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

  const asCol = {
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    date: findColumnIndex_(headers, ['date', '日付']),
    period: findColumnIndex_(headers, ['period', '時限'])
  };

  validateRequiredColumnsForTimetable_('attendanceSessions', asCol, ['classId', 'date', 'period']);

  if (typeof logPerf_ === 'function') {
    logPerf_('getSavedSessionKeySetByRangeCached_ resolve headers', headerStartedAt);
  }

  const loadStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const numRows = lastRow - 1;

  const values = sheet.getRange(2, 1, numRows, lastCol).getValues();

  // attendanceSessions の日付も表示値で読む
  const dateDisplayValues = sheet
    .getRange(2, asCol.date + 1, numRows, 1)
    .getDisplayValues();

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getSavedSessionKeySetByRangeCached_ load sheet direct',
      loadStartedAt,
      'rows=' + numRows
    );
  }

  const buildStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const keySet = {};

  values.forEach(function(row, index) {
    const rowDate = normalizeYmdDisplayText_(dateDisplayValues[index][0]);

    if (!rowDate || rowDate < normalizedStartYmd || rowDate > normalizedEndYmd) {
      return;
    }

    const classId = normalizeString_(row[asCol.classId]);
    const period = normalizeString_(row[asCol.period]);

    if (!classId || !period) return;

    const key = [classId, rowDate, period].join('__');
    keySet[key] = true;

    // 工学実験は G2_1/G2_2/G2_3... のどれに保存されても、同一 subjectId・日付・時限として保存済み扱いにする
    const experimentKey = buildExperimentSessionKeyForTimetable_(classId, rowDate, period);
    if (experimentKey) {
      keySet[experimentKey] = true;
    }
  });

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getSavedSessionKeySetByRangeCached_ build keySet',
      buildStartedAt,
      'keys=' + Object.keys(keySet).length
    );
  }

  putScriptCacheJson_(cacheKey, keySet, 60);

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getSavedSessionKeySetByRangeCached_ total',
      totalStartedAt,
      'cache=miss keys=' + Object.keys(keySet).length
    );
  }

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
  const effectiveClassDayIndex = getEffectiveClassDayIndex_();
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
      const classDayContext = getEffectiveClassDayContext_(ymd, effectiveClassDayIndex);
      const assignment = getTeachingAssignmentForSessionFromIndex_(context.assignmentIndex, classId, ymd, period, classDayContext);
      if (!assignment || !assignment.teacherIds.includes(teacherId)) return;

      const cls = context.classMap[classId] || {};
      const displayKey = buildTeacherUnsavedDisplayKey_(cls, classId, ymd, period);
      const saveKey = [classId, ymd, period].join('__');

      // 工学実験は同一 subjectId・日付・時限の保存ログでも保存済み扱いにする
      if (savedKeySet[saveKey] || savedKeySet[displayKey]) return;

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
  const effectiveClassDayIndex = getEffectiveClassDayIndex_();
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
      const classDayContext = getEffectiveClassDayContext_(ymd, effectiveClassDayIndex);
      const assignment = getTeachingAssignmentForSessionFromIndex_(context.assignmentIndex, classId, ymd, period, classDayContext);
      if (!assignment || !assignment.teacherIds.includes(teacherId)) return;

      const cls = context.classMap[classId] || {};
      const displayKey = buildTeacherUnsavedDisplayKey_(cls, classId, ymd, period);
      const saveKey = [classId, ymd, period].join('__');

      // 工学実験は同一 subjectId・日付・時限の保存ログでも保存済み扱いにする
      if (savedKeySet[saveKey] || savedKeySet[displayKey]) return;

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

  const classes = classesData.rows;
  const assignmentIndex = buildTeachingAssignmentIndex_(
    timetableData,
    teamData,
    buildTeacherTeamMember_
  );

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
    assignmentIndex: assignmentIndex,
    classMap: classMap
  };

  putScriptCacheJson_(cacheKey, result, 300);
  return result;
}

function buildTeacherUnsavedSummaryCacheKey_(teacherId, endYmd) {
  return 'teacherUnsavedSummary__v3__' + String(teacherId || '') + '__' + String(endYmd || '');
}

function buildTeacherUnsavedDetailsCacheKey_(teacherId, endYmd) {
  return 'teacherUnsavedDetails__v3__' + String(teacherId || '') + '__' + String(endYmd || '');
}

function buildTeacherUnsavedContextCacheKey_(teacherId) {
  // v4 adds the term-aware assignment index; v3 values have no assignmentIndex.
  return 'teacherUnsavedContext__v4__' + String(teacherId || '');
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

  const classDayInfo = getEffectiveClassDayInfo_(ymd);
  const weekday = classDayInfo.weekday;
  Logger.log(
    'targetDate=' + ymd +
    ', effectiveWeekday=' + weekday +
    ', isClassDay=' + classDayInfo.isClassDay +
    ', calendarEntry=' + classDayInfo.hasCalendarEntry +
    ', teacherId=' + currentTeacherId
  );

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

function getSaveStatusForTeacherSessions(sessionItems) {
  return getSaveStatusForTeacherSessionsDirect_(sessionItems);
}

function getSaveStatusForTeacherSessionsDirect_(sessionItems) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const items = Array.isArray(sessionItems) ? sessionItems : [];
  const result = {};
  const targetKeySet = {};
  const targetKeys = [];
  const targetExperimentKeyMap = {};

  items.forEach(function(item) {
    const classId = normalizeString_(item.classId);
    const date = formatDateToYmd(item.date);
    const period = normalizeString_(item.period);

    if (!classId || !date || !period) return;

    const key = [classId, date, period].join('__');

    if (!targetKeySet[key]) {
      targetKeySet[key] = true;
      targetKeys.push(key);
    }

    const experimentKey = buildExperimentSessionKeyForTimetable_(classId, date, period);
    if (experimentKey) {
      if (!targetExperimentKeyMap[experimentKey]) {
        targetExperimentKeyMap[experimentKey] = [];
      }
      if (targetExperimentKeyMap[experimentKey].indexOf(key) === -1) {
        targetExperimentKeyMap[experimentKey].push(key);
      }
    }

    result[key] = {
      isSaved: false,
      lastSavedInfo: null,
      saveStatusNotLoaded: false
    };
  });

  if (targetKeys.length === 0) {
    if (typeof logPerf_ === 'function') {
      logPerf_('getSaveStatusForTeacherSessionsDirect_ total', totalStartedAt, 'empty');
    }
    return result;
  }

  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_SESSIONS);

  if (!sheet) {
    throw new Error('attendanceSessions シートが見つかりません');
  }

  const loadStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const values = sheet.getDataRange().getValues();

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getSaveStatusForTeacherSessionsDirect_ load sheet',
      loadStartedAt,
      'rows=' + Math.max(values.length - 1, 0) + ' targets=' + targetKeys.length
    );
  }

  if (values.length <= 1) {
    if (typeof logPerf_ === 'function') {
      logPerf_(
        'getSaveStatusForTeacherSessionsDirect_ total',
        totalStartedAt,
        'no-data targets=' + targetKeys.length
      );
    }
    return result;
  }

  const headers = values[0];
  const rows = values.slice(1);

  const col = {
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    date: findColumnIndex_(headers, ['date', '日付']),
    period: findColumnIndex_(headers, ['period', '時限']),
    teacherEmail: findColumnIndex_(headers, ['teacherEmail', 'email']),
    accessedAt: findColumnIndex_(headers, ['accessedAt', 'savedAt']),
    actionType: findColumnIndex_(headers, ['actionType']),
    targetSessionKey: findColumnIndex_(headers, ['targetSessionKey']),
    savedModeLabel: findColumnIndex_(headers, ['savedModeLabel'])
  };

  ['classId', 'date', 'period'].forEach(function(key) {
    if (col[key] === -1) {
      throw new Error('attendanceSessions シートに ' + key + ' 列がありません');
    }
  });

  const scanStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  let foundCount = 0;
  const foundKeySet = {};

  // attendanceSessions は append 方式なので、下から見れば最新に近い
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];

    const rowClassId = String(row[col.classId] || '').trim();
    if (!rowClassId) continue;

    const rowDate = formatDateToYmd(row[col.date]);
    if (!rowDate) continue;

    const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();
    if (!rowPeriod) continue;

    const sessionKey = [rowClassId, rowDate, rowPeriod].join('__');
    const rowExperimentKey = buildExperimentSessionKeyForTimetable_(rowClassId, rowDate, rowPeriod);

    const matchedKeys = [];
    if (targetKeySet[sessionKey]) {
      matchedKeys.push(sessionKey);
    }

    if (rowExperimentKey && targetExperimentKeyMap[rowExperimentKey]) {
      targetExperimentKeyMap[rowExperimentKey].forEach(function(key) {
        if (matchedKeys.indexOf(key) === -1) matchedKeys.push(key);
      });
    }

    if (!matchedKeys.length) continue;

    const accessedAtRaw = col.accessedAt !== -1 ? row[col.accessedAt] : '';
    const teacherEmail = col.teacherEmail !== -1
      ? String(row[col.teacherEmail] || '').trim().toLowerCase()
      : '';

    const actionType = col.actionType !== -1
      ? String(row[col.actionType] || '').trim()
      : '';

    const targetSessionKey = col.targetSessionKey !== -1
      ? String(row[col.targetSessionKey] || '').trim()
      : sessionKey;

    const savedModeLabel = col.savedModeLabel !== -1
      ? String(row[col.savedModeLabel] || '').trim()
      : '';

    const accessedAtSerialized = accessedAtRaw instanceof Date
      ? accessedAtRaw.toISOString()
      : String(accessedAtRaw || '');

    const lastSavedInfo = {
      teacherEmail: teacherEmail,
      savedAt: accessedAtSerialized,
      savedAtText: formatDateTimeJst_(accessedAtRaw),
      actionType: actionType,
      targetSessionKey: targetSessionKey,
      savedModeLabel: savedModeLabel
    };

    matchedKeys.forEach(function(key) {
      if (foundKeySet[key]) return;

      result[key] = {
        isSaved: true,
        lastSavedInfo: lastSavedInfo,
        saveStatusNotLoaded: false
      };

      foundKeySet[key] = true;
      foundCount++;
    });

    if (foundCount >= targetKeys.length) {
      break;
    }
  }

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getSaveStatusForTeacherSessionsDirect_ scan rows',
      scanStartedAt,
      'found=' + foundCount + '/' + targetKeys.length
    );

    logPerf_(
      'getSaveStatusForTeacherSessionsDirect_ total',
      totalStartedAt,
      'targets=' + targetKeys.length + ' found=' + foundCount
    );
  }

  return result;
}
