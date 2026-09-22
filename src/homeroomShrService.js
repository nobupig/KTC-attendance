const HOMEROOM_SHR_CONFIG = {
  PERIOD: 0,
  ACTION_TYPE: 'homeroom-shr',
  MODE_LABEL: '担任SHR',
  ALLOWED_STATUS_CODES: ['', 'A', 'L', 'E']
};

const HOMEROOM_SHR_SCHEDULE_2026 = Object.freeze({
  PERIODS: Object.freeze([
    Object.freeze({
      start: '2026-04-07',
      end: '2026-09-19'
    }),
    Object.freeze({
      start: '2026-09-24',
      end: '2027-02-12'
    })
  ]),

  FORCE_INCLUDE_RANGES: Object.freeze([
    Object.freeze({
      start: '2026-09-14',
      end: '2026-09-19'
    }),
    Object.freeze({
      start: '2027-02-03',
      end: '2027-02-05'
    }),
    Object.freeze({
      start: '2027-02-08',
      end: '2027-02-12'
    })
  ]),

  FORCE_INCLUDE_DATES: Object.freeze([
    '2026-04-07'
  ]),

  FORCE_EXCLUDE_DATES: Object.freeze([
    '2026-07-27',
    '2027-02-06',
    '2027-02-07'
  ]),

  FINAL_DATE: '2027-02-12'
});

function toClientSafeLastSavedInfo_(info) {
  if (!info) return null;

  return {
    teacherEmail: String(info.teacherEmail || ''),
    savedAtText: String(info.savedAtText || ''),
    actionType: String(info.actionType || ''),
    targetSessionKey: String(info.targetSessionKey || ''),
    savedModeLabel: String(info.savedModeLabel || ''),
    savedByCurrentUser: !!info.savedByCurrentUser
  };
}

function getHomeroomShrInitialData() {
  const user = getCurrentUserContext();
  if (!user) {
    throw new Error('ユーザー情報を取得できませんでした');
  }

  const homeroomClasses = getMyHomeroomClasses();
  if (!homeroomClasses.length) {
    throw new Error('担任クラスが見つかりませんでした');
  }

  const firstClass = homeroomClasses[0];
  const today = formatDateToYmd(new Date());

  return {
    user: {
      name: user.name,
      email: user.email,
      roles: user.roles || []
    },
    homeroomClasses: homeroomClasses.map(function(item) {
      return {
        grade: String(item.grade || '').trim(),
        unit: String(item.unit || '').trim(),
        classLabel: buildHomeroomShrClassLabel_(item.grade, item.unit)
      };
    }),
    defaultGrade: String(firstClass.grade || '').trim(),
    defaultUnit: String(firstClass.unit || '').trim(),
    defaultDate: today
  };
}

function getHomeroomShrUnsavedSummary(grade, unit) {
  const targetGrade = String(grade || '').trim();
  const targetUnit = String(unit || '').trim();

  ensureHomeroomAccess_(targetGrade, targetUnit);

  const dateContext = getHomeroomShrUnsavedDateContext_(new Date());
  const startYmd = dateContext.startYmd;
  const endYmd = dateContext.endYmd;

  const cacheKey = buildHomeroomShrUnsavedSummaryCacheKey_(targetGrade, targetUnit, endYmd);
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const snapshot = buildHomeroomShrUnsavedSnapshot_(
    targetGrade,
    targetUnit,
    startYmd,
    endYmd
  );

  const classDayYmdList = snapshot.classDayYmdList;
  const unsavedCount = snapshot.unsavedYmdList.length;

  const result = {
    ok: true,
    classInfo: {
      grade: targetGrade,
      unit: targetUnit,
      classLabel: buildHomeroomShrClassLabel_(targetGrade, targetUnit)
    },
    count: unsavedCount,
    checkedDays: classDayYmdList.length,
    checkedRange: {
      start: startYmd,
      end: endYmd
    }
  };

  putScriptCacheJson_(cacheKey, result, 60);
  return result;
}

function getHomeroomShrUnsavedDetails(grade, unit) {
  const targetGrade = String(grade || '').trim();
  const targetUnit = String(unit || '').trim();

  ensureHomeroomAccess_(targetGrade, targetUnit);

  const dateContext = getHomeroomShrUnsavedDateContext_(new Date());
  const startYmd = dateContext.startYmd;
  const endYmd = dateContext.endYmd;

  const cacheKey = buildHomeroomShrUnsavedDetailsCacheKey_(targetGrade, targetUnit, endYmd);
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const snapshot = buildHomeroomShrUnsavedSnapshot_(
    targetGrade,
    targetUnit,
    startYmd,
    endYmd
  );

  const items = snapshot.unsavedYmdList.map(function(ymd) {
    return {
      date: ymd,
      weekday: getWeekdayJaFromYmd_(ymd)
    };
  });

  const result = {
    ok: true,
    classInfo: {
      grade: targetGrade,
      unit: targetUnit,
      classLabel: buildHomeroomShrClassLabel_(targetGrade, targetUnit)
    },
    items: items,
    checkedRange: {
      start: startYmd,
      end: endYmd
    }
  };

  putScriptCacheJson_(cacheKey, result, 60);
  return result;
}

function buildHomeroomShrUnsavedSnapshot_(grade, unit, startYmd, endYmd) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const targetGrade = String(grade || '').trim();
  const targetUnit = String(unit || '').trim();
  const cacheKey = buildHomeroomShrUnsavedSnapshotCacheKey_(
    targetGrade,
    targetUnit,
    endYmd
  );

  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    if (typeof logPerf_ === 'function') {
      logPerf_(
        'buildHomeroomShrUnsavedSnapshot_ total',
        totalStartedAt,
        'cache=hit'
      );
    }
    return cached;
  }

  const classId = buildHomeroomShrClassId_(targetGrade, targetUnit);
  const targetPeriod = String(HOMEROOM_SHR_CONFIG.PERIOD);

  const classDayStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const classDayYmdList = getHomeroomShrClassDayYmdList_(startYmd, endYmd);

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'buildHomeroomShrUnsavedSnapshot_ class days',
      classDayStartedAt,
      'days=' + classDayYmdList.length
    );
  }

  /*
   * attendanceSessions 全体は ScriptCache の1キー上限を大幅に超える。
   * 未保存SHR判定では必要列だけを直接読み、
   * classId / period / actionType を先に絞ってから日付を正規化する。
   */
  const loadStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_SESSIONS);

  if (!sheet) {
    throw new Error('attendanceSessions シートが見つかりません');
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const headers = lastColumn > 0
    ? sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    : [];

  const col = {
    classId: headers.indexOf('classId'),
    date: headers.indexOf('date'),
    period: headers.indexOf('period'),
    actionType: headers.indexOf('actionType')
  };

  ['classId', 'date', 'period'].forEach(function(key) {
    if (col[key] === -1) {
      throw new Error('attendanceSessions シートに ' + key + ' 列がありません');
    }
  });

  const scanIndexes = [
    col.classId,
    col.date,
    col.period
  ];

  if (col.actionType !== -1) {
    scanIndexes.push(col.actionType);
  }

  const scanStartIndex = Math.min.apply(null, scanIndexes);
  const scanEndIndex = Math.max.apply(null, scanIndexes);
  const scanWidth = scanEndIndex - scanStartIndex + 1;
  const dataRowCount = Math.max(lastRow - 1, 0);

  const rows = dataRowCount > 0
    ? sheet
        .getRange(
          2,
          scanStartIndex + 1,
          dataRowCount,
          scanWidth
        )
        .getDisplayValues()
    : [];

  const relativeCol = {
    classId: col.classId - scanStartIndex,
    date: col.date - scanStartIndex,
    period: col.period - scanStartIndex,
    actionType: col.actionType === -1
      ? -1
      : col.actionType - scanStartIndex
  };

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'buildHomeroomShrUnsavedSnapshot_ load attendanceSessions compact',
      loadStartedAt,
      'rows=' + rows.length + ' width=' + scanWidth
    );
  }

  const scanStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const savedDateMap = {};
  let classCandidateCount = 0;
  let shrCandidateCount = 0;

  rows.forEach(function(row) {
    const rowClassId = String(row[relativeCol.classId] || '').trim();
    if (rowClassId !== classId) return;

    classCandidateCount++;

    const rowPeriod = String(
      row[relativeCol.period] == null
        ? ''
        : row[relativeCol.period]
    ).trim();

    if (rowPeriod !== targetPeriod) return;

    const actionType = relativeCol.actionType !== -1
      ? String(row[relativeCol.actionType] || '').trim()
      : '';

    if (actionType && actionType !== HOMEROOM_SHR_CONFIG.ACTION_TYPE) return;

    shrCandidateCount++;

    const rawDate = row[relativeCol.date];
    const rowDate =
      typeof normalizeYmdDisplayText_ === 'function'
        ? normalizeYmdDisplayText_(rawDate)
        : formatDateToYmd(rawDate);

    if (!rowDate || rowDate < startYmd || rowDate > endYmd) return;

    savedDateMap[rowDate] = true;
  });

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'buildHomeroomShrUnsavedSnapshot_ scan attendanceSessions compact',
      scanStartedAt,
      'classCandidates=' + classCandidateCount +
        ' shrCandidates=' + shrCandidateCount +
        ' savedDates=' + Object.keys(savedDateMap).length
    );
  }

  const unsavedYmdList = classDayYmdList.filter(function(ymd) {
    return !savedDateMap[ymd];
  });

  const snapshot = {
    classDayYmdList: classDayYmdList,
    unsavedYmdList: unsavedYmdList
  };

  putScriptCacheJson_(cacheKey, snapshot, 60);

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'buildHomeroomShrUnsavedSnapshot_ total',
      totalStartedAt,
      'cache=miss unsaved=' + unsavedYmdList.length
    );
  }

  return snapshot;
}

function buildHomeroomShrUnsavedSnapshotCacheKey_(grade, unit, endYmd) {
  return 'homeroomShrUnsavedSnapshot__' +
    String(grade || '').trim() + '__' +
    String(unit || '').trim() + '__' +
    String(endYmd || '').trim();
}

function buildHomeroomShrUnsavedDetailsCacheKey_(grade, unit, endYmd) {
  return 'homeroomShrUnsavedDetails__' +
    String(grade || '').trim() + '__' +
    String(unit || '').trim() + '__' +
    String(endYmd || '').trim();
}

function getWeekdayJaFromYmd_(ymd) {
  const parts = String(ymd || '').split('-').map(Number);
  if (parts.length !== 3) return '';

  const d = new Date(parts[0], parts[1] - 1, parts[2]);
  const weekdays = ['日', '月', '火', '水', '木', '金', '土'];
  return weekdays[d.getDay()] || '';
}


function buildHomeroomShrUnsavedSummaryCacheKey_(grade, unit, endYmd) {
  return 'homeroomShrUnsavedSummary__' +
    String(grade || '').trim() + '__' +
    String(unit || '').trim() + '__' +
    String(endYmd || '').trim();
}

function getHomeroomShrUnsavedDateContext_(baseDate) {
  const todayYmd = formatDateToYmd(baseDate || new Date());
  if (!todayYmd) {
    throw new Error('SHR未保存判定の基準日を取得できませんでした');
  }

  const today = new Date(todayYmd + 'T12:00:00+09:00');
  const endDate = new Date(today);
  endDate.setDate(endDate.getDate() - 1);

  const periods = HOMEROOM_SHR_SCHEDULE_2026.PERIODS;
  const startYmd = periods[0].start;
  let endYmd = formatDateToYmd(endDate);

  if (endYmd > HOMEROOM_SHR_SCHEDULE_2026.FINAL_DATE) {
    endYmd = HOMEROOM_SHR_SCHEDULE_2026.FINAL_DATE;
  }

  return {
    startYmd: startYmd,
    endYmd: endYmd
  };
}

function isHomeroomShrDateInConfiguredPeriod_(ymd) {
  return HOMEROOM_SHR_SCHEDULE_2026.PERIODS.some(function(period) {
    return ymd >= period.start && ymd <= period.end;
  });
}

function isHomeroomShrForceIncluded_(ymd) {
  if (HOMEROOM_SHR_SCHEDULE_2026.FORCE_INCLUDE_DATES.indexOf(ymd) !== -1) {
    return true;
  }

  return HOMEROOM_SHR_SCHEDULE_2026.FORCE_INCLUDE_RANGES.some(function(range) {
    return ymd >= range.start && ymd <= range.end;
  });
}

function isHomeroomShrForceExcluded_(ymd) {
  return HOMEROOM_SHR_SCHEDULE_2026.FORCE_EXCLUDE_DATES.indexOf(ymd) !== -1;
}

function getHomeroomShrClassDayYmdList_(startYmd, endYmd) {
  if (!startYmd || !endYmd || startYmd > endYmd) {
    return [];
  }

  const calendarData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CALENDAR, 1800);
  const headers = Array.isArray(calendarData && calendarData.headers) ? calendarData.headers : [];
  const rows = Array.isArray(calendarData && calendarData.rows) ? calendarData.rows : [];

  const col = {
    date: headers.indexOf('date'),
    isClassDay: headers.indexOf('isClassDay')
  };

  if (col.date === -1 || col.isClassDay === -1 || !rows.length) {
    throw new Error('calendar シートに SHR未保存判定に必要な date / isClassDay データがありません');
  }

  const calendarMap = {};

  rows.forEach(function(row) {
    const rowDate = formatDateToYmd(row[col.date]);
    if (!rowDate) return;

    const raw = row[col.isClassDay];
    calendarMap[rowDate] =
      raw === true ||
      String(raw || '').trim().toUpperCase() === 'TRUE' ||
      String(raw || '').trim() === '1';
  });

  const startParts = String(startYmd).split('-').map(Number);
  const endParts = String(endYmd).split('-').map(Number);

  if (startParts.length !== 3 || endParts.length !== 3) {
    throw new Error('SHR未保存判定の日付範囲が不正です');
  }

  const current = new Date(startParts[0], startParts[1] - 1, startParts[2]);
  const end = new Date(endParts[0], endParts[1] - 1, endParts[2]);
  const result = [];

  while (current <= end) {
    const ymd = formatDateToYmd(current);

    if (!isHomeroomShrDateInConfiguredPeriod_(ymd)) {
      current.setDate(current.getDate() + 1);
      continue;
    }

    if (isHomeroomShrForceExcluded_(ymd)) {
      current.setDate(current.getDate() + 1);
      continue;
    }

    if (isHomeroomShrForceIncluded_(ymd)) {
      result.push(ymd);
      current.setDate(current.getDate() + 1);
      continue;
    }

    if (!Object.prototype.hasOwnProperty.call(calendarMap, ymd)) {
      throw new Error('calendar シートに SHR判定対象日の行がありません: ' + ymd);
    }

    if (calendarMap[ymd]) {
      result.push(ymd);
    }

    current.setDate(current.getDate() + 1);
  }

  return result;
}
function buildWeekdayYmdList_(startYmd, endYmd) {
  const startParts = String(startYmd || '').split('-').map(Number);
  const endParts = String(endYmd || '').split('-').map(Number);

  if (startParts.length !== 3 || endParts.length !== 3) {
    return [];
  }

  const current = new Date(startParts[0], startParts[1] - 1, startParts[2]);
  const end = new Date(endParts[0], endParts[1] - 1, endParts[2]);

  const result = [];

  while (current <= end) {
    const day = current.getDay();
    if (day !== 0 && day !== 6) {
      result.push(formatDateToYmd(current));
    }
    current.setDate(current.getDate() + 1);
  }

  return result;
}

function getHomeroomShrSummaryStartDate_(baseDate) {
  const d = new Date(baseDate);
  d.setHours(0, 0, 0, 0);

  // 4月〜12月ならその年の4/7、1月〜3月なら前年度の4/7
  const schoolYear = d.getMonth() >= 3 ? d.getFullYear() : d.getFullYear() - 1;

  const start = new Date(schoolYear, 3, 7); // 4月7日
  start.setHours(0, 0, 0, 0);
  return start;
}

function getHomeroomShrDailyData(grade, unit, date) {
  ensureHomeroomAccess_(grade, unit);

  const targetGrade = String(grade || '').trim();
  const targetUnit = String(unit || '').trim();
  const targetDate = formatDateToYmd(date || new Date());

  const students = getStudentsByHomeroomClass_(targetGrade, targetUnit);
  const attendanceMap = getHomeroomShrAttendanceMap_(targetGrade, targetUnit, targetDate);

  const classId = buildHomeroomShrClassId_(targetGrade, targetUnit);

  const lastSavedInfo = getLatestAttendanceSessionInfo_(
    classId,
    targetDate,
    HOMEROOM_SHR_CONFIG.PERIOD,
    HOMEROOM_SHR_CONFIG.ACTION_TYPE
  );

  const statusCounts = {
    present: 0,
    absent: 0,
    late: 0,
    early: 0
  };

  const studentRows = students.map(function(student) {
    const statusCode = String(attendanceMap[student.studentId] || '').trim();

    if (statusCode === 'A') {
      statusCounts.absent += 1;
    } else if (statusCode === 'L') {
      statusCounts.late += 1;
    } else if (statusCode === 'E') {
      statusCounts.early += 1;
    } else {
      statusCounts.present += 1;
    }

    return {
      studentId: String(student.studentId || '').trim(),
      attendanceNumber: student.attendanceNumber,
      name: String(student.name || '').trim(),
      unit: String(student.unit || '').trim(),
      statusCode: statusCode
    };
  });

return {
  classInfo: {
    grade: targetGrade,
    unit: targetUnit,
    classLabel: buildHomeroomShrClassLabel_(targetGrade, targetUnit)
  },
  date: targetDate,
  students: studentRows,
  statusCounts: statusCounts,
  hasSavedSession: !!lastSavedInfo,
  lastSavedInfo: toClientSafeLastSavedInfo_(lastSavedInfo)
};
}

function getHomeroomShrDailyStatus(grade, unit, date) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  ensureHomeroomAccess_(grade, unit);

  const targetGrade = String(grade || '').trim();
  const targetUnit = String(unit || '').trim();
  const targetDate = formatDateToYmd(date || new Date());

  const classId = buildHomeroomShrClassId_(targetGrade, targetUnit);
  const period = String(HOMEROOM_SHR_CONFIG.PERIOD);

  const loadStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_SESSIONS);

  if (!sheet) {
    throw new Error('attendanceSessions シートが見つかりません');
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const headers = lastColumn > 0
    ? sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
    : [];

  const col = {
    classId: headers.indexOf('classId'),
    date: headers.indexOf('date'),
    period: headers.indexOf('period'),
    teacherEmail: headers.indexOf('teacherEmail'),
    accessedAt: headers.indexOf('accessedAt'),
    actionType: headers.indexOf('actionType'),
    targetSessionKey: headers.indexOf('targetSessionKey'),
    savedModeLabel: headers.indexOf('savedModeLabel')
  };

  ['classId', 'date', 'period'].forEach(function(key) {
    if (col[key] === -1) {
      throw new Error('attendanceSessions シートに ' + key + ' 列がありません');
    }
  });

  const scanIndexes = [
    col.classId,
    col.date,
    col.period
  ];

  [
    col.teacherEmail,
    col.accessedAt,
    col.actionType,
    col.targetSessionKey,
    col.savedModeLabel
  ].forEach(function(index) {
    if (index !== -1) {
      scanIndexes.push(index);
    }
  });

  const scanStartIndex = Math.min.apply(null, scanIndexes);
  const scanEndIndex = Math.max.apply(null, scanIndexes);
  const scanWidth = scanEndIndex - scanStartIndex + 1;
  const dataRowCount = Math.max(lastRow - 1, 0);

  const rows = dataRowCount > 0
    ? sheet
        .getRange(
          2,
          scanStartIndex + 1,
          dataRowCount,
          scanWidth
        )
        .getValues()
    : [];

  const relativeCol = {};
  Object.keys(col).forEach(function(key) {
    relativeCol[key] =
      col[key] === -1
        ? -1
        : col[key] - scanStartIndex;
  });

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getHomeroomShrDailyStatus load attendanceSessions compact',
      loadStartedAt,
      'rows=' + rows.length + ' width=' + scanWidth
    );
  }

  const scanStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  let latest = null;
  let latestMs = 0;
  let classCandidateCount = 0;
  let matchedCount = 0;

  rows.forEach(function(row) {
    const rowClassId = String(row[relativeCol.classId] || '').trim();
    if (rowClassId !== classId) return;

    classCandidateCount++;

    const rowPeriod = String(
      row[relativeCol.period] == null
        ? ''
        : row[relativeCol.period]
    ).trim();

    if (rowPeriod !== period) return;

    const rowActionType = relativeCol.actionType !== -1
      ? String(row[relativeCol.actionType] || '').trim()
      : '';

    if (rowActionType && rowActionType !== HOMEROOM_SHR_CONFIG.ACTION_TYPE) return;

    const rowDate = formatDateToYmd(row[relativeCol.date]);
    if (rowDate !== targetDate) return;

    matchedCount++;

    const rawSavedAt = relativeCol.accessedAt !== -1
      ? row[relativeCol.accessedAt]
      : '';
    const savedAt = rawSavedAt instanceof Date ? rawSavedAt : new Date(rawSavedAt);
    const savedAtMs = isNaN(savedAt.getTime()) ? 0 : savedAt.getTime();

    if (!latest || savedAtMs >= latestMs) {
      latestMs = savedAtMs;
      latest = {
        teacherEmail: relativeCol.teacherEmail !== -1
          ? String(row[relativeCol.teacherEmail] || '').trim()
          : '',
        savedAtText: formatDateTimeJst_(rawSavedAt),
        actionType: rowActionType,
        targetSessionKey: relativeCol.targetSessionKey !== -1
          ? String(row[relativeCol.targetSessionKey] || '').trim()
          : '',
        savedModeLabel: relativeCol.savedModeLabel !== -1
          ? String(row[relativeCol.savedModeLabel] || '').trim()
          : HOMEROOM_SHR_CONFIG.MODE_LABEL,
        savedByCurrentUser: false
      };
    }
  });

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getHomeroomShrDailyStatus scan attendanceSessions compact',
      scanStartedAt,
      'classCandidates=' + classCandidateCount +
        ' matched=' + matchedCount
    );
  }

  const result = {
    classInfo: {
      grade: targetGrade,
      unit: targetUnit,
      classLabel: buildHomeroomShrClassLabel_(targetGrade, targetUnit)
    },
    date: targetDate,
    hasSavedSession: !!latest,
    lastSavedInfo: latest ? toClientSafeLastSavedInfo_(latest) : null
  };

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getHomeroomShrDailyStatus total',
      totalStartedAt,
      'hasSaved=' + (!!latest)
    );
  }

  return result;
}

function saveHomeroomShrAttendance(payload) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (!payload) {
      throw new Error('保存データがありません');
    }

    const grade = String(payload.grade || '').trim();
    const unit = String(payload.unit || '').trim();
    const targetDate = formatDateToYmd(payload.date || new Date());

    ensureHomeroomAccess_(grade, unit);

    const students = getStudentsByHomeroomClass_(grade, unit);
    const validStudentIds = {};
    students.forEach(function(student) {
      validStudentIds[String(student.studentId || '').trim()] = true;
    });

    const records = Array.isArray(payload.records) ? payload.records : [];
    if (!records.length) {
      throw new Error('保存対象の学生データがありません');
    }

    const classId = buildHomeroomShrClassId_(grade, unit);
    const period = String(HOMEROOM_SHR_CONFIG.PERIOD);
    const currentUserEmail = getCurrentUserEmail();
    const now = new Date();
    const sessionKey = [classId, targetDate, period].join('__');

    records.forEach(function(record) {
      const studentId = String(record.studentId || '').trim();
      const statusCode = String(record.statusCode || '').trim();

      if (!studentId || !validStudentIds[studentId]) {
        throw new Error('担任クラスに存在しない studentId が含まれています: ' + studentId);
      }

      if (HOMEROOM_SHR_CONFIG.ALLOWED_STATUS_CODES.indexOf(statusCode) === -1) {
        throw new Error('statusCode が不正です: ' + statusCode);
      }
    });

    const ss = getOperationSpreadsheet();
    const attendanceSessionsSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_SESSIONS);
    const attendanceSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);

    if (!attendanceSessionsSheet) {
      throw new Error('attendanceSessions シートが見つかりません');
    }
    if (!attendanceSheet) {
      throw new Error('attendance シートが見つかりません');
    }

    /*
     * 旧処理は attendance 全体を読み込み、対象SHR以外も含めて
     * 全行 clearContent() -> 全行 setValues() で再構築していた。
     *
     * ここでは対象クラス・日付・0限・対象学生の行だけを
     * clear / update / append し、他の出席データには触れない。
     */
    const loadStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

    const lastRow = attendanceSheet.getLastRow();
    const lastColumn = attendanceSheet.getLastColumn();
    const headers = lastColumn > 0
      ? attendanceSheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0]
      : [];

    const col = {
      classId: headers.indexOf('classId'),
      date: headers.indexOf('date'),
      period: headers.indexOf('period'),
      studentId: headers.indexOf('studentId'),
      statusCode: headers.indexOf('statusCode'),
      recordedAt: headers.indexOf('recordedAt')
    };

    Object.keys(col).forEach(function(key) {
      if (col[key] === -1) {
        throw new Error('attendance シートに ' + key + ' 列がありません');
      }
    });

    const scanIndexes = [
      col.classId,
      col.date,
      col.period,
      col.studentId
    ];
    const scanStartIndex = Math.min.apply(null, scanIndexes);
    const scanEndIndex = Math.max.apply(null, scanIndexes);
    const scanWidth = scanEndIndex - scanStartIndex + 1;
    const dataRowCount = Math.max(lastRow - 1, 0);

    const rows = dataRowCount > 0
      ? attendanceSheet
          .getRange(
            2,
            scanStartIndex + 1,
            dataRowCount,
            scanWidth
          )
          .getDisplayValues()
      : [];

    const relativeCol = {
      classId: col.classId - scanStartIndex,
      date: col.date - scanStartIndex,
      period: col.period - scanStartIndex,
      studentId: col.studentId - scanStartIndex
    };

    if (typeof logPerf_ === 'function') {
      logPerf_(
        'saveHomeroomShrAttendance load attendance compact',
        loadStartedAt,
        'rows=' + rows.length + ' width=' + scanWidth
      );
    }

    const targetStudentIds = {};
    const desiredByStudentId = {};

    records.forEach(function(record) {
      const studentId = String(record.studentId || '').trim();
      const statusCode = String(record.statusCode || '').trim();

      if (!studentId) return;

      targetStudentIds[studentId] = true;
      desiredByStudentId[studentId] = statusCode;
    });

    const scanStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
    const existingRowNumbersByStudentId = {};
    let classCandidateCount = 0;
    let matchedRowCount = 0;

    rows.forEach(function(row, index) {
      const rowClassId = String(row[relativeCol.classId] || '').trim();
      if (rowClassId !== classId) return;

      classCandidateCount++;

      const rowPeriod = String(
        row[relativeCol.period] == null
          ? ''
          : row[relativeCol.period]
      ).trim();

      if (rowPeriod !== period) return;

      const rowStudentId = String(row[relativeCol.studentId] || '').trim();
      if (!targetStudentIds[rowStudentId]) return;

      const rawDate = row[relativeCol.date];
      const rowDate =
        typeof normalizeYmdDisplayText_ === 'function'
          ? normalizeYmdDisplayText_(rawDate)
          : formatDateToYmd(rawDate);

      if (rowDate !== targetDate) return;

      if (!existingRowNumbersByStudentId[rowStudentId]) {
        existingRowNumbersByStudentId[rowStudentId] = [];
      }

      existingRowNumbersByStudentId[rowStudentId].push(index + 2);
      matchedRowCount++;
    });

    if (typeof logPerf_ === 'function') {
      logPerf_(
        'saveHomeroomShrAttendance scan attendance compact',
        scanStartedAt,
        'classCandidates=' + classCandidateCount +
          ' matchedRows=' + matchedRowCount
      );
    }

    const rowsToClear = [];
    const rowsToUpdate = [];
    const appendRows = [];

    Object.keys(targetStudentIds).forEach(function(studentId) {
      const statusCode = desiredByStudentId[studentId] || '';
      const existingRows = existingRowNumbersByStudentId[studentId] || [];

      if (!statusCode) {
        existingRows.forEach(function(rowNumber) {
          rowsToClear.push(rowNumber);
        });
        return;
      }

      const fullRow = buildAttendanceSheetRow_(headers.length, col, {
        classId: classId,
        date: targetDate,
        period: Number(period),
        studentId: studentId,
        statusCode: statusCode,
        recordedAt: now
      });

      if (existingRows.length > 0) {
        const updateRowNumber = existingRows[existingRows.length - 1];

        existingRows.slice(0, -1).forEach(function(rowNumber) {
          rowsToClear.push(rowNumber);
        });

        rowsToUpdate.push({
          rowNumber: updateRowNumber,
          values: fullRow
        });
      } else {
        appendRows.push(fullRow);
      }
    });

    rowsToClear.sort(function(a, b) {
      return a - b;
    });

    rowsToUpdate.sort(function(a, b) {
      return a.rowNumber - b.rowNumber;
    });

    const writeStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

    if (rowsToClear.length > 0) {
      let rangeStart = rowsToClear[0];
      let previousRow = rowsToClear[0];

      for (let i = 1; i <= rowsToClear.length; i++) {
        const currentRow = i < rowsToClear.length
          ? rowsToClear[i]
          : null;

        if (currentRow === previousRow + 1) {
          previousRow = currentRow;
          continue;
        }

        attendanceSheet
          .getRange(
            rangeStart,
            1,
            previousRow - rangeStart + 1,
            headers.length
          )
          .clearContent();

        if (currentRow !== null) {
          rangeStart = currentRow;
          previousRow = currentRow;
        }
      }
    }

    if (rowsToUpdate.length > 0) {
      const rowNumbers = rowsToUpdate.map(function(item) {
        return item.rowNumber;
      });

      if (isSequentialRows_(rowNumbers)) {
        attendanceSheet
          .getRange(
            rowNumbers[0],
            1,
            rowsToUpdate.length,
            headers.length
          )
          .setValues(
            rowsToUpdate.map(function(item) {
              return item.values;
            })
          );
      } else {
        rowsToUpdate.forEach(function(item) {
          attendanceSheet
            .getRange(
              item.rowNumber,
              1,
              1,
              headers.length
            )
            .setValues([item.values]);
        });
      }
    }

    if (appendRows.length > 0) {
      const appendStartRow =
        Math.max(attendanceSheet.getLastRow(), 1) + 1;

      attendanceSheet
        .getRange(
          appendStartRow,
          1,
          appendRows.length,
          headers.length
        )
        .setValues(appendRows);
    }

    if (typeof logPerf_ === 'function') {
      logPerf_(
        'saveHomeroomShrAttendance write targeted rows',
        writeStartedAt,
        'cleared=' + rowsToClear.length +
          ' updated=' + rowsToUpdate.length +
          ' appended=' + appendRows.length
      );
    }

    appendAttendanceSessionLog_(attendanceSessionsSheet, [
      classId,
      targetDate,
      HOMEROOM_SHR_CONFIG.PERIOD,
      currentUserEmail,
      now,
      HOMEROOM_SHR_CONFIG.ACTION_TYPE,
      sessionKey,
      HOMEROOM_SHR_CONFIG.MODE_LABEL
    ]);

    invalidateAttendanceCaches_(classId, targetDate, period);

    const summaryDateContext = getHomeroomShrUnsavedDateContext_(new Date());
    const summaryEndYmd = summaryDateContext.endYmd;

    removeScriptCacheKeys_([
      buildHomeroomShrDailyCacheKey_(grade, unit, targetDate),
      buildHomeroomShrUnsavedSummaryCacheKey_(grade, unit, summaryEndYmd),
      buildHomeroomShrUnsavedDetailsCacheKey_(grade, unit, summaryEndYmd),
      buildHomeroomShrUnsavedSnapshotCacheKey_(grade, unit, summaryEndYmd)
    ]);

    const result = {
      success: true,
      classId: classId,
      date: targetDate,
      savedCount: records.length,
      lastSavedInfo: toClientSafeLastSavedInfo_({
        teacherEmail: currentUserEmail,
        savedAtText: formatDateTimeJst_(now),
        actionType: HOMEROOM_SHR_CONFIG.ACTION_TYPE,
        targetSessionKey: sessionKey,
        savedModeLabel: HOMEROOM_SHR_CONFIG.MODE_LABEL,
        savedByCurrentUser: true
      })
    };

    if (typeof logPerf_ === 'function') {
      logPerf_(
        'saveHomeroomShrAttendance total',
        totalStartedAt,
        'records=' + records.length +
          ' cleared=' + rowsToClear.length +
          ' updated=' + rowsToUpdate.length +
          ' appended=' + appendRows.length
      );
    }

    return result;

  } finally {
    lock.releaseLock();
  }
}

function saveHomeroomShrNoAbsence(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (!payload) {
      throw new Error('保存データがありません');
    }

    const grade = String(payload.grade || '').trim();
    const unit = String(payload.unit || '').trim();
    const targetDate = formatDateToYmd(payload.date || new Date());

    ensureHomeroomAccess_(grade, unit);

    const classId = buildHomeroomShrClassId_(grade, unit);
    const period = String(HOMEROOM_SHR_CONFIG.PERIOD);

    const currentUserEmail = getCurrentUserEmail();
    const now = new Date();

    const sessionKey = [classId, targetDate, period].join('__');

    const ss = getOperationSpreadsheet();

    const attendanceSessionsSheet =
      ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_SESSIONS);

    if (!attendanceSessionsSheet) {
      throw new Error('attendanceSessions シートが見つかりません');
    }

    appendAttendanceSessionLog_(attendanceSessionsSheet, [
      classId,
      targetDate,
      HOMEROOM_SHR_CONFIG.PERIOD,
      currentUserEmail,
      now,
      HOMEROOM_SHR_CONFIG.ACTION_TYPE,
      sessionKey,
      HOMEROOM_SHR_CONFIG.MODE_LABEL + '（欠席者なし）'
    ]);

    invalidateAttendanceCaches_(classId, targetDate, period);

    const summaryDateContext = getHomeroomShrUnsavedDateContext_(new Date());
    const summaryEndYmd = summaryDateContext.endYmd;

    removeScriptCacheKeys_([
      buildHomeroomShrDailyCacheKey_(grade, unit, targetDate),
      buildHomeroomShrUnsavedSummaryCacheKey_(grade, unit, summaryEndYmd),
      buildHomeroomShrUnsavedDetailsCacheKey_(grade, unit, summaryEndYmd),
      buildHomeroomShrUnsavedSnapshotCacheKey_(grade, unit, summaryEndYmd)
    ]);

    return {
      success: true,
      classId: classId,
      date: targetDate,
      lastSavedInfo: toClientSafeLastSavedInfo_({
        teacherEmail: currentUserEmail,
        savedAtText: formatDateTimeJst_(now),
        actionType: HOMEROOM_SHR_CONFIG.ACTION_TYPE,
        targetSessionKey: sessionKey,
        savedModeLabel:
          HOMEROOM_SHR_CONFIG.MODE_LABEL + '（欠席者なし）',
        savedByCurrentUser: true
      })
    };

  } finally {
    lock.releaseLock();
  }
}

function getHomeroomShrSummary(grade, unit, startDate, endDate) {
  ensureHomeroomAccess_(grade, unit);

  const targetGrade = String(grade || '').trim();
  const targetUnit = String(unit || '').trim();
  const targetClassId = buildHomeroomShrClassId_(targetGrade, targetUnit);
  const targetPeriod = String(HOMEROOM_SHR_CONFIG.PERIOD);

  const students = getStudentsByHomeroomClass_(targetGrade, targetUnit);

  const sessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.ATTENDANCE_SESSIONS, 60);
  const sessionHeaders = sessionsData.headers;
  const sessionRows = sessionsData.rows;

  const sessionCol = {
    classId: findColumnIndex_(sessionHeaders, ['classId']),
    date: findColumnIndex_(sessionHeaders, ['date']),
    period: findColumnIndex_(sessionHeaders, ['period']),
    actionType: findColumnIndex_(sessionHeaders, ['actionType'])
  };

  const startYmd = startDate ? formatDateToYmd(startDate) : '';
  const endYmd = endDate ? formatDateToYmd(endDate) : '';

  const sessionDateMap = {};

  sessionRows.forEach(function(row) {
    const rowClassId = String(row[sessionCol.classId] || '').trim();
    const rowDate = formatDateToYmd(row[sessionCol.date]);
    const rowPeriod = String(row[sessionCol.period] || '').trim();
    const actionType = sessionCol.actionType !== -1 ? String(row[sessionCol.actionType] || '').trim() : '';

    if (rowClassId !== targetClassId) return;
    if (rowPeriod !== targetPeriod) return;
    if (actionType && actionType !== HOMEROOM_SHR_CONFIG.ACTION_TYPE) return;
    if (startYmd && rowDate < startYmd) return;
    if (endYmd && rowDate > endYmd) return;

    sessionDateMap[rowDate] = true;
  });

  const recordedDays = Object.keys(sessionDateMap).sort();

  const attendanceData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.ATTENDANCE, 60);
  const headers = attendanceData.headers;
  const rows = attendanceData.rows;

  const col = {
    classId: findColumnIndex_(headers, ['classId']),
    date: findColumnIndex_(headers, ['date']),
    period: findColumnIndex_(headers, ['period']),
    studentId: findColumnIndex_(headers, ['studentId']),
    statusCode: findColumnIndex_(headers, ['statusCode'])
  };

  const summaryMap = {};
  students.forEach(function(student) {
    const studentId = String(student.studentId || '').trim();
    summaryMap[studentId] = {
      studentId: studentId,
      attendanceNumber: student.attendanceNumber,
      name: student.name,
      recordDays: recordedDays.length,
      absentCount: 0,
      lateCount: 0,
      earlyCount: 0
    };
  });

  rows.forEach(function(row) {
    const rowClassId = String(row[col.classId] || '').trim();
    const rowDate = formatDateToYmd(row[col.date]);
    const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();
    const studentId = String(row[col.studentId] || '').trim();
    const statusCode = String(row[col.statusCode] || '').trim();

    if (rowClassId !== targetClassId) return;
    if (rowPeriod !== targetPeriod) return;
    if (!summaryMap[studentId]) return;
    if (startYmd && rowDate < startYmd) return;
    if (endYmd && rowDate > endYmd) return;
    if (!sessionDateMap[rowDate]) return;

    if (statusCode === 'A') summaryMap[studentId].absentCount += 1;
    if (statusCode === 'L') summaryMap[studentId].lateCount += 1;
    if (statusCode === 'E') summaryMap[studentId].earlyCount += 1;
  });

  return {
    classInfo: {
      grade: targetGrade,
      unit: targetUnit,
      classLabel: buildHomeroomShrClassLabel_(targetGrade, targetUnit)
    },
    recordedDays: recordedDays,
    students: Object.keys(summaryMap)
      .map(function(studentId) { return summaryMap[studentId]; })
      .sort(compareStudentsByAttendanceNumber_)
  };
}

function getHomeroomShrAttendanceMap_(grade, unit, date) {
  const targetGrade = String(grade || '').trim();
  const targetUnit = String(unit || '').trim();
  const targetDate = formatDateToYmd(date);

  const classId = buildHomeroomShrClassId_(targetGrade, targetUnit);
  const period = String(HOMEROOM_SHR_CONFIG.PERIOD);

  const ss = getOperationSpreadsheet();
  const attendanceSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (!attendanceSheet) {
    throw new Error('attendance シートが見つかりません');
  }

  const values = attendanceSheet.getDataRange().getValues();
  const headers = values.length > 0 ? values[0] : [];
  const rows = values.length > 1 ? values.slice(1) : [];

  const col = {
    classId: headers.indexOf('classId'),
    date: headers.indexOf('date'),
    period: headers.indexOf('period'),
    studentId: headers.indexOf('studentId'),
    statusCode: headers.indexOf('statusCode')
  };

  Object.keys(col).forEach(function(key) {
    if (col[key] === -1) {
      throw new Error('attendance シートに ' + key + ' 列がありません');
    }
  });

  const result = {};

  rows.forEach(function(row) {
    const rowClassId = String(row[col.classId] || '').trim();
    const rowDate = formatDateToYmd(row[col.date]);
    const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();
    const rowStudentId = String(row[col.studentId] || '').trim();
    const rowStatusCode = String(row[col.statusCode] || '').trim();

    if (rowClassId !== classId) return;
    if (rowDate !== targetDate) return;
    if (rowPeriod !== period) return;
    if (!rowStudentId) return;

    result[rowStudentId] = rowStatusCode;
  });

  return result;
}

function getStudentsByHomeroomClass_(grade, unit) {
  const targetGrade = String(grade || '').trim();
  const targetUnit = String(unit || '').trim();

  const cacheKey = 'studentsByHomeroomClass__' + targetGrade + '__' + targetUnit;
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const targetUnits = typeof expandTargetStudentUnitsForHomeroomUnit_ === 'function'
    ? expandTargetStudentUnitsForHomeroomUnit_(targetUnit)
    : [targetUnit];

  const studentsData = getSheetDataCached_('MASTER', CONFIG.SHEETS.STUDENTS, 300);
  const headers = studentsData.headers;
  const rows = studentsData.rows;

  const col = {
    studentId: findColumnIndex_(headers, ['studentId', 'StudentID']),
    grade: findColumnIndex_(headers, ['grade', '学年']),
    unit: findColumnIndex_(headers, ['unit', '組・コース']),
    attendanceNumber: findColumnIndex_(headers, ['attendanceNumber', '出席番号']),
    name: findColumnIndex_(headers, ['name', '氏名']),
    status: findColumnIndex_(headers, ['status', '在籍状態'])
  };

  const result = rows
    .filter(function(row) {
      const rowGrade = String(row[col.grade] || '').trim();
      const rowUnit = String(row[col.unit] || '').trim();
      const rowStatus = col.status !== -1 ? String(row[col.status] || '').trim().toLowerCase() : 'active';

      if (rowGrade !== targetGrade) return false;
      if (targetUnits.indexOf(rowUnit) === -1) return false;
      if (rowStatus && rowStatus !== 'active') return false;
      return true;
    })
    .map(function(row) {
      return {
        studentId: String(row[col.studentId] || '').trim(),
        grade: String(row[col.grade] || '').trim(),
        unit: String(row[col.unit] || '').trim(),
        attendanceNumber: row[col.attendanceNumber],
        name: String(row[col.name] || '').trim(),
        status: col.status !== -1 ? String(row[col.status] || '').trim() : 'active'
      };
    })
    .sort(compareStudentsByAttendanceNumber_);

  putScriptCacheJson_(cacheKey, result, 300);
  return result;
}

function buildHomeroomShrClassId_(grade, unit) {
  return 'HR_G' + String(grade || '').trim() + '_' + String(unit || '').trim();
}

function buildHomeroomShrClassLabel_(grade, unit) {
  return String(grade || '').trim() + '年 ' + String(unit || '').trim() + '組';
}

function buildHomeroomShrDailyCacheKey_(grade, unit, date) {
  return 'homeroomShrDaily__' +
    String(grade || '').trim() + '__' +
    String(unit || '').trim() + '__' +
    formatDateToYmd(date);
}

