const HOMEROOM_SHR_CONFIG = {
  PERIOD: 0,
  ACTION_TYPE: 'homeroom-shr',
  MODE_LABEL: '担任SHR',
  ALLOWED_STATUS_CODES: ['', 'A', 'L', 'E']
};

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

   const today = new Date();
  today.setHours(0, 0, 0, 0);

  // 昨日以前
  const endDate = new Date(today);
  endDate.setDate(endDate.getDate() - 1);

  // 当年度の 4/7 を開始日にする
  const schoolYearStart = getHomeroomShrSummaryStartDate_(today);
  const startDate = new Date(schoolYearStart);

  const startYmd = formatDateToYmd(startDate);
  const endYmd = formatDateToYmd(endDate);

  const cacheKey = buildHomeroomShrUnsavedSummaryCacheKey_(targetGrade, targetUnit, endYmd);
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const classId = buildHomeroomShrClassId_(targetGrade, targetUnit);
  const targetPeriod = String(HOMEROOM_SHR_CONFIG.PERIOD);

  const classDayYmdList = getHomeroomShrClassDayYmdList_(startYmd, endYmd);

  const sessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.ATTENDANCE_SESSIONS, 60);
  const headers = Array.isArray(sessionsData && sessionsData.headers) ? sessionsData.headers : [];
  const rows = Array.isArray(sessionsData && sessionsData.rows) ? sessionsData.rows : [];

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

  const savedDateMap = {};

  rows.forEach(function(row) {
    const rowClassId = String(row[col.classId] || '').trim();
    const rowDate = formatDateToYmd(row[col.date]);

    // period=0 対応。|| '' を使わない
    const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();

    const actionType = col.actionType !== -1
      ? String(row[col.actionType] || '').trim()
      : '';

    if (rowClassId !== classId) return;
    if (!rowDate || rowDate < startYmd || rowDate > endYmd) return;
    if (rowPeriod !== targetPeriod) return;
    if (actionType && actionType !== HOMEROOM_SHR_CONFIG.ACTION_TYPE) return;

    savedDateMap[rowDate] = true;
  });

  let unsavedCount = 0;
  classDayYmdList.forEach(function(ymd) {
    if (!savedDateMap[ymd]) {
      unsavedCount += 1;
    }
  });

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

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const endDate = new Date(today);
  endDate.setDate(endDate.getDate() - 1);

  const startDate = new Date(getHomeroomShrSummaryStartDate_(today));

  const startYmd = formatDateToYmd(startDate);
  const endYmd = formatDateToYmd(endDate);

  const cacheKey = buildHomeroomShrUnsavedDetailsCacheKey_(targetGrade, targetUnit, endYmd);
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const classId = buildHomeroomShrClassId_(targetGrade, targetUnit);
  const targetPeriod = String(HOMEROOM_SHR_CONFIG.PERIOD);

  const classDayYmdList = getHomeroomShrClassDayYmdList_(startYmd, endYmd);

  const sessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.ATTENDANCE_SESSIONS, 60);
  const headers = Array.isArray(sessionsData && sessionsData.headers) ? sessionsData.headers : [];
  const rows = Array.isArray(sessionsData && sessionsData.rows) ? sessionsData.rows : [];

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

  const savedDateMap = {};

  rows.forEach(function(row) {
    const rowClassId = String(row[col.classId] || '').trim();
    const rowDate = formatDateToYmd(row[col.date]);
    const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();
    const actionType = col.actionType !== -1 ? String(row[col.actionType] || '').trim() : '';

    if (rowClassId !== classId) return;
    if (!rowDate || rowDate < startYmd || rowDate > endYmd) return;
    if (rowPeriod !== targetPeriod) return;
    if (actionType && actionType !== HOMEROOM_SHR_CONFIG.ACTION_TYPE) return;

    savedDateMap[rowDate] = true;
  });

  const items = classDayYmdList
    .filter(function(ymd) {
      return !savedDateMap[ymd];
    })
    .map(function(ymd) {
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

function getHomeroomShrClassDayYmdList_(startYmd, endYmd) {
  const calendarData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CALENDAR, 300);
  const headers = Array.isArray(calendarData && calendarData.headers) ? calendarData.headers : [];
  const rows = Array.isArray(calendarData && calendarData.rows) ? calendarData.rows : [];

  const col = {
    date: headers.indexOf('date'),
    isClassDay: headers.indexOf('isClassDay')
  };

  // calendar が未整備でも止めず、平日ベースでフォールバック
  if (col.date === -1 || col.isClassDay === -1 || !rows.length) {
    return buildWeekdayYmdList_(startYmd, endYmd);
  }

  const result = [];

  rows.forEach(function(row) {
    const rowDate = formatDateToYmd(row[col.date]);
    if (!rowDate) return;
    if (rowDate < startYmd || rowDate > endYmd) return;

    const raw = row[col.isClassDay];
    const isClassDay = raw === true || String(raw || '').trim().toUpperCase() === 'TRUE' || String(raw || '').trim() === '1';

    if (isClassDay) {
      result.push(rowDate);
    }
  });

  result.sort();
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

function saveHomeroomShrAttendance(payload) {
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



    const values = attendanceSheet.getDataRange().getValues();
    const headers = values.length > 0 ? values[0] : [];
    const rows = values.length > 1 ? values.slice(1) : [];

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

    const newRows = records
      .filter(function(record) {
        return String(record.statusCode || '').trim() !== '';
      })
      .map(function(record) {
        return [
          classId,
          targetDate,
          HOMEROOM_SHR_CONFIG.PERIOD,
          String(record.studentId).trim(),
          String(record.statusCode).trim(),
          now
        ];
      });

    const targetStudentIds = {};
    records.forEach(function(record) {
      const studentId = String(record.studentId || '').trim();
      if (studentId) {
        targetStudentIds[studentId] = true;
      }
    });

    const keptRows = rows.filter(function(row) {
      const rowClassId = String(row[col.classId] || '').trim();
      const rowDate = formatDateToYmd(row[col.date]);
      const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();
      const rowStudentId = String(row[col.studentId] || '').trim();

      const isSameSession =
        rowClassId === classId &&
        rowDate === targetDate &&
        rowPeriod === period;

      const isTargetStudent = !!targetStudentIds[rowStudentId];

      return !(isSameSession && isTargetStudent);
    });

    const rebuiltRows = keptRows.concat(newRows);

    const lastRow = attendanceSheet.getLastRow();
    const lastColumn = attendanceSheet.getLastColumn();

    if (lastRow > 1) {
      attendanceSheet.getRange(2, 1, lastRow - 1, lastColumn).clearContent();
    }

    if (rebuiltRows.length > 0) {
      attendanceSheet
        .getRange(2, 1, rebuiltRows.length, rebuiltRows[0].length)
        .setValues(rebuiltRows);
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

    const summaryBaseDate = new Date();
    summaryBaseDate.setHours(0, 0, 0, 0);
    summaryBaseDate.setDate(summaryBaseDate.getDate() - 1);
    const summaryEndYmd = formatDateToYmd(summaryBaseDate);

      removeScriptCacheKeys_([
    buildHomeroomShrDailyCacheKey_(grade, unit, targetDate),
    buildHomeroomShrUnsavedSummaryCacheKey_(grade, unit, summaryEndYmd),
    buildHomeroomShrUnsavedDetailsCacheKey_(grade, unit, summaryEndYmd)
  ]);

return {
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

