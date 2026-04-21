function saveAttendance(payload) {
  return saveAttendanceInternal_(payload, false);
}

function savePastAttendance(payload) {
  return saveAttendanceInternal_(payload, true);
}

function saveAttendanceInternal_(payload, allowPastEdit) {
  const ss = getOperationSpreadsheet();
  const attendanceSessionsSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_SESSIONS);
  const attendanceSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (!payload) {
      throw new Error("保存データがありません");
    }

    const currentUserEmail = getCurrentUserEmail();
    const now = new Date();

    const targetClassId = String(payload.classId || "").trim();
    const targetDate = formatDateToYmd(payload.date);
    const targetPeriod = String(payload.period || "").trim();
    const targetSessionKey = [targetClassId, targetDate, targetPeriod].join("__");
    const actionType = allowPastEdit ? "past-edit" : "normal";
    const savedModeLabel = allowPastEdit ? "過去修正" : "通常入力";

    if (!targetClassId || !targetDate || !targetPeriod) {
      throw new Error("保存に必要な授業情報が不足しています");
    }

    const session = {
      classId: targetClassId,
      date: targetDate,
      period: targetPeriod
    };

    if (!canEditAttendance(session)) {
      throw new Error("この授業の出席を編集する権限がありません");
    }

    if (!allowPastEdit && !isAttendanceEditable(targetDate)) {
      throw new Error("出席入力の期限を過ぎています");
    }

    const attendance = Array.isArray(payload.attendance) ? payload.attendance : [];
    const allowedStatusCodes = ["P", "A", "L", "O", ""];

    attendance.forEach(function(record) {
      const studentId = String(record.studentId || "").trim();
      const statusCode = String(record.statusCode || "").trim();

      if (!studentId) {
        throw new Error("studentId が不正なデータがあります");
      }

      if (!allowedStatusCodes.includes(statusCode)) {
        throw new Error("statusCode が不正です: " + statusCode);
      }
    });



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

    const targetStudentIds = {};
    const desiredByStudentId = {};

    attendance.forEach(function(record) {
      const studentId = String(record.studentId || '').trim();
      const statusCode = String(record.statusCode || '').trim();
      if (!studentId) return;
      targetStudentIds[studentId] = true;
      desiredByStudentId[studentId] = statusCode;
    });

    const existingRowNumberByStudentId = {};
    rows.forEach(function(row, index) {
      const rowClassId = String(row[col.classId] || '').trim();
      const rowDate = formatDateToYmd(row[col.date]);
      const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();
      const rowStudentId = String(row[col.studentId] || '').trim();

      if (
        rowClassId === targetClassId &&
        rowDate === targetDate &&
        rowPeriod === targetPeriod &&
        targetStudentIds[rowStudentId]
      ) {
        existingRowNumberByStudentId[rowStudentId] = index + 2;
      }
    });

    const rowsToClear = [];
    const rowsToUpdate = [];
    const appendRows = [];

    Object.keys(targetStudentIds).forEach(function(studentId) {
      const statusCode = desiredByStudentId[studentId] || '';
      const existingRowNumber = existingRowNumberByStudentId[studentId];

      if (!statusCode) {
        if (existingRowNumber) {
          rowsToClear.push(existingRowNumber);
        }
        return;
      }

      const fullRow = buildAttendanceSheetRow_(headers.length, col, {
        classId: targetClassId,
        date: targetDate,
        period: Number(targetPeriod),
        studentId: studentId,
        statusCode: statusCode,
        recordedAt: now
      });

      if (existingRowNumber) {
        rowsToUpdate.push({ rowNumber: existingRowNumber, values: fullRow });
      } else {
        appendRows.push(fullRow);
      }
    });

    rowsToClear.sort(function(a, b) { return a - b; });
    rowsToUpdate.sort(function(a, b) { return a.rowNumber - b.rowNumber; });

    if (rowsToClear.length > 0) {
      if (isSequentialRows_(rowsToClear)) {
        attendanceSheet.getRange(rowsToClear[0], 1, rowsToClear.length, headers.length).clearContent();
      } else {
        rowsToClear.forEach(function(rowNumber) {
          attendanceSheet.getRange(rowNumber, 1, 1, headers.length).clearContent();
        });
      }
    }

    if (rowsToUpdate.length > 0) {
      const rowNumbers = rowsToUpdate.map(function(item) { return item.rowNumber; });
      if (isSequentialRows_(rowNumbers)) {
        attendanceSheet.getRange(rowNumbers[0], 1, rowsToUpdate.length, headers.length)
          .setValues(rowsToUpdate.map(function(item) { return item.values; }));
      } else {
        rowsToUpdate.forEach(function(item) {
          attendanceSheet.getRange(item.rowNumber, 1, 1, headers.length).setValues([item.values]);
        });
      }
    }

    if (appendRows.length > 0) {
      const startRow = Math.max(attendanceSheet.getLastRow(), 1) + 1;
      attendanceSheet.getRange(startRow, 1, appendRows.length, headers.length).setValues(appendRows);
    }
    appendAttendanceSessionLog_(attendanceSessionsSheet, [
  targetClassId,
  targetDate,
  Number(targetPeriod),
  currentUserEmail,
  now,
  actionType,
  targetSessionKey,
  savedModeLabel
]);
    invalidateAttendanceCaches_(targetClassId, targetDate, targetPeriod);

    const lastSavedInfo = {
      teacherEmail: currentUserEmail,
      savedAt: now,
      savedAtText: formatDateTimeJst_(now),
      actionType: actionType,
      targetSessionKey: targetSessionKey,
      savedModeLabel: savedModeLabel,
      savedByCurrentUser: true
    };

    return {
      success: true,
      savedCount: attendance.length,
      updatedCount: rowsToUpdate.length,
      appendedCount: appendRows.length,
      clearedCount: rowsToClear.length,
      mode: allowPastEdit ? 'past-edit' : 'normal',
      actionType: actionType,
      targetSessionKey: targetSessionKey,
      lastSavedInfo: lastSavedInfo
    };

  } finally {
    lock.releaseLock();
  }
}

function getAttendanceMap(classId, date, period) {
  const totalStartedAt = perfNow_();

  const targetClassId = String(classId || '').trim();
  const targetDate = formatDateToYmd(date);
  const targetPeriod = String(period || '').trim();

  const sessionCacheKey = buildAttendanceSessionCacheKey_(targetClassId, targetDate, targetPeriod);

  const cacheStartedAt = perfNow_();
  const cached = getScriptCacheJson_(sessionCacheKey);
  logPerf_(
    'getAttendanceMap getScriptCacheJson_',
    cacheStartedAt,
    'cacheKey=' + sessionCacheKey + ' hit=' + (!!cached)
  );

  if (cached) {
    logPerf_(
      'getAttendanceMap total',
      totalStartedAt,
      'cache=hit entries=' + Object.keys(cached).length + ' key=' + [targetClassId, targetDate, targetPeriod].join('__')
    );
    return cached;
  }

  const indexStartedAt = perfNow_();
  const attendanceIndex = buildAttendanceIndex_();
  logPerf_('getAttendanceMap buildAttendanceIndex_', indexStartedAt);

  const key = [targetClassId, targetDate, targetPeriod].join('__');
  const result = attendanceIndex[key] || {};

  const putCacheStartedAt = perfNow_();
  putScriptCacheJson_(sessionCacheKey, result, 60);
  logPerf_('getAttendanceMap putScriptCacheJson_', putCacheStartedAt, 'entries=' + Object.keys(result).length);

  logPerf_(
    'getAttendanceMap total',
    totalStartedAt,
    'cache=miss entries=' + Object.keys(result).length + ' key=' + key
  );

  return result;
}

/* =========================
 * 内部ヘルパー
 * ========================= */

function buildAttendanceSheetRow_(headerCount, col, record) {
  const row = new Array(headerCount).fill('');
  row[col.classId] = record.classId;
  row[col.date] = record.date;
  row[col.period] = record.period;
  row[col.studentId] = record.studentId;
  row[col.statusCode] = record.statusCode;
  row[col.recordedAt] = record.recordedAt;
  return row;
}

function appendAttendanceSessionLog_(sheet, baseRow) {
  const headerCount = sheet.getLastColumn();

  if (headerCount <= 5) {
    sheet.appendRow(baseRow.slice(0, 5));
    return;
  }

  const row = baseRow.slice();
  while (row.length < headerCount) {
    row.push("");
  }
  sheet.appendRow(row.slice(0, headerCount));
}

function isSequentialRows_(rowNumbers) {
  if (!rowNumbers || rowNumbers.length <= 1) {
    return true;
  }

  for (var i = 1; i < rowNumbers.length; i++) {
    if (rowNumbers[i] !== rowNumbers[i - 1] + 1) {
      return false;
    }
  }
  return true;
}


function invalidateAttendanceCaches_(classId, date, period) {
  removeScriptCacheKeys_([
    getAttendanceSheetCacheKey_(),
    getAttendanceSessionsSheetCacheKey_(),
    buildAttendanceSessionCacheKey_(classId, date, period),
    'attendanceIndex__all',
    'attendanceSessionLatestIndex__all'
  ]);
}

function normalizeAttendanceActionTypes_(actionTypes) {
  if (!actionTypes) return null;

  const list = Array.isArray(actionTypes) ? actionTypes : [actionTypes];
  const normalized = list
    .map(function(item) { return String(item || '').trim(); })
    .filter(Boolean);

  return normalized.length ? normalized : null;
}

function formatDateTimeJst_(value) {
  if (!value) return '';
  const d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return '';
  return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

function getLatestAttendanceSessionInfo_(classId, date, period, allowedActionTypes) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const targetClassId = String(classId || '').trim();
  const targetDate = formatDateToYmd(date);
  const targetPeriod = String(period || '').trim();

  if (!targetClassId || !targetDate || !targetPeriod) {
    if (typeof logPerf_ === 'function') {
      logPerf_('getLatestAttendanceSessionInfo_ total', totalStartedAt, 'invalid-args');
    }
    return null;
  }

  const mapStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const latestMap = getAttendanceSessionLatestMapByDateCached_(targetDate);
  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getLatestAttendanceSessionInfo_ getAttendanceSessionLatestMapByDateCached_',
      mapStartedAt,
      'date=' + targetDate + ' keys=' + Object.keys(latestMap).length
    );
  }

  const key = [targetClassId, targetDate, targetPeriod].join('__');
  const latest = latestMap[key];
  if (!latest) {
    if (typeof logPerf_ === 'function') {
      logPerf_('getLatestAttendanceSessionInfo_ total', totalStartedAt, 'not-found key=' + key);
    }
    return null;
  }

  const allowed = Array.isArray(allowedActionTypes) && allowedActionTypes.length
    ? allowedActionTypes.map(function(v) { return String(v || '').trim(); })
    : null;

  if (allowed && allowed.indexOf(String(latest.actionType || '').trim()) === -1) {
    if (typeof logPerf_ === 'function') {
      logPerf_(
        'getLatestAttendanceSessionInfo_ total',
        totalStartedAt,
        'filtered-out actionType=' + String(latest.actionType || '')
      );
    }
    return null;
  }

  const result = {
    teacherEmail: latest.teacherEmail,
    savedAt: latest.savedAt,
    savedAtText: latest.savedAtText,
    actionType: latest.actionType,
    targetSessionKey: latest.targetSessionKey,
    savedModeLabel: latest.savedModeLabel
  };

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getLatestAttendanceSessionInfo_ total',
      totalStartedAt,
      'found key=' + key + ' actionType=' + (result.actionType || '')
    );
  }

  return result;
}

function getAttendanceSessionLatestMapByDateCached_(ymd) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const cacheKey = 'attendanceSessionLatestMapByDate__' + ymd;
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    if (typeof logPerf_ === 'function') {
      logPerf_(
        'getAttendanceSessionLatestMapByDateCached_ total',
        totalStartedAt,
        'cache=hit date=' + ymd + ' keys=' + Object.keys(cached).length
      );
    }
    return cached;
  }

  const loadStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const attendanceSessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.ATTENDANCE_SESSIONS, 60);
  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getAttendanceSessionLatestMapByDateCached_ load attendanceSessionsData',
      loadStartedAt,
      'rows=' + attendanceSessionsData.rows.length
    );
  }

  const headersStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const headers = attendanceSessionsData.headers;
  const rows = attendanceSessionsData.rows;

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

  if (typeof logPerf_ === 'function') {
    logPerf_('getAttendanceSessionLatestMapByDateCached_ resolve headers', headersStartedAt);
  }

  const buildStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const latestMap = {};

  rows.forEach(function(row) {
    const rowDate = formatDateToYmd(row[col.date]);
    if (rowDate !== ymd) return;

    const rowClassId = String(row[col.classId] || '').trim();
    const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();
    if (!rowClassId || !rowPeriod) return;

    const key = [rowClassId, rowDate, rowPeriod].join('__');

    const teacherEmail = col.teacherEmail !== -1
      ? String(row[col.teacherEmail] || '').trim().toLowerCase()
      : '';

    const accessedAtRaw = col.accessedAt !== -1 ? row[col.accessedAt] : '';
    const accessedAt = accessedAtRaw instanceof Date ? accessedAtRaw : new Date(accessedAtRaw);
    const accessedAtMs = isNaN(accessedAt.getTime()) ? 0 : accessedAt.getTime();

    if (!latestMap[key] || accessedAtMs >= latestMap[key]._ms) {
      latestMap[key] = {
        teacherEmail: teacherEmail,
        savedAt: accessedAtRaw,
        savedAtText: formatDateTimeJst_(accessedAtRaw),
        actionType: col.actionType !== -1 ? String(row[col.actionType] || '').trim() : '',
        targetSessionKey: col.targetSessionKey !== -1 ? String(row[col.targetSessionKey] || '').trim() : '',
        savedModeLabel: col.savedModeLabel !== -1 ? String(row[col.savedModeLabel] || '').trim() : '',
        _ms: accessedAtMs
      };
    }
  });

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getAttendanceSessionLatestMapByDateCached_ build latestMap',
      buildStartedAt,
      'date=' + ymd + ' keys=' + Object.keys(latestMap).length
    );
  }

  putScriptCacheJson_(cacheKey, latestMap, 60);

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getAttendanceSessionLatestMapByDateCached_ total',
      totalStartedAt,
      'cache=miss date=' + ymd + ' keys=' + Object.keys(latestMap).length
    );
  }

  return latestMap;
}


function hasAttendanceSessionRecord_(classId, date, period, actionTypes) {
  return !!getLatestAttendanceSessionInfo_(classId, date, period, actionTypes);
}

function buildAttendanceIndex_() {
  const cacheKey = 'attendanceIndex__all';
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const attendanceData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.ATTENDANCE, 60);
  const headers = attendanceData.headers;
  const rows = attendanceData.rows;

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

  const index = {};

  rows.forEach(function(row) {
    const classId = String(row[col.classId] || '').trim();
    const date = formatDateToYmd(row[col.date]);
    const period = String(row[col.period] == null ? '' : row[col.period]).trim();
    const studentId = String(row[col.studentId] || '').trim();
    const statusCode = String(row[col.statusCode] || '').trim();

    if (!classId || !date || !period || !studentId) return;

    const sessionKey = [classId, date, period].join('__');
    if (!index[sessionKey]) {
      index[sessionKey] = {};
    }
    index[sessionKey][studentId] = statusCode;
  });

  putScriptCacheJson_(cacheKey, index, 60);
  return index;
}

function buildAttendanceSessionLatestIndex_() {
  const cacheKey = 'attendanceSessionLatestIndex__all';
  const cached = getScriptCacheJson_(cacheKey);
  if (cached) {
    return cached;
  }

  const data = getSheetDataCached_('OPERATION', CONFIG.SHEETS.ATTENDANCE_SESSIONS, 60);
  const headers = data.headers;
  const rows = data.rows;

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

  const index = {};

  rows.forEach(function(row) {
    const classId = normalizeString_(row[col.classId]);
    const date = formatDateToYmd(row[col.date]);
    const period = normalizeString_(row[col.period]);
    const actionType = col.actionType !== -1 ? normalizeString_(row[col.actionType]) : '';
    const teacherEmail = col.teacherEmail !== -1 ? normalizeString_(row[col.teacherEmail]).toLowerCase() : '';
    const accessedAtRaw = col.accessedAt !== -1 ? row[col.accessedAt] : '';
    const accessedAt = accessedAtRaw instanceof Date ? accessedAtRaw : new Date(accessedAtRaw);
    const accessedAtMs = isNaN(accessedAt.getTime()) ? 0 : accessedAt.getTime();

    if (!classId || !date || !period) return;

    const sessionKey = [classId, date, period].join('__');

    if (!index[sessionKey] || accessedAtMs >= index[sessionKey]._ms) {
      index[sessionKey] = {
        teacherEmail: teacherEmail,
        savedAt: accessedAtRaw,
        savedAtText: formatDateTimeJst_(accessedAtRaw),
        actionType: actionType,
        targetSessionKey: col.targetSessionKey !== -1 ? normalizeString_(row[col.targetSessionKey]) : '',
        savedModeLabel: col.savedModeLabel !== -1 ? normalizeString_(row[col.savedModeLabel]) : '',
        _ms: accessedAtMs
      };
    }
  });

  putScriptCacheJson_(cacheKey, index, 60);
  return index;
}