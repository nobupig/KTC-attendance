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

    const newRows = attendance
      .filter(function(record) {
        return String(record.statusCode || "").trim() !== "";
      })
      .map(function(record) {
        return [
          targetClassId,
          targetDate,
          Number(targetPeriod),
          String(record.studentId).trim(),
          String(record.statusCode).trim(),
          now
        ];
      });

    const matchedRowNumbers = [];
    rows.forEach(function(row, i) {
      const rowClassId = String(row[col.classId] || '').trim();
      const rowDate = formatDateToYmd(row[col.date]);
      const rowPeriod = String(row[col.period] || '').trim();

      if (
        rowClassId === targetClassId &&
        rowDate === targetDate &&
        rowPeriod === targetPeriod
      ) {
        matchedRowNumbers.push(i + 2);
      }
    });

    if (
      matchedRowNumbers.length > 0 &&
      matchedRowNumbers.length === newRows.length &&
      isSequentialRows_(matchedRowNumbers) &&
      newRows.length > 0
    ) {
      attendanceSheet
        .getRange(matchedRowNumbers[0], 1, newRows.length, newRows[0].length)
        .setValues(newRows);

    } else if (matchedRowNumbers.length === 0) {
      if (newRows.length > 0) {
        attendanceSheet
          .getRange(attendanceSheet.getLastRow() + 1, 1, newRows.length, newRows[0].length)
          .setValues(newRows);
      }

    } else {
      const filteredRows = rows.filter(function(row) {
        const rowClassId = String(row[col.classId] || '').trim();
        const rowDate = formatDateToYmd(row[col.date]);
        const rowPeriod = String(row[col.period] || '').trim();

        return !(
          rowClassId === targetClassId &&
          rowDate === targetDate &&
          rowPeriod === targetPeriod
        );
      });

      const rebuiltRows = filteredRows.concat(newRows);

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
    }

    invalidateAttendanceCaches_(targetClassId, targetDate, targetPeriod);

    return {
      success: true,
      savedCount: attendance.length,
      mode: allowPastEdit ? 'past-edit' : 'normal',
      actionType: actionType,
      targetSessionKey: targetSessionKey
    };

  } finally {
    lock.releaseLock();
  }
}

function getAttendanceMap(classId, date, period) {
  const targetClassId = String(classId || '').trim();
  const targetDate = formatDateToYmd(date);
  const targetPeriod = String(period || '').trim();

  const sessionCacheKey = buildAttendanceSessionCacheKey_(targetClassId, targetDate, targetPeriod);
  const cached = getScriptCacheJson_(sessionCacheKey);
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

  const result = {};

  rows.forEach(function(row) {
    const rowClassId = String(row[col.classId] || '').trim();
    const rowDate = formatDateToYmd(row[col.date]);
    const rowPeriod = String(row[col.period] || '').trim();

    if (
      rowClassId === targetClassId &&
      rowDate === targetDate &&
      rowPeriod === targetPeriod
    ) {
      const rowStudentId = String(row[col.studentId] || '').trim();
      const rowStatusCode = String(row[col.statusCode] || '').trim();
      result[rowStudentId] = rowStatusCode;
    }
  });

  putScriptCacheJson_(sessionCacheKey, result, 60);
  return result;
}

/* =========================
 * 内部ヘルパー
 * ========================= */

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
    buildAttendanceSessionCacheKey_(classId, date, period)
  ]);
}