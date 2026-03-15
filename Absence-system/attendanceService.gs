function saveAttendance(payload) {
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
    const targetDate = String(payload.date || "").trim();
    const targetPeriod = Number(payload.period);

    if (!targetClassId || !targetDate || Number.isNaN(targetPeriod)) {
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

    if (!isAttendanceEditable(targetDate)) {
      throw new Error("出席入力の期限を過ぎています");
    }

    const attendance = Array.isArray(payload.attendance) ? payload.attendance : [];
    const allowedStatusCodes = ["P","A","L","O"];

    attendance.forEach(record => {
      const studentId = String(record.studentId || "").trim();
      const statusCode = String(record.statusCode || "").trim();

      if (!studentId) {
        throw new Error("studentId が不正なデータがあります");
      }

      if (!allowedStatusCodes.includes(statusCode)) {
        throw new Error("statusCode が不正です: " + statusCode);
      }
    });

    attendanceSessionsSheet.appendRow([
      targetClassId,
      targetDate,
      targetPeriod,
      currentUserEmail,
      now
    ]);

    const values = attendanceSheet.getDataRange().getValues();

    if (values.length > 1) {
      const rows = values.slice(1);
      const deleteRows = [];

      rows.forEach((row, i) => {
        const rowClassId = String(row[0]).trim();
        const rowDate = formatDateToYmd(row[1]);
        const rowPeriod = Number(row[2]);

        if (
          rowClassId === targetClassId &&
          rowDate === targetDate &&
          rowPeriod === targetPeriod
        ) {
          deleteRows.push(i + 2);
        }
      });

      deleteRows.reverse().forEach(rowNumber => attendanceSheet.deleteRow(rowNumber));
    }

    if (attendance.length > 0) {
      const newRows = attendance.map(record => [
        targetClassId,
        targetDate,
        targetPeriod,
        String(record.studentId).trim(),
        String(record.statusCode).trim(),
        now
      ]);

      attendanceSheet
        .getRange(
          attendanceSheet.getLastRow() + 1,
          1,
          newRows.length,
          newRows[0].length
        )
        .setValues(newRows);
    }

    return {
      success: true,
      savedCount: attendance.length
    };

  } finally {
    lock.releaseLock();
  }
}

function getAttendanceMap(classId, date, period) {
  const ss = getOperationSpreadsheet();
  const attendanceSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);
  if (!attendanceSheet) {
    throw new Error('attendance シートがありません');
  }

  const values = attendanceSheet.getDataRange().getValues();
  if (values.length <= 1) {
    return {};
  }

  const headers = values[0];
  const rows = values.slice(1);

  const col = {
    classId: headers.indexOf('classId'),
    date: headers.indexOf('date'),
    period: headers.indexOf('period'),
    studentId: headers.indexOf('studentId'),
    statusCode: headers.indexOf('statusCode')
  };

  Object.keys(col).forEach(key => {
    if (col[key] === -1) {
      throw new Error('attendance シートに ' + key + ' 列がありません');
    }
  });

  const targetClassId = String(classId).trim();
  const targetDate = formatDateToYmd(date);
  const targetPeriod = String(period);

  const result = {};

  rows.forEach(row => {
    const rowClassId = String(row[col.classId]).trim();
    const rowDate = formatDateToYmd(row[col.date]);
    const rowPeriod = String(row[col.period]);
    const rowStudentId = String(row[col.studentId]).trim();
    const rowStatusCode = String(row[col.statusCode]).trim();

    if (
      rowClassId === targetClassId &&
      rowDate === targetDate &&
      rowPeriod === targetPeriod
    ) {
      result[rowStudentId] = rowStatusCode;
    }
  });

  return result;
}