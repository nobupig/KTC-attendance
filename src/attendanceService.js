function saveAttendance(payload) {
  return saveAttendanceInternal_(payload, false);
}

function savePastAttendance(payload) {
  return saveAttendanceInternal_(payload, true);
}

function saveNoAbsenceAttendance(payload) {
  return saveNoAbsenceAttendanceInternal_(payload, false);
}

function savePastNoAbsenceAttendance(payload) {
  return saveNoAbsenceAttendanceInternal_(payload, true);
}

function saveTeacherUnsavedBulkNoAbsence(payload) {
  return saveTeacherUnsavedBulkNoAbsenceInternal_(payload);
}

function saveTeacherUnsavedBulkNoAbsenceInternal_(payload) {

  const ss = getOperationSpreadsheet();

  const attendanceSessionsSheet =
    ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_SESSIONS);

  const attendanceSheet =
    ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);


  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    if (!payload || !Array.isArray(payload.items)) {
      throw new Error("一括保存データがありません");
    }

    const sourceItems = payload.items;

    if (sourceItems.length === 0) {
      throw new Error("一括保存する授業が選択されていません");
    }

    if (sourceItems.length > 100) {
      throw new Error("一度に保存できる授業は100件までです");
    }

    if (!attendanceSessionsSheet) {
      throw new Error("attendanceSessions シートが見つかりません");
    }

    if (!attendanceSheet) {
      throw new Error("attendance シートが見つかりません");
    }

    const currentUserEmail = getCurrentUserEmail();
    const now = new Date();
    const todayYmd = formatDateToYmd(now);

    const sessions = [];
    const targetKeySet = {};
    let duplicateInputCount = 0;

    sourceItems.forEach(function(item, index) {
      if (!item || typeof item !== "object") {
        throw new Error(
          "一括保存データの " + (index + 1) + " 件目が不正です"
        );
      }

      const classId = String(item.classId || "").trim();
      const date = formatDateToYmd(item.date);
      const period = String(
        item.period == null ? "" : item.period
      ).trim();

      if (!classId || !date || !period) {
        throw new Error(
          "一括保存データの " + (index + 1) +
          " 件目に必要な授業情報がありません"
        );
      }

      if (date >= todayYmd) {
        throw new Error(
          "一括登録できるのは過去の未保存授業だけです"
        );
      }

      if (
        typeof isExperimentGroupTargetClass_ === "function" &&
        isExperimentGroupTargetClass_(classId)
      ) {
        throw new Error(
          "工学実験は一括登録できません。名簿から個別に保存してください"
        );
      }

      const session = {
        classId: classId,
        date: date,
        period: period,
        sessionNumber: item.sessionNumber || ""
      };

      if (!canEditAttendance(session)) {
        throw new Error(
          "編集権限のない授業が選択されています: " +
          classId + " / " + date + " / " + period + "限"
        );
      }

      const targetSessionKey =
        [classId, date, period].join("__");

      if (targetKeySet[targetSessionKey]) {
        duplicateInputCount++;
        return;
      }

      targetKeySet[targetSessionKey] = true;

      session.targetSessionKey = targetSessionKey;
      sessions.push(session);
    });

    if (sessions.length === 0) {
      throw new Error("保存対象の授業がありません");
    }


    /*
     * 一覧表示後に別操作で保存された授業を
     * 「欠席者なし」で上書きしないための競合確認。
     * ScriptLock中にattendanceSessionsを1回だけ読み込む。
     */
    const sessionLastRow =
      attendanceSessionsSheet.getLastRow();

    const sessionLastColumn =
      attendanceSessionsSheet.getLastColumn();

    const sessionHeaders =
      sessionLastColumn > 0
        ? attendanceSessionsSheet
            .getRange(1, 1, 1, sessionLastColumn)
            .getDisplayValues()[0]
        : [];

    const sessionCol = {
      classId: findColumnIndex_(
        sessionHeaders,
        ["classId", "ClassID"]
      ),
      date: findColumnIndex_(
        sessionHeaders,
        ["date", "日付"]
      ),
      period: findColumnIndex_(
        sessionHeaders,
        ["period", "時限"]
      ),
      targetSessionKey: findColumnIndex_(
        sessionHeaders,
        ["targetSessionKey"]
      )
    };

    ["classId", "date", "period"].forEach(function(key) {
      if (sessionCol[key] === -1) {
        throw new Error(
          "attendanceSessions シートに " +
          key +
          " 列がありません"
        );
      }
    });

    const sessionScanIndexes = [
      sessionCol.classId,
      sessionCol.date,
      sessionCol.period
    ];

    if (sessionCol.targetSessionKey !== -1) {
      sessionScanIndexes.push(
        sessionCol.targetSessionKey
      );
    }

    const sessionScanStartIndex =
      Math.min.apply(null, sessionScanIndexes);

    const sessionScanEndIndex =
      Math.max.apply(null, sessionScanIndexes);

    const sessionScanWidth =
      sessionScanEndIndex -
      sessionScanStartIndex +
      1;

    const sessionDataRowCount =
      Math.max(sessionLastRow - 1, 0);

    const sessionScanValues =
      sessionDataRowCount > 0
        ? attendanceSessionsSheet
            .getRange(
              2,
              sessionScanStartIndex + 1,
              sessionDataRowCount,
              sessionScanWidth
            )
            .getDisplayValues()
        : [];


    const relativeSessionCol = {
      classId:
        sessionCol.classId -
        sessionScanStartIndex,
      date:
        sessionCol.date -
        sessionScanStartIndex,
      period:
        sessionCol.period -
        sessionScanStartIndex,
      targetSessionKey:
        sessionCol.targetSessionKey === -1
          ? -1
          : sessionCol.targetSessionKey -
            sessionScanStartIndex
    };

    const alreadySavedKeySet = {};

    sessionScanValues.forEach(function(row) {
      /*
       * 新しい行は targetSessionKey を直接比較する。
       * 旧データ等で targetSessionKey が空欄の場合だけ
       * classId/date/period からfallback keyを生成する。
       */
      if (
        relativeSessionCol.targetSessionKey !== -1
      ) {
        const storedKey = String(
          row[
            relativeSessionCol.targetSessionKey
          ] || ""
        ).trim();

        if (storedKey) {
          if (targetKeySet[storedKey]) {
            alreadySavedKeySet[storedKey] = true;
          }
          return;
        }
      }

      const rowClassId = String(
        row[relativeSessionCol.classId] || ""
      ).trim();

      const rowDate =
        normalizeYmdDisplayText_(
          row[relativeSessionCol.date]
        );

      const rowPeriod = String(
        row[relativeSessionCol.period] || ""
      ).trim();

      if (
        !rowClassId ||
        !rowDate ||
        !rowPeriod
      ) {
        return;
      }

      const fallbackKey =
        [
          rowClassId,
          rowDate,
          rowPeriod
        ].join("__");

      if (targetKeySet[fallbackKey]) {
        alreadySavedKeySet[fallbackKey] = true;
      }
    });

    const existingSavedKeyList =
      Object.keys(alreadySavedKeySet);

    if (existingSavedKeyList.length > 0) {
      throw new Error(
        "選択した授業のうち " +
        existingSavedKeyList.length +
        " 件が既に保存されています。" +
        "未保存授業一覧を再読み込みしてから、もう一度実行してください"
      );
    }


    /*
     * attendanceは1回だけ読み込み、
     * 選択セッションに属する既存例外行をまとめて消す。
     */
    const attendanceLastRow =
      attendanceSheet.getLastRow();

    const attendanceLastColumn =
      attendanceSheet.getLastColumn();

    const attendanceHeaders =
      attendanceLastColumn > 0
        ? attendanceSheet
            .getRange(
              1,
              1,
              1,
              attendanceLastColumn
            )
            .getDisplayValues()[0]
        : [];

    const attendanceCol = {
      classId:
        attendanceHeaders.indexOf(
          "classId"
        ),
      date:
        attendanceHeaders.indexOf(
          "date"
        ),
      period:
        attendanceHeaders.indexOf(
          "period"
        )
    };

    Object.keys(attendanceCol)
      .forEach(function(key) {
        if (attendanceCol[key] === -1) {
          throw new Error(
            "attendance シートに " +
            key +
            " 列がありません"
          );
        }
      });

    const attendanceScanIndexes = [
      attendanceCol.classId,
      attendanceCol.date,
      attendanceCol.period
    ];

    const attendanceScanStartIndex =
      Math.min.apply(
        null,
        attendanceScanIndexes
      );

    const attendanceScanEndIndex =
      Math.max.apply(
        null,
        attendanceScanIndexes
      );

    const attendanceScanWidth =
      attendanceScanEndIndex -
      attendanceScanStartIndex +
      1;

    const attendanceDataRowCount =
      Math.max(
        attendanceLastRow - 1,
        0
      );

    const attendanceScanValues =
      attendanceDataRowCount > 0
        ? attendanceSheet
            .getRange(
              2,
              attendanceScanStartIndex + 1,
              attendanceDataRowCount,
              attendanceScanWidth
            )
            .getDisplayValues()
        : [];


    const relativeAttendanceCol = {
      classId:
        attendanceCol.classId -
        attendanceScanStartIndex,
      date:
        attendanceCol.date -
        attendanceScanStartIndex,
      period:
        attendanceCol.period -
        attendanceScanStartIndex
    };

    const rowsToClear = [];

    attendanceScanValues.forEach(
      function(row, index) {
        const rowClassId = String(
          row[
            relativeAttendanceCol.classId
          ] || ""
        ).trim();

        const rowDate =
          normalizeYmdDisplayText_(
            row[
              relativeAttendanceCol.date
            ]
          );

        const rowPeriod = String(
          row[
            relativeAttendanceCol.period
          ] || ""
        ).trim();

        if (
          !rowClassId ||
          !rowDate ||
          !rowPeriod
        ) {
          return;
        }

        const rowKey =
          [
            rowClassId,
            rowDate,
            rowPeriod
          ].join("__");

        if (targetKeySet[rowKey]) {
          rowsToClear.push(index + 2);
        }
      }
    );


    clearAttendanceRowsByNumberGroups_(
      attendanceSheet,
      rowsToClear,
      attendanceHeaders.length
    );


    const actionType = "past-edit";
    const savedModeLabel = "過去修正（欠席者なし）";

    const logRows = sessions.map(function(session) {
      return [
        session.classId,
        session.date,
        Number(session.period),
        currentUserEmail,
        now,
        actionType,
        session.targetSessionKey,
        savedModeLabel
      ];
    });

    appendAttendanceSessionLogs_(
      attendanceSessionsSheet,
      logRows
    );


    const fastInvalidation =
      tryInvalidateTeacherUnsavedFastSnapshotsAfterBulkSaveUnderLock_(
        sessions,
        actionType
      );


    invalidateAttendanceCachesBulk_(sessions);


    /*
     * legacy / ScriptCache系の未保存キャッシュも
     * 一括処理後に1回だけ無効化する。
     */
    const currentUser = getCurrentUserContext();

    const currentTeacherId =
      currentUser && currentUser.teacherId
        ? normalizeString_(currentUser.teacherId)
        : "";

    const summaryBaseDate = new Date();
    summaryBaseDate.setHours(0, 0, 0, 0);
    summaryBaseDate.setDate(summaryBaseDate.getDate() - 1);

    const summaryEndYmd =
      formatDateToYmd(summaryBaseDate);

    const summaryStartYmd =
      formatDateToYmd(
        getTeacherUnsavedStartDate_(summaryBaseDate)
      );

    const teacherUnsavedCacheKeys = [
      "savedSessionKeySetByRange__" +
        summaryStartYmd + "__" + summaryEndYmd,

      "savedSessionKeySetByRange__v2__" +
        summaryStartYmd + "__" + summaryEndYmd,

      "savedSessionKeySetByRange__v4__" +
        summaryStartYmd + "__" + summaryEndYmd
    ];

    if (currentTeacherId) {
      teacherUnsavedCacheKeys.push(
        buildTeacherUnsavedSummaryCacheKey_(
          currentTeacherId,
          summaryEndYmd
        ),
        buildTeacherUnsavedDetailsCacheKey_(
          currentTeacherId,
          summaryEndYmd
        )
      );
    }

    removeScriptCacheKeys_(teacherUnsavedCacheKeys);



    return {
      success: true,
      mode: "past-edit",
      noAbsence: true,
      requestedCount: sourceItems.length,
      savedCount: sessions.length,
      duplicateInputCount: duplicateInputCount,
      clearedCount: rowsToClear.length,
      fastInvalidation: fastInvalidation,
      items: sessions.map(function(session) {
        return {
          classId: session.classId,
          date: session.date,
          period: session.period,
          sessionNumber: session.sessionNumber,
          targetSessionKey: session.targetSessionKey
        };
      })
    };

  } finally {
    lock.releaseLock();
  }
}
function saveNoAbsenceAttendanceInternal_(payload, allowPastEdit) {
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
    const savedModeLabel = allowPastEdit
      ? "過去修正（欠席者なし）"
      : "通常入力（欠席者なし）";

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

    if (!attendanceSessionsSheet) {
      throw new Error("attendanceSessions シートが見つかりません");
    }

    if (!attendanceSheet) {
      throw new Error("attendance シートが見つかりません");
    }

    const values = attendanceSheet.getDataRange().getValues();
    const headers = values.length > 0 ? values[0] : [];
    const rows = values.length > 1 ? values.slice(1) : [];

    const col = {
      classId: headers.indexOf("classId"),
      date: headers.indexOf("date"),
      period: headers.indexOf("period"),
      studentId: headers.indexOf("studentId"),
      statusCode: headers.indexOf("statusCode"),
      recordedAt: headers.indexOf("recordedAt")
    };

    Object.keys(col).forEach(function(key) {
      if (col[key] === -1) {
        throw new Error("attendance シートに " + key + " 列がありません");
      }
    });

    const rowsToClear = [];

    rows.forEach(function(row, index) {
      const rowClassId = String(row[col.classId] || "").trim();
      const rowPeriod = String(row[col.period] == null ? "" : row[col.period]).trim();

      if (
        rowClassId !== targetClassId ||
        rowPeriod !== targetPeriod
      ) {
        return;
      }

      const rowDate = formatDateToYmd(row[col.date]);

      if (rowDate === targetDate) {
        rowsToClear.push(index + 2);
      }
    });

    rowsToClear.sort(function(a, b) { return a - b; });

    if (rowsToClear.length > 0) {
      if (isSequentialRows_(rowsToClear)) {
        attendanceSheet
          .getRange(rowsToClear[0], 1, rowsToClear.length, headers.length)
          .clearContent();
      } else {
        rowsToClear.forEach(function(rowNumber) {
          attendanceSheet
            .getRange(rowNumber, 1, 1, headers.length)
            .clearContent();
        });
      }
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

    tryInvalidateTeacherUnsavedFastSnapshotAfterSaveUnderLock_(
      targetClassId,
      targetDate,
      targetPeriod,
      actionType
    );

    invalidateAttendanceCaches_(targetClassId, targetDate, targetPeriod);

    const currentUser = getCurrentUserContext();
    const currentTeacherId = currentUser && currentUser.teacherId
      ? normalizeString_(currentUser.teacherId)
      : "";

    const summaryBaseDate = new Date();
    summaryBaseDate.setHours(0, 0, 0, 0);
    summaryBaseDate.setDate(summaryBaseDate.getDate() - 1);

    const summaryEndYmd = formatDateToYmd(summaryBaseDate);
    const summaryStartYmd = formatDateToYmd(getTeacherUnsavedStartDate_(summaryBaseDate));
    const teacherUnsavedCacheKeys = [
      "savedSessionKeySetByRange__" + summaryStartYmd + "__" + summaryEndYmd,
      "savedSessionKeySetByRange__v2__" + summaryStartYmd + "__" + summaryEndYmd,
      "savedSessionKeySetByRange__v4__" + summaryStartYmd + "__" + summaryEndYmd
    ];

    if (currentTeacherId) {
      teacherUnsavedCacheKeys.push(
        buildTeacherUnsavedSummaryCacheKey_(currentTeacherId, summaryEndYmd),
        buildTeacherUnsavedDetailsCacheKey_(currentTeacherId, summaryEndYmd)
      );
    }
    removeScriptCacheKeys_(teacherUnsavedCacheKeys);

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
      savedCount: 0,
      updatedCount: 0,
      appendedCount: 0,
      clearedCount: rowsToClear.length,
      mode: allowPastEdit ? "past-edit" : "normal",
      actionType: actionType,
      targetSessionKey: targetSessionKey,
      noAbsence: true,
      lastSavedInfo: lastSavedInfo
    };

  } finally {
    lock.releaseLock();
  }
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
    const targetGroup = String(payload.group || "").trim();

    const baseSessionKey = [targetClassId, targetDate, targetPeriod].join("__");
    const targetSessionKey = targetGroup
      ? baseSessionKey + "__" + targetGroup
      : baseSessionKey;

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

    const isExperimentGroupSave =
      typeof isExperimentGroupTargetClass_ === 'function' &&
      isExperimentGroupTargetClass_(targetClassId);

    const relatedClassIds = isExperimentGroupSave
      ? getExperimentRelatedClassIdsByClassId_(targetClassId)
      : [targetClassId];

    const relatedClassIdSet = {};
    relatedClassIds.forEach(function(id) {
      const normalizedId = String(id || '').trim();
      if (normalizedId) relatedClassIdSet[normalizedId] = true;
    });

    if (!relatedClassIdSet[targetClassId]) {
      relatedClassIdSet[targetClassId] = true;
      relatedClassIds.push(targetClassId);
    }

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
      recordedAt: headers.indexOf('recordedAt'),
      group: findColumnIndex_(headers, ['group', '班'])
    };

    ['classId', 'date', 'period', 'studentId', 'statusCode', 'recordedAt'].forEach(function(key) {
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
      const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();
      const rowStudentId = String(row[col.studentId] || '').trim();

      const rowClassMatches = isExperimentGroupSave
        ? !!relatedClassIdSet[rowClassId]
        : rowClassId === targetClassId;

      if (
        !rowClassMatches ||
        rowPeriod !== targetPeriod ||
        !targetStudentIds[rowStudentId]
      ) {
        return;
      }

      const rowDate = formatDateToYmd(row[col.date]);

      if (rowDate === targetDate) {
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
        recordedAt: now,
        group: targetGroup
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
  savedModeLabel,
  targetGroup
]);

    tryInvalidateTeacherUnsavedFastSnapshotAfterSaveUnderLock_(
      targetClassId,
      targetDate,
      targetPeriod,
      actionType
    );

    relatedClassIds.forEach(function(classIdToClear) {
      invalidateAttendanceCaches_(classIdToClear, targetDate, targetPeriod);
    });

    const currentUser = getCurrentUserContext();
    const currentTeacherId = currentUser && currentUser.teacherId
      ? normalizeString_(currentUser.teacherId)
      : '';

    const summaryBaseDate = new Date();
    summaryBaseDate.setHours(0, 0, 0, 0);
    summaryBaseDate.setDate(summaryBaseDate.getDate() - 1);

    const summaryEndYmd = formatDateToYmd(summaryBaseDate);
    const summaryStartYmd = formatDateToYmd(getTeacherUnsavedStartDate_(summaryBaseDate));
    const teacherUnsavedCacheKeys = [
      'savedSessionKeySetByRange__' + summaryStartYmd + '__' + summaryEndYmd,
      'savedSessionKeySetByRange__v2__' + summaryStartYmd + '__' + summaryEndYmd,
      'savedSessionKeySetByRange__v4__' + summaryStartYmd + '__' + summaryEndYmd
    ];

    if (currentTeacherId) {
      teacherUnsavedCacheKeys.push(
        buildTeacherUnsavedSummaryCacheKey_(currentTeacherId, summaryEndYmd),
        buildTeacherUnsavedDetailsCacheKey_(currentTeacherId, summaryEndYmd)
      );
    }
    removeScriptCacheKeys_(teacherUnsavedCacheKeys);

    const lastSavedInfo = {
      teacherEmail: currentUserEmail,
      savedAt: now,
      savedAtText: formatDateTimeJst_(now),
      actionType: actionType,
      targetSessionKey: targetSessionKey,
      savedModeLabel: savedModeLabel,
      savedByCurrentUser: true,
      group: targetGroup
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
  return getAttendanceMapDirect_(classId, date, period);
}

function getAttendanceMapDirect_(classId, date, period) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const targetClassId = String(classId || '').trim();
  const targetDate = formatDateToYmd(date);
  const targetPeriod = String(period == null ? '' : period).trim();

  const sessionKey = [targetClassId, targetDate, targetPeriod].join('__');
  const sessionCacheKey = buildAttendanceSessionCacheKey_(targetClassId, targetDate, targetPeriod);

  const cacheStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const cached = getScriptCacheJson_(sessionCacheKey);

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getAttendanceMapDirect_ getScriptCacheJson_',
      cacheStartedAt,
      'cacheKey=' + sessionCacheKey + ' hit=' + (!!cached)
    );
  }

  if (cached) {
    if (typeof logPerf_ === 'function') {
      logPerf_(
        'getAttendanceMapDirect_ total',
        totalStartedAt,
        'cache=hit entries=' + Object.keys(cached).length + ' key=' + sessionKey
      );
    }
    return cached;
  }

  if (!targetClassId || !targetDate || !targetPeriod) {
    if (typeof logPerf_ === 'function') {
      logPerf_('getAttendanceMapDirect_ total', totalStartedAt, 'invalid-args');
    }
    return {};
  }

  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);

  if (!sheet) {
    throw new Error('attendance シートが見つかりません');
  }

  const loadStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const values = sheet.getDataRange().getValues();

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getAttendanceMapDirect_ load sheet',
      loadStartedAt,
      'rows=' + Math.max(values.length - 1, 0)
    );
  }

  if (values.length <= 1) {
    putScriptCacheJson_(sessionCacheKey, {}, 300);
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

  Object.keys(col).forEach(function(key) {
    if (col[key] === -1) {
      throw new Error('attendance シートに ' + key + ' 列がありません');
    }
  });

  const scanStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const result = {};

  // This detail path only needs one class/date/period.
  // After class/period filtering, repeated date conversion is memoized.
  const attendanceDetailDateMemo = {};

  function resolveAttendanceDetailRowYmd_(value) {
    if (value instanceof Date) {
      const timeValue = value.getTime();
      const memoKey = 'D:' + String(timeValue);

      if (Object.prototype.hasOwnProperty.call(attendanceDetailDateMemo, memoKey)) {
        return attendanceDetailDateMemo[memoKey];
      }

      const formatted = formatDateToYmd(value);
      attendanceDetailDateMemo[memoKey] = formatted;
      return formatted;
    }

    const normalizedText =
      typeof normalizeYmdDisplayText_ === 'function'
        ? normalizeYmdDisplayText_(value)
        : '';

    if (normalizedText) {
      return normalizedText;
    }

    const rawText = String(value == null ? '' : value).trim();
    if (!rawText) return '';

    const memoKey = 'S:' + rawText;

    if (Object.prototype.hasOwnProperty.call(attendanceDetailDateMemo, memoKey)) {
      return attendanceDetailDateMemo[memoKey];
    }

    const formatted = formatDateToYmd(value);
    attendanceDetailDateMemo[memoKey] = formatted;
    return formatted;
  }

  rows.forEach(function(row) {
    const rowClassId = String(row[col.classId] || '').trim();
    if (rowClassId !== targetClassId) return;

    const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();
    if (rowPeriod !== targetPeriod) return;

    const rowDate = resolveAttendanceDetailRowYmd_(row[col.date]);
    if (rowDate !== targetDate) return;

    const studentId = String(row[col.studentId] || '').trim();
    const statusCode = String(row[col.statusCode] || '').trim();

    if (!studentId) return;

    if (statusCode) {
      result[studentId] = statusCode;
    } else if (result[studentId]) {
      delete result[studentId];
    }
  });

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getAttendanceMapDirect_ scan rows',
      scanStartedAt,
      'entries=' + Object.keys(result).length + ' key=' + sessionKey + ' uniqueDates=' + Object.keys(attendanceDetailDateMemo).length
    );
  }

  const putCacheStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  putScriptCacheJson_(sessionCacheKey, result, 300);

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getAttendanceMapDirect_ putScriptCacheJson_',
      putCacheStartedAt,
      'entries=' + Object.keys(result).length
    );

    logPerf_(
      'getAttendanceMapDirect_ total',
      totalStartedAt,
      'cache=miss entries=' + Object.keys(result).length + ' key=' + sessionKey
    );
  }

  return result;
}


function getAttendanceMapForClassIds_(classIds, date, period, studentIds) {
  const targetDate = formatDateToYmd(date);
  const targetPeriod = String(period == null ? '' : period).trim();

  const classIdSet = {};
  (Array.isArray(classIds) ? classIds : [classIds]).forEach(function(classId) {
    const id = String(classId || '').trim();
    if (id) classIdSet[id] = true;
  });

  const studentIdSet = {};
  const hasStudentFilter = Array.isArray(studentIds) && studentIds.length > 0;
  (Array.isArray(studentIds) ? studentIds : []).forEach(function(studentId) {
    const id = String(studentId || '').trim();
    if (id) studentIdSet[id] = true;
  });

  if (!targetDate || !targetPeriod || Object.keys(classIdSet).length === 0) {
    return {};
  }

  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);

  if (!sheet) {
    throw new Error('attendance シートが見つかりません');
  }

  const values = sheet.getDataRange().getValues();

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
    statusCode: headers.indexOf('statusCode'),
    recordedAt: headers.indexOf('recordedAt')
  };

  ['classId', 'date', 'period', 'studentId', 'statusCode'].forEach(function(key) {
    if (col[key] === -1) {
      throw new Error('attendance シートに ' + key + ' 列がありません');
    }
  });

  const result = {};
  const latestOrderByStudentId = {};

  rows.forEach(function(row, index) {
    const rowClassId = String(row[col.classId] || '').trim();
    if (!classIdSet[rowClassId]) return;

    const rowDate = formatDateToYmd(row[col.date]);
    if (rowDate !== targetDate) return;

    const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();
    if (rowPeriod !== targetPeriod) return;

    const studentId = String(row[col.studentId] || '').trim();
    if (!studentId) return;
    if (hasStudentFilter && !studentIdSet[studentId]) return;

    const statusCode = String(row[col.statusCode] || '').trim();

    const recordedAtRaw = col.recordedAt !== -1 ? row[col.recordedAt] : '';
    const recordedAt = recordedAtRaw instanceof Date ? recordedAtRaw : new Date(recordedAtRaw);
    const order = isNaN(recordedAt.getTime()) ? index : recordedAt.getTime();

    if (latestOrderByStudentId[studentId] != null && order < latestOrderByStudentId[studentId]) {
      return;
    }

    latestOrderByStudentId[studentId] = order;

    if (statusCode) {
      result[studentId] = statusCode;
    } else if (result[studentId]) {
      delete result[studentId];
    }
  });

  return result;
}

function getLatestAttendanceSessionInfoForClassIds_(classIds, date, period, allowedActionTypes) {
  return getLatestAttendanceSessionInfoForClassIdsByDateCache_(
    classIds,
    date,
    period,
    allowedActionTypes
  );
}

function getLatestAttendanceSessionInfoForClassIdsDirect_(classIds, date, period, allowedActionTypes) {
  const targetDate = formatDateToYmd(date);
  const targetPeriod = String(period == null ? '' : period).trim();

  const classIdSet = {};
  (Array.isArray(classIds) ? classIds : [classIds]).forEach(function(classId) {
    const id = String(classId || '').trim();
    if (id) classIdSet[id] = true;
  });

  if (!targetDate || !targetPeriod || Object.keys(classIdSet).length === 0) {
    return null;
  }

  const allowed = normalizeAttendanceActionTypes_(allowedActionTypes);

  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_SESSIONS);
  if (!sheet) {
    throw new Error('attendanceSessions シートが見つかりません');
  }

  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return null;
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
    savedModeLabel: findColumnIndex_(headers, ['savedModeLabel']),
    group: findColumnIndex_(headers, ['group', '班'])
  };

  ['classId', 'date', 'period'].forEach(function(key) {
    if (col[key] === -1) {
      throw new Error('attendanceSessions シートに ' + key + ' 列がありません');
    }
  });

  let latest = null;
  let latestMs = -1;

  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];

    const rowClassId = String(row[col.classId] || '').trim();
    if (!classIdSet[rowClassId]) continue;

    const rowDate = formatDateToYmd(row[col.date]);
    if (rowDate !== targetDate) continue;

    const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();
    if (rowPeriod !== targetPeriod) continue;

    const actionType = col.actionType !== -1
      ? String(row[col.actionType] || '').trim()
      : '';

    if (allowed && allowed.indexOf(actionType) === -1) {
      continue;
    }

    const accessedAtRaw = col.accessedAt !== -1 ? row[col.accessedAt] : '';
    const accessedAt = accessedAtRaw instanceof Date ? accessedAtRaw : new Date(accessedAtRaw);
    const accessedAtMs = isNaN(accessedAt.getTime()) ? 0 : accessedAt.getTime();

    if (!latest || accessedAtMs >= latestMs) {
      latestMs = accessedAtMs;
      latest = {
        teacherEmail: col.teacherEmail !== -1
          ? String(row[col.teacherEmail] || '').trim().toLowerCase()
          : '',
        savedAt: accessedAtRaw,
        savedAtText: formatDateTimeJst_(accessedAtRaw),
        actionType: actionType,
        targetSessionKey: col.targetSessionKey !== -1
          ? String(row[col.targetSessionKey] || '').trim()
          : '',
        savedModeLabel: col.savedModeLabel !== -1
          ? String(row[col.savedModeLabel] || '').trim()
          : '',
        group: col.group !== -1
          ? String(row[col.group] || '').trim()
          : ''
      };

      break;
    }
  }

  return latest;
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

  if (col.group !== -1 && col.group != null) {
    row[col.group] = record.group || '';
  }

  return row;
}

function clearAttendanceRowsByNumberGroups_(
  sheet,
  rowNumbers,
  columnCount
) {
  if (
    !sheet ||
    !Array.isArray(rowNumbers) ||
    rowNumbers.length === 0
  ) {
    return;
  }

  const sorted = rowNumbers
    .slice()
    .sort(function(a, b) { return a - b; });

  let startRow = sorted[0];
  let previousRow = sorted[0];

  function clearCurrentGroup_() {
    const rowCount = previousRow - startRow + 1;

    sheet
      .getRange(
        startRow,
        1,
        rowCount,
        columnCount
      )
      .clearContent();
  }

  for (let i = 1; i < sorted.length; i++) {
    const rowNumber = sorted[i];

    if (rowNumber === previousRow + 1) {
      previousRow = rowNumber;
      continue;
    }

    clearCurrentGroup_();

    startRow = rowNumber;
    previousRow = rowNumber;
  }

  clearCurrentGroup_();
}

function appendAttendanceSessionLogs_(sheet, baseRows) {
  if (
    !sheet ||
    !Array.isArray(baseRows) ||
    baseRows.length === 0
  ) {
    return;
  }

  const headerCount = sheet.getLastColumn();

  if (headerCount <= 0) {
    throw new Error(
      "attendanceSessions シートのヘッダーがありません"
    );
  }

  const targetColumnCount =
    headerCount <= 5 ? 5 : headerCount;

  const rows = baseRows.map(function(baseRow) {
    const row = baseRow.slice();

    while (row.length < targetColumnCount) {
      row.push("");
    }

    return row.slice(0, targetColumnCount);
  });

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      rows.length,
      targetColumnCount
    )
    .setValues(rows);
}

function invalidateAttendanceCachesBulk_(sessions) {
  const cacheKeySet = {};

  cacheKeySet[getAttendanceSheetCacheKey_()] = true;
  cacheKeySet[getAttendanceSessionsSheetCacheKey_()] = true;
  cacheKeySet["attendanceIndex__all"] = true;
  cacheKeySet["attendanceSessionLatestIndex__all"] = true;

  (sessions || []).forEach(function(session) {
    const classId =
      String(session.classId || "").trim();

    const date =
      formatDateToYmd(session.date);

    const period =
      String(
        session.period == null ? "" : session.period
      ).trim();

    if (!classId || !date || !period) {
      return;
    }

    cacheKeySet[
      buildAttendanceSessionCacheKey_(
        classId,
        date,
        period
      )
    ] = true;

    cacheKeySet[
      "savedSessionMapByDate__" + date
    ] = true;

    cacheKeySet[
      "attendanceSessionLatestMapByDate__" + date
    ] = true;

    cacheKeySet[
      "attendanceSessionLatestMapByDate__v2__" + date
    ] = true;
  });

  removeScriptCacheKeys_(
    Object.keys(cacheKeySet)
  );
}

function tryInvalidateTeacherUnsavedFastSnapshotsAfterBulkSaveUnderLock_(
  sessions,
  actionType
) {
  try {
    return invalidateTeacherUnsavedFastSnapshotsAfterBulkSaveUnderLock_(
      sessions
    );
  } catch (error) {
    Logger.log(JSON.stringify({
      ok: false,
      event:
        "teacher-unsaved-fast-bulk-cache-invalidation-failed",
      warning: true,
      attendanceSaveSucceeded: true,
      actionType: actionType || "",
      sessionCount:
        Array.isArray(sessions) ? sessions.length : 0,
      errorName:
        error && error.name
          ? String(error.name)
          : "",
      errorMessage:
        error && error.message
          ? String(error.message)
          : String(error),
      errorStack:
        error && error.stack
          ? String(error.stack)
          : ""
    }));

    return {
      ok: false,
      warning: true,
      errorMessage:
        error && error.message
          ? String(error.message)
          : String(error)
    };
  }
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

function tryInvalidateTeacherUnsavedFastSnapshotAfterSaveUnderLock_(
  classId,
  date,
  period,
  actionType
) {
  try {
    return invalidateTeacherUnsavedFastSnapshotAfterSaveUnderLock_(
      classId,
      date,
      period
    );
  } catch (error) {
    Logger.log(JSON.stringify({
      ok: false,
      event: 'teacher-unsaved-fast-cache-invalidation-failed',
      warning: true,
      attendanceSaveSucceeded: true,
      actionType: actionType || '',
      classId: String(classId || ''),
      date: formatDateToYmd(date),
      period: String(period == null ? '' : period),
      errorName: error && error.name ? String(error.name) : '',
      errorMessage: error && error.message ? String(error.message) : String(error),
      errorStack: error && error.stack ? String(error.stack) : ''
    }));

    return {
      ok: false,
      warning: true,
      errorMessage: error && error.message ? String(error.message) : String(error)
    };
  }
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
  const targetDate = formatDateToYmd(date);

  removeScriptCacheKeys_([
    getAttendanceSheetCacheKey_(),
    getAttendanceSessionsSheetCacheKey_(),
    buildAttendanceSessionCacheKey_(classId, targetDate, period),
    'attendanceIndex__all',
    'attendanceSessionLatestIndex__all',
    'savedSessionMapByDate__' + targetDate,
    'attendanceSessionLatestMapByDate__' + targetDate,
    'attendanceSessionLatestMapByDate__v2__' + targetDate
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

function selectAttendanceSessionCacheEntryByAllowedActionTypes_(
  entry,
  allowedActionTypes
) {
  if (!entry) return null;

  const allowed = normalizeAttendanceActionTypes_(allowedActionTypes);
  if (!allowed) {
    return entry;
  }

  const byActionType =
    entry._byActionType && typeof entry._byActionType === 'object'
      ? entry._byActionType
      : {};

  let latest = null;
  let latestMs = -1;

  allowed.forEach(function(actionType) {
    const candidate = byActionType[actionType];
    if (!candidate) return;

    const candidateMs = Number(candidate._ms || 0);
    if (!latest || candidateMs >= latestMs) {
      latest = candidate;
      latestMs = candidateMs;
    }
  });

  return latest;
}

function buildAttendanceSessionInfoFromCacheEntry_(entry, fallbackSessionKey) {
  if (!entry) return null;

  return {
    teacherEmail: String(entry.teacherEmail || '').trim().toLowerCase(),
    savedAt: entry.savedAt || '',
    savedAtText: String(entry.savedAtText || ''),
    actionType: String(entry.actionType || ''),
    targetSessionKey: String(entry.targetSessionKey || fallbackSessionKey || ''),
    savedModeLabel: String(entry.savedModeLabel || ''),
    group: String(entry.group || '')
  };
}

function getLatestAttendanceSessionInfoByDateCache_(
  classId,
  date,
  period,
  allowedActionTypes
) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const targetClassId = String(classId || '').trim();
  const targetDate = formatDateToYmd(date);
  const targetPeriod = String(period == null ? '' : period).trim();

  if (!targetClassId || !targetDate || !targetPeriod) {
    if (typeof logPerf_ === 'function') {
      logPerf_(
        'getLatestAttendanceSessionInfoByDateCache_ total',
        totalStartedAt,
        'invalid-args'
      );
    }
    return null;
  }

  const sessionKey = [targetClassId, targetDate, targetPeriod].join('__');
  const latestMap = getAttendanceSessionLatestMapByDateCached_(targetDate) || {};
  const entry = selectAttendanceSessionCacheEntryByAllowedActionTypes_(
    latestMap[sessionKey] || null,
    allowedActionTypes
  );

  const result = buildAttendanceSessionInfoFromCacheEntry_(entry, sessionKey);

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getLatestAttendanceSessionInfoByDateCache_ total',
      totalStartedAt,
      (result ? 'found' : 'not-found') + ' key=' + sessionKey
    );
  }

  return result;
}

function getLatestAttendanceSessionInfoForClassIdsByDateCache_(
  classIds,
  date,
  period,
  allowedActionTypes
) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const targetDate = formatDateToYmd(date);
  const targetPeriod = String(period == null ? '' : period).trim();

  const ids = [];
  const seen = {};

  (Array.isArray(classIds) ? classIds : [classIds]).forEach(function(classId) {
    const id = String(classId || '').trim();
    if (!id || seen[id]) return;
    seen[id] = true;
    ids.push(id);
  });

  if (!targetDate || !targetPeriod || ids.length === 0) {
    return null;
  }

  const latestMap = getAttendanceSessionLatestMapByDateCached_(targetDate) || {};
  let latestEntry = null;
  let latestMs = -1;
  let latestSessionKey = '';

  ids.forEach(function(classId) {
    const sessionKey = [classId, targetDate, targetPeriod].join('__');
    const candidate = selectAttendanceSessionCacheEntryByAllowedActionTypes_(
      latestMap[sessionKey] || null,
      allowedActionTypes
    );

    if (!candidate) return;

    const candidateMs = Number(candidate._ms || 0);
    if (!latestEntry || candidateMs >= latestMs) {
      latestEntry = candidate;
      latestMs = candidateMs;
      latestSessionKey = sessionKey;
    }
  });

  const result = buildAttendanceSessionInfoFromCacheEntry_(
    latestEntry,
    latestSessionKey
  );

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getLatestAttendanceSessionInfoForClassIdsByDateCache_ total',
      totalStartedAt,
      'classIds=' + ids.length +
        ' result=' + (result ? 'found' : 'not-found') +
        ' date=' + targetDate +
        ' period=' + targetPeriod
    );
  }

  return result;
}

function getLatestAttendanceSessionInfo_(classId, date, period, allowedActionTypes) {
  return getLatestAttendanceSessionInfoByDateCache_(
    classId,
    date,
    period,
    allowedActionTypes
  );
}

function getLatestAttendanceSessionInfoDirect_(classId, date, period, allowedActionTypes) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const targetClassId = String(classId || '').trim();
  const targetDate = formatDateToYmd(date);
  const targetPeriod = String(period == null ? '' : period).trim();

  if (!targetClassId || !targetDate || !targetPeriod) {
    if (typeof logPerf_ === 'function') {
      logPerf_('getLatestAttendanceSessionInfoDirect_ total', totalStartedAt, 'invalid-args');
    }
    return null;
  }

  const allowed = normalizeAttendanceActionTypes_(allowedActionTypes);

  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_SESSIONS);
  if (!sheet) {
    throw new Error('attendanceSessions シートが見つかりません');
  }

  const loadStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const values = sheet.getDataRange().getValues();

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getLatestAttendanceSessionInfoDirect_ load sheet',
      loadStartedAt,
      'rows=' + Math.max(values.length - 1, 0)
    );
  }

  if (values.length <= 1) {
    return null;
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

  let latest = null;
  let latestMs = -1;

  // 新しい行ほど下にある前提なので、下から探す
  for (let i = rows.length - 1; i >= 0; i--) {
    const row = rows[i];

    const rowClassId = String(row[col.classId] || '').trim();
    if (rowClassId !== targetClassId) continue;

    const rowDate = formatDateToYmd(row[col.date]);
    if (rowDate !== targetDate) continue;

    const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();
    if (rowPeriod !== targetPeriod) continue;

    const actionType = col.actionType !== -1
      ? String(row[col.actionType] || '').trim()
      : '';

    if (allowed && allowed.indexOf(actionType) === -1) {
      continue;
    }

    const accessedAtRaw = col.accessedAt !== -1 ? row[col.accessedAt] : '';
    const accessedAt = accessedAtRaw instanceof Date ? accessedAtRaw : new Date(accessedAtRaw);
    const accessedAtMs = isNaN(accessedAt.getTime()) ? 0 : accessedAt.getTime();

    if (!latest || accessedAtMs >= latestMs) {
      latestMs = accessedAtMs;
      latest = {
        teacherEmail: col.teacherEmail !== -1
          ? String(row[col.teacherEmail] || '').trim().toLowerCase()
          : '',
        savedAt: accessedAtRaw,
        savedAtText: formatDateTimeJst_(accessedAtRaw),
        actionType: actionType,
        targetSessionKey: col.targetSessionKey !== -1
          ? String(row[col.targetSessionKey] || '').trim()
          : '',
        savedModeLabel: col.savedModeLabel !== -1
          ? String(row[col.savedModeLabel] || '').trim()
          : ''
      };

      // 基本は最下行が最新なので、見つかった時点で抜けてOK
      break;
    }
  }

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getLatestAttendanceSessionInfoDirect_ total',
      totalStartedAt,
      latest
        ? 'found key=' + [targetClassId, targetDate, targetPeriod].join('__')
        : 'not-found key=' + [targetClassId, targetDate, targetPeriod].join('__')
    );
  }

  return latest;
}

function getAttendanceSessionLatestMapByDateCached_(ymd) {
  const totalStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  const targetDate = formatDateToYmd(ymd);
  if (!targetDate) {
    return {};
  }

  const cacheKey = 'attendanceSessionLatestMapByDate__v2__' + targetDate;
  const cached = getScriptCacheJson_(cacheKey);

  if (cached) {
    if (typeof logPerf_ === 'function') {
      logPerf_(
        'getAttendanceSessionLatestMapByDateCached_ total',
        totalStartedAt,
        'cache=hit date=' + targetDate + ' keys=' + Object.keys(cached).length
      );
    }
    return cached;
  }

  const loadStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();

  // attendanceSessions は現在 ScriptCache の1キー上限を大きく超えるため、
  // 巨大な全シートJSON化・キャッシュ試行を避け、ここでは直接1回だけ読む。
  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_SESSIONS);
  if (!sheet) {
    throw new Error('attendanceSessions シートが見つかりません');
  }

  const values = sheet.getDataRange().getValues();
  const headers = values.length > 0 ? values[0] : [];
  const rows = values.length > 1 ? values.slice(1) : [];

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getAttendanceSessionLatestMapByDateCached_ load attendanceSessionsData',
      loadStartedAt,
      'rows=' + rows.length + ' source=direct'
    );
  }

  const col = {
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    date: findColumnIndex_(headers, ['date', '日付']),
    period: findColumnIndex_(headers, ['period', '時限']),
    teacherEmail: findColumnIndex_(headers, ['teacherEmail', 'email']),
    accessedAt: findColumnIndex_(headers, ['accessedAt', 'savedAt']),
    actionType: findColumnIndex_(headers, ['actionType']),
    targetSessionKey: findColumnIndex_(headers, ['targetSessionKey']),
    savedModeLabel: findColumnIndex_(headers, ['savedModeLabel']),
    group: findColumnIndex_(headers, ['group', '班'])
  };

  ['classId', 'date', 'period'].forEach(function(key) {
    if (col[key] === -1) {
      throw new Error('attendanceSessions シートに ' + key + ' 列がありません');
    }
  });

  const buildStartedAt = typeof perfNow_ === 'function' ? perfNow_() : Date.now();
  const latestMap = {};

  // formatDateToYmd() は Utilities.formatDate() を使うため、
  // 9,000行超で毎回呼ぶと数秒かかる。
  // 同一日付の Date 値は同じミリ秒値になるので、変換結果をメモ化する。
  const attendanceSessionDateMemo = {};

  function resolveAttendanceSessionRowYmd_(value) {
    if (value instanceof Date) {
      const timeValue = value.getTime();
      const memoKey = 'D:' + String(timeValue);

      if (Object.prototype.hasOwnProperty.call(attendanceSessionDateMemo, memoKey)) {
        return attendanceSessionDateMemo[memoKey];
      }

      const formatted = formatDateToYmd(value);
      attendanceSessionDateMemo[memoKey] = formatted;
      return formatted;
    }

    const normalizedText =
      typeof normalizeYmdDisplayText_ === 'function'
        ? normalizeYmdDisplayText_(value)
        : '';

    if (normalizedText) {
      return normalizedText;
    }

    const rawText = String(value == null ? '' : value).trim();
    if (!rawText) {
      return '';
    }

    const memoKey = 'S:' + rawText;

    if (Object.prototype.hasOwnProperty.call(attendanceSessionDateMemo, memoKey)) {
      return attendanceSessionDateMemo[memoKey];
    }

    const formatted = formatDateToYmd(value);
    attendanceSessionDateMemo[memoKey] = formatted;
    return formatted;
  }

  rows.forEach(function(row) {
    const rowDate = resolveAttendanceSessionRowYmd_(row[col.date]);
    if (rowDate !== targetDate) return;

    const rowClassId = String(row[col.classId] || '').trim();
    const rowPeriod = String(row[col.period] == null ? '' : row[col.period]).trim();
    if (!rowClassId || !rowPeriod) return;

    const key = [rowClassId, rowDate, rowPeriod].join('__');

    const actionType =
      col.actionType !== -1 ? String(row[col.actionType] || '').trim() : '';

    const accessedAtRaw = col.accessedAt !== -1 ? row[col.accessedAt] : '';
    const accessedAt =
      accessedAtRaw instanceof Date ? accessedAtRaw : new Date(accessedAtRaw);
    const accessedAtMs = isNaN(accessedAt.getTime()) ? 0 : accessedAt.getTime();

    const candidate = {
      teacherEmail:
        col.teacherEmail !== -1
          ? String(row[col.teacherEmail] || '').trim().toLowerCase()
          : '',
      savedAt: accessedAtRaw,
      savedAtText: formatDateTimeJst_(accessedAtRaw),
      actionType: actionType,
      targetSessionKey:
        col.targetSessionKey !== -1
          ? String(row[col.targetSessionKey] || '').trim()
          : '',
      savedModeLabel:
        col.savedModeLabel !== -1
          ? String(row[col.savedModeLabel] || '').trim()
          : '',
      group:
        col.group !== -1
          ? String(row[col.group] || '').trim()
          : '',
      _ms: accessedAtMs
    };

    if (!latestMap[key]) {
      latestMap[key] = Object.assign({}, candidate, {
        _byActionType: {}
      });
    }

    const current = latestMap[key];

    if (accessedAtMs >= Number(current._ms || 0)) {
      const existingByActionType = current._byActionType || {};
      latestMap[key] = Object.assign({}, candidate, {
        _byActionType: existingByActionType
      });
    }

    if (actionType) {
      const byActionType = latestMap[key]._byActionType || {};
      const currentForAction = byActionType[actionType];
      const currentForActionMs = currentForAction
        ? Number(currentForAction._ms || 0)
        : -1;

      if (!currentForAction || accessedAtMs >= currentForActionMs) {
        byActionType[actionType] = candidate;
      }

      latestMap[key]._byActionType = byActionType;
    }
  });

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getAttendanceSessionLatestMapByDateCached_ build latestMap',
      buildStartedAt,
      'date=' + targetDate +
        ' keys=' + Object.keys(latestMap).length +
        ' uniqueDates=' + Object.keys(attendanceSessionDateMemo).length
    );
  }

  putScriptCacheJson_(cacheKey, latestMap, 300);

  if (typeof logPerf_ === 'function') {
    logPerf_(
      'getAttendanceSessionLatestMapByDateCached_ total',
      totalStartedAt,
      'cache=miss date=' + targetDate + ' keys=' + Object.keys(latestMap).length
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
