const TEACHING_ASSIGNMENT_REVISION_PROPERTY_ =
  'KTC_TEACHING_ASSIGNMENT_REVISION';

const TEACHING_ASSIGNMENT_EDIT_TRIGGER_HANDLER_ =
  'handleTeachingAssignmentIntegrityEdit';

const TEACHING_ASSIGNMENT_CHANGE_TRIGGER_HANDLER_ =
  'handleTeachingAssignmentIntegrityChange';

function getTeachingAssignmentRevision_() {
  return PropertiesService.getScriptProperties()
    .getProperty(TEACHING_ASSIGNMENT_REVISION_PROPERTY_) || 'r0';
}

function bumpTeachingAssignmentRevision_() {
  const revision =
    Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmssSSS') +
    '__' +
    Utilities.getUuid();

  PropertiesService.getScriptProperties()
    .setProperty(TEACHING_ASSIGNMENT_REVISION_PROPERTY_, revision);

  return revision;
}

function getTeacherAssignmentCanonicalIndex_() {
  const rows = getTeachersSheetObjectsCached_(300);
  const byId = {};
  const byName = {};
  const ambiguousNames = {};

  rows.forEach(function(row) {
    const record = buildTeacherRecordFromRow_(row);
    if (!record || !record.teacherId) return;

    if (!byId[record.teacherId]) {
      byId[record.teacherId] = record;
    }

    if (!record.name) return;

    if (!byName[record.name]) {
      byName[record.name] = record;
      return;
    }

    if (byName[record.name].teacherId !== record.teacherId) {
      ambiguousNames[record.name] = true;
    }
  });

  return {
    byId: byId,
    byName: byName,
    ambiguousNames: ambiguousNames
  };
}

function resolveTeacherAssignmentCanonicalRecord_(teacherId, teacherName, index) {
  const normalizedId = normalizeString_(teacherId);
  const normalizedName = normalizeString_(teacherName);
  const sourceIndex = index || getTeacherAssignmentCanonicalIndex_();

  // teacherName が入力されている場合は、名前を人間が選択した意図として優先する。
  // 名前がマスタで一意に解決できない時は、古い teacherId へフォールバックしない。
  // これにより「名前だけ変更したがIDが旧担当のまま」の再発を fail-safe で防ぐ。
  if (normalizedName) {
    if (sourceIndex.ambiguousNames[normalizedName]) {
      return {
        record: null,
        resolution: 'ambiguous-name'
      };
    }

    if (sourceIndex.byName[normalizedName]) {
      return {
        record: sourceIndex.byName[normalizedName],
        resolution: 'name'
      };
    }

    return {
      record: null,
      resolution: 'name-not-found'
    };
  }

  // 名前が空欄の場合だけ、既存 teacherId から正規レコードを復元する。
  if (normalizedId && sourceIndex.byId[normalizedId]) {
    return {
      record: sourceIndex.byId[normalizedId],
      resolution: 'id'
    };
  }

  return {
    record: null,
    resolution: 'id-not-found'
  };
}

function isTeachingAssignmentSheet_(sheetName) {
  return (
    sheetName === CONFIG.SHEETS.TIMETABLE ||
    sheetName === CONFIG.SHEETS.CLASS_TEACHER_TEAMS
  );
}

function rangeTouchesTeachingAssignmentColumns_(range) {
  const firstColumn = range.getColumn();
  const lastColumn = firstColumn + range.getNumColumns() - 1;

  // classId / weekday / period / teacherName / teacherId
  return firstColumn <= 5 && lastColumn >= 1;
}

function syncTeacherAssignmentRows_(sheet, startRow, numRows, canonicalIndex) {
  if (!sheet || numRows <= 0) {
    return {
      changedCount: 0,
      mismatchCount: 0,
      unresolvedCount: 0,
      issues: []
    };
  }

  const index = canonicalIndex || getTeacherAssignmentCanonicalIndex_();
  const range = sheet.getRange(startRow, 4, numRows, 2);
  const values = range.getValues();
  const nextValues = [];
  const issues = [];
  let changedCount = 0;
  let mismatchCount = 0;
  let unresolvedCount = 0;

  values.forEach(function(row, offset) {
    const teacherName = normalizeString_(row[0]);
    const teacherId = normalizeString_(row[1]);
    const rowNumber = startRow + offset;

    if (!teacherName && !teacherId) {
      nextValues.push(['', '']);
      return;
    }

    const resolved = resolveTeacherAssignmentCanonicalRecord_(
      teacherId,
      teacherName,
      index
    );

    if (!resolved.record) {
      unresolvedCount++;
      issues.push({
        sheetName: sheet.getName(),
        rowNumber: rowNumber,
        teacherName: teacherName,
        teacherId: teacherId,
        reason: resolved.resolution
      });
      nextValues.push([row[0], row[1]]);
      return;
    }

    const canonicalName = resolved.record.name;
    const canonicalId = resolved.record.teacherId;

    if (
      teacherName &&
      teacherId &&
      (teacherName !== canonicalName || teacherId !== canonicalId)
    ) {
      mismatchCount++;
    }

    if (teacherName !== canonicalName || teacherId !== canonicalId) {
      changedCount++;
    }

    nextValues.push([canonicalName, canonicalId]);
  });

  if (changedCount > 0) {
    range.setValues(nextValues);
    SpreadsheetApp.flush();
  }

  return {
    changedCount: changedCount,
    mismatchCount: mismatchCount,
    unresolvedCount: unresolvedCount,
    issues: issues
  };
}

function invalidateTeachingAssignmentReadCaches_() {
  removeScriptCacheKeys_([
    'sheetData__OPERATION__' + CONFIG.SHEETS.TIMETABLE,
    'sheetData__OPERATION__' + CONFIG.SHEETS.CLASS_TEACHER_TEAMS,
    'classTeacherTeamRows__all'
  ]);
}

function markTeacherUnsavedFastCacheStaleAfterAssignmentEdit_() {
  if (
    typeof validateTeacherUnsavedCacheSheets_ !== 'function' ||
    typeof markTeacherUnsavedSummaryRowsStatus_ !== 'function'
  ) {
    return {
      ok: false,
      skipped: true,
      reason: 'teacher-unsaved-cache-service-unavailable'
    };
  }

  const validation = validateTeacherUnsavedCacheSheets_();
  if (!validation.ok) {
    return {
      ok: false,
      skipped: true,
      reason: 'teacher-unsaved-cache-sheet-invalid',
      errors: validation.errors || []
    };
  }

  markTeacherUnsavedSummaryRowsStatus_(
    validation.summary.sheet,
    'stale',
    '担当教員設定が変更されたためFastキャッシュを無効化しました。次回rebuildで再生成します。'
  );
  SpreadsheetApp.flush();

  return {
    ok: true,
    skipped: false
  };
}

function rebuildTeacherUnsavedFastCacheAfterAssignmentEdit_() {
  if (typeof rebuildTeacherUnsavedSummaryCache !== 'function') {
    return {
      ok: false,
      skipped: true,
      reason: 'rebuild-function-unavailable'
    };
  }

  try {
    const result = rebuildTeacherUnsavedSummaryCache();
    return {
      ok: true,
      skipped: false,
      result: result
    };
  } catch (error) {
    Logger.log(
      '[TEACHING_ASSIGNMENT_CACHE_REBUILD_FAILED] ' +
      (error && error.stack ? error.stack : error)
    );
    return {
      ok: false,
      skipped: false,
      errorMessage: error && error.message ? String(error.message) : String(error)
    };
  }
}

function handleTeachingAssignmentIntegrityEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (!isTeachingAssignmentSheet_(sheetName)) return;
  if (!rangeTouchesTeachingAssignmentColumns_(e.range)) return;

  const editStartRow = e.range.getRow();
  const editEndRow = editStartRow + e.range.getNumRows() - 1;
  const startRow = Math.max(2, editStartRow);

  if (editEndRow < 2 || startRow > editEndRow) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('担当教員整合性チェックのロックを取得できませんでした。');
  }

  let syncResult;
  let revision = '';

  try {
    const index = getTeacherAssignmentCanonicalIndex_();
    syncResult = syncTeacherAssignmentRows_(
      sheet,
      startRow,
      editEndRow - startRow + 1,
      index
    );

    revision = bumpTeachingAssignmentRevision_();
    invalidateTeachingAssignmentReadCaches_();
    markTeacherUnsavedFastCacheStaleAfterAssignmentEdit_();
  } finally {
    lock.releaseLock();
  }

  const rebuildResult = rebuildTeacherUnsavedFastCacheAfterAssignmentEdit_();

  Logger.log(JSON.stringify({
    event: 'teaching-assignment-edit',
    sheetName: sheetName,
    startRow: startRow,
    endRow: editEndRow,
    changedCount: syncResult ? syncResult.changedCount : 0,
    mismatchCount: syncResult ? syncResult.mismatchCount : 0,
    unresolvedCount: syncResult ? syncResult.unresolvedCount : 0,
    teachingAssignmentRevision: revision,
    cacheRebuildOk: !!(rebuildResult && rebuildResult.ok)
  }));

  return {
    ok: true,
    sheetName: sheetName,
    startRow: startRow,
    endRow: editEndRow,
    syncResult: syncResult,
    teachingAssignmentRevision: revision,
    cacheRebuild: rebuildResult
  };
}

function handleTeachingAssignmentIntegrityChange(e) {
  if (!e || !e.source) return;

  const changeType = normalizeString_(e.changeType).toUpperCase();
  const structuralChangeTypes = {
    INSERT_ROW: true,
    REMOVE_ROW: true,
    INSERT_COLUMN: true,
    REMOVE_COLUMN: true
  };

  // セル値の変更は onEdit 側で処理する。
  // 行・列の追加削除だけは onEdit では捕捉できないため、
  // onChange 側で担当割当キャッシュを無効化する。
  if (!structuralChangeTypes[changeType]) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('担当教員構造変更チェックのロックを取得できませんでした。');
  }

  let revision = '';
  try {
    revision = bumpTeachingAssignmentRevision_();
    invalidateTeachingAssignmentReadCaches_();
    markTeacherUnsavedFastCacheStaleAfterAssignmentEdit_();
  } finally {
    lock.releaseLock();
  }

  const rebuildResult = rebuildTeacherUnsavedFastCacheAfterAssignmentEdit_();

  Logger.log(JSON.stringify({
    event: 'teaching-assignment-structural-change',
    changeType: changeType,
    teachingAssignmentRevision: revision,
    cacheRebuildOk: !!(rebuildResult && rebuildResult.ok)
  }));

  return {
    ok: true,
    changeType: changeType,
    teachingAssignmentRevision: revision,
    cacheRebuild: rebuildResult
  };
}

function auditTeacherAssignmentIntegrity() {
  const ss = getOperationSpreadsheet();
  const index = getTeacherAssignmentCanonicalIndex_();
  const targets = [
    CONFIG.SHEETS.TIMETABLE,
    CONFIG.SHEETS.CLASS_TEACHER_TEAMS
  ];

  const issues = [];

  targets.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;

    const values = sheet
      .getRange(2, 4, sheet.getLastRow() - 1, 2)
      .getValues();

    values.forEach(function(row, offset) {
      const teacherName = normalizeString_(row[0]);
      const teacherId = normalizeString_(row[1]);
      if (!teacherName && !teacherId) return;

      const resolved = resolveTeacherAssignmentCanonicalRecord_(
        teacherId,
        teacherName,
        index
      );

      if (!resolved.record) {
        issues.push({
          sheetName: sheetName,
          rowNumber: offset + 2,
          teacherName: teacherName,
          teacherId: teacherId,
          reason: resolved.resolution,
          repairable: false
        });
        return;
      }

      if (
        teacherName !== resolved.record.name ||
        teacherId !== resolved.record.teacherId
      ) {
        issues.push({
          sheetName: sheetName,
          rowNumber: offset + 2,
          teacherName: teacherName,
          teacherId: teacherId,
          expectedTeacherName: resolved.record.name,
          expectedTeacherId: resolved.record.teacherId,
          reason: 'name-id-mismatch',
          repairable: true
        });
      }
    });
  });

  return {
    ok: true,
    issueCount: issues.length,
    repairableCount: issues.filter(function(item) {
      return item.repairable;
    }).length,
    issues: issues
  };
}

function repairTeacherAssignmentIntegrityNow() {
  const before = auditTeacherAssignmentIntegrity();
  const ss = getOperationSpreadsheet();
  const index = getTeacherAssignmentCanonicalIndex_();
  const targets = [
    CONFIG.SHEETS.TIMETABLE,
    CONFIG.SHEETS.CLASS_TEACHER_TEAMS
  ];

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('担当教員一括修復のロックを取得できませんでした。');
  }

  const results = [];
  let revision = '';

  try {
    targets.forEach(function(sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return;

      results.push({
        sheetName: sheetName,
        result: syncTeacherAssignmentRows_(
          sheet,
          2,
          sheet.getLastRow() - 1,
          index
        )
      });
    });

    revision = bumpTeachingAssignmentRevision_();
    invalidateTeachingAssignmentReadCaches_();
    markTeacherUnsavedFastCacheStaleAfterAssignmentEdit_();
  } finally {
    lock.releaseLock();
  }

  const rebuildResult = rebuildTeacherUnsavedFastCacheAfterAssignmentEdit_();
  const after = auditTeacherAssignmentIntegrity();

  return {
    ok: after.issueCount === 0,
    before: before,
    results: results,
    after: after,
    teachingAssignmentRevision: revision,
    cacheRebuild: rebuildResult
  };
}

function getTeachingAssignmentIntegrityTriggerStatus() {
  const allTriggers = ScriptApp.getProjectTriggers();
  const editTriggers = allTriggers.filter(function(trigger) {
    return trigger.getHandlerFunction() === TEACHING_ASSIGNMENT_EDIT_TRIGGER_HANDLER_;
  });
  const changeTriggers = allTriggers.filter(function(trigger) {
    return trigger.getHandlerFunction() === TEACHING_ASSIGNMENT_CHANGE_TRIGGER_HANDLER_;
  });

  return {
    ok: true,
    editHandlerFunction: TEACHING_ASSIGNMENT_EDIT_TRIGGER_HANDLER_,
    changeHandlerFunction: TEACHING_ASSIGNMENT_CHANGE_TRIGGER_HANDLER_,
    editTriggerCount: editTriggers.length,
    changeTriggerCount: changeTriggers.length,
    triggerCount: editTriggers.length + changeTriggers.length,
    teachingAssignmentRevision: getTeachingAssignmentRevision_()
  };
}

function installTeachingAssignmentIntegrityTrigger() {
  const targetHandlers = {};
  targetHandlers[TEACHING_ASSIGNMENT_EDIT_TRIGGER_HANDLER_] = true;
  targetHandlers[TEACHING_ASSIGNMENT_CHANGE_TRIGGER_HANDLER_] = true;

  const existing = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return !!targetHandlers[trigger.getHandlerFunction()];
  });

  existing.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });

  const spreadsheet = getOperationSpreadsheet();

  ScriptApp.newTrigger(TEACHING_ASSIGNMENT_EDIT_TRIGGER_HANDLER_)
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  ScriptApp.newTrigger(TEACHING_ASSIGNMENT_CHANGE_TRIGGER_HANDLER_)
    .forSpreadsheet(spreadsheet)
    .onChange()
    .create();

  const status = getTeachingAssignmentIntegrityTriggerStatus();
  status.removedCount = existing.length;
  status.createdCount = 2;
  return status;
}

function removeTeachingAssignmentIntegrityTrigger() {
  const targetHandlers = {};
  targetHandlers[TEACHING_ASSIGNMENT_EDIT_TRIGGER_HANDLER_] = true;
  targetHandlers[TEACHING_ASSIGNMENT_CHANGE_TRIGGER_HANDLER_] = true;

  const existing = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return !!targetHandlers[trigger.getHandlerFunction()];
  });

  existing.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });

  return {
    ok: true,
    deletedCount: existing.length,
    remainingCount: getTeachingAssignmentIntegrityTriggerStatus().triggerCount
  };
}
