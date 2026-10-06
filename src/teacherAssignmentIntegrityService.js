const TEACHING_ASSIGNMENT_INTEGRITY_EDIT_HANDLER_ =
  'handleTeachingAssignmentIntegrityEdit';
const TEACHING_ASSIGNMENT_INTEGRITY_CHANGE_HANDLER_ =
  'handleTeachingAssignmentIntegrityChange';

function getTeacherAssignmentIntegrityIndex_() {
  const rows = getTeachersSheetObjectsCached_(1800);
  const byId = {};
  const byName = {};
  const nameCounts = {};

  rows.forEach(function(row) {
    const teacherId = normalizeString_(row.teacherId);
    const teacherName = normalizeString_(row.name);
    if (teacherId && !byId[teacherId]) byId[teacherId] = row;
    if (!teacherName) return;

    nameCounts[teacherName] = (nameCounts[teacherName] || 0) + 1;
    if (!byName[teacherName]) byName[teacherName] = row;
  });

  return {
    byId: byId,
    byName: byName,
    nameCounts: nameCounts
  };
}

function getTeachingAssignmentIntegrityColumns_(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headers = lastColumn > 0
    ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    : [];

  return {
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    weekday: findColumnIndex_(headers, ['weekday', '曜日']),
    period: findColumnIndex_(headers, ['period', '時限']),
    teacherName: findColumnIndex_(headers, ['teacherName', '担当者名', 'name']),
    teacherId: findColumnIndex_(headers, ['teacherId', 'TeacherID'])
  };
}

function isTeachingAssignmentIntegrityTargetSheet_(sheetName) {
  return sheetName === CONFIG.SHEETS.TIMETABLE ||
    sheetName === CONFIG.SHEETS.CLASS_TEACHER_TEAMS;
}

function teachingAssignmentIntegrityRangeTouchesTarget_(range, columns) {
  if (!range) return false;
  const start = range.getColumn() - 1;
  const end = start + range.getNumColumns() - 1;
  return [
    columns.classId,
    columns.weekday,
    columns.period,
    columns.teacherName,
    columns.teacherId
  ].some(function(index) {
    return index >= 0 && index >= start && index <= end;
  });
}

function resolveTeacherAssignmentIntegrity_(teacherName, teacherId, index) {
  const name = normalizeString_(teacherName);
  const id = normalizeString_(teacherId);

  if (name) {
    const count = Number(index.nameCounts[name] || 0);
    if (count === 0) {
      return { ok: false, reason: 'name-not-found', expectedTeacherId: '' };
    }
    if (count > 1) {
      return { ok: false, reason: 'ambiguous-name', expectedTeacherId: '' };
    }

    const record = index.byName[name];
    const expectedTeacherId = record
      ? normalizeString_(record.teacherId)
      : '';

    if (!expectedTeacherId) {
      return { ok: false, reason: 'name-without-id', expectedTeacherId: '' };
    }

    return {
      ok: id === expectedTeacherId,
      reason: id === expectedTeacherId ? 'ok' : 'name-id-mismatch',
      expectedTeacherId: expectedTeacherId
    };
  }

  if (id && !index.byId[id]) {
    return { ok: false, reason: 'id-not-found', expectedTeacherId: '' };
  }

  return {
    ok: true,
    reason: id ? 'id-only' : 'blank',
    expectedTeacherId: id
  };
}

function invalidateTeachingAssignmentIntegrityCachesUnderLock_() {
  const lock = LockService.getScriptLock();
  if (!lock.hasLock()) {
    throw new Error('担当教員キャッシュ更新にはScriptLockが必要です。');
  }

  const sourceCacheKeys = getTeachingAssignmentSourceCacheKeys_();
  removeScriptCacheKeys_(sourceCacheKeys);
  const revision = bumpTeachingAssignmentRevisionUnderLock_();

  let fastSnapshotResult = null;
  try {
    fastSnapshotResult = invalidateAllTeacherUnsavedFastSnapshotsUnderLock_(
      '担当教員設定変更後にFastキャッシュを無効化しました。次回rebuildを待っています。'
    );
  } catch (error) {
    Logger.log(JSON.stringify({
      ok: false,
      event: 'teaching-assignment-fast-snapshot-invalidation-warning',
      warning: true,
      errorMessage: error && error.message ? String(error.message) : String(error)
    }));
  }

  return {
    revision: revision,
    invalidatedSourceCacheKeys: sourceCacheKeys.slice(),
    fastSnapshotResult: fastSnapshotResult
  };
}

function syncTeachingAssignmentIntegrityRows_(sheet, startRow, numRows, index) {
  const columns = getTeachingAssignmentIntegrityColumns_(sheet);
  if (columns.teacherName < 0 || columns.teacherId < 0) {
    throw new Error(
      sheet.getName() + ' に teacherName / teacherId 列がありません。'
    );
  }

  const nameRange = sheet.getRange(startRow, columns.teacherName + 1, numRows, 1);
  const idRange = sheet.getRange(startRow, columns.teacherId + 1, numRows, 1);

  SpreadsheetApp.flush();

  const names = nameRange.getValues();
  const ids = idRange.getValues();
  const formulas = idRange.getFormulas();
  const nextIds = ids.map(function(row) { return [row[0]]; });

  let changedCount = 0;
  const issues = [];

  for (let i = 0; i < numRows; i++) {
    const rowNumber = startRow + i;
    const teacherName = normalizeString_(names[i][0]);
    const teacherId = normalizeString_(ids[i][0]);
    const formula = formulas[i][0] || '';
    const resolved = resolveTeacherAssignmentIntegrity_(
      teacherName,
      teacherId,
      index
    );

    if (resolved.reason === 'name-id-mismatch') {
      if (formula) {
        issues.push({
          sheetName: sheet.getName(),
          rowNumber: rowNumber,
          teacherName: teacherName,
          storedTeacherId: teacherId,
          expectedTeacherId: resolved.expectedTeacherId,
          reason: 'formula-id-mismatch'
        });
        continue;
      }

      nextIds[i][0] = resolved.expectedTeacherId;
      changedCount++;
      continue;
    }

    if (!resolved.ok && resolved.reason !== 'blank') {
      issues.push({
        sheetName: sheet.getName(),
        rowNumber: rowNumber,
        teacherName: teacherName,
        storedTeacherId: teacherId,
        expectedTeacherId: resolved.expectedTeacherId || '',
        reason: resolved.reason
      });
    }
  }

  if (changedCount > 0) {
    idRange.setValues(nextIds);
    SpreadsheetApp.flush();
  }

  return {
    changedCount: changedCount,
    issueCount: issues.length,
    issues: issues
  };
}

function handleTeachingAssignmentIntegrityEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (!isTeachingAssignmentIntegrityTargetSheet_(sheet.getName())) return;

  const columns = getTeachingAssignmentIntegrityColumns_(sheet);
  if (!teachingAssignmentIntegrityRangeTouchesTarget_(e.range, columns)) return;

  const firstRow = Math.max(2, e.range.getRow());
  const lastRow = e.range.getRow() + e.range.getNumRows() - 1;
  if (lastRow < 2 || firstRow > lastRow) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('担当教員整合性チェックのLockを取得できませんでした。');
  }

  try {
    const result = syncTeachingAssignmentIntegrityRows_(
      sheet,
      firstRow,
      lastRow - firstRow + 1,
      getTeacherAssignmentIntegrityIndex_()
    );
    const invalidation = invalidateTeachingAssignmentIntegrityCachesUnderLock_();

    Logger.log(JSON.stringify({
      ok: true,
      event: 'teaching-assignment-integrity-edit',
      sheetName: sheet.getName(),
      firstRow: firstRow,
      lastRow: lastRow,
      changedCount: result.changedCount,
      issueCount: result.issueCount,
      revision: invalidation.revision
    }));

    return {
      ok: true,
      sheetName: sheet.getName(),
      result: result,
      invalidation: invalidation
    };
  } finally {
    lock.releaseLock();
  }
}

function handleTeachingAssignmentIntegrityChange(e) {
  if (!e || !e.source) return;

  const type = normalizeString_(e.changeType).toUpperCase();
  if (['INSERT_ROW', 'REMOVE_ROW', 'INSERT_COLUMN', 'REMOVE_COLUMN'].indexOf(type) === -1) {
    return;
  }

  const sheet = e.source.getActiveSheet();
  if (!sheet || !isTeachingAssignmentIntegrityTargetSheet_(sheet.getName())) {
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('担当教員構造変更処理のLockを取得できませんでした。');
  }

  try {
    const invalidation = invalidateTeachingAssignmentIntegrityCachesUnderLock_();
    return {
      ok: true,
      sheetName: sheet.getName(),
      changeType: type,
      invalidation: invalidation
    };
  } finally {
    lock.releaseLock();
  }
}

function auditTeacherAssignmentIntegrity() {
  const ss = getOperationSpreadsheet();
  const index = getTeacherAssignmentIntegrityIndex_();
  const issues = [];

  [CONFIG.SHEETS.TIMETABLE, CONFIG.SHEETS.CLASS_TEACHER_TEAMS]
    .forEach(function(sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return;

      const columns = getTeachingAssignmentIntegrityColumns_(sheet);
      if (columns.teacherName < 0 || columns.teacherId < 0) {
        issues.push({
          sheetName: sheetName,
          rowNumber: 1,
          teacherName: '',
          storedTeacherId: '',
          expectedTeacherId: '',
          reason: 'required-column-missing'
        });
        return;
      }

      const numRows = sheet.getLastRow() - 1;
      const names = sheet.getRange(2, columns.teacherName + 1, numRows, 1).getValues();
      const ids = sheet.getRange(2, columns.teacherId + 1, numRows, 1).getValues();
      const formulas = sheet.getRange(2, columns.teacherId + 1, numRows, 1).getFormulas();

      for (let i = 0; i < numRows; i++) {
        const teacherName = normalizeString_(names[i][0]);
        const teacherId = normalizeString_(ids[i][0]);
        const resolved = resolveTeacherAssignmentIntegrity_(
          teacherName,
          teacherId,
          index
        );

        if (resolved.ok || resolved.reason === 'blank' || resolved.reason === 'id-only') {
          continue;
        }

        issues.push({
          sheetName: sheetName,
          rowNumber: i + 2,
          teacherName: teacherName,
          storedTeacherId: teacherId,
          expectedTeacherId: resolved.expectedTeacherId || '',
          reason: formulas[i][0] && resolved.reason === 'name-id-mismatch'
            ? 'formula-id-mismatch'
            : resolved.reason
        });
      }
    });

  Logger.log(JSON.stringify({
    ok: issues.length === 0,
    issueCount: issues.length,
    issues: issues
  }));

  return {
    ok: issues.length === 0,
    issueCount: issues.length,
    issues: issues
  };
}

function repairTeacherAssignmentIntegrityNow() {
  const ss = getOperationSpreadsheet();
  const index = getTeacherAssignmentIntegrityIndex_();
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(20000)) {
    throw new Error('担当教員一括修復のLockを取得できませんでした。');
  }

  try {
    let repairedCount = 0;
    const results = [];

    [CONFIG.SHEETS.TIMETABLE, CONFIG.SHEETS.CLASS_TEACHER_TEAMS]
      .forEach(function(sheetName) {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet || sheet.getLastRow() < 2) return;

        const result = syncTeachingAssignmentIntegrityRows_(
          sheet,
          2,
          sheet.getLastRow() - 1,
          index
        );
        repairedCount += result.changedCount;
        results.push({ sheetName: sheetName, result: result });
      });

    const invalidation = invalidateTeachingAssignmentIntegrityCachesUnderLock_();
    const after = auditTeacherAssignmentIntegrity();

    return {
      ok: after.issueCount === 0,
      repairedCount: repairedCount,
      afterIssueCount: after.issueCount,
      remainingIssues: after.issues,
      invalidation: invalidation,
      results: results
    };
  } finally {
    lock.releaseLock();
  }
}

function getTeachingAssignmentIntegrityTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const editCount = triggers.filter(function(trigger) {
    return trigger.getHandlerFunction() === TEACHING_ASSIGNMENT_INTEGRITY_EDIT_HANDLER_;
  }).length;
  const changeCount = triggers.filter(function(trigger) {
    return trigger.getHandlerFunction() === TEACHING_ASSIGNMENT_INTEGRITY_CHANGE_HANDLER_;
  }).length;

  return {
    ok: editCount === 1 && changeCount === 1,
    editTriggerCount: editCount,
    changeTriggerCount: changeCount,
    triggerCount: editCount + changeCount,
    revision: getTeachingAssignmentRevision_()
  };
}

function installTeachingAssignmentIntegrityTrigger() {
  removeTeachingAssignmentIntegrityTrigger();

  const spreadsheet = getOperationSpreadsheet();

  ScriptApp.newTrigger(TEACHING_ASSIGNMENT_INTEGRITY_EDIT_HANDLER_)
    .forSpreadsheet(spreadsheet)
    .onEdit()
    .create();

  ScriptApp.newTrigger(TEACHING_ASSIGNMENT_INTEGRITY_CHANGE_HANDLER_)
    .forSpreadsheet(spreadsheet)
    .onChange()
    .create();

  return getTeachingAssignmentIntegrityTriggerStatus();
}

function removeTeachingAssignmentIntegrityTrigger() {
  const handlers = {};
  handlers[TEACHING_ASSIGNMENT_INTEGRITY_EDIT_HANDLER_] = true;
  handlers[TEACHING_ASSIGNMENT_INTEGRITY_CHANGE_HANDLER_] = true;

  let deletedCount = 0;
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (!handlers[trigger.getHandlerFunction()]) return;
    ScriptApp.deleteTrigger(trigger);
    deletedCount++;
  });

  return {
    ok: true,
    deletedCount: deletedCount
  };
}
