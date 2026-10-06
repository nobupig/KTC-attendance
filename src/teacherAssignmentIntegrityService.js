const TEACHING_ASSIGNMENT_REVISION_PROPERTY_ = 'KTC_TEACHING_ASSIGNMENT_REVISION';
const TEACHING_ASSIGNMENT_EDIT_TRIGGER_HANDLER_ = 'handleTeachingAssignmentIntegrityEdit';
const TEACHING_ASSIGNMENT_CHANGE_TRIGGER_HANDLER_ = 'handleTeachingAssignmentIntegrityChange';

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

function getTeacherAssignmentIntegrityIndex_() {
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

function isTeachingAssignmentTargetSheet_(sheetName) {
  return sheetName === CONFIG.SHEETS.TIMETABLE ||
    sheetName === CONFIG.SHEETS.CLASS_TEACHER_TEAMS;
}

function getTeachingAssignmentColumnIndexes_(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = lastCol > 0
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    : [];

  return {
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    weekday: findColumnIndex_(headers, ['weekday', '曜日']),
    period: findColumnIndex_(headers, ['period', '時限']),
    teacherName: findColumnIndex_(headers, ['teacherName', '担当者名', 'name']),
    teacherId: findColumnIndex_(headers, ['teacherId', 'TeacherID'])
  };
}

function rangeTouchesTeachingAssignmentColumns_(range, columns) {
  if (!range || !columns) return false;

  const first = range.getColumn() - 1;
  const last = first + range.getNumColumns() - 1;
  const targetColumns = [
    columns.classId,
    columns.weekday,
    columns.period,
    columns.teacherName,
    columns.teacherId
  ].filter(function(index) {
    return index >= 0;
  });

  return targetColumns.some(function(index) {
    return index >= first && index <= last;
  });
}

function resolveTeacherIntegrityRecord_(teacherId, teacherName, index) {
  const normalizedId = normalizeString_(teacherId);
  const normalizedName = normalizeString_(teacherName);

  if (normalizedName) {
    if (index.ambiguousNames[normalizedName]) {
      return {
        record: null,
        reason: 'ambiguous-name'
      };
    }

    if (!index.byName[normalizedName]) {
      return {
        record: null,
        reason: 'name-not-found'
      };
    }

    return {
      record: index.byName[normalizedName],
      reason: normalizedId === index.byName[normalizedName].teacherId
        ? 'ok'
        : 'name-id-mismatch'
    };
  }

  if (normalizedId) {
    if (!index.byId[normalizedId]) {
      return {
        record: null,
        reason: 'id-not-found'
      };
    }

    return {
      record: index.byId[normalizedId],
      reason: 'id-only'
    };
  }

  return {
    record: null,
    reason: 'blank'
  };
}

function invalidateTeachingAssignmentReadCaches_() {
  removeScriptCacheKeys_([
    'sheetData__OPERATION__' + CONFIG.SHEETS.TIMETABLE,
    'sheetData__OPERATION__' + CONFIG.SHEETS.CLASS_TEACHER_TEAMS
  ]);
}

function syncTeacherAssignmentRows_(sheet, startRow, numRows, index) {
  const columns = getTeachingAssignmentColumnIndexes_(sheet);
  if (columns.teacherName < 0 || columns.teacherId < 0) {
    throw new Error(
      sheet.getName() + ' の teacherName / teacherId 列を確認してください。'
    );
  }

  if (numRows <= 0) {
    return {
      changedCount: 0,
      mismatchCount: 0,
      unresolvedCount: 0,
      issues: []
    };
  }

  const teacherNameValues = sheet
    .getRange(startRow, columns.teacherName + 1, numRows, 1)
    .getValues();
  const teacherIdValues = sheet
    .getRange(startRow, columns.teacherId + 1, numRows, 1)
    .getValues();

  const nextTeacherIds = [];
  const issues = [];
  let changedCount = 0;
  let mismatchCount = 0;
  let unresolvedCount = 0;

  for (let i = 0; i < numRows; i++) {
    const teacherName = normalizeString_(teacherNameValues[i][0]);
    const teacherId = normalizeString_(teacherIdValues[i][0]);
    const rowNumber = startRow + i;
    const resolved = resolveTeacherIntegrityRecord_(
      teacherId,
      teacherName,
      index
    );

    if (resolved.reason === 'blank' || resolved.reason === 'id-only') {
      nextTeacherIds.push([teacherIdValues[i][0]]);
      continue;
    }

    if (!resolved.record) {
      unresolvedCount++;
      issues.push({
        sheetName: sheet.getName(),
        rowNumber: rowNumber,
        teacherName: teacherName,
        storedTeacherId: teacherId,
        expectedTeacherId: '',
        reason: resolved.reason
      });
      nextTeacherIds.push([teacherIdValues[i][0]]);
      continue;
    }

    if (resolved.reason === 'name-id-mismatch') {
      mismatchCount++;
    }

    const expectedTeacherId = resolved.record.teacherId;
    if (teacherId !== expectedTeacherId) {
      changedCount++;
      nextTeacherIds.push([expectedTeacherId]);
    } else {
      nextTeacherIds.push([teacherIdValues[i][0]]);
    }
  }

  if (changedCount > 0) {
    sheet
      .getRange(startRow, columns.teacherId + 1, numRows, 1)
      .setValues(nextTeacherIds);
    SpreadsheetApp.flush();
  }

  return {
    changedCount: changedCount,
    mismatchCount: mismatchCount,
    unresolvedCount: unresolvedCount,
    issues: issues
  };
}

function handleTeachingAssignmentIntegrityEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  if (!isTeachingAssignmentTargetSheet_(sheetName)) return;

  const columns = getTeachingAssignmentColumnIndexes_(sheet);
  if (!rangeTouchesTeachingAssignmentColumns_(e.range, columns)) return;

  const editStartRow = e.range.getRow();
  const editEndRow = editStartRow + e.range.getNumRows() - 1;
  const startRow = Math.max(2, editStartRow);
  if (editEndRow < 2 || startRow > editEndRow) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('担当教員整合性チェックのロックを取得できませんでした。');
  }

  let syncResult = null;
  let revision = '';

  try {
    const index = getTeacherAssignmentIntegrityIndex_();

    syncResult = syncTeacherAssignmentRows_(
      sheet,
      startRow,
      editEndRow - startRow + 1,
      index
    );

    revision = bumpTeachingAssignmentRevision_();
    invalidateTeachingAssignmentReadCaches_();
  } finally {
    lock.releaseLock();
  }

  Logger.log(JSON.stringify({
    event: 'teaching-assignment-edit-v193',
    sheetName: sheetName,
    startRow: startRow,
    endRow: editEndRow,
    changedCount: syncResult ? syncResult.changedCount : 0,
    mismatchCount: syncResult ? syncResult.mismatchCount : 0,
    unresolvedCount: syncResult ? syncResult.unresolvedCount : 0,
    revision: revision
  }));

  return {
    ok: true,
    sheetName: sheetName,
    syncResult: syncResult,
    revision: revision
  };
}

function handleTeachingAssignmentIntegrityChange(e) {
  if (!e || !e.source) return;

  const changeType = normalizeString_(e.changeType).toUpperCase();
  const structuralTypes = {
    INSERT_ROW: true,
    REMOVE_ROW: true,
    INSERT_COLUMN: true,
    REMOVE_COLUMN: true
  };

  if (!structuralTypes[changeType]) return;

  const activeSheet = e.source.getActiveSheet();
  if (!activeSheet || !isTeachingAssignmentTargetSheet_(activeSheet.getName())) {
    return;
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('担当教員構造変更チェックのロックを取得できませんでした。');
  }

  let revision = '';

  try {
    revision = bumpTeachingAssignmentRevision_();
    invalidateTeachingAssignmentReadCaches_();
  } finally {
    lock.releaseLock();
  }

  Logger.log(JSON.stringify({
    event: 'teaching-assignment-structural-change-v193',
    sheetName: activeSheet.getName(),
    changeType: changeType,
    revision: revision
  }));

  return {
    ok: true,
    sheetName: activeSheet.getName(),
    changeType: changeType,
    revision: revision
  };
}

function auditTeacherAssignmentIntegrity() {
  const ss = getOperationSpreadsheet();
  const index = getTeacherAssignmentIntegrityIndex_();
  const targets = [
    CONFIG.SHEETS.TIMETABLE,
    CONFIG.SHEETS.CLASS_TEACHER_TEAMS
  ];
  const issues = [];

  targets.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;

    const columns = getTeachingAssignmentColumnIndexes_(sheet);
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
    const teacherNames = sheet
      .getRange(2, columns.teacherName + 1, numRows, 1)
      .getValues();
    const teacherIds = sheet
      .getRange(2, columns.teacherId + 1, numRows, 1)
      .getValues();

    for (let i = 0; i < numRows; i++) {
      const teacherName = normalizeString_(teacherNames[i][0]);
      const teacherId = normalizeString_(teacherIds[i][0]);
      const resolved = resolveTeacherIntegrityRecord_(
        teacherId,
        teacherName,
        index
      );

      if (resolved.reason === 'blank' || resolved.reason === 'id-only' || resolved.reason === 'ok') {
        continue;
      }

      issues.push({
        sheetName: sheetName,
        rowNumber: i + 2,
        teacherName: teacherName,
        storedTeacherId: teacherId,
        expectedTeacherId: resolved.record ? resolved.record.teacherId : '',
        reason: resolved.reason
      });
    }
  });

  return {
    ok: true,
    issueCount: issues.length,
    issues: issues
  };
}

function repairTeacherAssignmentIntegrityNow() {
  const before = auditTeacherAssignmentIntegrity();
  const ss = getOperationSpreadsheet();
  const index = getTeacherAssignmentIntegrityIndex_();
  const targets = [
    CONFIG.SHEETS.TIMETABLE,
    CONFIG.SHEETS.CLASS_TEACHER_TEAMS
  ];

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    throw new Error('担当教員一括修復のロックを取得できませんでした。');
  }

  let repairedCount = 0;
  const results = [];
  let revision = '';

  try {
    targets.forEach(function(sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) return;

      const result = syncTeacherAssignmentRows_(
        sheet,
        2,
        sheet.getLastRow() - 1,
        index
      );

      repairedCount += result.changedCount;
      results.push({
        sheetName: sheetName,
        result: result
      });
    });

    revision = bumpTeachingAssignmentRevision_();
    invalidateTeachingAssignmentReadCaches_();
  } finally {
    lock.releaseLock();
  }

  const after = auditTeacherAssignmentIntegrity();

  return {
    ok: after.issueCount === 0,
    beforeIssueCount: before.issueCount,
    repairedCount: repairedCount,
    afterIssueCount: after.issueCount,
    revision: revision,
    results: results,
    remainingIssues: after.issues
  };
}

function getTeachingAssignmentIntegrityTriggerStatus() {
  const triggers = ScriptApp.getProjectTriggers();
  const editCount = triggers.filter(function(trigger) {
    return trigger.getHandlerFunction() === TEACHING_ASSIGNMENT_EDIT_TRIGGER_HANDLER_;
  }).length;
  const changeCount = triggers.filter(function(trigger) {
    return trigger.getHandlerFunction() === TEACHING_ASSIGNMENT_CHANGE_TRIGGER_HANDLER_;
  }).length;

  return {
    ok: true,
    editTriggerCount: editCount,
    changeTriggerCount: changeCount,
    triggerCount: editCount + changeCount,
    revision: getTeachingAssignmentRevision_()
  };
}

function installTeachingAssignmentIntegrityTrigger() {
  const handlers = {};
  handlers[TEACHING_ASSIGNMENT_EDIT_TRIGGER_HANDLER_] = true;
  handlers[TEACHING_ASSIGNMENT_CHANGE_TRIGGER_HANDLER_] = true;

  const existing = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return !!handlers[trigger.getHandlerFunction()];
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
  const handlers = {};
  handlers[TEACHING_ASSIGNMENT_EDIT_TRIGGER_HANDLER_] = true;
  handlers[TEACHING_ASSIGNMENT_CHANGE_TRIGGER_HANDLER_] = true;

  const existing = ScriptApp.getProjectTriggers().filter(function(trigger) {
    return !!handlers[trigger.getHandlerFunction()];
  });

  existing.forEach(function(trigger) {
    ScriptApp.deleteTrigger(trigger);
  });

  const status = getTeachingAssignmentIntegrityTriggerStatus();
  status.deletedCount = existing.length;
  return status;
}
