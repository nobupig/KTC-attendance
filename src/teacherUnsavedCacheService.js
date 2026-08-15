const TEACHER_UNSAVED_CACHE_SHEETS_ = Object.freeze({
  SUMMARY: 'teacherUnsavedSummaryCache',
  DETAIL: 'teacherUnsavedDetailCache'
});

const TEACHER_UNSAVED_SUMMARY_CACHE_HEADERS_ = Object.freeze([
  'snapshotId',
  'cacheDate',
  'teacherId',
  'teacherName',
  'teacherEmail',
  'startYmd',
  'endYmd',
  'unsavedCount',
  'firstDate',
  'lastDate',
  'detailStartRow',
  'detailCount',
  'checkedAt',
  'status',
  'errorMessage'
]);

const TEACHER_UNSAVED_DETAIL_CACHE_HEADERS_ = Object.freeze([
  'snapshotId',
  'cacheDate',
  'teacherId',
  'teacherName',
  'teacherEmail',
  'date',
  'period',
  'classId',
  'sessionNumber',
  'subjectName',
  'targetLabel',
  'isExperiment',
  'displayKey',
  'saveKey',
  'sortKey',
  'checkedAt'
]);

function getTeacherUnsavedCacheSchema() {
  return {
    summary: {
      sheetName: TEACHER_UNSAVED_CACHE_SHEETS_.SUMMARY,
      headers: TEACHER_UNSAVED_SUMMARY_CACHE_HEADERS_.slice()
    },
    detail: {
      sheetName: TEACHER_UNSAVED_CACHE_SHEETS_.DETAIL,
      headers: TEACHER_UNSAVED_DETAIL_CACHE_HEADERS_.slice()
    }
  };
}

function validateTeacherUnsavedCacheSheets() {
  const validation = validateTeacherUnsavedCacheSheets_();
  return {
    ok: validation.ok,
    summary: validation.summary.publicResult,
    detail: validation.detail.publicResult,
    errors: validation.errors.slice()
  };
}

function rebuildTeacherUnsavedSummaryCache() {
  const validation = validateTeacherUnsavedCacheSheets_();
  if (!validation.ok) {
    throw new Error(buildTeacherUnsavedCacheValidationError_(validation));
  }

  const startedAt = Date.now();

  try {
    const snapshot = buildTeacherUnsavedCacheSnapshot_();
    publishTeacherUnsavedCacheSnapshot_(snapshot);

    return buildTeacherUnsavedCacheBuildResult_(snapshot, Date.now() - startedAt, true);
  } catch (error) {
    try {
      markTeacherUnsavedCachePublishError_(error);
    } catch (markError) {
      Logger.log('Teacher unsaved cache error status update failed: ' + markError);
    }
    throw error;
  }
}

function debugBuildTeacherUnsavedCachePreview() {
  const startedAt = Date.now();
  const snapshot = buildTeacherUnsavedCacheSnapshot_();
  const result = buildTeacherUnsavedCacheBuildResult_(snapshot, Date.now() - startedAt, false);
  result.cacheSheets = validateTeacherUnsavedCacheSheets();
  return result;
}

function getTeacherUnsavedSummaryFast() {
  return getTeacherUnsavedSummaryFastContext_().result;
}

function getTeacherUnsavedSummaryFastContext_() {
  const user = getCurrentUserContext();
  if (!user || !user.teacherId) {
    throw new Error('ログインユーザー情報を取得できませんでした。');
  }

  const teacherId = normalizeString_(user.teacherId);
  const todayYmd = getTeacherUnsavedCacheDateContext_(new Date()).cacheDate;
  const validation = validateTeacherUnsavedCacheSheets_();

  if (!validation.ok) {
    return {
      result: buildTeacherUnsavedFastSummaryFailure_(
        'unavailable',
        teacherId,
        todayYmd,
        buildTeacherUnsavedCacheValidationError_(validation)
      ),
      validation: validation,
      row: null
    };
  }

  const summaryState = readTeacherUnsavedSummaryState_(
    validation.summary.sheet,
    teacherId,
    todayYmd
  );

  if (!summaryState.row) {
    return {
      result: buildTeacherUnsavedFastSummaryFailure_(
        'missing',
        teacherId,
        todayYmd,
        '当日分のサマリーキャッシュがありません。'
      ),
      validation: validation,
      row: null
    };
  }

  const row = summaryState.row;

  if (row.cacheDate !== todayYmd) {
    return {
      result: buildTeacherUnsavedFastSummaryFailureFromRow_(
        'stale',
        row,
        'サマリーキャッシュが当日分ではありません。'
      ),
      validation: validation,
      row: row
    };
  }

  if (row.status === 'error') {
    return {
      result: buildTeacherUnsavedFastSummaryFailureFromRow_(
        'error',
        row,
        row.errorMessage || 'キャッシュ再構築でエラーが発生しました。'
      ),
      validation: validation,
      row: row
    };
  }

  if (row.status !== 'ready') {
    return {
      result: buildTeacherUnsavedFastSummaryFailureFromRow_(
        'unavailable',
        row,
        row.status
          ? 'キャッシュの状態が ready ではありません: ' + row.status
          : 'キャッシュの状態が設定されていません。',
        row.status || ''
      ),
      validation: validation,
      row: row
    };
  }

  const integrity = validateTeacherUnsavedSummaryRowIntegrity_(
    row,
    validation.detail.sheet
  );

  if (!integrity.ok) {
    return {
      result: buildTeacherUnsavedFastSummaryFailureFromRow_(
        'unavailable',
        row,
        integrity.errorMessage
      ),
      validation: validation,
      row: row
    };
  }

  const result = {
    ok: true,
    status: 'ready',
    cacheStatus: 'ready',
    teacherId: row.teacherId,
    teacherName: row.teacherName,
    teacherEmail: row.teacherEmail,
    count: row.unsavedCount,
    unsavedCount: row.unsavedCount,
    firstDate: row.firstDate,
    lastDate: row.lastDate,
    detailCount: row.detailCount,
    checkedAt: formatTeacherUnsavedCacheDateTime_(row.checkedAt),
    startYmd: row.startYmd,
    endYmd: row.endYmd,
    checkedRange: {
      start: row.startYmd,
      end: row.endYmd
    },
    cacheDate: row.cacheDate,
    snapshotId: row.snapshotId,
    errorMessage: ''
  };

  return {
    result: result,
    validation: validation,
    row: row
  };
}

function getTeacherUnsavedDetailsFast(options) {
  const paging = validateTeacherUnsavedDetailPagingOptions_(options);
  const summaryContext = getTeacherUnsavedSummaryFastContext_();
  const summary = summaryContext.result;

  if (!summary.ok || summary.status !== 'ready') {
    return buildTeacherUnsavedFastDetailsFailure_(summary, paging);
  }

  const totalCount = summary.detailCount;
  const validation = summaryContext.validation;
  const currentRow = summaryContext.row;

  if (
    !currentRow ||
    currentRow.status !== 'ready' ||
    currentRow.snapshotId !== summary.snapshotId ||
    currentRow.detailCount !== totalCount
  ) {
    return buildTeacherUnsavedFastDetailsFailure_(
      buildTeacherUnsavedFastSummaryFailureFromRow_(
        'unavailable',
        currentRow || summary,
        'サマリーキャッシュが詳細取得中に更新されたか、整合性を確認できませんでした。'
      ),
      paging
    );
  }

  if (totalCount === 0 || paging.offset >= totalCount) {
    return buildTeacherUnsavedFastDetailsSuccess_(summary, paging, [], totalCount);
  }

  if (currentRow.detailStartRow === null || currentRow.detailStartRow < 2) {
    return buildTeacherUnsavedFastDetailsFailure_(
      buildTeacherUnsavedFastSummaryFailureFromRow_(
        'unavailable',
        currentRow,
        'detailStartRow が不正です。'
      ),
      paging
    );
  }

  const readCount = Math.min(paging.limit, totalCount - paging.offset);
  const startRow = currentRow.detailStartRow + paging.offset;
  const values = validation.detail.sheet
    .getRange(startRow, 1, readCount, TEACHER_UNSAVED_DETAIL_CACHE_HEADERS_.length)
    .getValues();

  const items = [];
  for (let i = 0; i < values.length; i++) {
    const item = buildTeacherUnsavedDetailObject_(values[i]);
    if (
      item.snapshotId !== summary.snapshotId ||
      item.teacherId !== summary.teacherId ||
      item.cacheDate !== summary.cacheDate
    ) {
      return buildTeacherUnsavedFastDetailsFailure_(
        buildTeacherUnsavedFastSummaryFailureFromRow_(
          'unavailable',
          currentRow,
          '詳細キャッシュの snapshotId、teacherId、または cacheDate が一致しません。'
        ),
        paging
      );
    }
    items.push(item);
  }

  return buildTeacherUnsavedFastDetailsSuccess_(summary, paging, items, totalCount);
}

function debugCompareTeacherUnsavedCacheForCurrentUser() {
  const user = getCurrentUserContext();
  if (!user || !user.teacherId) {
    throw new Error('ログインユーザー情報を取得できませんでした。');
  }

  const teacherId = normalizeString_(user.teacherId);
  const legacySummary = getTeacherUnsavedSummary();
  const legacyDetails = getTeacherUnsavedDetails();
  const fastSummary = getTeacherUnsavedSummaryFast();

  if (!fastSummary.ok || fastSummary.status !== 'ready') {
    return {
      ok: false,
      teacherId: teacherId,
      fastStatus: fastSummary.status,
      fastErrorMessage: fastSummary.errorMessage || '',
      legacyCount: legacySummary && typeof legacySummary.count === 'number'
        ? legacySummary.count
        : null,
      legacyDetailCount: legacyDetails && Array.isArray(legacyDetails.items)
        ? legacyDetails.items.length
        : null,
      message: 'Fast API が ready ではないため、キー集合の比較を実行できませんでした。'
    };
  }

  const fastDetails = getTeacherUnsavedDetailsFast({
    limit: Math.max(fastSummary.detailCount, 1),
    offset: 0
  });

  if (!fastDetails.ok) {
    return {
      ok: false,
      teacherId: teacherId,
      fastStatus: fastDetails.status,
      fastErrorMessage: fastDetails.errorMessage || '',
      message: 'Fast詳細APIが ready ではないため、キー集合の比較を実行できませんでした。'
    };
  }

  const context = getTeacherUnsavedContext_(teacherId);
  const legacyItems = legacyDetails && Array.isArray(legacyDetails.items)
    ? legacyDetails.items
    : [];
  const legacyKeys = buildTeacherUnsavedLegacyDebugKeySets_(legacyItems, context.classMap || {});
  const fastKeys = buildTeacherUnsavedFastDebugKeySets_(fastDetails.items || []);

  const onlyLegacyDisplayKeys = subtractTeacherUnsavedKeySets_(
    legacyKeys.displayKeys,
    fastKeys.displayKeys
  );
  const onlyFastDisplayKeys = subtractTeacherUnsavedKeySets_(
    fastKeys.displayKeys,
    legacyKeys.displayKeys
  );
  const onlyLegacySaveKeys = subtractTeacherUnsavedKeySets_(
    legacyKeys.saveKeys,
    fastKeys.saveKeys
  );
  const onlyFastSaveKeys = subtractTeacherUnsavedKeySets_(
    fastKeys.saveKeys,
    legacyKeys.saveKeys
  );

  const legacyCount = legacySummary && typeof legacySummary.count === 'number'
    ? legacySummary.count
    : null;

  return {
    ok: true,
    teacherId: teacherId,
    snapshotId: fastSummary.snapshotId,
    counts: {
      legacySummary: legacyCount,
      legacyDetails: legacyItems.length,
      fastSummary: fastSummary.count,
      fastDetails: fastDetails.totalCount
    },
    matches: {
      summaryCount: legacyCount === fastSummary.count,
      detailCount: legacyItems.length === fastDetails.totalCount,
      displayKeys: onlyLegacyDisplayKeys.length === 0 && onlyFastDisplayKeys.length === 0,
      saveKeys: onlyLegacySaveKeys.length === 0 && onlyFastSaveKeys.length === 0
    },
    differences: {
      onlyLegacyDisplayKeys: onlyLegacyDisplayKeys.slice(0, 10),
      onlyFastDisplayKeys: onlyFastDisplayKeys.slice(0, 10),
      onlyLegacySaveKeys: onlyLegacySaveKeys.slice(0, 10),
      onlyFastSaveKeys: onlyFastSaveKeys.slice(0, 10)
    }
  };
}

function debugLogTeacherUnsavedSummaryFast() {
  const result = getTeacherUnsavedSummaryFast();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function debugLogTeacherUnsavedDetailsFast() {
  const result = getTeacherUnsavedDetailsFast({ limit: 10, offset: 0 });
  const summary = {
    ok: result.ok,
    status: result.status,
    totalCount: result.totalCount,
    limit: result.limit,
    offset: result.offset,
    hasMore: result.hasMore,
    checkedAt: result.checkedAt,
    startYmd: result.startYmd,
    endYmd: result.endYmd,
    snapshotId: result.snapshotId
  };
  const items = (result.items || []).slice(0, 10).map(function(item) {
    return {
      date: item.date,
      period: item.period,
      classId: item.classId,
      subjectName: item.subjectName,
      displayKey: item.displayKey,
      saveKey: item.saveKey
    };
  });

  Logger.log(JSON.stringify(summary, null, 2));
  Logger.log(JSON.stringify(items, null, 2));
  return result;
}

function debugLogCompareTeacherUnsavedCacheForCurrentUser() {
  const result = debugCompareTeacherUnsavedCacheForCurrentUser();
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

function buildTeacherUnsavedCacheSnapshot_() {
  const checkedAt = new Date();
  const dateContext = getTeacherUnsavedCacheDateContext_(checkedAt);
  const snapshotId = buildTeacherUnsavedSnapshotId_(checkedAt);
  const warnings = [];
  const sources = loadTeacherUnsavedCacheSources_();
  const teacherIndex = buildTeacherUnsavedTeacherIndex_(sources.teachers, warnings);
  const classMap = buildTeacherUnsavedClassMap_(sources.classes);
  const assignmentMap = buildTeacherUnsavedAssignmentMap_(
    sources.timetable,
    sources.classTeacherTeams,
    teacherIndex,
    warnings
  );
  const savedKeySet = buildTeacherUnsavedSavedKeySet_(
    sources.attendanceSessions,
    dateContext.startYmd,
    dateContext.endYmd
  );

  const detailsByTeacherId = {};
  const seenByTeacherId = {};

  teacherIndex.teachers.forEach(function(teacher) {
    detailsByTeacherId[teacher.teacherId] = [];
    seenByTeacherId[teacher.teacherId] = {};
  });

  const csCol = resolveTeacherUnsavedRequiredColumns_(
    'classSessions',
    sources.classSessions.headers,
    {
      classId: ['classId', 'ClassID'],
      date: ['date', '日付'],
      period: ['period', '時限'],
      sessionNumber: ['sessionNumber', '回', '回数']
    },
    ['classId', 'date', 'period']
  );

  if (dateContext.endYmd >= dateContext.startYmd) {
    sources.classSessions.rows.forEach(function(row, index) {
      const ymd = normalizeTeacherUnsavedSourceYmd_(
        row[csCol.date],
        sources.classSessions.dateDisplayValues[index]
      );

      if (!ymd || ymd < dateContext.startYmd || ymd > dateContext.endYmd) return;

      const classId = normalizeString_(row[csCol.classId]);
      const period = normalizeString_(row[csCol.period]);
      if (!classId || !period) return;

      const weekday = getWeekdayFromYmdJst_(ymd);
      const assignmentKey = [classId, weekday, period].join('__');
      const assignment = assignmentMap[assignmentKey];
      if (!assignment || !assignment.teacherIds.length) return;

      const cls = classMap[classId] || {};
      const displayKey = buildTeacherUnsavedDisplayKey_(cls, classId, ymd, period);
      const saveKey = [classId, ymd, period].join('__');

      if (savedKeySet[saveKey] || savedKeySet[displayKey]) return;

      const isExperiment = isTeacherUnsavedExperimentClass_(classId);
      const targetLabel = buildTeacherUnsavedCacheTargetLabel_(cls, isExperiment);
      const sortKey = buildTeacherUnsavedDetailSortKey_(ymd, period, displayKey);

      assignment.teacherIds.forEach(function(teacherId) {
        const teacher = teacherIndex.byId[teacherId];
        if (!teacher) return;
        if (seenByTeacherId[teacherId][displayKey]) return;

        seenByTeacherId[teacherId][displayKey] = true;
        detailsByTeacherId[teacherId].push({
          snapshotId: snapshotId,
          cacheDate: dateContext.cacheDate,
          teacherId: teacher.teacherId,
          teacherName: teacher.teacherName,
          teacherEmail: teacher.teacherEmail,
          date: ymd,
          period: period,
          classId: classId,
          sessionNumber: csCol.sessionNumber !== -1 ? row[csCol.sessionNumber] : '',
          subjectName: cls.subjectName || '',
          targetLabel: targetLabel,
          isExperiment: isExperiment,
          displayKey: displayKey,
          saveKey: saveKey,
          sortKey: sortKey,
          checkedAt: checkedAt
        });
      });
    });
  }

  const summaryRows = [];
  const detailRows = [];
  let unsavedTeacherCount = 0;

  teacherIndex.teachers.forEach(function(teacher) {
    const items = detailsByTeacherId[teacher.teacherId] || [];
    items.sort(compareTeacherUnsavedDetailItems_);

    const detailCount = items.length;
    const detailStartRow = detailCount > 0 ? detailRows.length + 2 : '';
    const firstDate = detailCount > 0 ? items[0].date : '';
    const lastDate = detailCount > 0 ? items[detailCount - 1].date : '';

    if (detailCount > 0) unsavedTeacherCount += 1;

    items.forEach(function(item) {
      detailRows.push(buildTeacherUnsavedDetailRow_(item));
    });

    summaryRows.push([
      snapshotId,
      dateContext.cacheDate,
      teacher.teacherId,
      teacher.teacherName,
      teacher.teacherEmail,
      dateContext.startYmd,
      dateContext.endYmd,
      detailCount,
      firstDate,
      lastDate,
      detailStartRow,
      detailCount,
      checkedAt,
      'ready',
      ''
    ]);
  });

  return {
    snapshotId: snapshotId,
    cacheDate: dateContext.cacheDate,
    startYmd: dateContext.startYmd,
    endYmd: dateContext.endYmd,
    checkedAt: checkedAt,
    summaryRows: summaryRows,
    detailRows: detailRows,
    teacherCount: teacherIndex.teachers.length,
    unsavedTeacherCount: unsavedTeacherCount,
    warnings: warnings,
    sourceRowCounts: {
      teachers: sources.teachers.rows.length,
      timetable: sources.timetable.rows.length,
      classTeacherTeams: sources.classTeacherTeams.rows.length,
      classes: sources.classes.rows.length,
      classSessions: sources.classSessions.rows.length,
      attendanceSessions: sources.attendanceSessions.rows.length
    }
  };
}

function loadTeacherUnsavedCacheSources_() {
  const operation = getOperationSpreadsheet();
  const master = getMasterSpreadsheet();

  return {
    teachers: readTeacherUnsavedSourceSheet_(operation, CONFIG.SHEETS.TEACHERS, false),
    timetable: readTeacherUnsavedSourceSheet_(operation, CONFIG.SHEETS.TIMETABLE, false),
    classTeacherTeams: readTeacherUnsavedSourceSheet_(
      operation,
      CONFIG.SHEETS.CLASS_TEACHER_TEAMS,
      false
    ),
    classes: readTeacherUnsavedSourceSheet_(master, CONFIG.SHEETS.CLASSES, false),
    classSessions: readTeacherUnsavedSourceSheet_(
      operation,
      CONFIG.SHEETS.CLASS_SESSIONS,
      true
    ),
    attendanceSessions: readTeacherUnsavedSourceSheet_(
      operation,
      CONFIG.SHEETS.ATTENDANCE_SESSIONS,
      true
    )
  };
}

function readTeacherUnsavedSourceSheet_(spreadsheet, sheetName, includeDateDisplayValues) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(sheetName + ' シートが見つかりません。');
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 1 || lastColumn < 1) {
    throw new Error(sheetName + ' シートにヘッダーがありません。');
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const rows = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues()
    : [];
  let dateDisplayValues = [];

  if (includeDateDisplayValues && rows.length > 0) {
    const dateColumn = findColumnIndex_(headers, ['date', '日付']);
    if (dateColumn !== -1) {
      dateDisplayValues = sheet
        .getRange(2, dateColumn + 1, rows.length, 1)
        .getDisplayValues()
        .map(function(row) { return row[0]; });
    }
  }

  return {
    sheetName: sheetName,
    headers: headers,
    rows: rows,
    dateDisplayValues: dateDisplayValues
  };
}

function buildTeacherUnsavedTeacherIndex_(data, warnings) {
  const col = resolveTeacherUnsavedRequiredColumns_(
    'teachers',
    data.headers,
    {
      teacherId: ['teacherId', 'TeacherID'],
      name: ['name', '氏名', 'teacherName'],
      email: ['email', 'メールアドレス']
    },
    ['teacherId', 'name', 'email']
  );

  const teachers = [];
  const byId = {};
  const byName = {};

  data.rows.forEach(function(row, index) {
    const teacherId = normalizeString_(row[col.teacherId]);
    const teacherName = normalizeString_(row[col.name]);
    const teacherEmail = normalizeString_(row[col.email]).toLowerCase();

    const rowTeacher = {
      teacherId: teacherId,
      teacherName: teacherName,
      teacherEmail: teacherEmail
    };

    if (teacherName && !byName[teacherName]) {
      // 現行 teacherMasterBundle と同じく、名前索引は最初に現れた行を採用する。
      byName[teacherName] = rowTeacher;
    } else if (teacherName && byName[teacherName].teacherId !== teacherId) {
      addTeacherUnsavedWarning_(warnings, 'teachers: name が重複しています: ' + teacherName);
    }

    if (!teacherId) {
      addTeacherUnsavedWarning_(warnings, 'teachers row ' + (index + 2) + ': teacherId が空です。');
      return;
    }

    if (byId[teacherId]) {
      addTeacherUnsavedWarning_(warnings, 'teachers: teacherId が重複しています: ' + teacherId);
      return;
    }

    const teacher = rowTeacher;

    teachers.push(teacher);
    byId[teacherId] = teacher;
  });

  return {
    teachers: teachers,
    byId: byId,
    byName: byName
  };
}

function buildTeacherUnsavedClassMap_(data) {
  const col = resolveTeacherUnsavedRequiredColumns_(
    'classes',
    data.headers,
    {
      classId: ['classId', 'ClassID'],
      subjectId: ['subjectId', 'SubjectID'],
      subjectName: ['subjectName', '科目名'],
      grade: ['grade', '学年'],
      unit: ['unit', '対象区分', '組・コース']
    },
    ['classId', 'subjectId', 'subjectName']
  );

  const map = {};
  data.rows.forEach(function(row) {
    const classId = normalizeString_(row[col.classId]);
    if (!classId) return;

    map[classId] = {
      classId: classId,
      subjectId: col.subjectId !== -1 ? normalizeString_(row[col.subjectId]) : '',
      subjectName: col.subjectName !== -1 ? normalizeString_(row[col.subjectName]) : '',
      grade: col.grade !== -1 ? normalizeString_(row[col.grade]) : '',
      unit: col.unit !== -1 ? normalizeString_(row[col.unit]) : ''
    };
  });

  return map;
}

function buildTeacherUnsavedAssignmentMap_(timetableData, teamData, teacherIndex, warnings) {
  const ttCol = resolveTeacherUnsavedRequiredColumns_(
    'timetable',
    timetableData.headers,
    {
      classId: ['classId', 'ClassID'],
      weekday: ['weekday', '曜日'],
      period: ['period', '時限'],
      teacherName: ['teacherName', '担当者名', 'name'],
      teacherId: ['teacherId', 'TeacherID']
    },
    ['classId', 'weekday', 'period']
  );
  const teamCol = resolveTeacherUnsavedRequiredColumns_(
    'classTeacherTeams',
    teamData.headers,
    {
      classId: ['classId', 'ClassID'],
      weekday: ['weekday', '曜日'],
      period: ['period', '時限'],
      teacherName: ['teacherName', '担当者名', 'name'],
      teacherId: ['teacherId', 'TeacherID'],
      roleType: ['roleType', '役割']
    },
    []
  );

  const map = {};

  timetableData.rows.forEach(function(row, index) {
    const classId = normalizeString_(row[ttCol.classId]);
    const period = normalizeString_(row[ttCol.period]);
    const weekday = normalizeWeekday_(row[ttCol.weekday]);
    const teacherName = ttCol.teacherName !== -1
      ? normalizeString_(row[ttCol.teacherName])
      : '';
    let teacherId = ttCol.teacherId !== -1
      ? normalizeString_(row[ttCol.teacherId])
      : '';

    if (!teacherId && teacherName && teacherIndex.byName[teacherName]) {
      teacherId = teacherIndex.byName[teacherName].teacherId;
    }

    if (!classId || !period || !weekday) return;

    const key = [classId, weekday, period].join('__');

    // 現行 getTeacherUnsavedContext_ と同じく、同一キーの timetable 行は後勝ちにする。
    map[key] = {
      teacherIds: teacherId ? [teacherId] : []
    };

    if (teacherName && !teacherId) {
      addTeacherUnsavedWarning_(
        warnings,
        'timetable row ' + (index + 2) + ': 担当者名から teacherId を解決できません: ' + teacherName
      );
    }
  });

  teamData.rows.forEach(function(row, index) {
    const classId = teamCol.classId !== -1 ? normalizeString_(row[teamCol.classId]) : '';
    const period = teamCol.period !== -1 ? normalizeString_(row[teamCol.period]) : '';
    const weekday = teamCol.weekday !== -1 ? normalizeWeekday_(row[teamCol.weekday]) : '';
    const teacherName = teamCol.teacherName !== -1
      ? normalizeString_(row[teamCol.teacherName])
      : '';
    let teacherId = teamCol.teacherId !== -1
      ? normalizeString_(row[teamCol.teacherId])
      : '';

    if (!teacherId && teacherName && teacherIndex.byName[teacherName]) {
      teacherId = teacherIndex.byName[teacherName].teacherId;
    }

    if (!classId || !period || !weekday || !teacherId) {
      if (teacherName && !teacherId) {
        addTeacherUnsavedWarning_(
          warnings,
          'classTeacherTeams row ' + (index + 2) + ': 担当者名から teacherId を解決できません: ' + teacherName
        );
      }
      return;
    }

    const key = [classId, weekday, period].join('__');
    if (!map[key]) {
      map[key] = { teacherIds: [] };
    }

    if (map[key].teacherIds.indexOf(teacherId) === -1) {
      map[key].teacherIds.push(teacherId);
    }
  });

  return map;
}

function buildTeacherUnsavedSavedKeySet_(data, startYmd, endYmd) {
  const col = resolveTeacherUnsavedRequiredColumns_(
    'attendanceSessions',
    data.headers,
    {
      classId: ['classId', 'ClassID'],
      date: ['date', '日付'],
      period: ['period', '時限']
    },
    ['classId', 'date', 'period']
  );
  const keySet = {};

  if (!startYmd || !endYmd || endYmd < startYmd) return keySet;

  data.rows.forEach(function(row, index) {
    const ymd = normalizeTeacherUnsavedSourceYmd_(
      row[col.date],
      data.dateDisplayValues[index]
    );
    if (!ymd || ymd < startYmd || ymd > endYmd) return;

    const classId = normalizeString_(row[col.classId]);
    const period = normalizeString_(row[col.period]);
    if (!classId || !period) return;

    const saveKey = [classId, ymd, period].join('__');
    keySet[saveKey] = true;

    const experimentKey = buildExperimentSessionKeyForTimetable_(classId, ymd, period);
    if (experimentKey) {
      keySet[experimentKey] = true;
    }
  });

  return keySet;
}

function publishTeacherUnsavedCacheSnapshot_(snapshot) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    throw new Error('キャッシュ公開用ロックを取得できませんでした。');
  }

  try {
    const validation = validateTeacherUnsavedCacheSheets_();
    if (!validation.ok) {
      throw new Error(buildTeacherUnsavedCacheValidationError_(validation));
    }

    markTeacherUnsavedSummaryRowsStatus_(
      validation.summary.sheet,
      'building',
      ''
    );
    SpreadsheetApp.flush();

    replaceTeacherUnsavedCacheRows_(
      validation.detail.sheet,
      snapshot.detailRows,
      TEACHER_UNSAVED_DETAIL_CACHE_HEADERS_.length
    );
    SpreadsheetApp.flush();

    replaceTeacherUnsavedCacheRows_(
      validation.summary.sheet,
      snapshot.summaryRows,
      TEACHER_UNSAVED_SUMMARY_CACHE_HEADERS_.length
    );
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function replaceTeacherUnsavedCacheRows_(sheet, rows, columnCount) {
  const oldDataRowCount = Math.max(sheet.getLastRow() - 1, 0);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, columnCount).setValues(rows);
  }

  if (oldDataRowCount > rows.length) {
    sheet
      .getRange(2 + rows.length, 1, oldDataRowCount - rows.length, columnCount)
      .clearContent();
  }
}

function markTeacherUnsavedSummaryRowsStatus_(sheet, status, errorMessage) {
  const rowCount = Math.max(sheet.getLastRow() - 1, 0);
  if (rowCount === 0) return;

  const values = [];
  for (let i = 0; i < rowCount; i++) {
    values.push([status, errorMessage || '']);
  }

  const statusColumn = TEACHER_UNSAVED_SUMMARY_CACHE_HEADERS_.indexOf('status') + 1;
  sheet.getRange(2, statusColumn, rowCount, 2).setValues(values);
}

function markTeacherUnsavedCachePublishError_(error) {
  const validation = validateTeacherUnsavedCacheSheets_();
  if (!validation.ok) return;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return;

  try {
    const message = String(error && error.message ? error.message : error || '')
      .slice(0, 500);
    markTeacherUnsavedSummaryRowsStatus_(validation.summary.sheet, 'error', message);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function validateTeacherUnsavedCacheSheets_() {
  const operation = getOperationSpreadsheet();
  const summary = validateTeacherUnsavedCacheSheet_(
    operation,
    TEACHER_UNSAVED_CACHE_SHEETS_.SUMMARY,
    TEACHER_UNSAVED_SUMMARY_CACHE_HEADERS_
  );
  const detail = validateTeacherUnsavedCacheSheet_(
    operation,
    TEACHER_UNSAVED_CACHE_SHEETS_.DETAIL,
    TEACHER_UNSAVED_DETAIL_CACHE_HEADERS_
  );
  const errors = summary.errors.concat(detail.errors);

  return {
    ok: errors.length === 0,
    summary: summary,
    detail: detail,
    errors: errors
  };
}

function validateTeacherUnsavedCacheSheet_(spreadsheet, sheetName, expectedHeaders) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  const errors = [];
  let actualHeaders = [];

  if (!sheet) {
    errors.push(sheetName + ': シートが存在しません。');
  } else {
    const lastColumn = sheet.getLastColumn();
    if (lastColumn < 1) {
      errors.push(sheetName + ': ヘッダー行がありません。');
    } else {
      actualHeaders = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];

      if (actualHeaders.length !== expectedHeaders.length) {
        errors.push(
          sheetName + ': ヘッダー列数が一致しません。expected=' +
          expectedHeaders.length + ', actual=' + actualHeaders.length
        );
      }

      const maxLength = Math.max(actualHeaders.length, expectedHeaders.length);
      for (let i = 0; i < maxLength; i++) {
        const expected = expectedHeaders[i] == null ? '(なし)' : expectedHeaders[i];
        const actual = actualHeaders[i] == null ? '(なし)' : actualHeaders[i];
        if (actual !== expected) {
          errors.push(
            sheetName + ': ' + (i + 1) + '列目が一致しません。expected=' +
            expected + ', actual=' + actual
          );
        }
      }
    }
  }

  return {
    sheet: sheet,
    errors: errors,
    publicResult: {
      ok: errors.length === 0,
      sheetName: sheetName,
      exists: !!sheet,
      expectedHeaders: expectedHeaders.slice(),
      actualHeaders: actualHeaders,
      errors: errors.slice()
    }
  };
}

function buildTeacherUnsavedCacheValidationError_(validation) {
  return 'キャッシュシート検証エラー: ' + validation.errors.join(' / ');
}

function readTeacherUnsavedSummaryState_(sheet, teacherId, todayYmd) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return { row: null };

  const values = sheet
    .getRange(2, 1, lastRow - 1, TEACHER_UNSAVED_SUMMARY_CACHE_HEADERS_.length)
    .getValues();
  let current = null;
  let latest = null;

  values.forEach(function(valuesRow) {
    const row = buildTeacherUnsavedSummaryObject_(valuesRow);
    if (row.teacherId !== teacherId) return;

    if (!latest || getTeacherUnsavedComparableTime_(row.checkedAt) >= getTeacherUnsavedComparableTime_(latest.checkedAt)) {
      latest = row;
    }

    if (
      row.cacheDate === todayYmd &&
      (!current || getTeacherUnsavedComparableTime_(row.checkedAt) >= getTeacherUnsavedComparableTime_(current.checkedAt))
    ) {
      current = row;
    }
  });

  return { row: current || latest };
}

function buildTeacherUnsavedSummaryObject_(row) {
  return {
    snapshotId: normalizeString_(row[0]),
    cacheDate: normalizeTeacherUnsavedSourceYmd_(row[1], row[1]),
    teacherId: normalizeString_(row[2]),
    teacherName: normalizeString_(row[3]),
    teacherEmail: normalizeString_(row[4]).toLowerCase(),
    startYmd: normalizeTeacherUnsavedSourceYmd_(row[5], row[5]),
    endYmd: normalizeTeacherUnsavedSourceYmd_(row[6], row[6]),
    unsavedCount: parseTeacherUnsavedNonNegativeInteger_(row[7]),
    firstDate: normalizeTeacherUnsavedSourceYmd_(row[8], row[8]),
    lastDate: normalizeTeacherUnsavedSourceYmd_(row[9], row[9]),
    detailStartRow: parseTeacherUnsavedOptionalPositiveInteger_(row[10]),
    detailCount: parseTeacherUnsavedNonNegativeInteger_(row[11]),
    checkedAt: row[12],
    status: normalizeString_(row[13]).toLowerCase(),
    errorMessage: normalizeString_(row[14])
  };
}

function buildTeacherUnsavedDetailObject_(row) {
  return {
    snapshotId: normalizeString_(row[0]),
    cacheDate: normalizeTeacherUnsavedSourceYmd_(row[1], row[1]),
    teacherId: normalizeString_(row[2]),
    teacherName: normalizeString_(row[3]),
    teacherEmail: normalizeString_(row[4]).toLowerCase(),
    date: normalizeTeacherUnsavedSourceYmd_(row[5], row[5]),
    period: normalizeString_(row[6]),
    classId: normalizeString_(row[7]),
    sessionNumber: row[8] == null ? '' : row[8],
    subjectName: normalizeString_(row[9]),
    targetLabel: normalizeString_(row[10]),
    isExperiment: row[11] === true || normalizeString_(row[11]).toLowerCase() === 'true',
    displayKey: normalizeString_(row[12]),
    saveKey: normalizeString_(row[13]),
    sortKey: normalizeString_(row[14]),
    checkedAt: formatTeacherUnsavedCacheDateTime_(row[15])
  };
}

function validateTeacherUnsavedSummaryRowIntegrity_(row, detailSheet) {
  if (!row.snapshotId) {
    return { ok: false, errorMessage: 'サマリーキャッシュの snapshotId が空です。' };
  }
  if (!row.startYmd || !row.endYmd) {
    return { ok: false, errorMessage: 'サマリーキャッシュの対象期間が不正です。' };
  }
  if (getTeacherUnsavedComparableTime_(row.checkedAt) === 0) {
    return { ok: false, errorMessage: 'サマリーキャッシュの checkedAt が不正です。' };
  }
  if (row.unsavedCount === null || row.detailCount === null) {
    return { ok: false, errorMessage: 'サマリーキャッシュの件数が不正です。' };
  }
  if (row.unsavedCount !== row.detailCount) {
    return { ok: false, errorMessage: 'unsavedCount と detailCount が一致しません。' };
  }

  if (row.detailCount === 0) {
    return { ok: true, errorMessage: '' };
  }

  if (row.detailStartRow === null || row.detailStartRow < 2) {
    return { ok: false, errorMessage: 'detailStartRow が不正です。' };
  }
  if (!row.firstDate || !row.lastDate) {
    return { ok: false, errorMessage: 'firstDate または lastDate が不正です。' };
  }

  const endRow = row.detailStartRow + row.detailCount - 1;
  if (endRow > detailSheet.getLastRow()) {
    return { ok: false, errorMessage: '詳細キャッシュの参照範囲がシート末尾を超えています。' };
  }

  const firstIdentity = detailSheet.getRange(row.detailStartRow, 1, 1, 6).getValues()[0];
  const lastIdentity = row.detailCount === 1
    ? firstIdentity
    : detailSheet.getRange(endRow, 1, 1, 6).getValues()[0];

  const identities = [firstIdentity, lastIdentity];
  for (let i = 0; i < identities.length; i++) {
    const identity = identities[i];
    if (
      normalizeString_(identity[0]) !== row.snapshotId ||
      normalizeTeacherUnsavedSourceYmd_(identity[1], identity[1]) !== row.cacheDate ||
      normalizeString_(identity[2]) !== row.teacherId
    ) {
      return {
        ok: false,
        errorMessage: 'サマリーと詳細キャッシュの snapshotId、cacheDate、または teacherId が一致しません。'
      };
    }
  }

  if (
    normalizeTeacherUnsavedSourceYmd_(firstIdentity[5], firstIdentity[5]) !== row.firstDate ||
    normalizeTeacherUnsavedSourceYmd_(lastIdentity[5], lastIdentity[5]) !== row.lastDate
  ) {
    return {
      ok: false,
      errorMessage: 'サマリーと詳細キャッシュの firstDate または lastDate が一致しません。'
    };
  }

  return { ok: true, errorMessage: '' };
}

function validateTeacherUnsavedDetailPagingOptions_(options) {
  const source = options == null ? {} : options;
  if (typeof source !== 'object' || Array.isArray(source)) {
    throw new Error('options はオブジェクトで指定してください。');
  }

  const limit = source.limit == null || source.limit === '' ? 50 : Number(source.limit);
  const offset = source.offset == null || source.offset === '' ? 0 : Number(source.offset);

  if (!Number.isInteger(limit) || limit <= 0) {
    throw new Error('limit は1以上の整数で指定してください。');
  }
  if (!Number.isInteger(offset) || offset < 0) {
    throw new Error('offset は0以上の整数で指定してください。');
  }

  return { limit: limit, offset: offset };
}

function buildTeacherUnsavedFastSummaryFailure_(status, teacherId, cacheDate, message) {
  return {
    ok: false,
    status: status,
    cacheStatus: '',
    teacherId: teacherId || '',
    teacherName: '',
    teacherEmail: '',
    count: null,
    unsavedCount: null,
    firstDate: '',
    lastDate: '',
    detailCount: null,
    checkedAt: '',
    startYmd: '',
    endYmd: '',
    checkedRange: null,
    cacheDate: cacheDate || '',
    snapshotId: '',
    errorMessage: message || ''
  };
}

function buildTeacherUnsavedFastSummaryFailureFromRow_(status, row, message, cacheStatus) {
  const result = buildTeacherUnsavedFastSummaryFailure_(
    status,
    row && row.teacherId ? row.teacherId : '',
    row && row.cacheDate ? row.cacheDate : '',
    message
  );

  result.cacheStatus = cacheStatus != null
    ? cacheStatus
    : (row && row.status ? row.status : '');
  result.teacherName = row && row.teacherName ? row.teacherName : '';
  result.teacherEmail = row && row.teacherEmail ? row.teacherEmail : '';
  result.checkedAt = row && row.checkedAt
    ? formatTeacherUnsavedCacheDateTime_(row.checkedAt)
    : '';
  result.startYmd = row && row.startYmd ? row.startYmd : '';
  result.endYmd = row && row.endYmd ? row.endYmd : '';
  result.checkedRange = result.startYmd || result.endYmd
    ? { start: result.startYmd, end: result.endYmd }
    : null;
  result.snapshotId = row && row.snapshotId ? row.snapshotId : '';
  return result;
}

function buildTeacherUnsavedFastDetailsFailure_(summary, paging) {
  return {
    ok: false,
    status: summary.status || 'unavailable',
    cacheStatus: summary.cacheStatus || '',
    items: [],
    totalCount: null,
    limit: paging.limit,
    offset: paging.offset,
    hasMore: false,
    checkedAt: summary.checkedAt || '',
    startYmd: summary.startYmd || '',
    endYmd: summary.endYmd || '',
    checkedRange: summary.checkedRange || null,
    cacheDate: summary.cacheDate || '',
    snapshotId: summary.snapshotId || '',
    errorMessage: summary.errorMessage || ''
  };
}

function buildTeacherUnsavedFastDetailsSuccess_(summary, paging, items, totalCount) {
  return {
    ok: true,
    status: 'ready',
    cacheStatus: 'ready',
    items: items,
    totalCount: totalCount,
    limit: paging.limit,
    offset: paging.offset,
    hasMore: paging.offset + items.length < totalCount,
    checkedAt: summary.checkedAt,
    startYmd: summary.startYmd,
    endYmd: summary.endYmd,
    checkedRange: {
      start: summary.startYmd,
      end: summary.endYmd
    },
    cacheDate: summary.cacheDate,
    snapshotId: summary.snapshotId,
    errorMessage: ''
  };
}

function getTeacherUnsavedCacheDateContext_(baseDate) {
  const cacheDate = Utilities.formatDate(baseDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  const today = new Date(cacheDate + 'T12:00:00+09:00');
  const endDate = new Date(today);
  endDate.setDate(endDate.getDate() - 1);

  return {
    cacheDate: cacheDate,
    startYmd: formatDateToYmd(getTeacherUnsavedStartDate_(today)),
    endYmd: formatDateToYmd(endDate)
  };
}

function buildTeacherUnsavedSnapshotId_(checkedAt) {
  return [
    Utilities.formatDate(checkedAt, 'Asia/Tokyo', 'yyyyMMdd_HHmmss_SSS'),
    Utilities.getUuid()
  ].join('__');
}

function normalizeTeacherUnsavedSourceYmd_(rawValue, displayValue) {
  const displayYmd = normalizeYmdDisplayText_(displayValue);
  if (displayYmd) return displayYmd;

  const rawTextYmd = normalizeYmdDisplayText_(rawValue);
  if (rawTextYmd) return rawTextYmd;

  if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
    return Utilities.formatDate(rawValue, 'Asia/Tokyo', 'yyyy-MM-dd');
  }

  return '';
}

function resolveTeacherUnsavedRequiredColumns_(sheetName, headers, candidates, requiredKeys) {
  const result = {};
  Object.keys(candidates).forEach(function(key) {
    result[key] = findColumnIndex_(headers, candidates[key]);
  });

  requiredKeys.forEach(function(key) {
    if (result[key] === -1) {
      throw new Error(sheetName + ' シートに必要な列がありません: ' + key);
    }
  });

  return result;
}

function buildTeacherUnsavedCacheTargetLabel_(cls, isExperiment) {
  if (isExperiment) return '班選択して記録';

  const gradeText = cls && cls.grade ? String(cls.grade) + '年' : '';
  const unitText = cls && cls.unit ? String(cls.unit) + '組' : '';
  return (gradeText + ' ' + unitText).trim();
}

function buildTeacherUnsavedDetailSortKey_(ymd, period, displayKey) {
  const numericPeriod = Number(period);
  const periodKey = isNaN(numericPeriod)
    ? 'Z_' + normalizeString_(period)
    : 'N_' + String(numericPeriod).padStart(4, '0');
  return [ymd, periodKey, displayKey].join('__');
}

function compareTeacherUnsavedDetailItems_(a, b) {
  if (a.date !== b.date) return a.date < b.date ? -1 : 1;

  const periodDiff = Number(a.period) - Number(b.period);
  if (!isNaN(periodDiff) && periodDiff !== 0) return periodDiff;
  return 0;
}

function buildTeacherUnsavedDetailRow_(item) {
  return [
    item.snapshotId,
    item.cacheDate,
    item.teacherId,
    item.teacherName,
    item.teacherEmail,
    item.date,
    item.period,
    item.classId,
    item.sessionNumber,
    item.subjectName,
    item.targetLabel,
    item.isExperiment,
    item.displayKey,
    item.saveKey,
    item.sortKey,
    item.checkedAt
  ];
}

function buildTeacherUnsavedCacheBuildResult_(snapshot, elapsedMs, wroteSheets) {
  return {
    ok: true,
    wroteSheets: wroteSheets,
    snapshotId: snapshot.snapshotId,
    cacheDate: snapshot.cacheDate,
    startYmd: snapshot.startYmd,
    endYmd: snapshot.endYmd,
    checkedAt: formatTeacherUnsavedCacheDateTime_(snapshot.checkedAt),
    teacherCount: snapshot.teacherCount,
    unsavedTeacherCount: snapshot.unsavedTeacherCount,
    summaryRowCount: snapshot.summaryRows.length,
    detailRowCount: snapshot.detailRows.length,
    sourceRowCounts: snapshot.sourceRowCounts,
    warningCount: snapshot.warnings.length,
    warnings: snapshot.warnings.slice(0, 50),
    elapsedMs: elapsedMs
  };
}

function formatTeacherUnsavedCacheDateTime_(value) {
  if (!value) return '';

  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$/.test(text)) {
    return text;
  }

  const date = value instanceof Date ? value : new Date(value);
  if (isNaN(date.getTime())) return text;
  return Utilities.formatDate(date, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
}

function getTeacherUnsavedComparableTime_(value) {
  if (!value) return 0;
  const date = value instanceof Date ? value : new Date(value);
  return isNaN(date.getTime()) ? 0 : date.getTime();
}

function parseTeacherUnsavedNonNegativeInteger_(value) {
  if (value == null || String(value).trim() === '') return null;
  const number = Number(value);
  return Number.isInteger(number) && number >= 0 ? number : null;
}

function parseTeacherUnsavedOptionalPositiveInteger_(value) {
  if (value == null || String(value).trim() === '') return null;
  const number = Number(value);
  return Number.isInteger(number) && number >= 1 ? number : null;
}

function addTeacherUnsavedWarning_(warnings, message) {
  if (warnings.indexOf(message) === -1) {
    warnings.push(message);
  }
}

function buildTeacherUnsavedLegacyDebugKeySets_(items, classMap) {
  const displayKeys = {};
  const saveKeys = {};

  items.forEach(function(item) {
    const classId = normalizeString_(item.classId);
    const ymd = normalizeTeacherUnsavedSourceYmd_(item.date, item.date);
    const period = normalizeString_(item.period);
    if (!classId || !ymd || !period) return;

    const saveKey = [classId, ymd, period].join('__');
    const displayKey = buildTeacherUnsavedDisplayKey_(
      classMap[classId] || {},
      classId,
      ymd,
      period
    );
    saveKeys[saveKey] = true;
    displayKeys[displayKey] = true;
  });

  return { displayKeys: displayKeys, saveKeys: saveKeys };
}

function buildTeacherUnsavedFastDebugKeySets_(items) {
  const displayKeys = {};
  const saveKeys = {};

  items.forEach(function(item) {
    if (item.displayKey) displayKeys[item.displayKey] = true;
    if (item.saveKey) saveKeys[item.saveKey] = true;
  });

  return { displayKeys: displayKeys, saveKeys: saveKeys };
}

function subtractTeacherUnsavedKeySets_(left, right) {
  return Object.keys(left)
    .filter(function(key) { return !right[key]; })
    .sort();
}
