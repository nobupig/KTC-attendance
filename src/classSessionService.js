function generateClassSessions() {
  const ss = getOperationSpreadsheet();

  const timetableSheet = ss.getSheetByName(CONFIG.SHEETS.TIMETABLE);
  const calendarSheet = ss.getSheetByName(CONFIG.SHEETS.CALENDAR);
  const classSessionsSheet = ss.getSheetByName(CONFIG.SHEETS.CLASS_SESSIONS);

  if (!timetableSheet) {
    throw new Error('timetable シートが見つかりません。');
  }
  if (!calendarSheet) {
    throw new Error('calendar シートが見つかりません。');
  }
  if (!classSessionsSheet) {
    throw new Error('classSessions シートが見つかりません。');
  }

  const timetableValues = timetableSheet.getDataRange().getValues();
  const calendarValues = calendarSheet.getDataRange().getValues();

  if (timetableValues.length < 2) {
    throw new Error('timetable にデータがありません。');
  }
  if (calendarValues.length < 2) {
    throw new Error('calendar にデータがありません。');
  }

  const timetableHeaders = timetableValues[0];
  const calendarHeaders = calendarValues[0];

  const ttCol = {
    classId: findColumnIndex_(timetableHeaders, ['classId', 'ClassID']),
    weekday: findColumnIndex_(timetableHeaders, ['weekday', '曜日']),
    period: findColumnIndex_(timetableHeaders, ['period', '時限']),
    teacherName: findColumnIndex_(timetableHeaders, ['teacherName', '担当者名', 'name']),
    teacherId: findColumnIndex_(timetableHeaders, ['teacherId', 'TeacherID'])
  };

  const calCol = {
    date: findColumnIndex_(calendarHeaders, ['date', '日付']),
    weekday: findColumnIndex_(calendarHeaders, ['weekday', '曜日']),
    isClassDay: findColumnIndex_(calendarHeaders, ['isClassDay', '授業日'])
  };

  validateRequiredColumnsForClassSessions_('timetable', ttCol, ['classId', 'weekday', 'period']);
  validateRequiredColumnsForClassSessions_('calendar', calCol, ['date', 'weekday', 'isClassDay']);

  const timetableData = timetableValues.slice(1)
    .map(function(row) {
      const classId = normalizeString_(row[ttCol.classId]);
      const weekday = normalizeWeekday_(row[ttCol.weekday]);
      const period = normalizeString_(row[ttCol.period]);
      const teacherName = ttCol.teacherName !== -1 ? normalizeString_(row[ttCol.teacherName]) : '';
      const teacherId = ttCol.teacherId !== -1 ? normalizeString_(row[ttCol.teacherId]) : '';

      return {
        classId: classId,
        weekday: weekday,
        period: period,
        teacherName: teacherName,
        teacherId: teacherId
      };
    })
    .filter(function(item) {
      return item.classId && item.weekday && item.period;
    });

  const calendarData = calendarValues.slice(1)
    .map(function(row) {
      return {
        date: row[calCol.date],
        weekday: normalizeWeekday_(row[calCol.weekday]),
        isClassDay: row[calCol.isClassDay] === true
      };
    })
    .filter(function(item) {
      return item.date && item.weekday && item.isClassDay;
    });

  const sessionCountMap = {};
  const output = [];

  calendarData.forEach(function(cal) {
    timetableData.forEach(function(tt) {
      if (cal.weekday !== tt.weekday) {
        return;
      }

      if (!sessionCountMap[tt.classId]) {
        sessionCountMap[tt.classId] = 0;
      }

      sessionCountMap[tt.classId] += 1;

      output.push([
        tt.classId,
        formatDateToYmd(cal.date),
        Number(tt.period),
        sessionCountMap[tt.classId]
      ]);
    });
  });

  output.sort(function(a, b) {
    const dateA = normalizeString_(a[1]);
    const dateB = normalizeString_(b[1]);
    if (dateA !== dateB) {
      return dateA.localeCompare(dateB, 'ja');
    }

    const periodA = Number(a[2]);
    const periodB = Number(b[2]);
    if (periodA !== periodB) {
      return periodA - periodB;
    }

    return normalizeString_(a[0]).localeCompare(normalizeString_(b[0]), 'ja');
  });

  classSessionsSheet.clearContents();
  classSessionsSheet.getRange(1, 1, 1, 4).setValues([[
    'classId',
    'date',
    'period',
    'sessionNumber'
  ]]);

  if (output.length > 0) {
    classSessionsSheet.getRange(2, 1, output.length, 4).setValues(output);
  }
}

const SECOND_TERM_CLASS_SESSIONS_2026_ = Object.freeze({
  START_DATE: '2026-09-24',
  END_DATE: '2027-02-02',
  SAMPLE_SIZE: 10
});

function previewSecondTermClassSessions2026() {
  const sources = readClassSessionPlanningSources_();
  const plan = buildRangeLimitedClassSessionPlan_(
    SECOND_TERM_CLASS_SESSIONS_2026_.START_DATE,
    SECOND_TERM_CLASS_SESSIONS_2026_.END_DATE,
    sources
  );
  const preview = buildClassSessionPlanPreview_(plan);

  Logger.log(JSON.stringify(preview, null, 2));
  return preview;
}

function appendSecondTermClassSessions2026(expectedPlanHash) {
  const expectedHash = normalizeString_(expectedPlanHash);
  if (!expectedHash) {
    throw new Error('previewで取得したexpectedPlanHashが必要です。');
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const sources = readClassSessionPlanningSources_();
    const plan = buildRangeLimitedClassSessionPlan_(
      SECOND_TERM_CLASS_SESSIONS_2026_.START_DATE,
      SECOND_TERM_CLASS_SESSIONS_2026_.END_DATE,
      sources
    );

    if (plan.planHash !== expectedHash) {
      throw new Error(
        'preview後に元データが変更されました。再previewしてください。' +
        ' expected=' + expectedHash + ' actual=' + plan.planHash
      );
    }
    if (plan.blockingErrors.length > 0) {
      throw new Error('classSessions追加を中止しました: ' + plan.blockingErrors.join(' / '));
    }

    if (plan.insertRows.length === 0) {
      return {
        ok: true,
        wroteRows: 0,
        startDate: plan.startDate,
        endDate: plan.endDate,
        planHash: plan.planHash,
        cacheKeysInvalidated: [],
        fastUnsavedCacheStaled: false,
        message: '追加対象はありません。'
      };
    }

    const sheet = sources.classSessionsSheet;
    const beforeLastRow = sheet.getLastRow();
    const insertedDates = uniqueSortedStrings_(plan.insertRows.map(function(row) {
      return row[1];
    }));
    try {
      sheet.getRange(beforeLastRow + 1, 1, plan.insertRows.length, 4)
        .setValues(plan.insertRows);
      SpreadsheetApp.flush();

      const refreshedSources = readClassSessionPlanningSources_();
      verifyAppendedClassSessionRows_(sources, refreshedSources, plan, beforeLastRow);
    } catch (postWriteError) {
      try {
        invalidateCachesAfterClassSessionAppend_(insertedDates);
      } catch (invalidationError) {
        postWriteError.cacheInvalidationError = String(
          invalidationError && invalidationError.message
            ? invalidationError.message
            : invalidationError
        );
      }
      throw postWriteError;
    }
    const invalidation = invalidateCachesAfterClassSessionAppend_(insertedDates);

    return {
      ok: true,
      wroteRows: plan.insertRows.length,
      startDate: plan.startDate,
      endDate: plan.endDate,
      planHash: plan.planHash,
      insertedDates: insertedDates,
      cacheKeysInvalidated: invalidation.cacheKeys,
      fastUnsavedCacheStaled: invalidation.fastUnsavedCacheStaled,
      fastUnsavedCacheResult: invalidation.fastUnsavedCacheResult
    };
  } finally {
    lock.releaseLock();
  }
}

function invalidateCachesAfterClassSessionAppend_(insertedDates) {
  const cacheKeys = [
    getClassSessionsSheetCacheKey_(),
    getClassSessionsByDateIndexCacheKey_()
  ].concat(insertedDates.map(function(ymd) {
    return buildClassSessionsByDateCacheKey_(ymd);
  }));
  removeScriptCacheKeys_(cacheKeys);

  const yesterdayYmd = getTeacherUnsavedCacheDateContext_(new Date()).endYmd;
  const shouldStaleFastCache = insertedDates.some(function(ymd) {
    return ymd <= yesterdayYmd;
  });
  let fastCacheResult = null;
  if (shouldStaleFastCache) {
    const classSessionsRevision = advanceTeacherUnsavedClassSessionsRevisionUnderLock_();
    fastCacheResult = invalidateAllTeacherUnsavedFastSnapshotsUnderLock_(
      'classSessions追加後にFastキャッシュを無効化しました。次回rebuildを待っています。'
    );
    fastCacheResult.classSessionsRevision = classSessionsRevision;
  }

  return {
    cacheKeys: cacheKeys,
    fastUnsavedCacheStaled: shouldStaleFastCache,
    fastUnsavedCacheResult: fastCacheResult
  };
}

function buildRangeLimitedClassSessionPlan_(startDate, endDate, sources) {
  const startYmd = normalizeClassSessionPlanningYmd_(startDate, '');
  const endYmd = normalizeClassSessionPlanningYmd_(endDate, '');
  if (!startYmd || !endYmd || startYmd > endYmd) {
    throw new Error('classSessions生成範囲が不正です: ' + startDate + ' - ' + endDate);
  }

  const timetable = sources.timetable;
  const calendar = sources.calendar;
  const existing = sources.classSessions;
  const blockingErrors = [];
  const warnings = [];
  const validWeekdays = { Mon: true, Tue: true, Wed: true, Thu: true, Fri: true, Sat: true, Sun: true };

  const ttCol = resolveClassSessionPlanningColumns_(
    'timetable',
    timetable.headers,
    {
      classId: ['classId', 'ClassID'],
      weekday: ['weekday', '曜日'],
      period: ['period', '時限'],
      teacherName: ['teacherName', '担当者名', 'name'],
      teacherId: ['teacherId', 'TeacherID'],
      term: ['term', '学期']
    },
    ['classId', 'weekday', 'period']
  );
  const calCol = resolveClassSessionPlanningColumns_(
    'calendar',
    calendar.headers,
    {
      date: ['date', '日付'],
      weekday: ['weekday', '曜日'],
      isClassDay: ['isClassDay', '授業日'],
      term: ['term', '学期']
    },
    ['date', 'weekday', 'isClassDay']
  );
  const csCol = resolveClassSessionPlanningColumns_(
    'classSessions',
    existing.headers,
    {
      classId: ['classId', 'ClassID'],
      date: ['date', '日付'],
      period: ['period', '時限'],
      sessionNumber: ['sessionNumber', '回', '回数']
    },
    ['classId', 'date', 'period', 'sessionNumber']
  );
  if (
    csCol.classId !== 0 ||
    csCol.date !== 1 ||
    csCol.period !== 2 ||
    csCol.sessionNumber !== 3
  ) {
    blockingErrors.push(
      'classSessionsの先頭4列はclassId,date,period,sessionNumberの順である必要があります。'
    );
  }

  const hasTimetableTermColumn = ttCol.term !== -1;
  let invalidTimetableCount = 0;
  let invalidTimetableTermCount = 0;
  const timetableRowsByTerm = {};
  const timetableByKey = {};
  const timetableDuplicateKeys = {};
  timetable.rows.forEach(function(row, index) {
    const classId = normalizeString_(row[ttCol.classId]);
    const weekday = normalizeWeekday_(row[ttCol.weekday]);
    const period = normalizeClassSessionPeriod_(row[ttCol.period]);
    const rawTerm = hasTimetableTermColumn ? normalizeString_(row[ttCol.term]) : '';
    const term = normalizeAcademicTerm_(rawTerm);
    if (!classId || !validWeekdays[weekday] || !period) {
      invalidTimetableCount += 1;
      return;
    }

    // A missing term column is legacy-compatible for the existing SP+FY DEV
    // timetable. Once the column exists, every schedule row must declare it.
    if (hasTimetableTermColumn && !term) {
      invalidTimetableTermCount += 1;
      return;
    }

    const key = JSON.stringify([classId, term, weekday, period]);
    const item = {
      classId: classId,
      term: term,
      usesLegacyTermFallback: !hasTimetableTermColumn,
      weekday: weekday,
      period: period,
      sourceRow: index + 2
    };
    const termBucket = hasTimetableTermColumn ? term : 'legacy-no-column';
    timetableRowsByTerm[termBucket] = (timetableRowsByTerm[termBucket] || 0) + 1;
    if (timetableByKey[key]) {
      timetableDuplicateKeys[key] = (timetableDuplicateKeys[key] || 1) + 1;
    } else {
      timetableByKey[key] = item;
    }
  });
  const timetableDuplicateCount = Object.keys(timetableDuplicateKeys).reduce(function(total, key) {
    return total + timetableDuplicateKeys[key] - 1;
  }, 0);
  if (invalidTimetableCount > 0) {
    blockingErrors.push('timetableに不正行があります: ' + invalidTimetableCount + '件');
  }
  if (invalidTimetableTermCount > 0) {
    blockingErrors.push('timetableに空欄または不正なtermがあります: ' + invalidTimetableTermCount + '件');
  }
  if (timetableDuplicateCount > 0) {
    blockingErrors.push(
      'timetableにclassId + term + weekday + period重複があります: ' +
      timetableDuplicateCount + '件 keys=' + Object.keys(timetableDuplicateKeys).sort().join(',')
    );
  }

  let calendarRowsInRange = 0;
  let invalidCalendarCount = 0;
  let invalidCalendarTermCount = 0;
  let termFallbackCount = 0;
  const calendarByDate = {};
  const calendarDuplicateDates = {};
  const effectiveCalendarIndex = buildEffectiveClassDayIndex_(calendar);
  calendar.rows.forEach(function(row, index) {
    const ymd = normalizeClassSessionPlanningYmd_(
      row[calCol.date],
      calendar.dateDisplayValues[index]
    );
    if (!ymd) {
      invalidCalendarCount += 1;
      return;
    }
    if (ymd < startYmd || ymd > endYmd) return;

    calendarRowsInRange += 1;
    const isClassDay = row[calCol.isClassDay] === true;
    const weekday = normalizeWeekday_(row[calCol.weekday]);
    const rawTerm = calCol.term === -1 ? '' : normalizeString_(row[calCol.term]);
    const normalizedTerm = normalizeAcademicTerm_(rawTerm);
    if (isClassDay && !validWeekdays[weekday]) {
      invalidCalendarCount += 1;
    }
    if (isClassDay && rawTerm && normalizedTerm !== 'FA' && normalizedTerm !== 'SP') {
      invalidCalendarTermCount += 1;
    }
    if (calendarByDate[ymd]) {
      calendarDuplicateDates[ymd] = (calendarDuplicateDates[ymd] || 1) + 1;
    } else {
      calendarByDate[ymd] = {
        date: ymd,
        weekday: weekday,
        isClassDay: isClassDay,
        term: rawTerm,
        sourceRow: index + 2
      };
    }
  });
  const calendarDuplicateCount = Object.keys(calendarDuplicateDates).reduce(function(total, ymd) {
    return total + calendarDuplicateDates[ymd] - 1;
  }, 0);
  invalidCalendarCount += calendarDuplicateCount;
  if (invalidCalendarCount > 0) {
    blockingErrors.push('calendarに不正行または日付重複があります: ' + invalidCalendarCount + '件');
  }
  if (invalidCalendarTermCount > 0) {
    blockingErrors.push('calendarに不正なtermがあります: ' + invalidCalendarTermCount + '件');
  }

  const classDays = Object.keys(calendarByDate).filter(function(ymd) {
    const item = calendarByDate[ymd];
    const context = getEffectiveClassDayContext_(ymd, effectiveCalendarIndex);
    item.context = context;
    if (context.usedTermFallback) termFallbackCount += 1;
    return context.isClassDay && !!context.effectiveWeekday && !!context.term;
  }).sort();

  let invalidSessionNumberCount = 0;
  const existingIdentityCounts = {};
  const existingIdentityMap = {};
  const existingByClass = {};
  const existingHashRows = [];
  existing.rows.forEach(function(row, index) {
    const classId = normalizeString_(row[csCol.classId]);
    const ymd = normalizeClassSessionPlanningYmd_(
      row[csCol.date],
      existing.dateDisplayValues[index]
    );
    const period = normalizeClassSessionPeriod_(row[csCol.period]);
    const sessionNumber = normalizeClassSessionNumber_(row[csCol.sessionNumber]);
    const rowNumber = index + 2;

    existingHashRows.push([classId, ymd, period, sessionNumber, rowNumber]);
    if (!classId || !ymd || !period) {
      blockingErrors.push('classSessionsのidentityが不正です: row=' + rowNumber);
      return;
    }

    const identity = buildClassSessionIdentity_(classId, ymd, period);
    existingIdentityCounts[identity] = (existingIdentityCounts[identity] || 0) + 1;
    existingIdentityMap[identity] = {
      classId: classId,
      date: ymd,
      period: period,
      sessionNumber: sessionNumber,
      rowNumber: rowNumber
    };

    if (!sessionNumber) {
      invalidSessionNumberCount += 1;
      return;
    }
    if (!existingByClass[classId]) existingByClass[classId] = [];
    existingByClass[classId].push(existingIdentityMap[identity]);
  });

  const duplicateExistingKeys = Object.keys(existingIdentityCounts).filter(function(key) {
    return existingIdentityCounts[key] > 1;
  });
  const existingClassSessionDuplicateCount = duplicateExistingKeys.reduce(function(total, key) {
    return total + existingIdentityCounts[key] - 1;
  }, 0);
  if (existingClassSessionDuplicateCount > 0) {
    blockingErrors.push(
      'classSessionsにclassId + date + period重複があります: ' +
      existingClassSessionDuplicateCount + '件 keys=' + duplicateExistingKeys.sort().join(',')
    );
  }

  const maxSessionNumberByClassId = {};
  const latestChronologyByClassId = {};
  Object.keys(existingByClass).sort().forEach(function(classId) {
    const rows = existingByClass[classId].slice().sort(compareClassSessionChronology_);
    const seenNumbers = {};
    let previousNumber = 0;
    let maxNumber = 0;

    rows.forEach(function(item) {
      const number = Number(item.sessionNumber);
      if (seenNumbers[number]) {
        invalidSessionNumberCount += 1;
        blockingErrors.push(
          'classSessionsのsessionNumberが重複しています: classId=' + classId +
          ' sessionNumber=' + number
        );
      }
      seenNumbers[number] = true;
      if (previousNumber > 0 && number <= previousNumber) {
        invalidSessionNumberCount += 1;
        blockingErrors.push(
          'classSessionsのsessionNumberが時系列順ではありません: classId=' + classId +
          ' row=' + item.rowNumber
        );
      }
      previousNumber = number;
      maxNumber = Math.max(maxNumber, number);
    });

    maxSessionNumberByClassId[classId] = maxNumber;
    if (rows.length > 0) latestChronologyByClassId[classId] = rows[rows.length - 1];
  });
  if (invalidSessionNumberCount > 0) {
    blockingErrors.push('classSessionsの不正なsessionNumberがあります: ' + invalidSessionNumberCount + '件');
  }

  const timetableRows = Object.keys(timetableByKey).map(function(key) {
    return timetableByKey[key];
  });
  const candidates = [];
  const candidateCountByTerm = {};
  classDays.forEach(function(ymd) {
    const calendarItem = calendarByDate[ymd];
    const context = calendarItem.context || getEffectiveClassDayContext_(ymd, effectiveCalendarIndex);
    timetableRows.forEach(function(timetableItem) {
      if (timetableItem.weekday !== context.effectiveWeekday) return;
      if (!timetableItem.usesLegacyTermFallback && !isTimetableTermActive_(timetableItem.term, context.term)) {
        return;
      }

      candidateCountByTerm[context.term] = (candidateCountByTerm[context.term] || 0) + 1;
      candidates.push({
        classId: timetableItem.classId,
        date: ymd,
        period: timetableItem.period,
        weekday: context.effectiveWeekday,
        term: context.term,
        identity: buildClassSessionIdentity_(
          timetableItem.classId,
          ymd,
          timetableItem.period
        )
      });
    });
  });
  candidates.sort(compareClassSessionChronology_);

  const candidateIdentityCounts = {};
  candidates.forEach(function(candidate) {
    candidateIdentityCounts[candidate.identity] = (candidateIdentityCounts[candidate.identity] || 0) + 1;
  });
  const duplicateCandidateKeys = Object.keys(candidateIdentityCounts).filter(function(key) {
    return candidateIdentityCounts[key] > 1;
  });
  const candidateDuplicateCount = duplicateCandidateKeys.reduce(function(total, key) {
    return total + candidateIdentityCounts[key] - 1;
  }, 0);
  if (candidateDuplicateCount > 0) {
    blockingErrors.push(
      'term filter後のcandidateにclassId + date + period重複があります: ' +
      candidateDuplicateCount + '件 keys=' + duplicateCandidateKeys.sort().join(',')
    );
  }

  let existingExactCount = 0;
  const pendingCandidates = candidates.filter(function(candidate) {
    if (existingIdentityCounts[candidate.identity] === 1) {
      existingExactCount += 1;
      return false;
    }
    return true;
  });
  if (classDays.length === 0) {
    warnings.push('対象範囲に授業日がありません。');
  } else if (candidates.length === 0) {
    warnings.push('対象範囲の授業日に一致するtimetableがありません。');
  } else if (pendingCandidates.length === 0) {
    warnings.push('全候補が既存classSessionsと一致しているため追加対象はありません。');
  }

  let chronologicalInsertionConflictCount = 0;
  pendingCandidates.forEach(function(candidate) {
    const latest = latestChronologyByClassId[candidate.classId];
    if (latest && compareClassSessionChronology_(candidate, latest) <= 0) {
      chronologicalInsertionConflictCount += 1;
    }
  });
  if (chronologicalInsertionConflictCount > 0) {
    blockingErrors.push(
      '既存最新sessionより前へ挿入が必要な候補があります: ' +
      chronologicalInsertionConflictCount + '件'
    );
  }

  const insertRows = [];
  const newSessionNumberRangeByClassId = {};
  if (blockingErrors.length === 0) {
    const nextNumberByClass = Object.assign({}, maxSessionNumberByClassId);
    pendingCandidates.forEach(function(candidate) {
      const nextNumber = (nextNumberByClass[candidate.classId] || 0) + 1;
      nextNumberByClass[candidate.classId] = nextNumber;
      insertRows.push([
        candidate.classId,
        candidate.date,
        Number(candidate.period),
        nextNumber
      ]);

      if (!newSessionNumberRangeByClassId[candidate.classId]) {
        newSessionNumberRangeByClassId[candidate.classId] = {
          start: nextNumber,
          end: nextNumber,
          count: 0
        };
      }
      newSessionNumberRangeByClassId[candidate.classId].end = nextNumber;
      newSessionNumberRangeByClassId[candidate.classId].count += 1;
    });
  }

  const plan = {
    startDate: startYmd,
    endDate: endYmd,
    calendarRowsInRange: calendarRowsInRange,
    classDays: classDays.length,
    timetableRows: timetable.rows.length,
    timetableRowsByTerm: sortObjectByKey_(timetableRowsByTerm),
    candidateCount: candidates.length,
    candidateCountByTerm: sortObjectByKey_(candidateCountByTerm),
    candidateDuplicateCount: candidateDuplicateCount,
    existingExactCount: existingExactCount,
    insertCount: pendingCandidates.length,
    timetableDuplicateCount: timetableDuplicateCount,
    existingClassSessionDuplicateCount: existingClassSessionDuplicateCount,
    invalidCalendarCount: invalidCalendarCount,
    invalidCalendarTermCount: invalidCalendarTermCount,
    invalidTimetableCount: invalidTimetableCount,
    invalidTimetableTermCount: invalidTimetableTermCount,
    termFallbackCount: termFallbackCount,
    invalidSessionNumberCount: invalidSessionNumberCount,
    chronologicalInsertionConflictCount: chronologicalInsertionConflictCount,
    blockingErrors: uniqueSortedStrings_(blockingErrors),
    warnings: uniqueSortedStrings_(warnings),
    maxSessionNumberByClassId: sortObjectByKey_(maxSessionNumberByClassId),
    newSessionNumberRangeByClassId: sortObjectByKey_(newSessionNumberRangeByClassId),
    insertRows: insertRows,
    sourceSignature: {
      timetableHeaders: normalizeClassSessionHashRows_([timetable.headers])[0] || [],
      timetable: normalizeClassSessionHashRows_(timetable.rows),
      calendarHeaders: normalizeClassSessionHashRows_([calendar.headers])[0] || [],
      calendar: normalizeClassSessionHashRows_(calendar.rows),
      calendarDateDisplayValues: (calendar.dateDisplayValues || []).slice(),
      effectiveCalendarContexts: classDays.map(function(ymd) {
        const context = calendarByDate[ymd].context;
        return [ymd, context.effectiveWeekday, context.term, context.usedTermFallback];
      }),
      timetableTermContract: timetableRows.map(function(item) {
        return [item.classId, item.term, item.weekday, item.period, item.usesLegacyTermFallback];
      }),
      filteredCandidates: candidates.map(function(candidate) {
        return [candidate.classId, candidate.date, candidate.period, candidate.term];
      }),
      classSessionsHeaders: normalizeClassSessionHashRows_([existing.headers])[0] || [],
      classSessions: existingHashRows
    }
  };

  plan.planHash = buildClassSessionPlanHash_(plan);
  return plan;
}

function readClassSessionPlanningSources_() {
  const ss = getOperationSpreadsheet();
  const timetableSheet = ss.getSheetByName(CONFIG.SHEETS.TIMETABLE);
  const calendarSheet = ss.getSheetByName(CONFIG.SHEETS.CALENDAR);
  const classSessionsSheet = ss.getSheetByName(CONFIG.SHEETS.CLASS_SESSIONS);

  if (!timetableSheet) throw new Error('timetable シートが見つかりません。');
  if (!calendarSheet) throw new Error('calendar シートが見つかりません。');
  if (!classSessionsSheet) throw new Error('classSessions シートが見つかりません。');

  return {
    spreadsheet: ss,
    timetableSheet: timetableSheet,
    calendarSheet: calendarSheet,
    classSessionsSheet: classSessionsSheet,
    timetable: readClassSessionPlanningSheet_(timetableSheet, []),
    calendar: readClassSessionPlanningSheet_(calendarSheet, ['date', '日付']),
    classSessions: readClassSessionPlanningSheet_(classSessionsSheet, ['date', '日付'])
  };
}

function readClassSessionPlanningSheet_(sheet, dateCandidates) {
  const values = sheet.getDataRange().getValues();
  const headers = values.length > 0 ? values[0] : [];
  const rows = values.length > 1 ? values.slice(1) : [];
  let dateDisplayValues = [];

  if (rows.length > 0 && dateCandidates.length > 0) {
    const dateColumn = findColumnIndex_(headers, dateCandidates);
    if (dateColumn !== -1) {
      dateDisplayValues = sheet
        .getRange(2, dateColumn + 1, rows.length, 1)
        .getDisplayValues()
        .map(function(row) { return row[0]; });
    }
  }

  return {
    headers: headers,
    rows: rows,
    dateDisplayValues: dateDisplayValues
  };
}

function buildClassSessionPlanPreview_(plan) {
  const sampleSize = SECOND_TERM_CLASS_SESSIONS_2026_.SAMPLE_SIZE;
  const sampleRows = plan.insertRows.map(function(row) {
    return {
      classId: row[0],
      date: row[1],
      period: row[2],
      sessionNumber: row[3]
    };
  });

  return {
    startDate: plan.startDate,
    endDate: plan.endDate,
    calendarRowsInRange: plan.calendarRowsInRange,
    classDays: plan.classDays,
    timetableRows: plan.timetableRows,
    timetableRowsByTerm: plan.timetableRowsByTerm,
    candidateCount: plan.candidateCount,
    candidateCountByTerm: plan.candidateCountByTerm,
    candidateDuplicateCount: plan.candidateDuplicateCount,
    existingExactCount: plan.existingExactCount,
    insertCount: plan.insertCount,
    timetableDuplicateCount: plan.timetableDuplicateCount,
    existingClassSessionDuplicateCount: plan.existingClassSessionDuplicateCount,
    invalidCalendarCount: plan.invalidCalendarCount,
    invalidCalendarTermCount: plan.invalidCalendarTermCount,
    invalidTimetableCount: plan.invalidTimetableCount,
    invalidTimetableTermCount: plan.invalidTimetableTermCount,
    termFallbackCount: plan.termFallbackCount,
    invalidSessionNumberCount: plan.invalidSessionNumberCount,
    chronologicalInsertionConflictCount: plan.chronologicalInsertionConflictCount,
    blockingErrors: plan.blockingErrors.slice(),
    warnings: plan.warnings.slice(),
    planHash: plan.planHash,
    maxSessionNumberByClassId: plan.maxSessionNumberByClassId,
    newSessionNumberRangeByClassId: plan.newSessionNumberRangeByClassId,
    firstInsertSamples: sampleRows.slice(0, sampleSize),
    lastInsertSamples: sampleRows.slice(Math.max(sampleRows.length - sampleSize, 0))
  };
}

function verifyAppendedClassSessionRows_(beforeSources, afterSources, plan, beforeLastRow) {
  const before = beforeSources.classSessions;
  const after = afterSources.classSessions;
  if (afterSources.classSessionsSheet.getLastRow() !== beforeLastRow + plan.insertRows.length) {
    throw new Error('classSessions追加後の行数検証に失敗しました。');
  }

  const beforeRows = normalizeClassSessionHashRows_(before.rows);
  const unchangedRows = normalizeClassSessionHashRows_(after.rows.slice(0, before.rows.length));
  if (JSON.stringify(beforeRows) !== JSON.stringify(unchangedRows)) {
    throw new Error('classSessionsの既存行が変更されています。');
  }

  const appendedRows = after.rows.slice(before.rows.length).map(function(row, index) {
    return [
      normalizeString_(row[0]),
      normalizeClassSessionPlanningYmd_(
        row[1],
        after.dateDisplayValues[before.rows.length + index]
      ),
      Number(normalizeClassSessionPeriod_(row[2])),
      Number(normalizeClassSessionNumber_(row[3]))
    ];
  });
  if (JSON.stringify(appendedRows) !== JSON.stringify(plan.insertRows)) {
    throw new Error('classSessions追加行の再読込検証に失敗しました。');
  }
}

function buildClassSessionIdentity_(classId, ymd, period) {
  return JSON.stringify([
    normalizeString_(classId),
    normalizeClassSessionPlanningYmd_(ymd, ''),
    normalizeClassSessionPeriod_(period)
  ]);
}

function normalizeClassSessionPlanningYmd_(rawValue, displayValue) {
  try {
    return normalizeEffectiveCalendarYmd_(rawValue, displayValue);
  } catch (error) {
    return '';
  }
}

function normalizeClassSessionPeriod_(value) {
  const text = normalizeString_(value);
  const number = Number(text);
  if (!text || !Number.isInteger(number) || number <= 0) return '';
  return String(number);
}

function normalizeClassSessionNumber_(value) {
  const text = normalizeString_(value);
  const number = Number(text);
  if (!text || !Number.isInteger(number) || number <= 0) return '';
  return String(number);
}

function compareClassSessionChronology_(a, b) {
  if (a.date !== b.date) return a.date < b.date ? -1 : 1;
  const periodDifference = Number(a.period) - Number(b.period);
  if (periodDifference !== 0) return periodDifference;
  return normalizeString_(a.classId).localeCompare(normalizeString_(b.classId), 'ja');
}

function resolveClassSessionPlanningColumns_(sheetName, headers, candidates, requiredKeys) {
  const columns = {};
  Object.keys(candidates).forEach(function(key) {
    columns[key] = findColumnIndex_(headers, candidates[key]);
  });
  validateRequiredColumnsForClassSessions_(sheetName, columns, requiredKeys);
  return columns;
}

function normalizeClassSessionHashRows_(rows) {
  return (rows || []).map(function(row) {
    return row.map(function(value) {
      if (value instanceof Date && !isNaN(value.getTime())) return value.toISOString();
      return normalizeString_(value);
    });
  });
}

function buildClassSessionPlanHash_(plan) {
  const hashInput = {
    version: 2,
    startDate: plan.startDate,
    endDate: plan.endDate,
    sourceSignature: plan.sourceSignature,
    insertRows: plan.insertRows,
    timetableRowsByTerm: plan.timetableRowsByTerm,
    candidateCountByTerm: plan.candidateCountByTerm,
    candidateDuplicateCount: plan.candidateDuplicateCount,
    termFallbackCount: plan.termFallbackCount,
    blockingErrors: plan.blockingErrors,
    warnings: plan.warnings
  };
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    JSON.stringify(hashInput),
    Utilities.Charset.UTF_8
  );
  return digest.map(function(byte) {
    return ('0' + ((byte + 256) % 256).toString(16)).slice(-2);
  }).join('');
}

function uniqueSortedStrings_(values) {
  const seen = {};
  (values || []).forEach(function(value) {
    const text = normalizeString_(value);
    if (text) seen[text] = true;
  });
  return Object.keys(seen).sort();
}

function sortObjectByKey_(value) {
  const result = {};
  Object.keys(value || {}).sort().forEach(function(key) {
    result[key] = value[key];
  });
  return result;
}

function validateRequiredColumnsForClassSessions_(sheetName, colMap, requiredKeys) {
  requiredKeys.forEach(function(key) {
    if (colMap[key] === -1) {
      throw new Error(sheetName + ' シートに必要な列がありません: ' + key);
    }
  });
}

function debugClassSessionService() {
  const ss = getOperationSpreadsheet();
  const timetableSheet = ss.getSheetByName(CONFIG.SHEETS.TIMETABLE);
  const calendarSheet = ss.getSheetByName(CONFIG.SHEETS.CALENDAR);

  const timetableValues = timetableSheet.getDataRange().getValues();
  const calendarValues = calendarSheet.getDataRange().getValues();

  Logger.log('timetableRows=' + Math.max(0, timetableValues.length - 1));
  Logger.log('calendarRows=' + Math.max(0, calendarValues.length - 1));

  const timetableHeaders = timetableValues[0];
  const calendarHeaders = calendarValues[0];

  Logger.log('timetableHeaders=' + JSON.stringify(timetableHeaders));
  Logger.log('calendarHeaders=' + JSON.stringify(calendarHeaders));

  const timetablePreview = timetableValues.slice(1, 6).map(function(row) {
    return {
      classId: row[0],
      weekday: row[1],
      period: row[2],
      teacherName: row[3],
      teacherId: row[4]
    };
  });

  const calendarPreview = calendarValues.slice(1, 6).map(function(row) {
    return {
      date: row[0],
      weekday: row[1],
      isClassDay: row[2]
    };
  });

  Logger.log('timetablePreview=' + JSON.stringify(timetablePreview, null, 2));
  Logger.log('calendarPreview=' + JSON.stringify(calendarPreview, null, 2));
}
