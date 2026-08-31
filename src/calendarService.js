const CALENDAR_SERVICE_CONFIG = {
  EXCEPTION_SHEET_NAME: 'calendarExceptions',
  CALENDAR_HEADER: ['date', 'weekday', 'isClassDay'],
  WEEKDAY_LABELS: ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'],

  FIRST_TERM_2026: {
    start: '2026-04-08',
    end: '2026-09-11',
    closedRanges: [
      {
        start: '2026-07-28',
        end: '2026-09-06',
        isClassDay: false,
        note: '夏季休暇'
      }
    ]
  }
};

/**
 * 汎用: 指定期間の calendar を生成・更新する
 * - 指定期間内だけ上書き
 * - 期間外の既存データは保持
 * - 土日FALSE / 平日TRUE
 * - calendarExceptions で個別上書き
 *
 * @param {string} startDateStr YYYY-MM-DD
 * @param {string} endDateStr YYYY-MM-DD
 * @returns {{start:string,end:string,updatedCount:number}}
 */
function generateCalendar(startDateStr, endDateStr) {
  return upsertCalendarRange_(startDateStr, endDateStr, []);
}

/**
 * 2026年前期用
 * - 2026-04-08 ～ 2026-09-11
 * - 夏季休暇 2026-07-28 ～ 2026-09-06 を FALSE
 */
function generateFirstTermCalendar2026() {
  const cfg = CALENDAR_SERVICE_CONFIG.FIRST_TERM_2026;
  return upsertCalendarRange_(cfg.start, cfg.end, cfg.closedRanges);
}

/**
 * テスト用
 */
function testGenerateFirstTermCalendar2026() {
  const result = generateFirstTermCalendar2026();
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * calendar を全消去してヘッダーだけ残す
 */
function clearCalendarSheet() {
  return runCalendarWriteWithLock_(function() {
    return clearCalendarSheetUnderLock_();
  });
}

function clearCalendarSheetUnderLock_() {
  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CALENDAR);

  if (!sheet) {
    throw new Error('calendar シートがありません');
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, CALENDAR_SERVICE_CONFIG.CALENDAR_HEADER.length)
    .setValues([CALENDAR_SERVICE_CONFIG.CALENDAR_HEADER]);

  invalidateEffectiveCalendarCache_();
}

/* =========================
 * 内部処理
 * ========================= */

function upsertCalendarRange_(startDateStr, endDateStr, closedRanges) {
  return runCalendarWriteWithLock_(function() {
    return upsertCalendarRangeUnderLock_(startDateStr, endDateStr, closedRanges);
  });
}

function upsertCalendarRangeUnderLock_(startDateStr, endDateStr, closedRanges) {
  const ss = getOperationSpreadsheet();
  const calendarSheet = ss.getSheetByName(CONFIG.SHEETS.CALENDAR);

  if (!calendarSheet) {
    throw new Error('calendar シートがありません');
  }

  const startDate = parseYmdToDate_(startDateStr);
  const endDate = parseYmdToDate_(endDateStr);

  if (startDate > endDate) {
    throw new Error('開始日は終了日以前にしてください');
  }

  const exceptionMap = buildCalendarExceptionMap_();
  const generatedMap = buildGeneratedCalendarMap_(startDate, endDate, closedRanges, exceptionMap);
  const existingMap = readExistingCalendarMap_(calendarSheet);

  const startKey = formatDateToYmd(startDate);
  const endKey = formatDateToYmd(endDate);

  // 指定期間内は generatedMap で上書き、期間外は既存を保持
  const mergedMap = {};

  Object.keys(existingMap).forEach(dateKey => {
    if (dateKey < startKey || dateKey > endKey) {
      mergedMap[dateKey] = existingMap[dateKey];
    }
  });

  Object.keys(generatedMap).forEach(dateKey => {
    mergedMap[dateKey] = generatedMap[dateKey];
  });

  const sortedKeys = Object.keys(mergedMap).sort();

  const rows = sortedKeys.map(dateKey => mergedMap[dateKey]);

  calendarSheet.clearContents();
  calendarSheet
    .getRange(1, 1, 1, CALENDAR_SERVICE_CONFIG.CALENDAR_HEADER.length)
    .setValues([CALENDAR_SERVICE_CONFIG.CALENDAR_HEADER]);

  if (rows.length > 0) {
    calendarSheet
      .getRange(2, 1, rows.length, rows[0].length)
      .setValues(rows);
  }

  invalidateEffectiveCalendarCache_();

  return {
    start: startKey,
    end: endKey,
    updatedCount: Object.keys(generatedMap).length
  };
}

function buildGeneratedCalendarMap_(startDate, endDate, closedRanges, exceptionMap) {
  const map = {};
  const current = new Date(startDate);

  while (current <= endDate) {
    const dateKey = formatDateToYmd(current);
    const day = current.getDay(); // 0:Sun ... 6:Sat
    const weekday = CALENDAR_SERVICE_CONFIG.WEEKDAY_LABELS[day];

    // 原則: 平日TRUE / 土日FALSE
    let isClassDay = !(day === 0 || day === 6);

    // 長期休暇・期間例外を適用
    (closedRanges || []).forEach(range => {
      const rangeStart = String(range.start || '').trim();
      const rangeEnd = String(range.end || '').trim();
      const rangeFlag = toBooleanForCalendar_(range.isClassDay);

      if (!rangeStart || !rangeEnd) {
        return;
      }

      if (dateKey >= rangeStart && dateKey <= rangeEnd) {
        isClassDay = rangeFlag;
      }
    });

    // calendarExceptions で最終上書き
    if (Object.prototype.hasOwnProperty.call(exceptionMap, dateKey)) {
      isClassDay = exceptionMap[dateKey];
    }

    map[dateKey] = [dateKey, weekday, isClassDay];

    current.setDate(current.getDate() + 1);
  }

  return map;
}

function buildCalendarExceptionMap_() {
  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CALENDAR_SERVICE_CONFIG.EXCEPTION_SHEET_NAME);

  if (!sheet || sheet.getLastRow() < 2) {
    return {};
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = values.slice(1);

  const col = {
    date: headers.indexOf('date'),
    isClassDay: headers.indexOf('isClassDay')
  };

  if (col.date === -1 || col.isClassDay === -1) {
    throw new Error('calendarExceptions シートに date または isClassDay 列がありません');
  }

  const map = {};

  rows.forEach(row => {
    const rawDate = row[col.date];
    if (!rawDate) return;

    const dateKey = formatDateToYmd(rawDate);
    map[dateKey] = toBooleanForCalendar_(row[col.isClassDay]);
  });

  return map;
}

function readExistingCalendarMap_(sheet) {
  const map = {};

  if (sheet.getLastRow() < 2) {
    return map;
  }

  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const rows = values.slice(1);

  const col = {
    date: headers.indexOf('date'),
    weekday: headers.indexOf('weekday'),
    isClassDay: headers.indexOf('isClassDay')
  };

  if (col.date === -1 || col.weekday === -1 || col.isClassDay === -1) {
    throw new Error('calendar シートに必要な列がありません');
  }

  rows.forEach(row => {
    const rawDate = row[col.date];
    if (!rawDate) return;

    const dateKey = formatDateToYmd(rawDate);
    map[dateKey] = [
      dateKey,
      String(row[col.weekday] || '').trim(),
      toBooleanForCalendar_(row[col.isClassDay])
    ];
  });

  return map;
}

function parseYmdToDate_(ymd) {
  const value = String(ymd || '').trim();
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if (!match) {
    throw new Error('日付形式が不正です: ' + value + '（YYYY-MM-DD 形式にしてください）');
  }

  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);

  return new Date(year, month - 1, day);
}

function toBooleanForCalendar_(value) {
  if (typeof value === 'boolean') {
    return value;
  }

  const normalized = String(value || '').trim().toUpperCase();

  return normalized === 'TRUE' || normalized === '1';
}

/**
 * 対象日に実施する授業曜日を calendar から解決する。
 * calendar に行がない日付だけは、従来互換として実曜日へフォールバックする。
 *
 * @param {*} value 日付
 * @param {Object=} calendarIndex buildEffectiveClassDayIndex_ の戻り値
 * @returns {{date:string,hasCalendarEntry:boolean,isClassDay:boolean,weekday:string,usedActualWeekdayFallback:boolean}}
 */
function getEffectiveClassDayInfo_(value, calendarIndex) {
  const ymd = normalizeEffectiveCalendarYmd_(value, '');
  const index = calendarIndex || getEffectiveClassDayIndex_();
  const hasCalendarEntry = Object.prototype.hasOwnProperty.call(index, ymd);

  if (hasCalendarEntry) {
    const entry = index[ymd] || {};
    const isClassDay = entry.isClassDay === true;

    return {
      date: ymd,
      hasCalendarEntry: true,
      isClassDay: isClassDay,
      weekday: isClassDay ? normalizeWeekday_(entry.weekday) : '',
      usedActualWeekdayFallback: false
    };
  }

  return {
    date: ymd,
    hasCalendarEntry: false,
    isClassDay: true,
    weekday: getWeekdayFromYmdJst_(ymd),
    usedActualWeekdayFallback: true
  };
}

function getEffectiveWeekdayForDate_(value, calendarIndex) {
  const info = getEffectiveClassDayInfo_(value, calendarIndex);
  return info.isClassDay ? info.weekday : '';
}

function getEffectiveClassDayIndex_() {
  const calendarData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CALENDAR, 300);
  return buildEffectiveClassDayIndex_(calendarData);
}

function buildEffectiveClassDayIndex_(calendarData) {
  const headers = Array.isArray(calendarData && calendarData.headers)
    ? calendarData.headers
    : [];
  const rows = Array.isArray(calendarData && calendarData.rows)
    ? calendarData.rows
    : [];
  const dateDisplayValues = Array.isArray(calendarData && calendarData.dateDisplayValues)
    ? calendarData.dateDisplayValues
    : [];

  const col = {
    date: findColumnIndex_(headers, ['date', '日付']),
    weekday: findColumnIndex_(headers, ['weekday', '曜日']),
    isClassDay: findColumnIndex_(headers, ['isClassDay', '授業日'])
  };

  ['date', 'weekday', 'isClassDay'].forEach(function(key) {
    if (col[key] === -1) {
      throw new Error('calendar シートに必要な列がありません: ' + key);
    }
  });

  const index = {};

  rows.forEach(function(row, rowIndex) {
    const rawDate = row[col.date];
    const displayDate = dateDisplayValues[rowIndex];
    const ymd = normalizeEffectiveCalendarYmd_(rawDate, displayDate);

    if (!ymd) return;

    index[ymd] = {
      weekday: normalizeWeekday_(row[col.weekday]),
      isClassDay: row[col.isClassDay] === true
    };
  });

  return index;
}

function normalizeEffectiveCalendarYmd_(rawDate, displayDate) {
  const displayYmd = normalizeYmdDisplayText_(displayDate);
  if (displayYmd) return displayYmd;

  if (!rawDate) return '';
  if (rawDate instanceof Date) return formatDateToYmd(rawDate);

  const rawText = String(rawDate).trim();
  if (!rawText) return '';

  // getSheetDataCached_ 経由の Date は UTC の ISO 文字列になるため、
  // 先頭10文字ではなく JST に戻して日付を確定する。
  if (/^\d{4}-\d{2}-\d{2}T/.test(rawText)) {
    return formatDateToYmd(rawText);
  }

  return normalizeYmdDisplayText_(rawText) || formatDateToYmd(rawText);
}

function invalidateEffectiveCalendarCache_() {
  removeScriptCacheKeys_([
    'sheetData__OPERATION__' + CONFIG.SHEETS.CALENDAR
  ]);

  invalidateTeacherUnsavedFastSnapshotsAfterCalendarChange_();
}

function runCalendarWriteWithLock_(callback) {
  const lock = LockService.getScriptLock();
  const alreadyLocked = lock.hasLock();

  if (!alreadyLocked) {
    lock.waitLock(10000);
  }

  try {
    return callback();
  } finally {
    if (!alreadyLocked) {
      lock.releaseLock();
    }
  }
}

function testEffectiveWeekdayContract() {
  const cases = [
    { name: 'A', date: '2026-10-12', weekday: '月', isClassDay: true, expected: 'Mon', actual: 'Mon' },
    { name: 'B', date: '2026-10-15', weekday: '月', isClassDay: true, expected: 'Mon', actual: 'Thu' },
    { name: 'C', date: '2026-11-27', weekday: '月', isClassDay: true, expected: 'Mon', actual: 'Fri' },
    { name: 'D', date: '2027-01-14', weekday: '火', isClassDay: true, expected: 'Tue', actual: 'Thu' },
    { name: 'E', date: '2026-10-16', weekday: '金', isClassDay: false, expected: '', actual: 'Fri' }
  ];

  const results = cases.map(function(testCase) {
    const index = {};
    index[testCase.date] = {
      weekday: testCase.weekday,
      isClassDay: testCase.isClassDay
    };

    const actual = getEffectiveClassDayInfo_(testCase.date, index);
    const actualCalendarWeekday = getWeekdayFromYmdJst_(testCase.date);
    const passed = actual.isClassDay === testCase.isClassDay &&
      actual.weekday === testCase.expected &&
      actualCalendarWeekday === testCase.actual;

    if (!passed) {
      throw new Error(
        'Effective Weekday test ' + testCase.name + ' failed: ' + JSON.stringify(actual)
      );
    }

    return {
      name: testCase.name,
      date: testCase.date,
      isClassDay: actual.isClassDay,
      weekday: actual.weekday,
      actualCalendarWeekday: actualCalendarWeekday,
      passed: true
    };
  });

  const cachedCalendarIndex = buildEffectiveClassDayIndex_({
    headers: ['date', 'weekday', 'isClassDay'],
    rows: [['2026-10-14T15:00:00.000Z', '月', true]]
  });
  if (!cachedCalendarIndex['2026-10-15']) {
    throw new Error('Effective Weekday cached calendar date normalization failed');
  }

  const missingCalendarFallback = getEffectiveClassDayInfo_('2026-10-15', {});
  if (
    missingCalendarFallback.weekday !== 'Thu' ||
    missingCalendarFallback.usedActualWeekdayFallback !== true
  ) {
    throw new Error('Effective Weekday missing-row fallback failed');
  }

  const directIsoInput = getEffectiveClassDayInfo_('2026-10-14T15:00:00.000Z', {
    '2026-10-15': { weekday: '月', isClassDay: true }
  });
  if (directIsoInput.date !== '2026-10-15' || directIsoInput.weekday !== 'Mon') {
    throw new Error('Effective Weekday direct ISO normalization failed');
  }

  Logger.log(JSON.stringify(results, null, 2));
  return results;
}
