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
  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CALENDAR);

  if (!sheet) {
    throw new Error('calendar シートがありません');
  }

  sheet.clearContents();
  sheet.getRange(1, 1, 1, CALENDAR_SERVICE_CONFIG.CALENDAR_HEADER.length)
    .setValues([CALENDAR_SERVICE_CONFIG.CALENDAR_HEADER]);
}

/* =========================
 * 内部処理
 * ========================= */

function upsertCalendarRange_(startDateStr, endDateStr, closedRanges) {
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