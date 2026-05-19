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