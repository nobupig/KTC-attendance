function generateClassSessions() {
  const ss = getOperationSpreadsheet();

  const timetableSheet = ss.getSheetByName(CONFIG.SHEETS.TIMETABLE);
  const calendarSheet = ss.getSheetByName(CONFIG.SHEETS.CALENDAR);
  const classSessionsSheet = ss.getSheetByName(CONFIG.SHEETS.CLASS_SESSIONS);

  const timetableValues = timetableSheet.getDataRange().getValues();
  const calendarValues = calendarSheet.getDataRange().getValues();

  if (timetableValues.length < 2 || calendarValues.length < 2) {
    throw new Error('timetable または calendar にデータがありません。');
  }

  const timetableData = timetableValues.slice(1).map(row => ({
    classId: row[0],
    weekday: row[1],
    period: row[2],
    teacherEmail: row[3]
  }));

  const calendarData = calendarValues.slice(1)
    .filter(row => row[2] === true)
    .map(row => ({
      date: row[0],
      weekday: row[1]
    }));

  const sessionCountMap = {};
  const output = [];

  calendarData.forEach(cal => {
    timetableData.forEach(tt => {
      if (cal.weekday === tt.weekday) {
        if (!sessionCountMap[tt.classId]) {
          sessionCountMap[tt.classId] = 0;
        }

        sessionCountMap[tt.classId] += 1;

        output.push([
          tt.classId,
          formatDateToYmd(cal.date),
          tt.period,
          sessionCountMap[tt.classId]
        ]);
      }
    });
  });

  classSessionsSheet.clear();
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