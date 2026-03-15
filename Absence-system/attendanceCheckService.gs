function getUnsubmittedClasses(targetDate) {

  const ss = getOperationSpreadsheet();

  const sessionsSheet = ss.getSheetByName(CONFIG.SHEETS.CLASS_SESSIONS);
  const attendanceSheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE);

  const sessions = sessionsSheet.getDataRange().getValues();
  const attendance = attendanceSheet.getDataRange().getValues();

  const sessionHeaders = sessions.shift();
  const attendanceHeaders = attendance.shift();

  const colSession = {
    classId: sessionHeaders.indexOf("classId"),
    date: sessionHeaders.indexOf("date"),
    period: sessionHeaders.indexOf("period")
  };

  const colAttendance = {
    classId: attendanceHeaders.indexOf("classId"),
    date: attendanceHeaders.indexOf("date"),
    period: attendanceHeaders.indexOf("period")
  };

  const target = formatDateToYmd(targetDate);

  const sessionList = sessions.filter(row => {
    return formatDateToYmd(row[colSession.date]) === target;
  });

  const attendanceSet = new Set();

  attendance.forEach(row => {

    const key = [
      String(row[colAttendance.classId]).trim(),
      formatDateToYmd(row[colAttendance.date]),
      String(row[colAttendance.period]).trim()
    ].join("_");

    attendanceSet.add(key);

  });

  const unsubmitted = [];

  sessionList.forEach(row => {

    const classId = String(row[colSession.classId]).trim();
    const date = formatDateToYmd(row[colSession.date]);
    const period = String(row[colSession.period]).trim();

    const key = [classId, date, period].join("_");

    if (!attendanceSet.has(key)) {

      unsubmitted.push({
        classId,
        date,
        period
      });

    }

  });

  return unsubmitted;

}

function checkYesterdayAttendance() {

  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);

  const unsubmitted = getUnsubmittedClasses(yesterday);

  if (unsubmitted.length === 0) {
    return;
  }

  const grouped = {};
  const logData = [];

  unsubmitted.forEach(item => {

    const teacherEmail = getTeacherByClassId(item.classId);

    if (!teacherEmail) return;

    if (!grouped[teacherEmail]) {
      grouped[teacherEmail] = [];
    }

    grouped[teacherEmail].push(item);

  });

  Object.keys(grouped).forEach(email => {

    const classes = grouped[email];

    let message = "【出席未入力通知】\n\n";

    const teacherName = getTeacherNameByEmail(email);

    message += `担当教員: ${teacherName}\n\n`;

    classes.forEach(c => {

      const className = getClassDisplayName(c.classId);
      const url =
`${CONFIG.APP.BASE_URL}?page=teacher&classId=${encodeURIComponent(c.classId)}&date=${encodeURIComponent(c.date)}&period=${encodeURIComponent(c.period)}`;

      message += `授業: ${className}\n`;
      message += `日付: ${c.date}\n`;
      message += `時限: ${c.period}\n\n`;
      message += `▶ 出席入力はこちら\n${url}\n\n`;

        logData.push({
        classId: c.classId,
        className: className,
        date: c.date,
        period: c.period,
        teacherEmail: email,
        teacherName: teacherName
  });

    });

    sendSlackMessage(message);
    
  });
  recordMissingAttendanceLog(logData);

}

function recordMissingAttendanceLog(data) {

  if (!data || data.length === 0) {
    return;
  }

  const ss = getOperationSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.ATTENDANCE_MISSING_LOG);

  const now = new Date();

  const rows = data.map(item => [
    now,
    item.classId,
    item.className,
    item.date,
    item.period,
    item.teacherEmail,
    item.teacherName
  ]);

  sheet.getRange(
    sheet.getLastRow() + 1,
    1,
    rows.length,
    rows[0].length
  ).setValues(rows);
}