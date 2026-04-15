function getClassesForCurrentUserByDate(targetDate) {
  const user = getCurrentUserContext();
  if (!user || !user.teacherId) {
    throw new Error('ログインユーザー情報を取得できませんでした。');
  }

 const currentTeacherId = normalizeString_(user.teacherId);
const currentUserEmail = normalizeString_(user.email).toLowerCase();
const ymd = targetDate
  ? formatDateToYmd(targetDate)
  : formatDateToYmd(new Date());

const classSessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_SESSIONS, 300);
const timetableData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.TIMETABLE, 300);
const classesData = getSheetDataCached_('MASTER', CONFIG.SHEETS.CLASSES, 300);
const teamData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_TEACHER_TEAMS, 300);
const attendanceSessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.ATTENDANCE_SESSIONS, 60);

const classSessions = classSessionsData.rows;
const timetable = timetableData.rows;
const classes = classesData.rows;
const teamRows = teamData.rows;
const attendanceSessionRows = attendanceSessionsData.rows;
const attendanceSessionHeaders = attendanceSessionsData.headers;

  const timetableHeaders = timetableData.headers;
  const ttCol = {
    classId: findColumnIndex_(timetableHeaders, ['classId', 'ClassID']),
    weekday: findColumnIndex_(timetableHeaders, ['weekday', '曜日']),
    period: findColumnIndex_(timetableHeaders, ['period', '時限']),
    teacherName: findColumnIndex_(timetableHeaders, ['teacherName', '担当者名', 'name']),
    teacherId: findColumnIndex_(timetableHeaders, ['teacherId', 'TeacherID'])
  };
  validateRequiredColumnsForTimetable_('timetable', ttCol, ['classId', 'weekday', 'period']);

  const teamHeaders = teamData.headers;
  const teamCol = {
    classId: findColumnIndex_(teamHeaders, ['classId', 'ClassID']),
    weekday: findColumnIndex_(teamHeaders, ['weekday', '曜日']),
    period: findColumnIndex_(teamHeaders, ['period', '時限']),
    teacherName: findColumnIndex_(teamHeaders, ['teacherName', '担当者名', 'name']),
    teacherId: findColumnIndex_(teamHeaders, ['teacherId', 'TeacherID']),
    roleType: findColumnIndex_(teamHeaders, ['roleType', '役割'])
  };

  const asCol = {
  classId: findColumnIndex_(attendanceSessionHeaders, ['classId', 'ClassID']),
  date: findColumnIndex_(attendanceSessionHeaders, ['date', '日付']),
  period: findColumnIndex_(attendanceSessionHeaders, ['period', '時限']),
  teacherEmail: findColumnIndex_(attendanceSessionHeaders, ['teacherEmail', 'email']),
  accessedAt: findColumnIndex_(attendanceSessionHeaders, ['accessedAt', 'savedAt']),
  actionType: findColumnIndex_(attendanceSessionHeaders, ['actionType']),
  targetSessionKey: findColumnIndex_(attendanceSessionHeaders, ['targetSessionKey']),
  savedModeLabel: findColumnIndex_(attendanceSessionHeaders, ['savedModeLabel'])
};

validateRequiredColumnsForTimetable_('attendanceSessions', asCol, ['classId', 'date', 'period']);

  const timetableMap = {};
  timetable.forEach(function(row) {
    const classId = normalizeString_(row[ttCol.classId]);
    const period = normalizeString_(row[ttCol.period]);
    const weekday = ttCol.weekday !== -1 ? normalizeWeekday_(row[ttCol.weekday]) : '';
    const teacherName = ttCol.teacherName !== -1 ? normalizeString_(row[ttCol.teacherName]) : '';
    let teacherId = ttCol.teacherId !== -1 ? normalizeString_(row[ttCol.teacherId]) : '';

    if (!teacherId && teacherName) {
      const teacher = getTeacherRecordByName_(teacherName);
      teacherId = teacher ? teacher.teacherId : '';
    }

    if (!classId || !period || !weekday) return;

    const key = classId + '__' + weekday + '__' + period;

    timetableMap[key] = {
      classId: classId,
      weekday: weekday,
      period: period,
      teacherId: teacherId,
      teacherName: teacherName,
      teacherIds: teacherId ? [teacherId] : [],
      teachers: teacherId ? [{
        teacherId: teacherId,
        teacherName: teacherName,
        roleType: 'main'
      }] : []
    };
  });

  teamRows.forEach(function(row) {
    const classId = teamCol.classId !== -1 ? normalizeString_(row[teamCol.classId]) : '';
    const period = teamCol.period !== -1 ? normalizeString_(row[teamCol.period]) : '';
    const weekday = teamCol.weekday !== -1 ? normalizeWeekday_(row[teamCol.weekday]) : '';
    const teacherName = teamCol.teacherName !== -1 ? normalizeString_(row[teamCol.teacherName]) : '';
    let teacherId = teamCol.teacherId !== -1 ? normalizeString_(row[teamCol.teacherId]) : '';
    const roleType = teamCol.roleType !== -1
      ? normalizeString_(row[teamCol.roleType]).toLowerCase()
      : 'support';

    if (!teacherId && teacherName) {
      const teacher = getTeacherRecordByName_(teacherName);
      teacherId = teacher ? teacher.teacherId : '';
    }

    if (!classId || !period || !weekday || !teacherId) return;

    const key = classId + '__' + weekday + '__' + period;

    if (!timetableMap[key]) {
      timetableMap[key] = {
        classId: classId,
        weekday: weekday,
        period: period,
        teacherId: '',
        teacherName: '',
        teacherIds: [],
        teachers: []
      };
    }

    if (!timetableMap[key].teacherIds.includes(teacherId)) {
      timetableMap[key].teacherIds.push(teacherId);
      timetableMap[key].teachers.push({
        teacherId: teacherId,
        teacherName: teacherName,
        roleType: roleType || 'support'
      });
    }
  });

  const savedSessionMap = {};

attendanceSessionRows.forEach(function(row) {
  const rowClassId = normalizeString_(row[asCol.classId]);
  const rowDate = formatDateToYmd(row[asCol.date]);
  const rowPeriod = normalizeString_(row[asCol.period]);

  if (!rowClassId || !rowDate || !rowPeriod) return;
  if (rowDate !== ymd) return;

  const key = [rowClassId, rowDate, rowPeriod].join('__');

  const teacherEmail = asCol.teacherEmail !== -1
    ? normalizeString_(row[asCol.teacherEmail]).toLowerCase()
    : '';

  const accessedAtRaw = asCol.accessedAt !== -1 ? row[asCol.accessedAt] : '';
  const accessedAt = accessedAtRaw instanceof Date ? accessedAtRaw : new Date(accessedAtRaw);
  const accessedAtMs = isNaN(accessedAt.getTime()) ? 0 : accessedAt.getTime();

  if (!savedSessionMap[key] || accessedAtMs >= savedSessionMap[key]._ms) {
savedSessionMap[key] = {
  teacherEmail: teacherEmail,
  savedAtText: formatDateTimeJst_(accessedAtRaw),
  actionType: asCol.actionType !== -1 ? normalizeString_(row[asCol.actionType]) : '',
  targetSessionKey: asCol.targetSessionKey !== -1 ? normalizeString_(row[asCol.targetSessionKey]) : '',
  savedModeLabel: asCol.savedModeLabel !== -1 ? normalizeString_(row[asCol.savedModeLabel]) : '',
  savedByCurrentUser: !!teacherEmail && teacherEmail === currentUserEmail,
  _ms: accessedAtMs
};
  }
});

  const classHeaders = classesData.headers;
  const clsCol = {
    classId: findColumnIndex_(classHeaders, ['classId', 'ClassID']),
    subjectId: findColumnIndex_(classHeaders, ['subjectId', 'SubjectID']),
    subjectName: findColumnIndex_(classHeaders, ['subjectName', '科目名']),
    grade: findColumnIndex_(classHeaders, ['grade', '学年']),
    unit: findColumnIndex_(classHeaders, ['unit', '対象区分', '組・コース']),
    term: findColumnIndex_(classHeaders, ['term', '開設期']),
    curriculumUnit: findColumnIndex_(classHeaders, ['curriculumUnit', '組・コース']),
    allowedAbsences: findColumnIndex_(classHeaders, ['allowedAbsences', '欠席可能コマ数'])
  };
  validateRequiredColumnsForTimetable_('classes', clsCol, ['classId', 'subjectId', 'subjectName', 'grade', 'unit']);

  const classMap = {};
  classes.forEach(function(row) {
    const classId = normalizeString_(row[clsCol.classId]);
    if (!classId) return;

    classMap[classId] = {
      classId: classId,
      subjectId: clsCol.subjectId !== -1 ? normalizeString_(row[clsCol.subjectId]) : '',
      subjectName: clsCol.subjectName !== -1 ? normalizeString_(row[clsCol.subjectName]) : '',
      grade: clsCol.grade !== -1 ? normalizeString_(row[clsCol.grade]) : '',
      unit: clsCol.unit !== -1 ? normalizeString_(row[clsCol.unit]) : '',
      term: clsCol.term !== -1 ? normalizeString_(row[clsCol.term]) : '',
      curriculumUnit: clsCol.curriculumUnit !== -1 ? normalizeString_(row[clsCol.curriculumUnit]) : '',
      allowedAbsences: clsCol.allowedAbsences !== -1 ? row[clsCol.allowedAbsences] : ''
    };
  });

  const classSessionHeaders = classSessionsData.headers;
  const csCol = {
    classId: findColumnIndex_(classSessionHeaders, ['classId', 'ClassID']),
    date: findColumnIndex_(classSessionHeaders, ['date', '日付']),
    period: findColumnIndex_(classSessionHeaders, ['period', '時限']),
    sessionNumber: findColumnIndex_(classSessionHeaders, ['sessionNumber', '回', '回数'])
  };
  validateRequiredColumnsForTimetable_('classSessions', csCol, ['classId', 'date', 'period']);

  return classSessions
    .filter(function(row) {
      return formatDateToYmd(row[csCol.date]) === ymd;
    })
    .map(function(row) {
      const classId = normalizeString_(row[csCol.classId]);
      const period = normalizeString_(row[csCol.period]);
      const sessionNumber = csCol.sessionNumber !== -1 ? row[csCol.sessionNumber] : '';
      const sessionYmd = formatDateToYmd(row[csCol.date]);
      const weekday = getWeekdayFromYmdJst_(sessionYmd);

      const tt = timetableMap[classId + '__' + weekday + '__' + period];
      const cls = classMap[classId];
      const saveInfoRaw = savedSessionMap[[classId, sessionYmd, period].join('__')] || null;

const saveInfo = saveInfoRaw ? {
  teacherEmail: saveInfoRaw.teacherEmail || '',
  savedAtText: saveInfoRaw.savedAtText || '',
  actionType: saveInfoRaw.actionType || '',
  targetSessionKey: saveInfoRaw.targetSessionKey || '',
  savedModeLabel: saveInfoRaw.savedModeLabel || '',
  savedByCurrentUser: !!saveInfoRaw.savedByCurrentUser
} : null;

      if (!tt || !Array.isArray(tt.teacherIds) || !tt.teacherIds.includes(currentTeacherId)) {
        return null;
      }

     return {
  classId: classId,
  date: sessionYmd,
  period: period,
  sessionNumber: sessionNumber,
  subjectId: cls ? cls.subjectId : '',
  subjectName: cls ? cls.subjectName : '',
  grade: cls ? cls.grade : '',
  unit: cls ? cls.unit : '',
  term: cls ? cls.term : '',
  curriculumUnit: cls ? cls.curriculumUnit : '',
  allowedAbsences: cls ? cls.allowedAbsences : '',
  teacherId: tt.teacherId,
  teacherName: tt.teacherName,
  teacherIds: tt.teacherIds,
  teachers: tt.teachers,
  weekday: tt.weekday,
  isSaved: !!saveInfo,
  lastSavedInfo: saveInfo
};
    })
    .filter(Boolean)
    .sort(function(a, b) {
      return Number(a.period) - Number(b.period);
    });
}

function getWeekdayFromDate_(value) {
  const date = value instanceof Date ? value : new Date(value);
  return Utilities.formatDate(date, 'Asia/Tokyo', 'EEE'); // Sun, Mon, Tue...
}

function getTodayClassesForCurrentUser(targetDate) {
  return getClassesForCurrentUserByDate(targetDate);
}

function testGetTodayClassesForCurrentUser() {
  const result = getTodayClassesForCurrentUser('2026-04-06');
  Logger.log(JSON.stringify(result, null, 2));
}

function debugTodayClassMatching(targetDate) {
  const user = getCurrentUserContext();
  Logger.log('currentUser=' + JSON.stringify(user, null, 2));

  const timezone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const today = targetDate
    ? formatDateToYmd(targetDate)
    : Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');

  const classSessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_SESSIONS, 60);
  const timetableData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.TIMETABLE, 60);
  const classesData = getSheetDataCached_('MASTER', CONFIG.SHEETS.CLASSES, 60);

  Logger.log('today=' + today);
  Logger.log('classSessionsHeaders=' + JSON.stringify(classSessionsData.headers));
  Logger.log('timetableHeaders=' + JSON.stringify(timetableData.headers));
  Logger.log('classesHeaders=' + JSON.stringify(classesData.headers));

  const classSessionHeaders = classSessionsData.headers;
  const csDateCol = findColumnIndex_(classSessionHeaders, ['date', '日付']);
  if (csDateCol === -1) {
    throw new Error('classSessions の date 列が見つかりません');
  }

  const filtered = classSessionsData.rows.filter(function(row) {
    return formatDateToYmd(row[csDateCol]) === today;
  });

  Logger.log('filteredClassSessions=' + JSON.stringify(filtered));
}

function validateRequiredColumnsForTimetable_(sheetName, colMap, requiredKeys) {
  requiredKeys.forEach(function(key) {
    if (colMap[key] === -1) {
      throw new Error(sheetName + ' シートに必要な列がありません: ' + key);
    }
  });
}

function debugTeacherClassesByDate() {
  targetDate = '2026-04-09';
  const user = getCurrentUserContext();
  Logger.log('user=' + JSON.stringify(user, null, 2));

  const currentTeacherId = normalizeString_(user.teacherId);
  const ymd = formatDateToYmd(targetDate);

  const classSessionsData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_SESSIONS, 60);
  const timetableData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.TIMETABLE, 60);
  const teamData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.CLASS_TEACHER_TEAMS, 60);

  const classSessions = classSessionsData.rows;
  const timetable = timetableData.rows;
  const teamRows = teamData.rows;

  const ttCol = {
    classId: findColumnIndex_(timetableData.headers, ['classId', 'ClassID']),
    weekday: findColumnIndex_(timetableData.headers, ['weekday', '曜日']),
    period: findColumnIndex_(timetableData.headers, ['period', '時限']),
    teacherId: findColumnIndex_(timetableData.headers, ['teacherId', 'TeacherID']),
    teacherName: findColumnIndex_(timetableData.headers, ['teacherName', '担当者名', 'name'])
  };

  const teamCol = {
    classId: findColumnIndex_(teamData.headers, ['classId', 'ClassID']),
    weekday: findColumnIndex_(teamData.headers, ['weekday', '曜日']),
    period: findColumnIndex_(teamData.headers, ['period', '時限']),
    teacherId: findColumnIndex_(teamData.headers, ['teacherId', 'TeacherID']),
    teacherName: findColumnIndex_(teamData.headers, ['teacherName', '担当者名', 'name']),
    roleType: findColumnIndex_(teamData.headers, ['roleType', '役割'])
  };

  const csCol = {
    classId: findColumnIndex_(classSessionsData.headers, ['classId', 'ClassID']),
    date: findColumnIndex_(classSessionsData.headers, ['date', '日付']),
    period: findColumnIndex_(classSessionsData.headers, ['period', '時限'])
  };

  const weekday = getWeekdayFromYmdJst_(ymd);
  Logger.log('targetDate=' + ymd + ', weekday=' + weekday + ', teacherId=' + currentTeacherId);

  const targetSessions = classSessions.filter(row => formatDateToYmd(row[csCol.date]) === ymd);
  Logger.log('targetSessions=' + JSON.stringify(targetSessions, null, 2));

  targetSessions.forEach(function(row) {
    const classId = normalizeString_(row[csCol.classId]);
    const period = normalizeString_(row[csCol.period]);

    const timetableHit = timetable.filter(t =>
      normalizeString_(t[ttCol.classId]) === classId &&
      normalizeWeekday_(t[ttCol.weekday]) === weekday &&
      normalizeString_(t[ttCol.period]) === period
    );

    const teamHit = teamRows.filter(t =>
      normalizeString_(t[teamCol.classId]) === classId &&
      normalizeWeekday_(t[teamCol.weekday]) === weekday &&
      normalizeString_(t[teamCol.period]) === period
    );

    Logger.log('---');
    Logger.log('session classId=' + classId + ', period=' + period);
    Logger.log('timetableHit=' + JSON.stringify(timetableHit, null, 2));
    Logger.log('teamHit=' + JSON.stringify(teamHit, null, 2));
  });
}

function getWeekdayFromYmdJst_(ymd) {
  const m = String(ymd || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';

  // JSTの正午を明示して、実行環境のタイムゾーン影響を受けないようにする
  const date = new Date(ymd + 'T12:00:00+09:00');
  const weekday = Utilities.formatDate(date, 'Asia/Tokyo', 'E'); // Mon, Tue, Wed...

  return normalizeWeekday_(weekday);
}