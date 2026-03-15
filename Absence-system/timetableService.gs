function getTodayClassesForCurrentUser(targetDate) {
  Logger.log('getTodayClassesForCurrentUser called');

  const opSs = getOperationSpreadsheet();
  const masterSs = getMasterSpreadsheet();

  const classSessionsSheet = opSs.getSheetByName(CONFIG.SHEETS.CLASS_SESSIONS);
  const timetableSheet = opSs.getSheetByName(CONFIG.SHEETS.TIMETABLE);
  const subjectsSheet = masterSs.getSheetByName(CONFIG.SHEETS.SUBJECTS);
  const classesSheet = masterSs.getSheetByName(CONFIG.SHEETS.CLASSES);

  if (!classSessionsSheet) throw new Error('classSessions シートがありません');
  if (!timetableSheet) throw new Error('timetable シートがありません');
  if (!subjectsSheet) throw new Error('Subjects シートがありません');
  if (!classesSheet) throw new Error('classes シートがありません');

  const currentUserEmail = String(getCurrentUserEmail()).trim().toLowerCase();
  Logger.log('currentUserEmail=' + currentUserEmail);

  const timezone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const today = targetDate
    ? formatDateToYmd(targetDate)
    : Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');

  Logger.log('today=' + today);

  const classSessions = getSheetBodyValues_(classSessionsSheet);
  const timetable = getSheetBodyValues_(timetableSheet);
  const subjects = getSheetBodyValues_(subjectsSheet);
  const classes = getSheetBodyValues_(classesSheet);

  Logger.log('classSessions=' + JSON.stringify(classSessions));
  Logger.log('timetable=' + JSON.stringify(timetable));

  // ===== timetable（OPERATION）=====
  // 想定ヘッダー例: classId / weekday / period / teacherEmail
  const timetableHeaders = getSheetHeaders_(timetableSheet);
  const ttCol = {
    classId: findColumnIndex_(timetableHeaders, ['classId', 'ClassID']),
    weekday: findColumnIndex_(timetableHeaders, ['weekday', '曜日']),
    period: findColumnIndex_(timetableHeaders, ['period', '時限']),
    teacherEmail: findColumnIndex_(timetableHeaders, ['teacherEmail', 'email', '担当者メール'])
  };
  validateColumns_('timetable', ttCol);

  const timetableMap = {};
  timetable.forEach(row => {
    const classId = String(row[ttCol.classId] || '').trim();
    const period = String(row[ttCol.period] || '').trim();

    if (!classId || !period) return;

    const key = classId + '__' + period;
    timetableMap[key] = {
      classId: classId,
      weekday: row[ttCol.weekday],
      period: period,
      teacherEmail: String(row[ttCol.teacherEmail] || '').trim().toLowerCase()
    };
  });

  // ===== Subjects（MASTER）=====
  // 想定ヘッダー例:
  // SubjectID, 科目名, 開設学年, 組・コース, 開設期, 履修区分, 欠席可能コマ数
  const subjectHeaders = getSheetHeaders_(subjectsSheet);
  const subjCol = {
    subjectId: findColumnIndex_(subjectHeaders, ['SubjectID', 'subjectId']),
    subjectName: findColumnIndex_(subjectHeaders, ['科目名', 'subjectName'])
  };
  validateColumns_('Subjects', subjCol);

  const subjectMap = {};
  subjects.forEach(row => {
    const subjectId = String(row[subjCol.subjectId] || '').trim();
    if (!subjectId) return;

    subjectMap[subjectId] = {
      subjectId: subjectId,
      subjectName: String(row[subjCol.subjectName] || '').trim()
    };
  });

  // ===== classes（MASTER）=====
  // 想定ヘッダー例:
  // ClassID, SubjectID, 学年, 対象区分
  const classHeaders = getSheetHeaders_(classesSheet);
  const clsCol = {
    classId: findColumnIndex_(classHeaders, ['ClassID', 'classId']),
    subjectId: findColumnIndex_(classHeaders, ['SubjectID', 'subjectId']),
    grade: findColumnIndex_(classHeaders, ['学年', 'grade']),
    unit: findColumnIndex_(classHeaders, ['対象区分', 'unit', '組・コース'])
  };
  validateColumns_('classes', clsCol);

  const classMap = {};
  classes.forEach(row => {
    const classId = String(row[clsCol.classId] || '').trim();
    if (!classId) return;

    classMap[classId] = {
      classId: classId,
      subjectId: String(row[clsCol.subjectId] || '').trim(),
      grade: row[clsCol.grade],
      unit: row[clsCol.unit]
    };
  });

  // ===== classSessions（OPERATION）=====
  // 想定ヘッダー例: classId / date / period / sessionNumber
  const classSessionHeaders = getSheetHeaders_(classSessionsSheet);
  const csCol = {
    classId: findColumnIndex_(classSessionHeaders, ['classId', 'ClassID']),
    date: findColumnIndex_(classSessionHeaders, ['date', '日付']),
    period: findColumnIndex_(classSessionHeaders, ['period', '時限']),
    sessionNumber: findColumnIndex_(classSessionHeaders, ['sessionNumber', '回', '回数'])
  };
  validateColumns_('classSessions', csCol);

  const result = classSessions
    .filter(row => formatDateToYmd(row[csCol.date]) === today)
    .map(row => {
      const classId = String(row[csCol.classId] || '').trim();
      const period = String(row[csCol.period] || '').trim();
      const sessionNumber = row[csCol.sessionNumber];

      const tt = timetableMap[classId + '__' + period];
      const cls = classMap[classId];
      const subj = cls ? subjectMap[cls.subjectId] : null;

      Logger.log(
        'checking classId=' + classId +
        ', period=' + period +
        ', ttTeacher=' + (tt ? tt.teacherEmail : 'NONE') +
        ', currentUser=' + currentUserEmail
      );

      if (!tt || tt.teacherEmail !== currentUserEmail) {
        return null;
      }

      return {
        classId: classId,
        date: formatDateToYmd(row[csCol.date]),
        period: period,
        sessionNumber: sessionNumber,
        subjectName: subj ? subj.subjectName : '',
        grade: cls ? cls.grade : '',
        unit: cls ? cls.unit : ''
      };
    })
    .filter(Boolean)
    .sort((a, b) => Number(a.period) - Number(b.period));

  Logger.log('result=' + JSON.stringify(result));
  return result;
}

function testGetTodayClassesForCurrentUser() {
  const result = getTodayClassesForCurrentUser('2026-04-06');
  Logger.log(JSON.stringify(result));
}

function debugTodayClassMatching(targetDate) {
  const opSs = getOperationSpreadsheet();

  const classSessionsSheet = opSs.getSheetByName(CONFIG.SHEETS.CLASS_SESSIONS);
  const timetableSheet = opSs.getSheetByName(CONFIG.SHEETS.TIMETABLE);

  if (!classSessionsSheet) throw new Error('classSessions シートがありません');
  if (!timetableSheet) throw new Error('timetable シートがありません');

  const currentUserEmail = String(getCurrentUserEmail()).trim().toLowerCase();
  const timezone = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const today = targetDate
    ? formatDateToYmd(targetDate)
    : Utilities.formatDate(new Date(), timezone, 'yyyy-MM-dd');

  const classSessions = getSheetBodyValues_(classSessionsSheet);
  const timetable = getSheetBodyValues_(timetableSheet);

  Logger.log('currentUserEmail=' + currentUserEmail);
  Logger.log('today=' + today);
  Logger.log('classSessions=' + JSON.stringify(classSessions));
  Logger.log('timetable=' + JSON.stringify(timetable));

  const classSessionHeaders = getSheetHeaders_(classSessionsSheet);
  const csDateCol = findColumnIndex_(classSessionHeaders, ['date', '日付']);
  if (csDateCol === -1) {
    throw new Error('classSessions の date 列が見つかりません');
  }

  const filtered = classSessions.filter(row => formatDateToYmd(row[csDateCol]) === today);
  Logger.log('filteredClassSessions=' + JSON.stringify(filtered));
  Logger.log('spreadsheetId=' + opSs.getId());
  Logger.log('spreadsheetName=' + opSs.getName());
}

/**
 * ヘッダー行を返す
 */
function getSheetHeaders_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) return [];
  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
}

/**
 * データ本体（2行目以降）を返す
 */
function getSheetBodyValues_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn === 0) {
    return [];
  }

  return sheet.getRange(2, 1, lastRow - 1, lastColumn).getValues();
}

/**
 * 候補ヘッダー名から列番号を探す
 */
function findColumnIndex_(headers, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const idx = headers.indexOf(candidates[i]);
    if (idx !== -1) return idx;
  }
  return -1;
}

/**
 * 必須列チェック
 */
function validateColumns_(sheetName, colMap) {
  Object.keys(colMap).forEach(key => {
    if (colMap[key] === -1) {
      throw new Error(sheetName + ' シートに必要な列がありません: ' + key);
    }
  });
}