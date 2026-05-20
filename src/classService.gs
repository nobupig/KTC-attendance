function getClassRecordById_(classId) {
  const targetClassId = normalizeString_(classId);
  if (!targetClassId) {
    return null;
  }

  const cacheKey = 'classRecordById__' + targetClassId;
  const cached = getScriptCacheJson_(cacheKey);
  if (cached !== null) {
    return cached;
  }

  const classesData = getSheetDataCached_('MASTER', CONFIG.SHEETS.CLASSES, 300);
  const headers = classesData.headers;
  const rows = classesData.rows;

  const col = {
    classId: findColumnIndex_(headers, ['classId', 'ClassID']),
    subjectId: findColumnIndex_(headers, ['subjectId', 'SubjectID']),
    subjectName: findColumnIndex_(headers, ['subjectName', '科目名']),
    grade: findColumnIndex_(headers, ['grade', '学年']),
    unit: findColumnIndex_(headers, ['unit', '対象区分', '組・コース']),
    term: findColumnIndex_(headers, ['term', '開設期']),
    curriculumUnit: findColumnIndex_(headers, ['curriculumUnit', '組・コース']),
    allowedAbsences: findColumnIndex_(headers, ['allowedAbsences', '欠席可能コマ数'])
  };

  validateRequiredColumnsForClassService_('classes', col, [
    'classId',
    'subjectId',
    'subjectName',
    'grade',
    'unit'
  ]);

  const row = rows.find(function(r) {
    return normalizeString_(r[col.classId]) === targetClassId;
  });

  const result = row ? {
    classId: normalizeString_(row[col.classId]),
    subjectId: col.subjectId !== -1 ? normalizeString_(row[col.subjectId]) : '',
    subjectName: col.subjectName !== -1 ? normalizeString_(row[col.subjectName]) : '',
    grade: col.grade !== -1 ? normalizeString_(row[col.grade]) : '',
    unit: col.unit !== -1 ? normalizeString_(row[col.unit]) : '',
    term: col.term !== -1 ? normalizeString_(row[col.term]) : '',
    curriculumUnit: col.curriculumUnit !== -1 ? normalizeString_(row[col.curriculumUnit]) : '',
    allowedAbsences: col.allowedAbsences !== -1 ? row[col.allowedAbsences] : ''
  } : null;

  putScriptCacheJson_(cacheKey, result, 300);
  return result;
}

function getClassInfoById(classId) {
  return getClassRecordById_(classId);
}

function getClassDisplayName(classId) {
  const cls = getClassRecordById_(classId);
  if (!cls) {
    return normalizeString_(classId);
  }

const gradeUnitLabel = (cls.grade && cls.unit)
  ? cls.grade + '年' + cls.unit
  : ((cls.grade ? cls.grade + '年' : '') + (cls.unit || ''));

const subjectLabel = cls.subjectName || cls.classId;

return [gradeUnitLabel, subjectLabel]
  .filter(Boolean)
  .join(' ');
}

function getClassAndTeacherAssignment_(classId) {
  const cls = getClassRecordById_(classId);
  const assignment = getTeacherAssignmentByClassId_(classId);

  if (!cls && !assignment) {
    return null;
  }

  return {
    classId: normalizeString_(classId),
    classInfo: cls,
    teacherAssignment: assignment
  };
}

function validateRequiredColumnsForClassService_(sheetName, colMap, requiredKeys) {
  requiredKeys.forEach(function(key) {
    if (colMap[key] === -1) {
      throw new Error(sheetName + ' シートに必要な列がありません: ' + key);
    }
  });
}

/**
 * 動作確認用
 * classId を1つ指定して、classes から正しく取れるか確認します
 */
function debugClassService() {
  const ss = getOperationSpreadsheet();
  const timetableSheet = ss.getSheetByName(CONFIG.SHEETS.TIMETABLE);
  if (!timetableSheet) {
    throw new Error('timetable シートが見つかりません');
  }

  const values = timetableSheet.getDataRange().getValues();
  if (values.length < 2) {
    Logger.log('timetable にデータがありません');
    return;
  }

  const headers = values[0];
  const rows = values.slice(1);
  const classIdCol = findColumnIndex_(headers, ['classId', 'ClassID']);
  if (classIdCol === -1) {
    throw new Error('timetable シートに classId 列がありません');
  }

  const firstClassId = rows
    .map(function(r) { return normalizeString_(r[classIdCol]); })
    .find(Boolean);

  if (!firstClassId) {
    Logger.log('timetable に classId が入っていません');
    return;
  }

  Logger.log('classId=' + firstClassId);
  Logger.log('classRecord=' + JSON.stringify(getClassRecordById_(firstClassId), null, 2));
  Logger.log('displayName=' + getClassDisplayName(firstClassId));
  Logger.log('classAndTeacher=' + JSON.stringify(getClassAndTeacherAssignment_(firstClassId), null, 2));
}