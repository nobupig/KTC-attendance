function getTeacherByClassId(classId) {
  const targetClassId = String(classId || '').trim();
  if (!targetClassId) {
    return null;
  }

  const cacheKey = 'teacherByClassId__' + targetClassId;
  const cached = getScriptCacheJson_(cacheKey);
  if (cached !== null) {
    return cached;
  }

  const timetableData = getSheetDataCached_('OPERATION', CONFIG.SHEETS.TIMETABLE, 300);
  const headers = timetableData.headers;
  const rows = timetableData.rows;

  const classIndex = headers.indexOf("classId");
  const teacherIndex = headers.indexOf("teacherEmail");

  if (classIndex === -1 || teacherIndex === -1) {
    throw new Error('timetable シートに classId / teacherEmail 列がありません');
  }

  const row = rows.find(function(r) {
    return String(r[classIndex]).trim() === targetClassId;
  });

  const result = row ? row[teacherIndex] : null;
  putScriptCacheJson_(cacheKey, result, 300);
  return result;
}