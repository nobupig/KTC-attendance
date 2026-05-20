function parseRoles_(rolesText) {
  return String(rolesText || '')
    .split(',')
    .map(role => normalizeString_(role))
    .filter(Boolean);
}


function getDevMasqueradeConfig_() {
  return {
    enabled: true,
    developerEmail: 'nyasui@ktc.ac.jp',
    propertyKey: 'devMasqueradeEmail'
  };
}

function getMasqueradePropertyKey_() {
  const config = getDevMasqueradeConfig_();
  return normalizeString_(config.propertyKey) || 'devMasqueradeEmail';
}

function isMasqueradeDeveloper_(actualEmail) {
  const config = getDevMasqueradeConfig_();
  return !!(
    config &&
    config.enabled === true &&
    normalizeString_(config.developerEmail).toLowerCase() === normalizeString_(actualEmail).toLowerCase()
  );
}

/**
 * 実際のログインメールアドレスを取得
 */
function getActualCurrentUserEmail_() {
  const activeUserEmail = normalizeString_(Session.getActiveUser().getEmail()).toLowerCase();
  const effectiveUserEmail = normalizeString_(Session.getEffectiveUser().getEmail()).toLowerCase();

  const email = activeUserEmail || effectiveUserEmail;

  if (!email) {
    throw new Error(
      'ログイン中のメールアドレスを取得できませんでした。' +
      'ウェブアプリの公開設定と、組織アカウントでのログイン状態を確認してください。'
    );
  }

  return email;
}

function getStoredMasqueradeEmail_() {
  const actualEmail = getActualCurrentUserEmail_();
  if (!isMasqueradeDeveloper_(actualEmail)) {
    return '';
  }

  const value = PropertiesService.getUserProperties().getProperty(getMasqueradePropertyKey_()) || '';
  return normalizeString_(value).toLowerCase();
}

function getMasqueradeSelectableTeachers() {
  const actualEmail = getActualCurrentUserEmail_();
  if (!isMasqueradeDeveloper_(actualEmail)) {
    return [];
  }

  const seen = new Set();

  return getTeachersSheetObjectsCached_(300)
    .map(function(row) {
      return {
        teacherId: normalizeString_(row.teacherId),
        name: normalizeString_(row.name),
        email: normalizeString_(row.email).toLowerCase(),
        roles: parseRoles_(row.roles)
      };
    })
    .filter(function(row) {
      if (!row.email) return false;
      if (row.email === actualEmail) return false;
      if (seen.has(row.email)) return false;
      seen.add(row.email);
      return true;
    })
    .sort(function(a, b) {
      const nameCompare = String(a.name || '').localeCompare(String(b.name || ''), 'ja');
      if (nameCompare !== 0) return nameCompare;
      return String(a.email || '').localeCompare(String(b.email || ''), 'ja');
    });
}

function isValidMasqueradeTeacherEmail_(email) {
  const target = normalizeString_(email).toLowerCase();
  if (!target) return false;

  return getMasqueradeSelectableTeachers().some(function(row) {
    return row.email === target;
  });
}

/**
 * 開発者なりすまし適用後のメールアドレスを返す
 */
function getCurrentUserEmail() {
  const actualEmail = getActualCurrentUserEmail_();

  if (!isMasqueradeDeveloper_(actualEmail)) {
    return actualEmail;
  }

  const storedEmail = getStoredMasqueradeEmail_();
  if (storedEmail && isValidMasqueradeTeacherEmail_(storedEmail)) {
    return storedEmail;
  }

  return actualEmail;
}

function buildCurrentUserContextCacheKey_(actualEmail, resolvedEmail) {
  return 'currentUserContext__' +
    normalizeString_(actualEmail).toLowerCase() + '__' +
    normalizeString_(resolvedEmail).toLowerCase();
}

function getOperationSheetObjectsCached_(sheetName, ttlSeconds) {
  const data = getSheetDataCached_('OPERATION', sheetName, ttlSeconds || 300);
  const headers = Array.isArray(data && data.headers) ? data.headers : [];
  const rows = Array.isArray(data && data.rows) ? data.rows : [];

  if (headers.length === 0 || rows.length === 0) {
    return [];
  }

  return rows.map(function(row) {
    return getRowObject_(headers, row);
  });
}

function getTeachersSheetObjectsCached_(ttlSeconds) {
  return getOperationSheetObjectsCached_(CONFIG.SHEETS.TEACHERS, ttlSeconds || 300);
}

function getHomeroomAssignmentsSheetObjectsCached_(ttlSeconds) {
  return getOperationSheetObjectsCached_(CONFIG.SHEETS.HOMEROOM_ASSIGNMENTS, ttlSeconds || 300);
}

/**
 * 現在ユーザーのコンテキストを返す
 */
function getCurrentUserContext() {
  const actualEmail = getActualCurrentUserEmail_();
  const resolvedEmail = getCurrentUserEmail();
  const cacheKey = buildCurrentUserContextCacheKey_(actualEmail, resolvedEmail);
  const cached = getScriptCacheJson_(cacheKey);

  if (cached !== null) {
    return cached;
  }

  const teachers = getTeachersSheetObjectsCached_(300);

  const teacher = teachers.find(function(row) {
    return normalizeString_(row.email).toLowerCase() === resolvedEmail;
  });

  if (!teacher) {
    putScriptCacheJson_(cacheKey, null, 120);
    return null;
  }

  const teacherId = normalizeString_(teacher.teacherId);
  const name = normalizeString_(teacher.name);
  const roles = parseRoles_(teacher.roles);
  const homeroomClasses = getHomeroomClassesByTeacherId_(teacherId);

  const context = {
    teacherId: teacherId,
    name: name,
    email: resolvedEmail,
    actualEmail: actualEmail,
    isMasquerading: resolvedEmail !== actualEmail,
    canMasquerade: isMasqueradeDeveloper_(actualEmail),
    roles: roles,
    homeroomClasses: homeroomClasses,
    isTeacher: roles.includes('teacher'),
    isHomeroom: roles.includes('homeroom'),
    isAdmin: roles.includes('admin')
  };

  putScriptCacheJson_(cacheKey, context, 300);
  return context;
}

function getHomeroomClassesByTeacherId_(teacherId) {
  const targetTeacherId = normalizeString_(teacherId);
  if (!targetTeacherId) return [];

  const rows = getHomeroomAssignmentsSheetObjectsCached_(300);

  const results = [];
  const seen = new Set();

  rows.forEach(function(row) {
    const rowTeacherId = normalizeString_(row.teacherId);
    const grade = normalizeString_(row.grade);
    const unit = normalizeString_(row.unit);

    if (!rowTeacherId || !grade || !unit) return;
    if (rowTeacherId !== targetTeacherId) return;

    const key = grade + '|' + unit;
    if (seen.has(key)) return;

    seen.add(key);
    results.push({
      grade: grade,
      unit: unit
    });
  });

  results.sort(function(a, b) {
    const gradeA = Number(a.grade);
    const gradeB = Number(b.grade);
    if (gradeA !== gradeB) return gradeA - gradeB;
    return String(a.unit).localeCompare(String(b.unit), 'ja');
  });

  return results;
}

function openOperationSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEETS.OPERATION);
}

function readSheetAsObjects_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('シートが見つかりません: ' + sheetName);
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0].map(function(h) {
    return normalizeString_(h);
  });

  return values.slice(1).map(function(row) {
    const obj = {};
    headers.forEach(function(header, i) {
      obj[header] = row[i];
    });
    return obj;
  });
}

function getMasqueradeUiState() {
  const actualEmail = getActualCurrentUserEmail_();
  const context = getCurrentUserContext();
  const canMasquerade = isMasqueradeDeveloper_(actualEmail);

  return {
    canMasquerade: canMasquerade,
    actualEmail: actualEmail,
    currentEmail: context ? context.email : getCurrentUserEmail(),
    currentName: context ? context.name : '',
    isMasquerading: context ? !!context.isMasquerading : false,
    options: canMasquerade ? getMasqueradeSelectableTeachers() : []
  };
}

function setMasqueradeEmail(email) {
  const actualEmail = getActualCurrentUserEmail_();
  if (!isMasqueradeDeveloper_(actualEmail)) {
    throw new Error('なりすまし変更権限がありません。');
  }

  const nextEmail = normalizeString_(email).toLowerCase();
  if (!nextEmail) {
    throw new Error('なりすまし先の教員を選択してください。');
  }

  if (!isValidMasqueradeTeacherEmail_(nextEmail)) {
    throw new Error('指定した教員メールが教員名簿に存在しません。');
  }

  const oldResolvedEmail = getCurrentUserEmail();

  PropertiesService.getUserProperties().setProperty(getMasqueradePropertyKey_(), nextEmail);

  removeScriptCacheKeys_([
    buildCurrentUserContextCacheKey_(actualEmail, actualEmail),
    buildCurrentUserContextCacheKey_(actualEmail, oldResolvedEmail),
    buildCurrentUserContextCacheKey_(actualEmail, nextEmail)
  ]);

  return getMasqueradeUiState();
}

function clearMasqueradeEmail() {
  const actualEmail = getActualCurrentUserEmail_();
  if (!isMasqueradeDeveloper_(actualEmail)) {
    throw new Error('なりすまし解除権限がありません。');
  }

  const oldResolvedEmail = getCurrentUserEmail();

  PropertiesService.getUserProperties().deleteProperty(getMasqueradePropertyKey_());

  removeScriptCacheKeys_([
    buildCurrentUserContextCacheKey_(actualEmail, actualEmail),
    buildCurrentUserContextCacheKey_(actualEmail, oldResolvedEmail)
  ]);

  return getMasqueradeUiState();
}

function clearCurrentUserContextCacheForTest() {
  const actualEmail = getActualCurrentUserEmail_();
  const resolvedEmail = getCurrentUserEmail();

  removeScriptCacheKeys_([
    buildCurrentUserContextCacheKey_(actualEmail, actualEmail),
    buildCurrentUserContextCacheKey_(actualEmail, resolvedEmail)
  ]);
}

/**
 * 動作確認用
 */
function debugCurrentUserContext() {
  const user = getCurrentUserContext();
  Logger.log(JSON.stringify(user, null, 2));
}
