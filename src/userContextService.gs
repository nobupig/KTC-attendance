function parseRoles_(rolesText) {
  return String(rolesText || '')
    .split(',')
    .map(role => normalizeString_(role))
    .filter(Boolean);
}

/**
 * 開発者なりすまし設定
 *
 * enabled: true にすると有効
 * developerEmail: 実際にログインしている開発者メール
 * masqueradeEmail: テスト対象の教員メール
 *
 * 例:
 * enabled: true →通常利用の場合: falseにすればいい
 * developerEmail: 'nyasui@ktc.ac.jp'
 * masqueradeEmail: 'oono@ktc.ac.jp'
 */
function getDevMasqueradeConfig_() {
  return {
    enabled: false, 
    developerEmail: 'nyasui@ktc.ac.jp',
    masqueradeEmail: 'nakahira@ktc.ac.jp'
  };
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

/**
 * 開発者なりすまし適用後のメールアドレスを返す
 */
function getCurrentUserEmail() {
  const actualEmail = getActualCurrentUserEmail_();
  const config = getDevMasqueradeConfig_();

  if (
    config &&
    config.enabled === true &&
    normalizeString_(config.developerEmail).toLowerCase() === actualEmail
  ) {
    const masqueradeEmail = normalizeString_(config.masqueradeEmail).toLowerCase();
    if (masqueradeEmail) {
      return masqueradeEmail;
    }
  }

  return actualEmail;
}

/**
 * 現在ユーザーのコンテキストを返す
 *
 * 返却例:
 * {
 *   teacherId: 'T018',
 *   name: '安井 宣仁',
 *   email: 'nyasui@ktc.ac.jp',
 *   actualEmail: 'nyasui@ktc.ac.jp',
 *   isMasquerading: false,
 *   roles: ['teacher','homeroom','admin'],
 *   homeroomClasses: [{ grade:'3', unit:'CA' }],
 *   isTeacher: true,
 *   isHomeroom: true,
 *   isAdmin: true
 * }
 */
function getCurrentUserContext() {
  const actualEmail = getActualCurrentUserEmail_();
  const resolvedEmail = getCurrentUserEmail();

  const ss = openOperationSpreadsheet_();
  const teachers = readSheetAsObjects_(ss, CONFIG.SHEETS.TEACHERS);

  const teacher = teachers.find(row =>
    normalizeString_(row.email).toLowerCase() === resolvedEmail
  );

  if (!teacher) {
    return null;
  }

  const teacherId = normalizeString_(teacher.teacherId);
  const name = normalizeString_(teacher.name);
  const roles = parseRoles_(teacher.roles);
  const homeroomClasses = getHomeroomClassesByTeacherId_(ss, teacherId);

  return {
    teacherId: teacherId,
    name: name,
    email: resolvedEmail,
    actualEmail: actualEmail,
    isMasquerading: resolvedEmail !== actualEmail,
    roles: roles,
    homeroomClasses: homeroomClasses,
    isTeacher: roles.includes('teacher'),
    isHomeroom: roles.includes('homeroom'),
    isAdmin: roles.includes('admin')
  };
}

function getHomeroomClassesByTeacherId_(ss, teacherId) {
  const rows = readSheetAsObjects_(ss, CONFIG.SHEETS.HOMEROOM_ASSIGNMENTS);

  const results = [];
  const seen = new Set();

  rows.forEach(row => {
    const rowTeacherId = normalizeString_(row.teacherId);
    const grade = normalizeString_(row.grade);
    const unit = normalizeString_(row.unit);

    if (!rowTeacherId || !grade || !unit) return;
    if (rowTeacherId !== teacherId) return;

    const key = grade + '|' + unit;
    if (seen.has(key)) return;

    seen.add(key);
    results.push({
      grade: grade,
      unit: unit
    });
  });

  results.sort((a, b) => {
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

  const headers = values[0].map(h => normalizeString_(h));

  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((header, i) => {
      obj[header] = row[i];
    });
    return obj;
  });
}

/**
 * 動作確認用
 */
function debugCurrentUserContext() {
  const user = getCurrentUserContext();
  Logger.log(JSON.stringify(user, null, 2));
}