function doGet(e) {
  const page = (e && e.parameter && e.parameter.page)
    ? String(e.parameter.page).trim()
    : '';

  const user = getCurrentUserContext();
  if (!user) {
    throw new Error('ユーザー情報を取得できませんでした');
  }

  const email = normalizeString_(user.email).toLowerCase();
  const teacherId = normalizeString_(user.teacherId);
  const roles = Array.isArray(user.roles) ? user.roles : [];
  const homeroomClasses = Array.isArray(user.homeroomClasses) ? user.homeroomClasses : [];

  const isAdmin = roles.includes('admin');
  const hasHomeroom = homeroomClasses.length > 0;
  const hasTeacherRole = roles.includes('teacher');
  const hasTeachingAssignment = hasAnyTeachingAssignmentByTeacherId_(teacherId);

  // teacher権限がなくても、時間割に担当授業があるなら teacher 扱いで通す
  const canUseTeacher = hasTeacherRole || hasTeachingAssignment;

  // page指定が無いときの初期遷移
  // admin または担任あり → ポータル
  // それ以外で teacher 利用可 → teacher 直行
  // どれにも該当しない → 権限案内
  let targetPage = page;
  if (!targetPage) {
    if (isAdmin || hasHomeroom) {
      targetPage = 'index';
    } else if (canUseTeacher) {
      targetPage = 'teacher';
    } else {
      return buildUnregisteredUserHtml_(email, hasTeachingAssignment);
    }
  }

const pageToFileMap = {
  index: 'html/index',
  teacher: 'html/teacher',
  homeroom: 'html/homeroom',
  homeroomShr: 'html/homeroomShr',
  admin: 'html/admin'
};

  if (!pageToFileMap[targetPage]) {
    targetPage = (isAdmin || hasHomeroom)
      ? 'index'
      : (canUseTeacher ? 'teacher' : '');

    if (!targetPage) {
      return buildUnregisteredUserHtml_(email, hasTeachingAssignment);
    }
  }

  // 権限チェック
   // 権限チェック
  if (targetPage === 'index') {
    if (!(isAdmin || hasHomeroom)) {
      // ポータル利用対象外の教員は teacher へ直行
      if (canUseTeacher) {
        targetPage = 'teacher';
      } else {
        return buildUnregisteredUserHtml_(email, hasTeachingAssignment);
      }
    }
  }

  if (targetPage === 'teacher' && !canUseTeacher) {
    return buildUnregisteredUserHtml_(email, hasTeachingAssignment);
  }

  if (targetPage === 'homeroom') {
    const canUseHomeroom = hasHomeroom;
    if (!canUseHomeroom) {
      return HtmlService.createHtmlOutput('<h2>担任画面の権限がありません。</h2>');
    }
  }

  if (targetPage === 'homeroomShr') {
    const canUseHomeroom = hasHomeroom;
    if (!canUseHomeroom) {
      return HtmlService.createHtmlOutput('<h2>担任SHR画面の権限がありません。</h2>');
    }
  }

  if (targetPage === 'admin' && !isAdmin) {
    return HtmlService.createHtmlOutput('<h2>教務画面の権限がありません。</h2>');
  }

  const template = HtmlService.createTemplateFromFile(pageToFileMap[targetPage]);

  template.classId = e && e.parameter && e.parameter.classId ? e.parameter.classId : '';
  template.date = e && e.parameter && e.parameter.date ? e.parameter.date : '';
  template.period = e && e.parameter && e.parameter.period ? e.parameter.period : '';

  template.hasTeacherRole = hasTeacherRole;
  template.hasTeachingAssignment = hasTeachingAssignment;
  template.needsTeacherRoleNotice = (!hasTeacherRole && hasTeachingAssignment);

  return template
    .evaluate()
    .setTitle('出席管理システム')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAppBaseUrl() {
  return ScriptApp.getService().getUrl();
}

function getPortalInitialData() {
  const user = getCurrentUserContext();
  if (!user) {
    throw new Error('ユーザー情報を取得できませんでした');
  }

  return {
    user: {
      name: user.name || '',
      email: user.email || '',
      teacherId: user.teacherId || '',
      roles: Array.isArray(user.roles) ? user.roles : [],
      homeroomClasses: Array.isArray(user.homeroomClasses) ? user.homeroomClasses : []
    },
    appBaseUrl: getAppBaseUrl()
  };
}

function getUserRoleInfo() {
  const user = getCurrentUserContext();
  return {
    email: user.email,
    teacherId: user.teacherId,
    roles: user.roles,
    homeroomClasses: user.homeroomClasses,
    name: user.name
  };
}

/**
 * 時間割に担当授業が1件でもあるか確認
 */
/**
 * 時間割または classTeacherTeams に担当授業が1件でもあるか確認
 */
function hasAnyTeachingAssignmentByTeacherId_(teacherId) {
  const targetTeacherId = normalizeString_(teacherId);
  if (!targetTeacherId) return false;

  const ss = openOperationSpreadsheet_();
  const timetable = readSheetAsObjects_(ss, CONFIG.SHEETS.TIMETABLE);

  const hasTimetableAssignment = timetable.some(function(row) {
    return normalizeString_(row.teacherId) === targetTeacherId;
  });

  if (hasTimetableAssignment) {
    return true;
  }

  try {
    const teamRows = readSheetAsObjects_(ss, CONFIG.SHEETS.CLASS_TEACHER_TEAMS);
    return teamRows.some(function(row) {
      return normalizeString_(row.teacherId) === targetTeacherId;
    });
  } catch (e) {
    return false;
  }
}

/**
 * 権限未設定ユーザー向け案内画面
 */
function buildUnregisteredUserHtml_(email, hasTeachingAssignment) {
  const safeEmail = escapeHtml_(email || '');

  let message = '';
  if (hasTeachingAssignment) {
    message =
      '<p>このメールアドレスは時間割上の担当授業を持っていますが、権限設定に反映されていない可能性があります。</p>' +
      '<p>教務担当者に、<strong>teachers / 権限設定 / 担当者情報</strong> の登録状況をご確認ください。</p>';
  } else {
    message =
      '<p>このメールアドレスでは利用可能な権限が見つかりませんでした。</p>' +
      '<p>教務担当者に、<strong>メールアドレス登録・teacher 権限・担任設定</strong> の確認をご依頼ください。</p>';
  }

  const html =
    '<!DOCTYPE html>' +
    '<html><head><meta charset="UTF-8"><title>権限未設定</title>' +
    '<style>' +
    'body{font-family:sans-serif;background:#f5f7fb;padding:32px;color:#223;}' +
    '.box{max-width:760px;margin:0 auto;background:#fff;border-radius:16px;padding:24px;box-shadow:0 4px 14px rgba(0,0,0,.08);}' +
    'h1{margin-top:0;color:#a12622;}' +
    '.email{margin:12px 0;padding:10px 12px;background:#f7f9fc;border-radius:10px;color:#334;}' +
    '.note{margin-top:16px;font-size:13px;color:#667;}' +
    '</style></head><body>' +
    '<div class="box">' +
    '<h1>利用権限を確認してください</h1>' +
    '<div class="email">ログイン中: ' + safeEmail + '</div>' +
    message +
    '<div class="note">登録修正後、再度アクセスしてください。</div>' +
    '</div></body></html>';

  return HtmlService.createHtmlOutput(html).setTitle('権限未設定');
}

function escapeHtml_(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}