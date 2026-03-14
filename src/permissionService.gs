function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function hasRole_(user, role) {
  if (!user || !Array.isArray(user.roles)) {
    return false;
  }
  return user.roles.includes(role);
}

function canAccessTeacherPage() {
  const user = getCurrentUserContext();
  return hasRole_(user, 'admin') || hasRole_(user, 'teacher');
}

function canAccessHomeroomPage() {
  const user = getCurrentUserContext();

  if (hasRole_(user, 'admin')) {
    return true;
  }

  return Array.isArray(user.homeroomClasses) && user.homeroomClasses.length > 0;
}

function canAccessAdminPage() {
  const user = getCurrentUserContext();
  return hasRole_(user, 'admin');
}

function canEditAttendance(session) {
  const user = getCurrentUserContext();

  // admin は常に許可
  if (hasRole_(user, 'admin')) {
    return true;
  }

  // teacher 権限がない場合は不可
  if (!hasRole_(user, 'teacher')) {
    return false;
  }

  if (!session || !session.classId) {
    return false;
  }

  const teacherEmail = normalizeEmail_(getTeacherByClassId(session.classId));
  const userEmail = normalizeEmail_(user.email);

  if (!teacherEmail || !userEmail) {
    return false;
  }

  return teacherEmail === userEmail;
}