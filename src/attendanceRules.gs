function isAttendanceEditable(date) {

  const classDate = new Date(date);

  const limit = new Date(classDate);
  limit.setDate(limit.getDate() + 1);

  limit.setHours(8);
  limit.setMinutes(40);
  limit.setSeconds(0);

  const now = new Date();

  return now <= limit;
}