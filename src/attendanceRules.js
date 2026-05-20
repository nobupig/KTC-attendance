function isAttendanceEditable(date) {
  const tz = "Asia/Tokyo";

  const targetYmd = normalizeAttendanceRuleYmd_(date);
  if (!targetYmd) {
    console.warn("[attendanceRules] invalid attendance date:", date);
    return false;
  }

  const now = new Date();
  const todayYmd = Utilities.formatDate(now, tz, "yyyy-MM-dd");
  const nowHm = Utilities.formatDate(now, tz, "HHmm");

  // 未来日は保存不可
  if (targetYmd > todayYmd) {
    return false;
  }

  // 授業当日は終日保存可能
  if (targetYmd === todayYmd) {
    return true;
  }

  // 授業翌日の 8:40 までは保存可能
  const deadlineYmd = addDaysYmdForAttendanceRule_(targetYmd, 1);
  if (todayYmd === deadlineYmd) {
    return nowHm <= "0840";
  }

  // それ以前の過去日は通常保存不可
  return false;
}

function normalizeAttendanceRuleYmd_(value) {
  const tz = "Asia/Tokyo";

  if (value === null || value === undefined || value === "") {
    return "";
  }

  if (Object.prototype.toString.call(value) === "[object Date]") {
    if (isNaN(value.getTime())) return "";
    return Utilities.formatDate(value, tz, "yyyy-MM-dd");
  }

  const text = String(value).trim();

  // 例:
  // 2026-05-20
  // 2026/05/20
  // 2026年5月20日
  // 2026-05-20 / 第19回
  const match = text.match(/(\d{4})[-\/年](\d{1,2})[-\/月](\d{1,2})/);
  if (!match) {
    return "";
  }

  const y = match[1];
  const m = String(Number(match[2])).padStart(2, "0");
  const d = String(Number(match[3])).padStart(2, "0");

  return y + "-" + m + "-" + d;
}

function addDaysYmdForAttendanceRule_(ymd, days) {
  const tz = "Asia/Tokyo";
  const parts = String(ymd).split("-");
  const y = Number(parts[0]);
  const m = Number(parts[1]);
  const d = Number(parts[2]);

  const date = new Date(y, m - 1, d);
  date.setDate(date.getDate() + Number(days || 0));

  return Utilities.formatDate(date, tz, "yyyy-MM-dd");
}