function normalizeWeekday_(value) {
  const s = String(value || '').trim();

  const map = {
    '月': 'Mon',
    '火': 'Tue',
    '水': 'Wed',
    '木': 'Thu',
    '金': 'Fri',
    '土': 'Sat',
    '日': 'Sun',

    '月曜': 'Mon',
    '火曜': 'Tue',
    '水曜': 'Wed',
    '木曜': 'Thu',
    '金曜': 'Fri',
    '土曜': 'Sat',
    '日曜': 'Sun',

    '月曜日': 'Mon',
    '火曜日': 'Tue',
    '水曜日': 'Wed',
    '木曜日': 'Thu',
    '金曜日': 'Fri',
    '土曜日': 'Sat',
    '日曜日': 'Sun',

    'Mon': 'Mon',
    'Tue': 'Tue',
    'Wed': 'Wed',
    'Thu': 'Thu',
    'Fri': 'Fri',
    'Sat': 'Sat',
    'Sun': 'Sun',

    'Monday': 'Mon',
    'Tuesday': 'Tue',
    'Wednesday': 'Wed',
    'Thursday': 'Thu',
    'Friday': 'Fri',
    'Saturday': 'Sat',
    'Sunday': 'Sun'
  };

  return map[s] || '';
}

function normalizeString_(value) {
  return String(value == null ? '' : value).trim();
}

function findColumnIndex_(headers, candidates) {
  const normalizedHeaders = headers.map(h => normalizeString_(h));
  for (const candidate of candidates) {
    const idx = normalizedHeaders.indexOf(candidate);
    if (idx !== -1) return idx;
  }
  return -1;
}

function getRowObject_(headers, row) {
  const obj = {};
  headers.forEach((header, i) => {
    obj[normalizeString_(header)] = row[i];
  });
  return obj;
}