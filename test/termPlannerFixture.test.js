const assert = require('assert');
const crypto = require('crypto');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const weekdayMap = {
  '月': 'Mon', Mon: 'Mon', '火': 'Tue', Tue: 'Tue', '水': 'Wed', Wed: 'Wed',
  '木': 'Thu', Thu: 'Thu', '金': 'Fri', Fri: 'Fri', '土': 'Sat', Sat: 'Sat',
  '日': 'Sun', Sun: 'Sun'
};

const sandbox = {
  console,
  Logger: { log() {} },
  normalizeString_(value) { return value == null ? '' : String(value).trim(); },
  findColumnIndex_(headers, candidates) {
    for (const candidate of candidates || []) {
      const index = (headers || []).indexOf(candidate);
      if (index !== -1) return index;
    }
    return -1;
  },
  normalizeWeekday_(value) { return weekdayMap[String(value == null ? '' : value).trim()] || ''; },
  normalizeYmdDisplayText_(value) {
    const match = String(value == null ? '' : value).trim().match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
    return match ? `${match[1]}-${String(match[2]).padStart(2, '0')}-${String(match[3]).padStart(2, '0')}` : '';
  },
  formatDateToYmd(value) {
    const text = String(value == null ? '' : value);
    return sandbox.normalizeYmdDisplayText_(text) || '';
  },
  getWeekdayFromYmdJst_(ymd) {
    return { '2026-10-15': 'Thu' }[ymd] || '';
  },
  Utilities: {
    DigestAlgorithm: { SHA_256: 'sha256' },
    Charset: { UTF_8: 'utf8' },
    computeDigest(_algorithm, input) {
      return Array.from(crypto.createHash('sha256').update(input, 'utf8').digest());
    }
  }
};
vm.createContext(sandbox);
['src/calendarService.js', 'src/classSessionService.js'].forEach(file => {
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), sandbox, { filename: file });
});

function assertFixtureEqual(actual, expected) {
  assert.deepStrictEqual(JSON.parse(JSON.stringify(actual)), expected);
}

function sources(timetableHeaders, timetableRows, calendarHeaders, calendarRows, existingRows) {
  return {
    timetable: { headers: timetableHeaders, rows: timetableRows, dateDisplayValues: [] },
    calendar: { headers: calendarHeaders, rows: calendarRows, dateDisplayValues: calendarRows.map(row => row[0]) },
    classSessions: {
      headers: ['classId', 'date', 'period', 'sessionNumber'],
      rows: existingRows || [],
      dateDisplayValues: (existingRows || []).map(row => row[1])
    }
  };
}

const termHeaders = ['classId', 'weekday', 'period', 'teacherName', 'teacherId', 'term'];
const calendarHeaders = ['date', 'weekday', 'isClassDay', 'term'];

{
  const substitute = sandbox.getEffectiveClassDayContext_('2026-10-15', {
    '2026-10-15': { weekday: '月', isClassDay: true, term: 'SP' }
  });
  assertFixtureEqual(substitute, {
    ymd: '2026-10-15', date: '2026-10-15', hasCalendarEntry: true,
    isClassDay: true, effectiveWeekday: 'Mon', weekday: 'Mon', term: 'SP',
    usedActualWeekdayFallback: false, usedTermFallback: false
  });

  const legacy = sandbox.getEffectiveClassDayContext_('2026-10-15', {
    '2026-10-15': { weekday: 'Mon', isClassDay: true, term: '' }
  });
  assert.strictEqual(legacy.term, 'SP');
  assert.strictEqual(legacy.usedTermFallback, true);
  assert.strictEqual(sandbox.getEffectiveClassDayContext_('2026-10-15', {
    '2026-10-15': { weekday: 'Mon', isClassDay: false, term: 'SP' }
  }).isClassDay, false);
  assert.strictEqual(sandbox.getEffectiveClassDayContext_('2026-10-15', {
    '2026-10-15': { weekday: 'Mon', isClassDay: true, term: 'BAD' }
  }).isClassDay, false);
}

{
  const plan = sandbox.buildRangeLimitedClassSessionPlan_('2026-04-08', '2026-04-08', sources(
    termHeaders,
    [['fa', 'Wed', 1, '', '', 'FA'], ['sp', 'Wed', 1, '', '', 'SP'], ['fy', 'Wed', 1, '', '', 'FY']],
    calendarHeaders,
    [['2026-04-08', 'Wed', true, 'FA']]
  ));
  assertFixtureEqual(plan.insertRows.map(row => row[0]), ['fa', 'fy']);
  assertFixtureEqual(plan.candidateCountByTerm, { FA: 2 });
}

['2026-10-15', '2026-11-27'].forEach(date => {
  const plan = sandbox.buildRangeLimitedClassSessionPlan_(date, date, sources(
    termHeaders,
    [['fa', 'Mon', 1, '', '', 'FA'], ['sp', 'Mon', 1, '', '', 'SP'], ['fy', 'Mon', 1, '', '', 'FY']],
    calendarHeaders,
    [[date, '月', true, 'SP']]
  ));
  assertFixtureEqual(plan.insertRows.map(row => row[0]), ['fy', 'sp']);
});

{
  const plan = sandbox.buildRangeLimitedClassSessionPlan_('2027-01-14', '2027-01-14', sources(
    termHeaders,
    [['sp', 'Tue', 1, '', '', 'SP'], ['fy', 'Tue', 1, '', '', 'FY'], ['fa', 'Thu', 1, '', '', 'FA']],
    calendarHeaders,
    [['2027-01-14', '火', true, 'SP']]
  ));
  assertFixtureEqual(plan.insertRows.map(row => row[0]), ['fy', 'sp']);
}

{
  const plan = sandbox.buildRangeLimitedClassSessionPlan_('2026-10-16', '2026-10-16', sources(
    termHeaders,
    [['sp', 'Fri', 1, '', '', 'SP']],
    calendarHeaders,
    [['2026-10-16', 'Fri', false, 'SP']]
  ));
  assert.strictEqual(plan.candidateCount, 0);
}

{
  const invalidCalendar = sandbox.buildRangeLimitedClassSessionPlan_('2026-10-15', '2026-10-15', sources(
    termHeaders, [['sp', 'Mon', 1, '', '', 'SP']], calendarHeaders,
    [['2026-10-15', 'Mon', true, 'BAD']]
  ));
  assert.strictEqual(invalidCalendar.invalidCalendarTermCount, 1);
  assert.ok(invalidCalendar.blockingErrors.some(error => error.includes('calendarに不正なterm')));

  const invalidTimetable = sandbox.buildRangeLimitedClassSessionPlan_('2026-10-15', '2026-10-15', sources(
    termHeaders, [['sp', 'Mon', 1, '', '', 'BAD']], calendarHeaders,
    [['2026-10-15', 'Mon', true, 'SP']]
  ));
  assert.strictEqual(invalidTimetable.invalidTimetableTermCount, 1);
  assert.ok(invalidTimetable.blockingErrors.some(error => error.includes('timetableに空欄または不正なterm')));
}

{
  const collision = sandbox.buildRangeLimitedClassSessionPlan_('2026-10-15', '2026-10-15', sources(
    termHeaders,
    [['same', 'Mon', 1, '', '', 'SP'], ['same', 'Mon', 1, '', '', 'FY']],
    calendarHeaders,
    [['2026-10-15', 'Mon', true, 'SP']]
  ));
  assert.strictEqual(collision.candidateDuplicateCount, 1);
  assert.ok(collision.blockingErrors.some(error => error.includes('candidateにclassId + date + period重複')));
}

{
  const plan = sandbox.buildRangeLimitedClassSessionPlan_('2026-10-15', '2026-10-15', sources(
    termHeaders,
    [['fy', 'Mon', 1, '', '', 'FY'], ['sp', 'Mon', 2, '', '', 'SP']],
    calendarHeaders,
    [['2026-10-15', 'Mon', true, 'SP']],
    [['fy', '2026-09-10', 1, 5]]
  ));
  assertFixtureEqual(plan.insertRows, [['fy', '2026-10-15', 1, 6], ['sp', '2026-10-15', 2, 1]]);
}

{
  const legacy = sandbox.buildRangeLimitedClassSessionPlan_('2026-10-15', '2026-10-15', sources(
    ['classId', 'weekday', 'period', 'teacherName', 'teacherId'],
    [['legacy', 'Mon', 1, '', '']],
    ['date', 'weekday', 'isClassDay'],
    [['2026-10-15', 'Mon', true]]
  ));
  assert.strictEqual(legacy.insertCount, 1);
  assert.strictEqual(legacy.termFallbackCount, 1);
  assertFixtureEqual(legacy.timetableRowsByTerm, { 'legacy-no-column': 1 });
}

console.log('termPlannerFixture.test.js: PASS');
