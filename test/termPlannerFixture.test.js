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
['src/calendarService.js', 'src/classSessionService.js', 'src/teacherService.js'].forEach(file => {
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

{
  const timetable = {
    headers: ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'term'],
    rows: [
      ['faOnly', 'Mon', 1, 'T_FA', 'FA Teacher', 'FA'],
      ['spOnly', 'Mon', 1, 'T_SP', 'SP Teacher', 'SP'],
      ['yearRound', 'Mon', 1, 'T_FY', 'FY Teacher', 'FY']
    ]
  };
  const teams = {
    headers: ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'roleType', 'term'],
    rows: [
      ['spOnly', 'Mon', 1, 'T_SUPPORT', 'Support', 'support', 'SP'],
      ['orphan', 'Mon', 1, 'T_ORPHAN', 'Orphan', 'support', 'SP']
    ]
  };
  const index = sandbox.buildTeachingAssignmentIndex_(timetable, teams);
  const faContext = sandbox.getEffectiveClassDayContext_('2026-04-13', {
    '2026-04-13': { weekday: 'Mon', isClassDay: true, term: 'FA' }
  });
  const spContext = sandbox.getEffectiveClassDayContext_('2026-10-15', {
    '2026-10-15': { weekday: 'Mon', isClassDay: true, term: 'SP' }
  });
  assert.strictEqual(sandbox.getTeachingAssignmentForSessionFromIndex_(index, 'spOnly', '2026-04-13', 1, faContext), null);
  assert.deepStrictEqual(
    JSON.parse(JSON.stringify(sandbox.getTeachingAssignmentForSessionFromIndex_(index, 'spOnly', '2026-10-15', 1, spContext).teacherIds)),
    ['T_SP', 'T_SUPPORT']
  );
  assert.strictEqual(sandbox.getTeachingAssignmentForSessionFromIndex_(index, 'yearRound', '2026-04-13', 1, faContext).teacherId, 'T_FY');
  assert.ok(index.warnings.some(warning => warning.includes('対応するtimetable行がありません')));

  const strict = sandbox.buildTeachingAssignmentIndex_({
    headers: timetable.headers,
    rows: [['bad', 'Mon', 1, 'T_BAD', 'Bad', '']]
  }, teams);
  assert.strictEqual(Object.keys(strict.byKey).length, 0);
  assert.ok(strict.warnings.some(warning => warning.includes('termが空欄または不正')));
}

{
  const timetable = {
    headers: ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'term'],
    rows: [['canonical', 'Mon', '01', 'T_CANON', 'Canonical', 'FA']]
  };
  const teams = { headers: ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'roleType', 'term'], rows: [] };
  const index = sandbox.buildTeachingAssignmentIndex_(timetable, teams);
  const context = sandbox.getEffectiveClassDayContext_('2026-04-13', {
    '2026-04-13': { weekday: 'Mon', isClassDay: true, term: 'FA' }
  });
  assert.strictEqual(sandbox.normalizeTeachingAssignmentPeriod_('001'), '1');
  assert.strictEqual(sandbox.normalizeTeachingAssignmentPeriod_(1), '1');
  [0, -1, '1.5', '', 'x'].forEach(value => assert.strictEqual(sandbox.normalizeTeachingAssignmentPeriod_(value), ''));
  assert.strictEqual(
    sandbox.getTeachingAssignmentForSessionFromIndex_(index, 'canonical', '2026-04-13', 1, context).teacherId,
    'T_CANON'
  );
}

{
  let teacherBundleLoadCount = 0;
  sandbox.getTeacherMasterBundle_ = () => {
    teacherBundleLoadCount += 1;
    return {
      byId: {
        T_FA: { teacherId: 'T_FA', name: 'Normalized FA', email: 'fa@example.com', roles: ['teacher'] },
        T_FY: { teacherId: 'T_FY', name: 'Normalized FY', email: 'fy@example.com', roles: ['teacher'] },
        T_SUPPORT: { teacherId: 'T_SUPPORT', name: 'Normalized Support', email: 'support@example.com', roles: ['teacher'] }
      },
      byName: {
        'Fallback Teacher': { teacherId: 'T_SP', name: 'Normalized SP', email: 'sp@example.com', roles: ['teacher'] }
      }
    };
  };
  const resolver = sandbox.createTeacherTeamMemberResolver_();
  const index = sandbox.buildTeachingAssignmentIndex_({
    headers: ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'term'],
    rows: [
      ['faResolver', 'Mon', 1, 'T_FA', 'Raw FA', 'FA'],
      ['spResolver', 'Mon', '01', '', 'Fallback Teacher', 'SP'],
      ['fyResolver', 'Mon', 1, 'T_FY', 'Raw FY', 'FY']
    ]
  }, {
    headers: ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'roleType', 'term'],
    rows: [
      ['faResolver', 'Mon', 1, 'T_SUPPORT', 'Raw Support', 'support', 'FA'],
      ['spResolver', 'Mon', 1, 'T_SUPPORT', 'Raw Support', 'support', 'SP']
    ]
  }, resolver);
  assert.strictEqual(teacherBundleLoadCount, 1);
  assert.strictEqual(resolver.teacherBundleLoadCount, 1);

  const faContext = sandbox.getEffectiveClassDayContext_('2026-04-13', {
    '2026-04-13': { weekday: 'Mon', isClassDay: true, term: 'FA' }
  });
  const spContext = sandbox.getEffectiveClassDayContext_('2026-10-15', {
    '2026-10-15': { weekday: 'Mon', isClassDay: true, term: 'SP' }
  });
  const fa = sandbox.getTeachingAssignmentForSessionFromIndex_(index, 'faResolver', '2026-04-13', 1, faContext);
  const sp = sandbox.getTeachingAssignmentForSessionFromIndex_(index, 'spResolver', '2026-10-15', 1, spContext);
  const fy = sandbox.getTeachingAssignmentForSessionFromIndex_(index, 'fyResolver', '2026-04-13', 1, faContext);
  assert.strictEqual(fa.teacherName, 'Normalized FA');
  assert.deepStrictEqual(JSON.parse(JSON.stringify(fa.teacherIds)), ['T_FA', 'T_SUPPORT']);
  assert.strictEqual(sp.teacherId, 'T_SP');
  assert.strictEqual(sp.teacherName, 'Normalized SP');
  assert.strictEqual(fy.teacherId, 'T_FY');
}

{
  const headers = ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'term'];
  const teamHeaders = ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'roleType', 'term'];
  const index = sandbox.buildTeachingAssignmentIndex_({
    headers,
    rows: [
      ['duplicate', 'Mon', '1', 'T_A', 'Teacher A', 'SP'],
      ['duplicate', 'Mon', '01', 'T_B', 'Teacher B', 'SP'],
      ['unrelated', 'Mon', 1, 'T_OK', 'Teacher OK', 'SP']
    ]
  }, {
    headers: teamHeaders,
    rows: [['duplicate', 'Mon', 1, 'T_SUPPORT', 'Support', 'support', 'SP']]
  });
  const context = sandbox.getEffectiveClassDayContext_('2026-10-15', {
    '2026-10-15': { weekday: 'Mon', isClassDay: true, term: 'SP' }
  });
  assert.strictEqual(sandbox.getTeachingAssignmentForSessionFromIndex_(index, 'duplicate', '2026-10-15', 1, context), null);
  assert.strictEqual(
    sandbox.getTeachingAssignmentForSessionFromIndex_(index, 'unrelated', '2026-10-15', 1, context).teacherId,
    'T_OK'
  );
  assert.strictEqual(Object.keys(index.duplicateKeys).length, 1);
  assert.ok(index.warnings.some(warning => warning.includes('duplicate assignment key')));
  assert.ok(index.warnings.some(warning => warning.includes('duplicateで無効')));
}

{
  const legacyHeaders = ['classId', 'weekday', 'period', 'teacherId', 'teacherName'];
  const legacyTeamHeaders = ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'roleType'];
  const legacy = sandbox.buildTeachingAssignmentIndex_({
    headers: legacyHeaders,
    rows: [['legacy', 'Mon', 1, 'T_A', 'Teacher A'], ['legacy', 'Mon', '01', 'T_B', 'Teacher B']]
  }, {
    headers: legacyTeamHeaders,
    rows: [['legacy', 'Mon', 1, 'T_SUPPORT', 'Support', 'support']]
  });
  const context = sandbox.getEffectiveClassDayContext_('2026-10-15', {
    '2026-10-15': { weekday: 'Mon', isClassDay: true, term: 'SP' }
  });
  assert.strictEqual(sandbox.getTeachingAssignmentForSessionFromIndex_(legacy, 'legacy', '2026-10-15', 1, context), null);
  assert.strictEqual(Object.keys(legacy.duplicateKeys).length, 1);

  const termful = sandbox.buildTeachingAssignmentIndex_({
    headers: ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'term'],
    rows: [['termful', 'Mon', 1, 'T_MAIN', 'Main', 'FA']]
  }, {
    headers: legacyTeamHeaders,
    rows: [['termful', 'Mon', 1, 'T_BAD_TEAM', 'Bad Team', 'support']]
  });
  const faContext = sandbox.getEffectiveClassDayContext_('2026-04-13', {
    '2026-04-13': { weekday: 'Mon', isClassDay: true, term: 'FA' }
  });
  assert.deepStrictEqual(
    JSON.parse(JSON.stringify(sandbox.getTeachingAssignmentForSessionFromIndex_(termful, 'termful', '2026-04-13', 1, faContext).teacherIds)),
    ['T_MAIN']
  );
  assert.ok(termful.warnings.some(warning => warning.includes('term契約がtimetableと一致しません')));
}

{
  sandbox.CONFIG = {
    SHEETS: { TIMETABLE: 'timetable', CLASS_TEACHER_TEAMS: 'classTeacherTeams', CALENDAR: 'calendar' },
    APP: { BASE_URL: 'https://example.invalid' }
  };
  sandbox.getTeacherRecordById_ = () => null;
  sandbox.getTeacherRecordByName_ = () => null;
  const cachedSources = {
    timetable: { headers: ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'term'], rows: [
      ['a', 'Mon', 1, 'T_A', 'Teacher A', 'FA'], ['b', 'Mon', 1, 'T_B', 'Teacher B', 'FA']
    ] },
    classTeacherTeams: { headers: ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'roleType', 'term'], rows: [] },
    calendar: { headers: ['date', 'weekday', 'isClassDay', 'term'], rows: [
      ['2026-04-13', 'Mon', true, 'FA']
    ], dateDisplayValues: ['2026-04-13'] }
  };
  sandbox.getSheetDataCached_ = (_scope, sheetName) => cachedSources[sheetName];
  vm.runInContext(fs.readFileSync(path.join(root, 'src/attendanceCheckService.js'), 'utf8'), sandbox, { filename: 'src/attendanceCheckService.js' });
  const originalBuildIndex = sandbox.buildTeachingAssignmentIndex_;
  let assignmentIndexBuildCount = 0;
  sandbox.buildTeachingAssignmentIndex_ = function() {
    assignmentIndexBuildCount += 1;
    return originalBuildIndex.apply(this, arguments);
  };
  sandbox.getUnsubmittedClasses = () => [
    { classId: 'a', date: '2026-04-13', period: 1 },
    { classId: 'b', date: '2026-04-13', period: 1 }
  ];
  sandbox.getClassDisplayName = classId => classId;
  sandbox.sendSlackMessage = () => {};
  sandbox.recordMissingAttendanceLog = () => {};
  sandbox.checkYesterdayAttendance();
  assert.strictEqual(assignmentIndexBuildCount, 1);
}

{
  const permissionSources = {
    timetable: { headers: ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'term'], rows: [
      ['switch', 'Mon', 1, 'T_FA', 'FA Teacher', 'FA'],
      ['switch', 'Mon', 1, 'T_SP', 'SP Teacher', 'SP'],
      ['both', 'Mon', 1, 'T_FY', 'FY Teacher', 'FY']
    ] },
    classTeacherTeams: { headers: ['classId', 'weekday', 'period', 'teacherId', 'teacherName', 'roleType', 'term'], rows: [] },
    calendar: { headers: ['date', 'weekday', 'isClassDay', 'term'], rows: [
      ['2026-04-13', 'Mon', true, 'FA'], ['2026-10-15', 'Mon', true, 'SP']
    ], dateDisplayValues: ['2026-04-13', '2026-10-15'] }
  };
  sandbox.getSheetDataCached_ = (_scope, sheetName) => permissionSources[sheetName];
  vm.runInContext(fs.readFileSync(path.join(root, 'src/permissionService.js'), 'utf8'), sandbox, { filename: 'src/permissionService.js' });
  sandbox.hasAnyTeachingAssignmentByTeacherId_ = () => true;
  sandbox.getCurrentUserContext = () => ({ teacherId: 'T_FA', roles: ['teacher'] });
  assert.strictEqual(sandbox.canEditAttendance({ classId: 'switch', date: '2026-04-13', period: 1 }), true);
  assert.strictEqual(sandbox.canEditAttendance({ classId: 'switch', date: '2026-10-15', period: 1 }), false);
  sandbox.getCurrentUserContext = () => ({ teacherId: 'T_SP', roles: ['teacher'] });
  assert.strictEqual(sandbox.canEditAttendance({ classId: 'switch', date: '2026-04-13', period: 1 }), false);
  assert.strictEqual(sandbox.canEditAttendance({ classId: 'switch', date: '2026-10-15', period: 1 }), true);
  sandbox.getCurrentUserContext = () => ({ teacherId: 'T_FY', roles: ['teacher'] });
  assert.strictEqual(sandbox.canEditAttendance({ classId: 'both', date: '2026-04-13', period: 1 }), true);
  assert.strictEqual(sandbox.canEditAttendance({ classId: 'both', date: '2026-10-15', period: 1 }), true);
}

console.log('termPlannerFixture.test.js: PASS');
