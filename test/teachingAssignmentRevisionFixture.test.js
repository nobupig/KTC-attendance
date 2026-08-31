const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const root = path.resolve(__dirname, '..');
const properties = {};
const events = [];
let uuidCounter = 0;
let lockHeld = true;

const lock = {
  hasLock() { return lockHeld; },
  tryLock() { lockHeld = true; return true; },
  waitLock() { lockHeld = true; },
  releaseLock() { lockHeld = false; }
};

const sandbox = {
  console,
  CONFIG: {
    SHEETS: {
      TIMETABLE: 'timetable',
      CLASS_TEACHER_TEAMS: 'classTeacherTeams'
    }
  },
  Logger: { log() {} },
  normalizeString_(value) { return value == null ? '' : String(value).trim(); },
  normalizeYmdDisplayText_(value) {
    const match = String(value == null ? '' : value).match(/^(\d{4}-\d{2}-\d{2})/);
    return match ? match[1] : '';
  },
  formatDateToYmd(value) { return value == null ? '' : String(value).slice(0, 10); },
  PropertiesService: {
    getScriptProperties() {
      return {
        getProperty(key) { return properties[key] || null; },
        setProperty(key, value) { properties[key] = String(value); }
      };
    }
  },
  LockService: { getScriptLock() { return lock; } },
  Utilities: {
    getUuid() { uuidCounter += 1; return `revision-${uuidCounter}`; },
    formatDate(value) { return String(value); }
  }
};

vm.createContext(sandbox);
['src/teacherUnsavedCacheService.js', 'src/timetableService.js'].forEach(file => {
  vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), sandbox, { filename: file });
});

const cacheSchema = sandbox.getTeacherUnsavedCacheSchema();
let summaryHeaders = cacheSchema.summary.legacyHeaders.slice();
const detailHeaders = cacheSchema.detail.headers.slice();
function makeHeaderSheet(getHeaders) {
  return {
    getLastColumn() { return getHeaders().length; },
    getLastRow() { return 1; },
    getRange() {
      return {
        getDisplayValues() { return [getHeaders().slice()]; },
        getValues() { return []; }
      };
    }
  };
}
const summarySheet = makeHeaderSheet(() => summaryHeaders);
const detailSheet = makeHeaderSheet(() => detailHeaders);
const operationSpreadsheet = {
  getSheetByName(name) {
    if (name === 'teacherUnsavedSummaryCache') return summarySheet;
    if (name === 'teacherUnsavedDetailCache') return detailSheet;
    return null;
  }
};
sandbox.getOperationSpreadsheet = () => operationSpreadsheet;

let validation = sandbox.validateTeacherUnsavedCacheSheets_();
assert.strictEqual(validation.ok, true);
assert.strictEqual(validation.summary.schemaMode, 'legacy');
summaryHeaders = cacheSchema.summary.revisionedHeaders.slice();
validation = sandbox.validateTeacherUnsavedCacheSheets_();
assert.strictEqual(validation.ok, true);
assert.strictEqual(validation.summary.schemaMode, 'revisioned');
summaryHeaders = ['invalid'];
assert.strictEqual(sandbox.validateTeacherUnsavedCacheSheets_().ok, false);
summaryHeaders = cacheSchema.summary.legacyHeaders.slice();

assert.strictEqual(sandbox.getTeachingAssignmentRevision_(), '');
lockHeld = false;
assert.throws(
  () => sandbox.bumpTeachingAssignmentRevisionUnderLock_(),
  error => error.message.includes('ScriptLock')
);
lockHeld = true;
const contextKeyBefore = sandbox.buildTeacherUnsavedContextCacheKey_('T1');
const summaryKeyBefore = sandbox.buildTeacherUnsavedSummaryCacheKey_('T1', '2026-10-01');
const firstRevision = sandbox.bumpTeachingAssignmentRevisionUnderLock_();
assert.strictEqual(firstRevision, 'revision-1');
assert.strictEqual(sandbox.getTeachingAssignmentRevision_(), firstRevision);
assert.notStrictEqual(sandbox.buildTeacherUnsavedContextCacheKey_('T1'), contextKeyBefore);
assert.notStrictEqual(sandbox.buildTeacherUnsavedSummaryCacheKey_('T1', '2026-10-01'), summaryKeyBefore);

const summaryRow = [
  'snapshot-1', '2026-10-02', 'T1', 'Teacher', 'teacher@example.com',
  '2026-04-07', '2026-10-01', 0, '', '', '', 0, new Date(), 'ready', '', firstRevision
];
const summary = sandbox.buildTeacherUnsavedSummaryObject_(summaryRow);
assert.strictEqual(summary.teachingAssignmentRevision, firstRevision);
assert.strictEqual(
  sandbox.buildTeacherUnsavedSummaryRowForSchema_(summaryRow.slice(0, 15), 'legacy', '').length,
  15
);
const revisionedSummaryRow = sandbox.buildTeacherUnsavedSummaryRowForSchema_(
  summaryRow.slice(0, 15),
  'revisioned',
  firstRevision
);
assert.strictEqual(revisionedSummaryRow.length, 16);
assert.strictEqual(revisionedSummaryRow[15], firstRevision);
assert.strictEqual(
  sandbox.getTeacherUnsavedCacheSchema().summary.headers.includes('teachingAssignmentRevision'),
  true
);

assert.strictEqual(
  sandbox.getTeacherUnsavedTeachingAssignmentRevisionState_(firstRevision, firstRevision, 'revisioned').status,
  'ready'
);
assert.strictEqual(
  sandbox.getTeacherUnsavedTeachingAssignmentRevisionState_(firstRevision, 'revision-new', 'revisioned').status,
  'stale'
);
assert.strictEqual(
  sandbox.getTeacherUnsavedTeachingAssignmentRevisionState_('', '', 'legacy').status,
  'ready'
);
assert.strictEqual(
  sandbox.getTeacherUnsavedTeachingAssignmentRevisionState_('', firstRevision, 'legacy').status,
  'stale'
);
assert.strictEqual(
  sandbox.getTeacherUnsavedTeachingAssignmentRevisionState_('', '', 'revisioned').status,
  'stale'
);
assert.strictEqual(
  sandbox.getTeacherUnsavedTeachingAssignmentRevisionState_('', firstRevision, 'revisioned').status,
  'stale'
);
assert.strictEqual(
  sandbox.assertTeacherUnsavedTeachingAssignmentRevisionCurrent_(firstRevision, firstRevision, 'revisioned'),
  true
);
assert.throws(
  () => sandbox.assertTeacherUnsavedTeachingAssignmentRevisionCurrent_(firstRevision, 'revision-new', 'revisioned'),
  error => error.reason === 'teaching-assignment-revision-changed' && error.retryable === true
);

const originalBump = sandbox.bumpTeachingAssignmentRevisionUnderLock_;
sandbox.removeScriptCacheKeys_ = keys => {
  events.push('source-cache-invalidation');
  assert.ok(keys.includes('sheetData__OPERATION__timetable'));
  assert.ok(keys.includes('sheetData__OPERATION__classTeacherTeams'));
  assert.ok(keys.includes('classTeacherTeamRows__all'));
};
sandbox.bumpTeachingAssignmentRevisionUnderLock_ = function() {
  events.push('revision-bump');
  return originalBump();
};
sandbox.invalidateAllTeacherUnsavedFastSnapshotsUnderLock_ = message => {
  events.push('fast-snapshot-stale');
  assert.ok(message.includes('timetable/classTeacherTeams'));
  return { ok: true, invalidatedSummaryCount: 2 };
};

delete properties.teachingAssignmentRevision;
events.length = 0;
summaryHeaders = cacheSchema.summary.legacyHeaders.slice();
lockHeld = false;
const legacyMaintenanceResult = sandbox.notifyTeachingAssignmentDataChanged();
assert.strictEqual(legacyMaintenanceResult.requiresSummarySchemaMigration, true);
assert.strictEqual(legacyMaintenanceResult.summarySchemaMode, 'legacy');
assert.strictEqual(sandbox.getTeachingAssignmentRevision_(), '');
assert.deepStrictEqual(events, []);
assert.strictEqual(lockHeld, false);

properties.teachingAssignmentRevision = firstRevision;
events.length = 0;
lockHeld = true;
const legacyInitializedMaintenance = sandbox.notifyTeachingAssignmentDataChangedUnderLock_();
assert.strictEqual(legacyInitializedMaintenance.requiresSummarySchemaMigration, true);
assert.strictEqual(sandbox.getTeachingAssignmentRevision_(), firstRevision);
assert.deepStrictEqual(events, []);
delete properties.teachingAssignmentRevision;

const publishedColumnCounts = [];
sandbox.SpreadsheetApp = { flush() {} };
sandbox.getTeacherUnsavedAttendanceSessionsRowCount_ = () => 0;
sandbox.markTeacherUnsavedSummaryRowsStatus_ = () => {};
sandbox.replaceTeacherUnsavedCacheRows_ = (sheet, rows, columnCount) => {
  publishedColumnCounts.push({ sheet, rows, columnCount });
};
sandbox.publishTeacherUnsavedCacheSnapshot_({
  summarySchemaMode: 'legacy',
  teachingAssignmentRevision: '',
  calendarRevision: '',
  classSessionsRevision: '',
  attendanceSessionsRowCount: 0,
  summaryRows: [summaryRow.slice(0, 15)],
  detailRows: []
});
assert.strictEqual(
  publishedColumnCounts.find(item => item.sheet === summarySheet).columnCount,
  15
);

summaryHeaders = cacheSchema.summary.revisionedHeaders.slice();
lockHeld = true;
const maintenanceResult = sandbox.notifyTeachingAssignmentDataChangedUnderLock_();
assert.deepStrictEqual(events, [
  'source-cache-invalidation',
  'revision-bump',
  'fast-snapshot-stale'
]);
assert.strictEqual(maintenanceResult.teachingAssignmentRevision, 'revision-2');
assert.strictEqual(sandbox.getTeachingAssignmentRevision_(), 'revision-2');

events.length = 0;
lockHeld = false;
const publicMaintenanceResult = sandbox.notifyTeachingAssignmentDataChanged();
assert.strictEqual(publicMaintenanceResult.teachingAssignmentRevision, 'revision-3');
assert.strictEqual(lockHeld, false);
assert.deepStrictEqual(events, [
  'source-cache-invalidation',
  'revision-bump',
  'fast-snapshot-stale'
]);

publishedColumnCounts.length = 0;
sandbox.publishTeacherUnsavedCacheSnapshot_({
  summarySchemaMode: 'revisioned',
  teachingAssignmentRevision: 'revision-3',
  calendarRevision: '',
  classSessionsRevision: '',
  attendanceSessionsRowCount: 0,
  summaryRows: [summaryRow.slice(0, 15).concat(['revision-3'])],
  detailRows: []
});
assert.strictEqual(
  publishedColumnCounts.find(item => item.sheet === summarySheet).columnCount,
  16
);
assert.throws(
  () => sandbox.publishTeacherUnsavedCacheSnapshot_({
    summarySchemaMode: 'revisioned',
    teachingAssignmentRevision: 'revision-2',
    calendarRevision: '',
    classSessionsRevision: '',
    attendanceSessionsRowCount: 0,
    summaryRows: [],
    detailRows: []
  }),
  error => error.reason === 'teaching-assignment-revision-changed'
);

const matchingDetailsRevision = sandbox.buildTeacherUnsavedFastDetailsRevisionFailureIfAny_(
  { teachingAssignmentRevision: 'revision-3', summarySchemaMode: 'revisioned' },
  { teachingAssignmentRevision: 'revision-3' },
  { limit: 10, offset: 0 }
);
assert.strictEqual(matchingDetailsRevision, null);
const staleDetailsRevision = sandbox.buildTeacherUnsavedFastDetailsRevisionFailureIfAny_(
  { teachingAssignmentRevision: 'revision-1', summarySchemaMode: 'revisioned' },
  { teachingAssignmentRevision: 'revision-1' },
  { limit: 10, offset: 0 }
);
assert.strictEqual(staleDetailsRevision.status, 'stale');

let fastSchemaMode = 'legacy';
let currentSummaryRow = Object.assign({}, summary, {
  cacheDate: '2026-10-02',
  status: 'ready',
  teachingAssignmentRevision: ''
});
sandbox.getCurrentUserContext = () => ({ teacherId: 'T1' });
sandbox.getTeacherUnsavedCacheDateContext_ = () => ({ cacheDate: '2026-10-02' });
sandbox.validateTeacherUnsavedCacheSheets_ = () => ({
  ok: true,
  summary: { sheet: {}, schemaMode: fastSchemaMode },
  detail: { sheet: {} }
});
sandbox.readTeacherUnsavedSummaryState_ = () => ({ row: currentSummaryRow });
sandbox.validateTeacherUnsavedSummaryRowIntegrity_ = () => ({ ok: true });

delete properties.teachingAssignmentRevision;
assert.strictEqual(sandbox.getTeacherUnsavedSummaryFastContext_().result.status, 'ready');
properties.teachingAssignmentRevision = 'revision-3';
assert.strictEqual(sandbox.getTeacherUnsavedSummaryFastContext_().result.status, 'stale');

fastSchemaMode = 'revisioned';
delete properties.teachingAssignmentRevision;
assert.strictEqual(sandbox.getTeacherUnsavedSummaryFastContext_().result.status, 'stale');
properties.teachingAssignmentRevision = 'revision-3';
assert.strictEqual(sandbox.getTeacherUnsavedSummaryFastContext_().result.status, 'stale');
currentSummaryRow = Object.assign({}, currentSummaryRow, {
  teachingAssignmentRevision: 'revision-3'
});
assert.strictEqual(sandbox.getTeacherUnsavedSummaryFastContext_().result.status, 'ready');
currentSummaryRow = Object.assign({}, currentSummaryRow, {
  teachingAssignmentRevision: 'revision-1'
});
const staleRead = sandbox.getTeacherUnsavedSummaryFastContext_().result;
assert.strictEqual(staleRead.ok, false);
assert.strictEqual(staleRead.status, 'stale');
assert.ok(staleRead.errorMessage.includes('利用可能条件'));

vm.runInContext(fs.readFileSync(path.join(root, 'src/attendanceService.js'), 'utf8'), sandbox, {
  filename: 'src/attendanceService.js'
});
sandbox.invalidateTeacherUnsavedFastSnapshotAfterSaveUnderLock_ = () => {
  throw new Error('fixture invalidation failure');
};
const hotfixResult = sandbox.tryInvalidateTeacherUnsavedFastSnapshotAfterSaveUnderLock_(
  'class-1', '2026-10-01', 1, 'save'
);
assert.strictEqual(hotfixResult.ok, false);

console.log('teachingAssignmentRevisionFixture.test.js: PASS');
