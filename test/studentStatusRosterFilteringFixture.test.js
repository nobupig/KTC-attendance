const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

function read(rel) {
  return fs.readFileSync(path.join(__dirname, '..', rel), 'utf8');
}

function extractFunction(source, name) {
  const start = source.indexOf('function ' + name + '(');
  assert.notStrictEqual(start, -1, name + ' must exist');
  const bodyStart = source.indexOf('{', start);
  let depth = 0;
  for (let i = bodyStart; i < source.length; i += 1) {
    if (source[i] === '{') depth += 1;
    if (source[i] === '}') depth -= 1;
    if (depth === 0) return source.slice(start, i + 1);
  }
  throw new Error(name + ' body was not closed');
}

const studentService = read('src/studentService.js');
const absenceCalculator = read('src/absenceCalculator.js');
const homeroomShrService = read('src/homeroomShrService.js');
const spreadsheetUtils = read('src/spreadsheetUtils.js');

const activeHelper = extractFunction(studentService, 'isActiveStudentStatus_');
const activeSandbox = {
  normalizeString_: value => String(value == null ? '' : value).trim()
};
vm.createContext(activeSandbox);
vm.runInContext(activeHelper + '\nthis.isActiveStudentStatus_ = isActiveStudentStatus_;', activeSandbox);

['', 'active', 'ACTIVE', '在籍', '有効'].forEach(status => {
  assert.strictEqual(activeSandbox.isActiveStudentStatus_(status), true, status + ' should be active');
});

['withdrawn', 'inactive', 'leave', '休学', '退学', '卒業'].forEach(status => {
  assert.strictEqual(activeSandbox.isActiveStudentStatus_(status), false, status + ' should be inactive');
});

const groupRoster = extractFunction(studentService, 'getStudentsByClassIdAndGroup');
assert.ok(groupRoster.includes("status: findColumnIndex_(studentHeaders, ['status', '在籍状態'])"));
assert.ok(groupRoster.includes('if (!isActiveStudentStatus_(studentStatus)) return;'));

const riskRoster = extractFunction(studentService, 'getStudentRiskMapForTeacherSession');
assert.ok(riskRoster.includes('isActiveStudentStatus_(student.status)'));

const absenceActive = extractFunction(absenceCalculator, 'isStudentActive_');
assert.ok(absenceActive.includes("typeof isActiveStudentStatus_ === 'function'"));

const shrRoster = extractFunction(homeroomShrService, 'getStudentsByHomeroomClass_');
assert.ok(shrRoster.includes('isActiveStudentStatus_(rowStatus)'));
assert.ok(homeroomShrService.includes("studentsByHomeroomClass__statusV2__"));

assert.ok(studentService.includes("studentsByClassId__statusV2__"));
assert.ok(spreadsheetUtils.includes("return baseKey + '__statusV2';"));

console.log('studentStatusRosterFilteringFixture.test.js: PASS');
