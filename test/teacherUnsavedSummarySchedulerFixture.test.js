const assert = require('assert');
const fs = require('fs');
const path = require('path');
const vm = require('vm');

const teacherHtml = fs.readFileSync(
  path.join(__dirname, '..', 'html', 'teacher.html'),
  'utf8'
);

function extractFunction(name) {
  const start = teacherHtml.indexOf('function ' + name + '(');
  assert.notStrictEqual(start, -1, name + ' must exist');

  const bodyStart = teacherHtml.indexOf('{', start);
  let depth = 0;
  for (let index = bodyStart; index < teacherHtml.length; index += 1) {
    if (teacherHtml[index] === '{') depth += 1;
    if (teacherHtml[index] === '}') depth -= 1;
    if (depth === 0) return teacherHtml.slice(start, index + 1);
  }
  throw new Error(name + ' body was not closed');
}

const schedulerSource = [
  extractFunction('loadTeacherUnsavedSummaryManual'),
  extractFunction('loadTeacherUnsavedSummary'),
  extractFunction('scheduleInitialTeacherUnsavedSummaryLoad')
].join('\n');

function createFixture() {
  const timers = [];
  const requests = [];
  const renders = [];
  const scheduleTimer = function(callback, delay) {
    timers.push({ callback, delay });
  };
  const sandbox = {
    console: { error() {} },
    setTimeout: scheduleTimer,
    window: {
      setTimeout: scheduleTimer
    },
    google: {
      script: {
        run: {
          withSuccessHandler(success) {
            return {
              withFailureHandler(failure) {
                return {
                  getTeacherUnsavedSummaryFast() {
                    requests.push({ success, failure });
                  }
                };
              }
            };
          }
        }
      }
    }
  };

  const harness = `
    let teacherUnsavedSummaryRequestSeq = 0;
    let teacherUnsavedDetailsLoadedKey = 'existing';
    let teacherUnsavedSummaryData = null;
    let currentMode = 'past';
    function renderTeacherUnsavedSummaryLoading() { renders.push('loading'); }
    function renderTeacherUnsavedSummary(data) { renders.push(data.status || 'ready'); }
    function renderTeacherUnsavedSummaryError() { renders.push('failed'); }
    function normalizeTeacherUnsavedSummaryFastData(data) { return data; }
    ${schedulerSource}
    globalThis.fixtureApi = {
      scheduleInitialTeacherUnsavedSummaryLoad,
      loadTeacherUnsavedSummary,
      state: function() {
        return {
          seq: teacherUnsavedSummaryRequestSeq,
          data: teacherUnsavedSummaryData,
          detailsKey: teacherUnsavedDetailsLoadedKey
        };
      }
    };
  `;

  sandbox.renders = renders;
  vm.createContext(sandbox);
  vm.runInContext(harness, sandbox, { filename: 'teacher-unsaved-summary-scheduler' });
  return { timers, requests, renders, api: sandbox.fixtureApi };
}

function runNextTimer(fixture, expectedDelay) {
  const timer = fixture.timers.shift();
  assert.ok(timer, 'expected a scheduled timer');
  assert.strictEqual(timer.delay, expectedDelay);
  timer.callback();
}

// CASE 1: initial load must start after seven seconds even if the screen is Past.
{
  const fixture = createFixture();
  fixture.api.scheduleInitialTeacherUnsavedSummaryLoad();
  runNextTimer(fixture, 7000);
  runNextTimer(fixture, 800);
  assert.strictEqual(fixture.requests.length, 1);
  fixture.requests[0].success({ ok: true, status: 'ready', count: 0 });
  assert.strictEqual(fixture.api.state().data.status, 'ready');
  assert.strictEqual(fixture.renders.at(-1), 'ready');
}

// CASE 3/4: stale A success or failure cannot overwrite newer B.
{
  const fixture = createFixture();
  fixture.api.loadTeacherUnsavedSummary();
  fixture.api.loadTeacherUnsavedSummary();
  runNextTimer(fixture, 800);
  runNextTimer(fixture, 800);
  assert.strictEqual(fixture.requests.length, 1);
  fixture.requests[0].success({ ok: true, status: 'ready', count: 2 });
  assert.strictEqual(fixture.api.state().data.count, 2);
}

// Dispatch A before B so both callbacks exist; A must then be stale.
{
  const fixture = createFixture();
  fixture.api.loadTeacherUnsavedSummary();
  runNextTimer(fixture, 800);
  fixture.api.loadTeacherUnsavedSummary();
  runNextTimer(fixture, 800);
  assert.strictEqual(fixture.requests.length, 2);
  fixture.requests[0].success({ ok: true, status: 'ready', count: 1 });
  assert.strictEqual(fixture.api.state().data.loading, true);
  fixture.requests[0].failure(new Error('old request failed'));
  assert.strictEqual(fixture.api.state().data.loading, true);
  fixture.requests[1].success({ ok: true, status: 'ready', count: 2 });
  assert.strictEqual(fixture.api.state().data.count, 2);
}

// CASE 5/6: latest failure exits loading, while all Fast result statuses remain renderable.
{
  const fixture = createFixture();
  fixture.api.loadTeacherUnsavedSummary();
  runNextTimer(fixture, 800);
  fixture.requests[0].failure(new Error('network failure'));
  assert.strictEqual(fixture.api.state().data.status, 'unavailable');
  assert.strictEqual(fixture.renders.at(-1), 'failed');
}

['ready', 'missing', 'stale', 'building', 'unavailable'].forEach(status => {
  const fixture = createFixture();
  fixture.api.loadTeacherUnsavedSummary();
  runNextTimer(fixture, 800);
  fixture.requests[0].success({ ok: status === 'ready', status, count: 0 });
  assert.strictEqual(fixture.api.state().data.status, status);
});

// CASE 7: Past search remains on its own tokens and cannot change summary sequence.
const pastSearchStart = teacherHtml.indexOf('function loadPastClassesByDate(');
const pastSearchEnd = teacherHtml.indexOf('function initButtons()', pastSearchStart);
const pastSearchSource = teacherHtml.slice(pastSearchStart, pastSearchEnd);
assert.ok(!pastSearchSource.includes('teacherUnsavedSummaryRequestSeq'));
assert.ok(!teacherHtml.includes('teacherUnsavedSummaryRequestKey'));
assert.ok(!schedulerSource.includes('currentMode'));

console.log('teacherUnsavedSummarySchedulerFixture.test.js: PASS');
