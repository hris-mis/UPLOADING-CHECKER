const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const appSource = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const validationSource = appSource.slice(
  appSource.indexOf('function normalizeEmployeeNo'),
  appSource.indexOf('function getImportedPreviewConflicts')
);
const conflictGroupingSource = appSource.slice(
  appSource.indexOf('function getScheduleConflictType'),
  appSource.indexOf('function renderImportSummaryDashboard')
);
const context = {};

vm.runInNewContext(
  `${validationSource}\n${conflictGroupingSource}\nthis.getRestDayViolations = getRestDayViolations; this.getMissingScheduleRecords = getMissingScheduleRecords; this.groupScheduleConflictsByEmployee = groupScheduleConflictsByEmployee;`,
  context
);

const { getRestDayViolations, getMissingScheduleRecords, groupScheduleConflictsByEmployee } = context;
const employeeNo = '1001';

function row(date, extra = {}) {
  return { employeeNo, date, position: 'Cashier', ...extra };
}

function buildWeek(dates, restDates) {
  const restDateSet = new Set(restDates);
  return {
    workRows: dates.filter(date => !restDateSet.has(date)).map(date => row(date, { shiftCode: 'RBT-001' })),
    restRows: restDates.map(date => row(date))
  };
}

function weeklyViolations(schedule, validateOperationWeekends = false) {
  return getRestDayViolations(
    schedule.workRows,
    schedule.restRows,
    validateOperationWeekends
  ).filter(violation => violation.type === 'weeklyRestDays');
}

const crossingMonthDates = [
  '2026-10-26',
  '2026-10-27',
  '2026-10-28',
  '2026-10-29',
  '2026-10-30',
  '2026-10-31',
  '2026-11-01'
];

// Test 1: a boundary week cannot be judged when Sunday is outside the import.
const octoberOnly = buildWeek(crossingMonthDates.slice(0, 6), ['2026-10-31']);
assert.equal(weeklyViolations(octoberOnly, true).length, 0);
assert.equal(weeklyViolations(octoberOnly, false).length, 0);

// Test 2: Saturday and Sunday remain in one Monday-Sunday week across months.
const twoBoundaryRestDays = buildWeek(crossingMonthDates, ['2026-10-31', '2026-11-01']);
assert.equal(weeklyViolations(twoBoundaryRestDays).length, 0);

// Test 3: a fully covered cross-month week with only one RD is still invalid.
const oneBoundaryRestDay = buildWeek(crossingMonthDates, ['2026-10-31']);
const boundaryViolations = weeklyViolations(oneBoundaryRestDay);
assert.equal(boundaryViolations.length, 1);
assert.match(boundaryViolations[0].reason, /Only 1 of 2 required rest days/);

// Test 4: complete weeks contained by one month retain the existing rule.
const withinMonth = buildWeek([
  '2026-10-05',
  '2026-10-06',
  '2026-10-07',
  '2026-10-08',
  '2026-10-09',
  '2026-10-10',
  '2026-10-11'
], ['2026-10-10']);
assert.equal(weeklyViolations(withinMonth).length, 1);

// Test 5: the 15/16 cut-off does not split a calendar week.
const crossingCutoff = buildWeek([
  '2026-10-12',
  '2026-10-13',
  '2026-10-14',
  '2026-10-15',
  '2026-10-16',
  '2026-10-17',
  '2026-10-18'
], ['2026-10-13', '2026-10-17']);
assert.equal(weeklyViolations(crossingCutoff).length, 0);

// Test 6: a WS-only employee receives one missing-RD issue, not weekly 0-of-2 issues.
const twoWorkWeeks = [
  ...Array.from({ length: 7 }, (_, index) => `2026-10-${String(index + 5).padStart(2, '0')}`),
  ...Array.from({ length: 7 }, (_, index) => `2026-10-${String(index + 12).padStart(2, '0')}`)
].map(date => row(date));
const missingRestIssues = getMissingScheduleRecords(twoWorkWeeks, []);
assert.equal(missingRestIssues.length, 1);
assert.equal(missingRestIssues[0].type, 'missingRest');
assert.equal(
  missingRestIssues[0].reason,
  'Employee 1001 — No Rest Day Record\nCheck RD entry or Employee Number.'
);
assert.equal(getRestDayViolations(twoWorkWeeks, [], true).length, 0);
assert.equal(getRestDayViolations(twoWorkWeeks, [], false).length, 0);

// Test 7: an RD-only employee receives one missing-WS issue despite multiple rows.
const restOnlyRows = ['2026-10-05', '2026-10-06', '2026-10-12'].map(date => row(date));
const missingWorkIssues = getMissingScheduleRecords([], restOnlyRows);
assert.equal(missingWorkIssues.length, 1);
assert.equal(missingWorkIssues[0].type, 'missingWork');
assert.equal(
  missingWorkIssues[0].reason,
  'Employee 1001 — No Work Schedule Record\nCheck WS entry or Employee Number.'
);

// Test 8: matching is based only on normalized employee number, never name.
const sameNameDifferentNumbers = getMissingScheduleRecords(
  [row('2026-10-05', { employeeNo: '001001', name: 'Same Name' })],
  [row('2026-10-06', { employeeNo: '2002', name: 'Same Name' })]
);
assert.equal(sameNameDifferentNumbers.length, 2);

// Test 9: once an employee exists in both datasets, normal weekly validation remains.
assert.equal(weeklyViolations(withinMonth).length, 1);

function weekendViolations(dates, extra = {}) {
  return getRestDayViolations([], dates.map(date => row(date, extra)), true)
    .filter(violation => violation.type === 'weekend');
}

// Test 10: weekend allowances reset between the first and second cut-offs.
assert.equal(weekendViolations(['09/13/2026', '09/27/2026']).length, 0);

// Test 11: two weekend Rest Days in cut-off 1 conflict and identify both dates.
const firstCutoffWeekend = weekendViolations(['09/06/2026', '09/13/2026']);
assert.equal(firstCutoffWeekend.length, 1);
assert.deepEqual(Array.from(firstCutoffWeekend[0].rows, item => item.date), ['09/06/2026', '09/13/2026']);

// Test 12: two weekend Rest Days in cut-off 2 conflict and identify both dates.
const secondCutoffWeekend = weekendViolations(['09/20/2026', '09/27/2026']);
assert.equal(secondCutoffWeekend.length, 1);
assert.deepEqual(Array.from(secondCutoffWeekend[0].rows, item => item.date), ['09/20/2026', '09/27/2026']);

// Test 13: different reasons for one normalized employee share one employee group.
const oneEmployee = groupScheduleConflictsByEmployee([
  {
    employeeNo: '2619.0',
    reason: 'Employee 2619 — No Work Schedule Record\nCheck WS entry or Employee Number.'
  },
  {
    employeeNo: '2619',
    date: '09/13/2026',
    reason: 'Weekend RD Limit — Only one Saturday or Sunday rest day is allowed per cut-off for this position.'
  }
]);
assert.deepEqual(Object.keys(oneEmployee), ['2619']);
assert.deepEqual(Object.keys(oneEmployee['2619'].groups), ['missing-work', 'weekend']);
assert.equal(oneEmployee['2619'].conflictCount, 2);
assert.deepEqual(Array.from(oneEmployee['2619'].groups['missing-work'].dates), []);
assert.deepEqual(Array.from(oneEmployee['2619'].groups['missing-work'].reasons), ['Check WS entry or Employee Number.']);
assert.deepEqual(Array.from(oneEmployee['2619'].groups.weekend.dates), ['09/13/2026']);

// Test 14: conflicts for different employees remain in separate employee groups.
const twoEmployees = groupScheduleConflictsByEmployee([
  { employeeNo: '2619', reason: 'Employee 2619 has duplicate date in Rest Day Schedule: 09/13/2026' },
  { employeeNo: '3001', reason: 'Employee 3001 has duplicate date in Rest Day Schedule: 09/13/2026' }
]);
assert.deepEqual(Object.keys(twoEmployees), ['2619', '3001']);

console.log('Rest Day validation and conflict grouping tests passed.');
