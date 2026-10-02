const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const appSource = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const parserSource = appSource.slice(
  appSource.indexOf('function parseMixedScheduleRows'),
  appSource.indexOf('function validateMixedRows')
);
const templateEmployeeSource = appSource.slice(
  appSource.indexOf('function isTemplateEmployeeNumber'),
  appSource.indexOf('function parseMixedScheduleRows')
);
const normalizerSource = appSource.slice(
  appSource.indexOf('function normalizeShiftCode'),
  appSource.indexOf('function getRestDayName')
);
const dateSource = appSource.slice(
  appSource.indexOf('function excelDateToJS'),
  appSource.indexOf('function normalizeDateForExport')
);
const context = {};

vm.runInNewContext(
  `${templateEmployeeSource}\n${parserSource}\n${normalizerSource}\n${dateSource}\nthis.parseMixedScheduleRows = parseMixedScheduleRows;`,
  context
);

const { parseMixedScheduleRows } = context;
const reusedWorkHeaders = [
  'EMPLOYEE NO.',
  'NAME',
  'WORK DATE',
  'SHIFT CODE',
  'Day of Week',
  'POSITION'
];

// Regression: RD - LIP in the uploaded workbook uses work-oriented labels, but
// its blank shift-code cells and sheet context identify these as rest-day rows.
const restRows = parseMixedScheduleRows([
  reusedWorkHeaders,
  ['878', 'Aaron Jaycob Malaluan', '46296', '', 'Thursday', 'Inventory Analyst'],
  ['878', 'Aaron Jaycob Malaluan', '46302', '', 'Wednesday', 'Inventory Analyst']
], null, 'RD - LIP');

assert.equal(restRows.length, 2);
assert.deepEqual(Array.from(restRows, row => row.type), ['rest', 'rest']);
assert.deepEqual(Array.from(restRows, row => row.date), ['10/1/2026', '10/7/2026']);

// The same labels on the actual WS sheet must continue to produce work rows.
const workRows = parseMixedScheduleRows([
  reusedWorkHeaders,
  ['878', 'Aaron Jaycob Malaluan', '46297', 'RBT-032', 'Friday', 'Inventory Analyst']
], null, 'WS - LIP');

assert.equal(workRows.length, 1);
assert.equal(workRows[0].type, 'work');
assert.equal(workRows[0].shiftCode, 'RBT-032');
