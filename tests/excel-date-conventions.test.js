const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const appSource = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const importerSource = appSource.slice(
  appSource.indexOf('function detectDateConventionsFromRows'),
  appSource.indexOf('function validateMixedRows')
);
const normalizerSource = appSource.slice(
  appSource.indexOf('function normalizeShiftCode'),
  appSource.indexOf('function getRestDayName')
);
const dateSource = appSource.slice(
  appSource.indexOf('function excelDateToJS'),
  appSource.indexOf('function normalizeDateForExport')
);
const conflictSource = appSource.slice(
  appSource.indexOf('function normalizeEmployeeNo'),
  appSource.indexOf('function getScheduleConflictType')
);
const context = { importedFiles: [], IS_SUPPORT_GROUP: true };

vm.runInNewContext(
  `${importerSource}\n${normalizerSource}\n${dateSource}\n${conflictSource}\n` +
  'this.api = { detectDateConventionsFromRows, excludeTemplateEmployeeRows, parseMixedScheduleRows, excelDateToJS, getImportedPreviewConflicts };',
  context
);

const {
  detectDateConventionsFromRows,
  excludeTemplateEmployeeRows,
  parseMixedScheduleRows,
  excelDateToJS,
  getImportedPreviewConflicts
} = context.api;

const workHeader = ['EMPLOYEE NO.', 'NAME', 'WORK DATE', 'SHIFT CODE', 'Day of Week', 'POSITION'];
const restHeader = ['EMPLOYEE NO.', 'NAME', 'REST DAY DATE', 'Day of Week', 'POSITION'];

function parse(rows, sheetName) {
  return parseMixedScheduleRows(rows, detectDateConventionsFromRows(rows), sheetName);
}

// DD/MM/YYYY: the unambiguous 15/10 establishes the convention for 02/10.
const dmyRows = parse([
  workHeader,
  ['7101', 'Example One', '02/10/2026', 'RBT-001', 'Friday', 'Cashier'],
  ['7101', 'Example One', '15/10/2026', 'RBT-001', 'Thursday', 'Cashier']
], 'Work Schedule');
assert.deepEqual(Array.from(dmyRows, row => row.date), ['10/2/2026', '10/15/2026']);

// MM/DD/YYYY: the unambiguous 10/15 establishes the convention for 10/02.
const mdyRows = parse([
  workHeader,
  ['7102', 'Example Two', '10/02/2026', 'RBT-001', 'Friday', 'Cashier'],
  ['7102', 'Example Two', '10/15/2026', 'RBT-001', 'Thursday', 'Cashier']
], 'Work Schedule');
assert.deepEqual(Array.from(mdyRows, row => row.date), ['10/2/2026', '10/15/2026']);

// Generic non-regression data (different organization-neutral names, month, and year).
const genericWork = parse([
  workHeader,
  ['8301', 'Taylor North', '04/03/2027', 'RBT-001', 'Thursday', 'Cashier'],
  ['8301', 'Taylor North', '17/03/2027', 'RBT-001', 'Wednesday', 'Cashier']
], 'Work Schedule');
const genericRest = parse([
  restHeader,
  ['8301', 'Taylor North', '03/04/2027', 'Thursday', 'Cashier'],
  ['8301', 'Taylor North', '03/17/2027', 'Wednesday', 'Cashier']
], 'Rest Day Schedule');
assert.deepEqual(Array.from(genericWork, row => row.date), ['3/4/2027', '3/17/2027']);
assert.deepEqual(Array.from(genericRest, row => row.date), ['3/4/2027', '3/17/2027']);

// An ambiguous-only import is rejected rather than silently guessed.
assert.throws(
  () => parse([workHeader, ['8302', 'Morgan Lake', '04/03/2027', 'RBT-001', '', 'Cashier']], 'Work Schedule'),
  /Ambiguous date/
);

// Template employee 1010 is removed before convention detection and parsing.
// Its ambiguous MDY-looking guide date therefore cannot override the real DMY evidence.
const workWithGuide = excludeTemplateEmployeeRows([
  workHeader,
  ['1010.0', 'Juan Dela Cruz', '10/2/25', 'RBT-001', 'Thursday', 'Guide'],
  ['8401', 'Real Employee', '02/10/2026', 'RBT-001', 'Friday', 'Cashier'],
  ['8401', 'Real Employee', '15/10/2026', 'RBT-001', 'Thursday', 'Cashier']
]);
const importedWork = parse(workWithGuide, 'Work Schedule');
assert.deepEqual(Array.from(importedWork, row => row.employeeNo), ['8401', '8401']);
assert.deepEqual(Array.from(importedWork, row => row.date), ['10/2/2026', '10/15/2026']);
assert.equal(
  parse(excludeTemplateEmployeeRows([
    workHeader,
    ['1010', 'Juan Dela Cruz', '10/2/25', 'RBT-001', 'Thursday', 'Guide']
  ]), 'Work Schedule').length,
  0
);

const restWithGuide = excludeTemplateEmployeeRows([
  restHeader,
  ['001010', 'Juan Dela Cruz', '10/1/25', 'Wednesday', 'Guide'],
  ['8401', 'Real Employee', '02/10/2026', 'Friday', 'Cashier'],
  ['8401', 'Real Employee', '15/10/2026', 'Thursday', 'Cashier']
]);
const importedRest = parse(restWithGuide, 'Rest Day Schedule');
assert.deepEqual(Array.from(importedRest, row => row.employeeNo), ['8401', '8401']);

// A name match alone is deliberately not excluded.
const legitimateSameName = excludeTemplateEmployeeRows([
  workHeader,
  ['8402', 'Juan Dela Cruz', '02/10/2026', 'RBT-001', 'Friday', 'Cashier'],
  ['8402', 'Juan Dela Cruz', '15/10/2026', 'RBT-001', 'Thursday', 'Cashier']
]);
assert.equal(parse(legitimateSameName, 'Work Schedule').length, 2);

// Genuine Excel Date objects and serial values bypass slash-format guessing.
assert.equal(excelDateToJS(new Date(2027, 2, 17), { requireConvention: true }), '3/17/2027');
assert.equal(excelDateToJS(46463, { requireConvention: true }), '3/17/2027');

// Existing normalized-date conflict rules still identify overlap and duplicates.
context.importedFiles = [{
  fileName: 'generic-schedule.xlsx',
  importFileKey: 'generic-schedule.xlsx-1',
  sheetName: 'Combined Schedule',
  rows: [
    ...genericWork.map((row, index) => ({ ...row, rowNumber: index + 2 })),
    { ...genericWork[0], rowNumber: 20 },
    ...genericRest.map((row, index) => ({ ...row, rowNumber: index + 30 }))
  ]
}];
context.importedFiles[0].rows.push(
  { employeeNo: '1010', name: 'Juan Dela Cruz', date: '3/4/2027', type: 'work', rowNumber: 50 },
  { employeeNo: '1010.0', name: 'Juan Dela Cruz', date: '3/4/2027', type: 'rest', rowNumber: 51 }
);
const conflicts = getImportedPreviewConflicts();
assert.ok(conflicts.some(item => item.reason.includes('Work Schedule and Rest Day on the same date: 3/4/2027')));
assert.ok(conflicts.some(item => item.reason.includes('duplicate date in Work Schedule: 3/4/2027')));
assert.equal(conflicts.some(item => String(item.employeeNo).includes('1010')), false);

console.log('Excel date convention and normalized conflict tests passed.');
