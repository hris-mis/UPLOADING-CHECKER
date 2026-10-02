const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const appSource = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const persistenceSource = appSource.slice(
  appSource.indexOf('function withoutDerivedConflictState'),
  appSource.indexOf('function saveState()')
);
const employeeNumberSource = appSource.slice(
  appSource.indexOf('function normalizeEmployeeNo'),
  appSource.indexOf('function normalizeDayName')
);
const templateEmployeeSource = appSource.slice(
  appSource.indexOf('function isTemplateEmployeeNumber'),
  appSource.indexOf('function parseMixedScheduleRows')
).split('function excludeTemplateEmployeeRows')[0];
const context = {};

vm.runInNewContext(
  `const CHECKER_STORAGE_VERSION = 1; ${employeeNumberSource}\n${templateEmployeeSource}\n${persistenceSource}\n` +
  'this.api = { createPersistedGroupState, parsePersistedCheckerState };',
  context
);

const { createPersistedGroupState, parsePersistedCheckerState } = context.api;
const work = [{
  employeeNo: '1001',
  date: '10/1/2026',
  shiftCode: 'RBT-001',
  conflict: true,
  conflictReasons: ['stale conflict'],
  conflictReason: 'stale conflict',
  conflictType: 'duplicate',
  _rowNum: 1
}];
const rest = [{ employeeNo: '1001', date: '10/2/2026', conflict: false }];
const files = [{
  fileName: 'schedule.xlsx',
  importFileKey: 'schedule.xlsx-123',
  sheetName: 'Schedule',
  rows: [
    { ...work[0], type: 'work' },
    { ...rest[0], type: 'rest' }
  ],
  conflicts: [{ row: 2, reason: 'stale imported conflict' }]
}];

function roundTrip(groupState) {
  const restored = parsePersistedCheckerState(JSON.stringify({
    version: 1,
    groups: { operation: groupState }
  })).groups.operation;
  return JSON.parse(JSON.stringify(restored));
}

// Imported Work/RD rows, file history, source identity and branch fields survive a round trip.
let restored = roundTrip(createPersistedGroupState(work, rest, files, {
  work: 'Branch A',
  rest: 'Branch A'
}));
assert.deepEqual(restored.workScheduleData, [{
  employeeNo: '1001',
  date: '10/1/2026',
  shiftCode: 'RBT-001'
}]);
assert.deepEqual(restored.restDayData, [{ employeeNo: '1001', date: '10/2/2026' }]);
assert.deepEqual(restored.importedFiles, [{
  fileName: 'schedule.xlsx',
  importFileKey: 'schedule.xlsx-123',
  sheetName: 'Schedule',
  rows: [
    { employeeNo: '1001', date: '10/1/2026', shiftCode: 'RBT-001', type: 'work' },
    { employeeNo: '1001', date: '10/2/2026', type: 'rest' }
  ]
}]);
assert.equal(restored.branchNames.work, 'Branch A');

// The known template employee is never written to current or staged import data.
const guideRow = { employeeNo: '001010.0', name: 'Juan Dela Cruz', date: '10/2/2025' };
restored = roundTrip(createPersistedGroupState(
  [guideRow, ...work],
  [guideRow, ...rest],
  [{ ...files[0], rows: [guideRow, ...files[0].rows] }]
));
assert.deepEqual(restored.workScheduleData.map(row => row.employeeNo), ['1001']);
assert.deepEqual(restored.restDayData.map(row => row.employeeNo), ['1001']);
assert.equal(restored.importedFiles[0].rows.some(row => row.employeeNo === '001010.0'), false);

// Saving current arrays after edits and deletes cannot resurrect the prior records.
const editedWork = [{ ...work[0], shiftCode: 'RBT-099' }];
restored = roundTrip(createPersistedGroupState(editedWork, [], files));
assert.equal(restored.workScheduleData[0].shiftCode, 'RBT-099');
assert.deepEqual(restored.restDayData, []);

// Clear All is represented by empty persisted arrays and remains empty after restore.
restored = roundTrip(createPersistedGroupState([], [], []));
assert.deepEqual(restored.workScheduleData, []);
assert.deepEqual(restored.restDayData, []);
assert.deepEqual(restored.importedFiles, []);

// Missing, corrupt, incompatible, and structurally unsafe payloads safely fall back.
assert.equal(parsePersistedCheckerState(null), null);
assert.equal(parsePersistedCheckerState('{not-json'), null);
assert.equal(parsePersistedCheckerState(JSON.stringify({ version: 0, groups: {} })), null);
assert.equal(parsePersistedCheckerState(JSON.stringify({
  version: 1,
  groups: { operation: { workScheduleData: 'bad', restDayData: [], importedFiles: [] } }
})), null);

console.log('Local storage persistence tests passed.');
