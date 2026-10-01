const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const appSource = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const persistenceSource = appSource.slice(
  appSource.indexOf('function createPersistedGroupState'),
  appSource.indexOf('function saveState()')
);
const context = {};

vm.runInNewContext(
  `const CHECKER_STORAGE_VERSION = 1; ${persistenceSource}\n` +
  'this.api = { createPersistedGroupState, parsePersistedCheckerState };',
  context
);

const { createPersistedGroupState, parsePersistedCheckerState } = context.api;
const work = [{ employeeNo: '1001', date: '10/1/2026', shiftCode: 'RBT-001' }];
const rest = [{ employeeNo: '1001', date: '10/2/2026' }];
const files = [{
  fileName: 'schedule.xlsx',
  importFileKey: 'schedule.xlsx-123',
  sheetName: 'Schedule',
  rows: [
    { ...work[0], type: 'work' },
    { ...rest[0], type: 'rest' }
  ],
  conflicts: []
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
assert.deepEqual(restored.workScheduleData, work);
assert.deepEqual(restored.restDayData, rest);
assert.deepEqual(restored.importedFiles, files);
assert.equal(restored.branchNames.work, 'Branch A');

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
