const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const appSource = fs.readFileSync(path.join(__dirname, '..', 'app.js'), 'utf8');
const employeeNumberSource = appSource.slice(
  appSource.indexOf('function normalizeEmployeeNo'),
  appSource.indexOf('function normalizeDayName')
);
const templateEmployeeSource = appSource.slice(
  appSource.indexOf('function isTemplateEmployeeNumber'),
  appSource.indexOf('function parseMixedScheduleRows')
).split('function excludeTemplateEmployeeRows')[0];
const exportSource = appSource.slice(
  appSource.indexOf('function dateForExcel'),
  appSource.indexOf('/*************************\\\n       * UI & RENDERING')
);
const exportDateSource = appSource.slice(
  appSource.indexOf('function normalizeDateForExport'),
  appSource.indexOf('function jsDateToExcel')
);
const exportedSheets = [];
const context = {
  GROUP_KEY: 'operation',
  workScheduleData: [
    { employeeNo: '1010.0', name: 'Juan Dela Cruz', date: '10/2/2025', shiftCode: 'RBT-001' },
    { employeeNo: '8401', name: 'Real Employee', date: '10/2/2026', shiftCode: 'RBT-001' }
  ],
  restDayData: [
    { employeeNo: '001010', name: 'Juan Dela Cruz', date: '10/1/2025' },
    { employeeNo: '8401', name: 'Real Employee', date: '10/3/2026' }
  ],
  XLSX: {
    utils: {
      book_new: () => ({}),
      json_to_sheet: rows => ({ '!ref': `A1:F${rows.length + 1}`, rows }),
      decode_range: () => ({ e: { r: 1 } }),
      encode_cell: () => 'unused',
      book_append_sheet: (_workbook, sheet, name) => exportedSheets.push({ name, rows: sheet.rows })
    },
    writeFile: () => {}
  },
  showWarning: () => {},
  showSuccess: () => {}
};

vm.runInNewContext(
  `${employeeNumberSource}\n${templateEmployeeSource}\n${exportDateSource}\n${exportSource}\nthis.runExport = exportCurrentSession;`,
  context
);

context.runExport();
assert.equal(exportedSheets.length, 2);
assert.deepEqual(
  exportedSheets.map(sheet => Array.from(sheet.rows, row => row['Employee Number'])),
  [['8401'], ['8401']]
);

console.log('Template employee export exclusion test passed.');
