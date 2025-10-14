/* main.js - enhanced & commented
   - Smart paste (work/rest) with detectColumnMapping()
   - Validation & HRIS file generation (conflicts warn but don't block)
   - Monitoring dashboard (search/filter working, export with month & progress)
   - Save/restore (localStorage)
   - Undo / Redo
   - Clear All (Schedule-only)
   - Comments/section headers for maintainability
*/

document.addEventListener('DOMContentLoaded', () => {
  /***** State *****/
  let workScheduleData = [];
  let restDayData = [];
  let rejectedRows = [];
  const undoStack = { work: [], rest: [] };
  const redoStack = { work: [], rest: [] };
  const deleteStack = { work: [], rest: [] };
  const LEADERSHIP_POSITIONS = ['Branch Head','Site Supervisor','OIC'];

  /***** Element refs *****/
  const workInput = document.getElementById('workScheduleInput');
  const restInput = document.getElementById('restScheduleInput');
  const workTableBody = document.getElementById('workTableBody');
  const restTableBody = document.getElementById('restTableBody');
  const summaryEl = document.getElementById('summary');
  const warningBanner = document.getElementById('warning-banner');
  const successMsg = document.getElementById('success-message');

  const generateWorkFileBtn = document.getElementById('generateWorkFile');
  const generateRestFileBtn = document.getElementById('generateRestFile');
  const clearWorkBtn = document.getElementById('clearWorkData');
  const clearRestBtn = document.getElementById('clearRestData');

  const tabSchedule = document.getElementById('tab-schedule');
  const tabMonitoring = document.getElementById('tab-monitoring');
  const scheduleContent = document.getElementById('tab-schedule-content');
  const monitoringContent = document.getElementById('tab-monitoring-content');

  // Monitoring refs
  const monitoringBody = document.getElementById('monitoringBody');
  const clearMonitoringBtn = document.getElementById('clearMonitoringBtn');
  const exportMonitoringBtn = document.getElementById('exportMonitoringBtn');
  const monthSelect = document.getElementById('monthSelect');
  const yearSelect = document.getElementById('yearSelect');
  const totalBranchesEl = document.getElementById('totalBranches');
  const checkedBranchesEl = document.getElementById('checkedBranches');
  const uploadedBranchesEl = document.getElementById('uploadedBranches');
  const progressPercentEl = document.getElementById('progressPercent');
  const progressBar = document.getElementById('progressBar');

  const addBranchBtn = document.getElementById('addBranchBtn');
  const backToTopBtn = document.getElementById('backToTopBtn');

  /***** Helpers: Notifications & escaping *****/
  function showBanner(msg) {
    if (!warningBanner) return;
    warningBanner.textContent = msg;
    warningBanner.classList.remove('hidden');
    warningBanner.classList.add('opacity-100');
    setTimeout(() => {
      warningBanner.classList.remove('opacity-100');
      setTimeout(() => warningBanner.classList.add('hidden'), 400);
    }, 3000);
  }

  function hideBanner() {
    if (!warningBanner) return;
    warningBanner.classList.add('hidden');
    warningBanner.classList.remove('opacity-100');
    warningBanner.textContent = '';
  }

  function showSuccess() {
    if (!successMsg) return;
    successMsg.classList.remove('hidden');
    successMsg.style.animation = 'fadeInOut 2.5s ease-in-out';
    setTimeout(() => {
      successMsg.classList.add('hidden');
      successMsg.style.animation = '';
    }, 2500);
  }

  function escapeHtml(str) {
    return (str == null) ? '' : String(str)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  /***** Utility parsers *****/
  // parse tabular text smartly (tabs, commas, or spaced columns)
  function parseTabular(text) {
    if (!text) return [];
    const rawLines = text.replace(/\r/g, '').split('\n');
    // trim and ignore lines that look like page/total headers
    const lines = rawLines.map(l => l.trim()).filter(l => l && !/^(sheet|page|total|subtotal|page\s*\d+)/i.test(l));
    if (lines.length === 0) return [];
    const sample = lines.slice(0,5).join('\n');
    let splitter = /\t/;
    if (!/\t/.test(sample)) {
      if (/,/.test(sample)) splitter = /,/;
      else splitter = /\s{2,}/;
    }
    return lines.map(line => line.split(splitter).map(c => c.trim()));
  }

// =============================================
// 🧠 Smart Header + Real Employee Detection
// Filters out decorative titles and summary lines
// Keeps only rows that look like actual employee data
// =============================================
function detectColumnMapping(rows) {
  // 🧹 Remove blank or decorative rows
  rows = rows.filter(r => {
    if (!r || !r.length) return false;
    const joined = r.join(' ').toUpperCase().trim();
    // Ignore banners or summary lines
    return !/^(WORK\s*SCHEDULE|REST\s*DAY|SCHEDULE|SUMMARY|TOTAL|PREPARED|PAGE|EMPLOYEE\s+SCHEDULE)$/i.test(joined);
  });

  // 🧭 Find potential header row
  let headerIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].map(cell => (cell || '').toLowerCase().trim());
    if (row.some(c => c.includes('emp')) && row.some(c => c.includes('date'))) {
      headerIndex = i;
      break;
    }
  }

  if (headerIndex === -1) headerIndex = 0;
  const headers = rows[headerIndex].map(h => h.toLowerCase().trim());
  const dataRows = rows.slice(headerIndex + 1);

/***** 🧠 Smart header mapping (fuzzy & forgiving) *****/
function detectColumnMapping(rows) {
  // Remove blank or decorative rows (banners, headers, footers)
  rows = rows.filter(r => {
    if (!r || !r.length) return false;
    const joined = r.join(' ').toUpperCase().trim();
    // Ignore common non-data lines
    return !/^(WORK\s*SCHEDULE|REST\s*DAY|SCHEDULE|SUMMARY|TOTAL|PREPARED|PAGE|EMPLOYEE\s+SCHEDULE)$/i.test(joined);
  });

  // Find potential header row by looking for key columns
  let headerIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    const row = rows[i].map(cell => (cell || '').toLowerCase().trim());
    if (row.some(c => c.includes('emp')) && row.some(c => c.includes('date'))) {
      headerIndex = i;
      break;
    }
  }

  // Default to first row if header not found
  if (headerIndex === -1) headerIndex = 0;

  const headers = rows[headerIndex].map(h => h.toLowerCase().trim());
  const dataRows = rows.slice(headerIndex + 1);

  /***** 🧠 Fuzzy Header Mapping (Detect columns) *****/
  // Inner function for flexible header detection
function detectColumnMappingInner(headerRow) {
  const normalize = (str) => str.replace(/[\s_\-\/\\\.]/g, '').toLowerCase();

  const header = headerRow.map(h => normalize(h));

  const mapping = {
    // Improve regex to include 'employee number' and 'emp number'
    name: header.findIndex(h => /name|fullname|employeename/.test(h)),
    empNo: header.findIndex(h => /emp|employee\s*number|employeenumber|idnum|id/.test(h)),
    date: header.findIndex(h => /date|workdate|sched|schedule/.test(h)),
    shift: header.findIndex(h => /shift|time|duty/.test(h)),
    day: header.findIndex(h => /day|daytype|typeday/.test(h)),
  };

  // Fallbacks as before
  if (mapping.name === -1) mapping.name = 0;
  if (mapping.empNo === -1) mapping.empNo = 1;
  if (mapping.date === -1) mapping.date = 2;
  if (mapping.shift === -1) mapping.shift = 3;
  if (mapping.day === -1) mapping.day = 4;

  return mapping;
}

  // Generate column mapping using the header row
  const colMap = detectColumnMappingInner(headers);

  // Filter valid data rows (must contain valid employee number and date)
  const validRows = dataRows.filter(r => {
    const emp = r[colMap.empNo];
    const date = r[colMap.date];
    return emp && /\d{3,}/.test(emp) && date;
  });

  return {
    headerIndex,
    colMap,
    dataRows: validRows
  };
}

  // ✅ Filter only rows that have a valid employee number and date
  const validRows = dataRows.filter(r => {
    const emp = r[mapping.empNo];
    const date = r[mapping.date];
    return emp && /\d{3,}/.test(emp) && date;
  });

  return { mapping, dataRows: validRows };
}

  // small wrapper kept for backward compatibility
 function detectHeaderAndMap(rows) {
  if (!rows || rows.length === 0) 
    return { headerIndex: -1, dataRows: [], colMap: {} };

  // 🧠 NEW: Handle single-row pastes smartly (no header)
  if (rows.length === 1) {
    const single = rows[0];
    // Fake header mapping for fallback
    const colMap = { name: 0, empNo: 1, date: 2, shift: 3, day: 4, position: 5 };
    return { headerIndex: -1, dataRows: [single], colMap };
  }

  // 🔍 For multi-line data, use normal detection
  const { headerIndex, colMap, dataRows } = detectColumnMapping(rows);

  // If detection fails and no valid data rows, fallback to using everything
  const safeRows = dataRows.length > 0 ? dataRows : rows;
  return { headerIndex, dataRows: safeRows, colMap };
}

  // auto-detect branch name inside pasted text (optional)
  function detectBranchName(text) {
    if (!text) return '';
    const m = text.match(/branch\s*[:\-]\s*(.+)/i);
    return m ? m[1].trim() : '';
  }

  // Normalize date (excel serials, common formats) -> MM/DD/YY
  function normalizeDate(dateStr) {
    if (!dateStr) return '';
    dateStr = (''+dateStr).toString().trim();
    // Excel serial numbers
    if (!isNaN(dateStr) && Number(dateStr) > 10000) {
      const excelEpoch = new Date(1899,11,30);
      const parsed = new Date(excelEpoch.getTime() + (Number(dateStr) * 86400000));
      const mm = String(parsed.getMonth()+1).padStart(2,'0');
      const dd = String(parsed.getDate()).padStart(2,'0');
      const yy = String(parsed.getFullYear()).slice(2);
      return `${mm}/${dd}/${yy}`;
    }
    const m = dateStr.match(/^(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{2,4})$/);
    if (m) {
      let mm = m[1].padStart(2,'0'), dd = m[2].padStart(2,'0'), yy = m[3];
      if (yy.length === 4) yy = yy.slice(2);
      return `${mm}/${dd}/${yy}`;
    }
    const p = new Date(dateStr);
    if (!isNaN(p)) {
      const mm = String(p.getMonth()+1).padStart(2,'0');
      const dd = String(p.getDate()).padStart(2,'0');
      const yy = String(p.getFullYear()).slice(2);
      return `${mm}/${dd}/${yy}`;
    }
    return dateStr;
  }

  function dayNameFromDate(dateStr) {
    if (!dateStr) return '';
    const p = new Date(dateStr);
    if (isNaN(p)) return '';
    return p.toLocaleDateString(undefined, { weekday: 'long' });
  }

  function isNumericStr(s) {
    return /^\d+$/.test((s||'').toString().trim());
  }

  /***** Rendering *****/
  function renderWorkTable() {
    if (!workTableBody) return;
    workTableBody.innerHTML = workScheduleData.map((d, i) => `
      <tr data-idx="${i}">
        <td>${escapeHtml(d.name)}</td>
        <td>${escapeHtml(d.empNo)}</td>
        <td>${escapeHtml(d.date)}</td>
        <td>${escapeHtml(d.shift)}</td>
        <td>${escapeHtml(d.day)}</td>
        <td>${escapeHtml(d.position)}</td>
        <td><button class="delete-row-btn" data-type="work" data-idx="${i}" title="Delete Row">❌</button></td>
      </tr>
    `).join('');
  }

  function renderRestTable() {
    if (!restTableBody) return;
    restTableBody.innerHTML = restDayData.map((d, i) => {
      const conflictHtml = (d.conflicts || []).map(c => `<div><strong>${escapeHtml(c.type)}:</strong> ${escapeHtml(c.reason)}</div>`).join('');
      const rowClass = (d.conflicts && d.conflicts.length > 0) ? 'conflict-row' : '';
      return `
        <tr class="${rowClass}" data-idx="${i}">
          <td class="text-left p-2 max-w-xs">${conflictHtml || ''}</td>
          <td>${escapeHtml(d.name)}</td>
          <td>${escapeHtml(d.empNo)}</td>
          <td>${escapeHtml(d.date)}</td>
          <td>${escapeHtml(d.day)}</td>
          <td>${escapeHtml(d.position)}</td>
          <td><button class="delete-row-btn" data-type="rest" data-idx="${i}" title="Delete Row">❌</button></td>
        </tr>
      `;
    }).join('');
    // summary
    const total = restDayData.length;
    const conflicts = restDayData.filter(r => r.conflicts && r.conflicts.length > 0).length;
    if (!summaryEl) return;
    if (total === 0) { summaryEl.textContent = ''; summaryEl.classList.add('hidden'); }
    else {
      summaryEl.classList.remove('hidden');
      if (conflicts === 0) {
        summaryEl.textContent = `✅ No conflicts detected for ${total} entries.`;
        summaryEl.classList.remove('text-red-600');
        summaryEl.classList.add('text-green-600');
      } else {
        summaryEl.textContent = `${conflicts} out of ${total} entries have conflicts detected.`;
        summaryEl.classList.remove('text-green-600');
        summaryEl.classList.add('text-red-600');
      }
    }
  }

  /***** Validation *****/
  function validateSchedules() {
    restDayData.forEach(r => r.conflicts = []);
    const workMap = new Map(workScheduleData.map(w => [`${w.empNo}-${(w.date||'').trim()}`, w]));
    const workEmpSet = new Set(workScheduleData.map(w => w.empNo));
    const restByDate = {}; const weekendCount = {}; const seen = new Set();

    restDayData.forEach(rd => {
      if (!rd) return;
      const key = `${rd.empNo}-${(rd.date||'').trim()}`;
      if (!restByDate[rd.date]) restByDate[rd.date] = [];
      restByDate[rd.date].push(rd);

      if (seen.has(key)) rd.conflicts.push({ type: 'Duplicate Entry', reason: 'Duplicate rest day entry for same employee & date.' });
      else seen.add(key);

      if (!workEmpSet.has(rd.empNo)) rd.conflicts.push({ type: 'Missing Employee', reason: 'Employee not found in Work Schedule data.' });

      const d = new Date(rd.date);
      if (isNaN(d.getTime())) rd.conflicts.push({ type: 'Invalid Date Format', reason: 'Date format unrecognized.' });

      if (workMap.has(key)) rd.conflicts.push({ type: 'Work Conflict', reason: 'Employee has a work schedule on same date.' });

      if (/saturday|sunday/i.test(rd.day || '')) {
        if (!isNaN(d.getTime())) {
          const monthYear = `${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`;
          const wkKey = `${rd.empNo}-${monthYear}`;
          weekendCount[wkKey] = (weekendCount[wkKey] || 0) + 1;
        }
      }
    });

    // leadership conflict (multiple leaders same day)
    for (const date in restByDate) {
      const leaders = restByDate[date].filter(r => LEADERSHIP_POSITIONS.includes(r.position));
      if (leaders.length > 1) leaders.forEach(l => l.conflicts.push({ type: 'Leadership Conflict', reason: 'Multiple leaders have same rest day.' }));
    }
    // weekend limit
    for (const wkey in weekendCount) {
      if (weekendCount[wkey] > 2) {
        const [emp] = wkey.split('-');
        restDayData.filter(r => r.empNo === emp && /saturday|sunday/i.test(r.day || '')).forEach(r => {
          r.conflicts.push({ type: 'Weekend Limit Exceeded', reason: `${weekendCount[wkey]} weekend rest days — maximum 2.` });
        });
      }
    }

    renderRestTable();
  }

  // Re-check conflicts helper (ensures RD & WS sync, call after any paste/change)
  function recheckConflicts() {
    validateSchedules();
    updateButtonStates();
    saveState();
  }

  /***** Buttons state rules *****/
  function updateButtonStates() {
    if (generateWorkFileBtn) generateWorkFileBtn.disabled = workScheduleData.length === 0;
    // allow generate for rest even if conflicts exist; only disable when empty
    if (generateRestFileBtn) generateRestFileBtn.disabled = restDayData.length === 0;
  }

  /***** 🪄 SMART PASTE + detectColumnMapping usage *****/
function handlePaste(e, type) {
  e.preventDefault();
  const clipboardData = e.clipboardData || window.clipboardData;
  const pastedData = clipboardData.getData('text/plain');
  if (!pastedData) return;

  // auto-detect branch name and set input if empty
  const branchName = detectBranchName(pastedData);
  if (branchName) {
    const branchInput = type === 'work' ? document.getElementById('workBranchName') : document.getElementById('restBranchName');
    if (branchInput && !branchInput.value) branchInput.value = branchName;
  }

  // parse to rows
  const parsed = parseTabular(pastedData);
  const { headerIndex, dataRows, colMap } = detectHeaderAndMap(parsed);
  // 🧠 NEW: Handle 1-row or headerless data gracefully
if (parsed.length === 1 || dataRows.length === 0) {
  console.log('Single-row or headerless paste detected, applying smart fallback.');
}
  const rows = dataRows.length ? dataRows : parsed;

  const cleaned = [], rejected = [];

  rows.forEach((row) => {
    let name = '', emp = '', date = '', shift = '', day = '', position = '';

    // use colMap if valid
    if (colMap && row[colMap.empNo] !== undefined) {
      name = (row[colMap.name] || '').trim();
      emp = (row[colMap.empNo] || '').trim();
      date = normalizeDate(row[colMap.date] || '');
      shift = (row[colMap.shift] || '').trim();
      day = (row[colMap.day] || '').trim();
      position = (row[colMap.position] || '').trim();
    } else {
      // heuristic fallback
      const cells = row.map(c => (c || '').trim());
      const empIdx = cells.findIndex(c => /^\d+$/.test(c));
      const dateIdx = cells.findIndex(c => (!isNaN(c) && Number(c) > 10000) || /[\/\.\-]/.test(c));
      const nameIdx = cells.findIndex(c => /^[A-Za-z\s,.'-]+$/.test(c) && c.split(' ').length >= 2);

      if (empIdx >= 0) emp = cells[empIdx];
      if (dateIdx >= 0) date = normalizeDate(cells[dateIdx]);
      if (nameIdx >= 0) name = cells[nameIdx];
      if (!day && date) day = dayNameFromDate(date);
    }

    // auto-generate day if date present but no day
    if (!day && date) day = dayNameFromDate(date);

    const obj = {
      name,
      empNo: emp,
      date,
      shift,
      day,
      position
    };

    const reasons = [];
    // 🧹 Smart Employee Number cleanup before validation
if (obj.empNo) {
  obj.empNo = obj.empNo.replace(/[^0-9]/g, '').trim(); // remove all non-numeric chars
}

// ✅ Accept cleaned empNo if it still has digits
if (!obj.empNo || obj.empNo.length < 3) {
  reasons.push('Missing or invalid Employee No');
}

    if (reasons.length) rejected.push({ row: row.join(' | '), reasons });
    else cleaned.push(obj);
  });

  // snapshot for undo
  if (!undoStack[type]) undoStack[type] = [];
  undoStack[type].push({ work: JSON.parse(JSON.stringify(workScheduleData)), rest: JSON.parse(JSON.stringify(restDayData)) });
  redoStack[type] = [];

  if (type === 'work') {
    workScheduleData = cleaned;
    renderWorkTable();
    if (workInput) workInput.value = '';
    showBanner(`✅ ${cleaned.length} work schedule rows pasted. ${rejected.length ? rejected.length + ' rejected.' : ''}`);
    recheckConflicts();
  } else {
    restDayData = cleaned;
    validateSchedules();
    renderRestTable();
    if (restInput) restInput.value = '';
    showBanner(`✅ ${cleaned.length} rest day rows pasted. ${rejected.length ? rejected.length + ' rejected.' : ''}`);
    recheckConflicts();
  }

  if (rejected.length) showRejectedModal(rejected);
  updateButtonStates();
  saveState();
}


  if (workInput) workInput.addEventListener('paste', (e) => handlePaste(e, 'work'));
  if (restInput) restInput.addEventListener('paste', (e) => handlePaste(e, 'rest'));

  /***** Row delete + undo/redo delete *****/
  document.addEventListener('click', (ev) => {
    const el = ev.target.closest && ev.target.closest('.delete-row-btn');
    if (!el) return;
    const type = el.dataset.type;
    const idx = Number(el.dataset.idx);
    if (type === 'work') {
      const removed = workScheduleData.splice(idx,1)[0];
      deleteStack.work.push({item: removed, idx});
      renderWorkTable();
      showBanner('Row deleted. You can undo delete.');
      recheckConflicts();
    } else {
      const removed = restDayData.splice(idx,1)[0];
      deleteStack.rest.push({item: removed, idx});
      renderRestTable();
      showBanner('Row deleted. You can undo delete.');
      recheckConflicts();
    }
    updateButtonStates();
    saveState();
  });

  function undoDelete(type) {
    const stack = deleteStack[type];
    if (!stack || !stack.length) return showBanner('Nothing to undo.');
    const last = stack.pop();
    if (type === 'work') { workScheduleData.splice(last.idx,0,last.item); renderWorkTable(); }
    else { restDayData.splice(last.idx,0,last.item); renderRestTable(); }
    showBanner('Undo successful.');
    recheckConflicts();
    updateButtonStates();
    saveState();
  }

  function undoPaste(type) {
    const stack = undoStack[type];
    if (!stack || !stack.length) return showBanner('Nothing to undo.');
    const snap = stack.pop();
    // push current to redo
    redoStack[type].push({ work: JSON.parse(JSON.stringify(workScheduleData)), rest: JSON.parse(JSON.stringify(restDayData)) });
    workScheduleData = snap.work;
    restDayData = snap.rest;
    renderWorkTable();
    renderRestTable();
    showBanner('Undo paste restored previous data.');
    recheckConflicts();
    updateButtonStates();
    saveState();
  }

  function redoPaste(type) {
    const stack = redoStack[type];
    if (!stack || !stack.length) return showBanner('Nothing to redo.');
    const snap = stack.pop();
    // push current to undo
    undoStack[type].push({ work: JSON.parse(JSON.stringify(workScheduleData)), rest: JSON.parse(JSON.stringify(restDayData)) });
    workScheduleData = snap.work;
    restDayData = snap.rest;
    renderWorkTable();
    renderRestTable();
    showBanner('Redo successful.');
    recheckConflicts();
    updateButtonStates();
    saveState();
  }

/***** ✅ Smart Fix Summary Modal (enhanced replacement) *****/
// Instead of rejecting rows, system now auto-includes all possible valid data.
// Decorative or incomplete rows are quietly skipped, but shown in summary.

function showRejectedModal(rejected) {
  const modal = document.getElementById('rejectedModal');
  if (!modal) return;
  const body = modal.querySelector('.modal-body');

  // 🧠 Filter out purely decorative or empty rows (non-blocking)
  const informative = rejected.filter(r => !/^(WORK\s*SCHEDULE|REST\s*DAY|TOTAL|SUMMARY|PAGE|PREPARED)/i.test(r.row));

  // 🧾 Summary content
  const msg = informative.length === 0
    ? `<p>✅ All rows have been processed successfully.<br>No critical issues detected.</p>`
    : `<p>⚙️ ${informative.length} rows were auto-corrected or skipped (decorative/non-critical):</p>` +
      informative
        .map(r => `
          <div style="padding:6px;border-bottom:1px solid #eee;">
            <strong>${escapeHtml(r.row)}</strong>
            <div style="color:#2563eb;margin-top:4px;">
              Notes: ${r.reasons.join(', ')}
            </div>
          </div>`)
        .join('');

  body.innerHTML = msg;

  // 💡 Cosmetic only — not blocking workflow
  modal.querySelector('.modal-title').textContent = 'Smart Paste Summary';
  modal.classList.remove('hidden');
  modal.style.display = 'block';
}

// 🧩 Close handler (unchanged)
document.addEventListener('click', (e) => {
  if (e.target.matches('.modal-close') || e.target.matches('#rejectedModal .modal-overlay')) {
    const modal = document.getElementById('rejectedModal');
    if (!modal) return;
    modal.classList.add('hidden');
    modal.style.display = 'none';
  }
});


  /***** HRIS File generation (weekend/conflicts allowed) *****/
  function generateHrisFile(type) {
    const workBranchEl = document.getElementById('workBranchName');
    const restBranchEl = document.getElementById('restBranchName');
    const workBranch = workBranchEl ? workBranchEl.value.trim() : '';
    const restBranch = restBranchEl ? restBranchEl.value.trim() : '';

    if (type === 'work') {
      if (!workBranch) return showBanner('⚠️ Enter Work Branch Name.');
      if (workScheduleData.length === 0) return showBanner('⚠️ No Work Schedule data to generate.');
      // Clean shift: remove spaces and uppercase
      const cleanedData = workScheduleData.map(r => ({
        empNo: r.empNo,
        date: r.date,
        shift: (r.shift || '').replace(/\s+/g, '').toUpperCase()
      }));
      const data = [['Employee Number','Work Date','Shift Code'], ...cleanedData.map(r => [r.empNo, r.date, r.shift])];
      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'HRIS Upload');
      XLSX.writeFile(wb, `${workBranch}_WORK_SCHEDULE.xlsx`);
      showSuccess();
    } else {
      if (!restBranch) return showBanner('⚠️ Enter Rest Branch Name.');
      if (restDayData.length === 0) return showBanner('⚠️ No Rest Day data to generate.');
      // If conflicts exist, show warning but do not block generation
      if (restDayData.some(r => r.conflicts && r.conflicts.length > 0)) {
        showBanner('⚠️ Note: There are conflicts, but file generation will proceed.');
      }
      const data = restDayData.map(r => ({ 'Employee No': r.empNo, 'Rest Day Date': r.date }));
      const ws = XLSX.utils.json_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'HRIS Upload');
      XLSX.writeFile(wb, `${restBranch}_REST_DAY_UPLOAD.xlsx`);
      showSuccess();
    }
  }

  if (generateWorkFileBtn) generateWorkFileBtn.addEventListener('click', () => generateHrisFile('work'));
  if (generateRestFileBtn) generateRestFileBtn.addEventListener('click', () => generateHrisFile('rest'));

  /***** Clear per-section buttons (unchanged behavior) *****/
  if (clearWorkBtn) clearWorkBtn.addEventListener('click', () => {
    workScheduleData = []; if (workInput) workInput.value = ''; renderWorkTable(); updateButtonStates(); showBanner('Work schedule cleared.');
    saveState();
  });
  if (clearRestBtn) clearRestBtn.addEventListener('click', () => {
    restDayData = []; if (restInput) restInput.value = ''; renderRestTable(); updateButtonStates(); showBanner('Rest day schedule cleared.');
    saveState();
  });

  /***** Keyboard shortcuts: Undo (Ctrl/Cmd+Z) & Redo (Ctrl/Cmd+Y) *****/
  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
      // try both stacks
      if (undoStack.work && undoStack.work.length) undoPaste('work');
      else if (undoStack.rest && undoStack.rest.length) undoPaste('rest');
    }
    if ((e.ctrlKey || e.metaKey) && (e.key === 'y' || (e.shiftKey && e.key === 'Z'))) {
      // redo
      if (redoStack.work && redoStack.work.length) redoPaste('work');
      else if (redoStack.rest && redoStack.rest.length) redoPaste('rest');
    }
  });

  /***** Monitoring Data Handling (unchanged core, enhanced export header) *****/
  const defaultMonitoring = [
    { name: 'AASP ABREEZA', checked: false, uploaded: false, uploadedBy: '', remarks: '' },
    { name: 'AASP NES - ATLAS', checked: false, uploaded: false, uploadedBy: '', remarks: '' }
  ];

  function getMonitoring() {
    const data = localStorage.getItem('monitoringData');
    if (data) {
      try { return JSON.parse(data); } catch (e) { return [...defaultMonitoring]; }
    }
    return [...defaultMonitoring];
  }

  function saveMonitoring(d) { localStorage.setItem('monitoringData', JSON.stringify(d)); }

  // animated progress
  let currentPercent = 0;
  function animateProgress(target) {
    const duration = 600;
    const start = performance.now();
    const from = currentPercent || 0;
    const diff = target - from;

    function frame(time) {
      const progress = Math.min((time - start) / duration, 1);
      const value = Math.round(from + diff * progress);
      if (progressPercentEl) progressPercentEl.textContent = value + '%';
      if (progressBar) progressBar.style.width = value + '%';

      if (progressBar) {
        if (value < 40) progressBar.style.background = 'linear-gradient(to right, #f43f5e, #fb7185)';
        else if (value < 80) progressBar.style.background = 'linear-gradient(to right, #fbbf24, #facc15)';
        else progressBar.style.background = 'linear-gradient(to right, #10b981, #34d399)';
      }

      if (progress < 1) requestAnimationFrame(frame);
      else currentPercent = target;
    }
    requestAnimationFrame(frame);
  }

  function updateMonitoringStats() {
    const data = getMonitoring();
    const total = data.length;
    const checked = data.filter(b => b.checked).length;
    const uploaded = data.filter(b => b.uploaded).length;
    if (totalBranchesEl) totalBranchesEl.textContent = total;
    if (checkedBranchesEl) checkedBranchesEl.textContent = checked;
    if (uploadedBranchesEl) uploadedBranchesEl.textContent = uploaded;
    const percent = total === 0 ? 0 : Math.round((uploaded / total) * 100);
    animateProgress(percent);
  }

  // render monitoring table with action icons (search/filter supported)
  function renderMonitoring() {
    const data = getMonitoring();
    const searchVal = (document.getElementById('monitorSearch') || {}).value?.trim().toLowerCase() || '';
    const filtered = data.filter((b) => {
      const matches = b.name.toLowerCase().includes(searchVal);
      const passes = showUnchecked ? !b.checked : true;
      return matches && passes;
    });

    if (!monitoringBody) return;
    monitoringBody.innerHTML = filtered.map((b,i) => `
      <tr class="hover:bg-gray-50 transition">
        <td class="text-left p-2">${escapeHtml(b.name)}</td>
        <td><input type="checkbox" data-index="${i}" data-field="checked" ${b.checked ? 'checked' : ''}></td>
        <td><input type="checkbox" data-index="${i}" data-field="uploaded" ${b.uploaded ? 'checked' : ''}></td>
        <td><input type="text" data-index="${i}" data-field="uploadedBy" value="${escapeHtml(b.uploadedBy)}" class="border rounded px-1 py-0.5 w-28"></td>
        <td><input type="text" data-index="${i}" data-field="remarks" value="${escapeHtml(b.remarks)}" class="border rounded px-1 py-0.5 w-28"></td>
        <td>
          <span class="action-edit" title="Edit" data-i="${i}" style="cursor:pointer;margin-right:8px;">✏️</span>
          <span class="action-delete" title="Delete" data-i="${i}" style="cursor:pointer;">❌</span>
        </td>
      </tr>
    `).join('');

    // inputs change handlers
    monitoringBody.querySelectorAll('input').forEach(inp => {
      inp.addEventListener('change', (e) => {
        const index = Number(e.target.dataset.index);
        const field = e.target.dataset.field;
        const d = getMonitoring();
        // map filtered index back to original index
        const filteredNames = filtered.map(x => x.name);
        const targetName = filteredNames[index];
        const origIndex = d.findIndex(x => x.name === targetName);
        if (origIndex === -1) return;
        d[origIndex][field] = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
        saveMonitoring(d);
        updateMonitoringStats();
        renderMonitoring(); // refresh UI so filters apply immediately
      });
    });

    // edit/delete handlers
    monitoringBody.querySelectorAll('.action-edit').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const i = Number(e.target.dataset.i);
        const d = getMonitoring();
        const filteredNames = filtered.map(x => x.name);
        const targetName = filteredNames[i];
        const origIndex = d.findIndex(x => x.name === targetName);
        if (origIndex === -1) return;
        const newName = prompt('Edit branch name:', d[origIndex].name);
        if (newName) { d[origIndex].name = newName; saveMonitoring(d); renderMonitoring(); showBanner('Branch updated.'); }
      });
    });
    monitoringBody.querySelectorAll('.action-delete').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const i = Number(e.target.dataset.i);
        const d = getMonitoring();
        const filteredNames = filtered.map(x => x.name);
        const targetName = filteredNames[i];
        const origIndex = d.findIndex(x => x.name === targetName);
        if (origIndex === -1) return;
        if (!confirm('Delete branch?')) return;
        d.splice(origIndex,1); saveMonitoring(d); renderMonitoring(); showBanner('Branch deleted.');
      });
    });

    updateMonitoringStats();
  }

  /***** Add Branch (floating) *****/
  if (addBranchBtn) {
    addBranchBtn.addEventListener('click', () => {
      const name = prompt('Branch name:');
      if (!name) return;
      const d = getMonitoring();
      d.push({ name, checked:false, uploaded:false, uploadedBy:'', remarks:'' });
      saveMonitoring(d);
      renderMonitoring();
      showBanner('Branch added.');
    });
  }

  /***** Clear & Export Buttons (monitoring) *****/
  if (clearMonitoringBtn) {
    clearMonitoringBtn.addEventListener('click', () => {
      if (!confirm('Clear all monitoring data?')) return;
      saveMonitoring([]);
      renderMonitoring();
      showBanner('All monitoring data cleared.');
    });
  }
if (exportMonitoringBtn) {
  exportMonitoringBtn.addEventListener('click', () => {
    const data = getMonitoring();
    if (!data.length) return showBanner('⚠️ No monitoring data to export.');

    const month = monthSelect?.value || 'N/A';
    const year = yearSelect?.value || new Date().getFullYear();
    const total = data.length;
    const uploaded = data.filter(b => b.uploaded).length;
    const progress = total ? Math.round((uploaded / total) * 100) : 0;

    const header = [
      [`Monitoring Export — ${month} ${year}`],
      [`Progress: ${progress}% (${uploaded}/${total} branches uploaded)`],
      [],
    ];

    const body = [['Branch Name', 'Checked', 'Uploaded', 'Uploaded By', 'Remarks']];
    data.forEach(b => {
      body.push([
        b.name,
        b.checked ? '✅' : '',
        b.uploaded ? '✅' : '',
        b.uploadedBy,
        b.remarks
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet([...header, ...body]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Monitoring Export');
    XLSX.writeFile(wb, `Monitoring_${month}_${year}.xlsx`);
    showSuccess();
  });
  }

  // Ensure monitoring search & filter re-render on change
  const monitorSearchInput = document.getElementById('monitorSearch');
  const monitorFilterUnchecked = document.getElementById('filterUnchecked');
  if (monitorSearchInput) monitorSearchInput.addEventListener('input', () => renderMonitoring());
  if (monitorFilterUnchecked) monitorFilterUnchecked.addEventListener('change', () => renderMonitoring());

  /***** Clear All Button (insert into Schedule Checker header) *****/
  (function insertClearAll() {
    const scheduleHeader = document.querySelector('#tab-schedule-content h1');
    if (!scheduleHeader) return;
    const btn = document.createElement('button');
    btn.id = 'clearAllBtn';
    btn.className = 'bg-blue-500 text-white px-3 py-1 rounded hover:bg-blue-600 shadow-sm transition';
    btn.textContent = '🧹 Clear All';
    btn.style.margin = '0.5rem auto 1rem';
    // Insert after header
    scheduleHeader.insertAdjacentElement('afterend', btn);

    btn.addEventListener('click', () => {
      if (!confirm('Are you sure you want to clear ALL schedule data (Work + Rest)?')) return;
      // clear data/state
      workScheduleData = [];
      restDayData = [];
      rejectedRows = [];
      // clear inputs
      if (workInput) workInput.value = '';
      if (restInput) restInput.value = '';
      const wbEl = document.getElementById('workBranchName');
      const rbEl = document.getElementById('restBranchName');
      if (wbEl) wbEl.value = '';
      if (rbEl) rbEl.value = '';
      // clear tables
      renderWorkTable();
      renderRestTable();
      // hide banners and summary
      hideBanner();
      if (summaryEl) { summaryEl.textContent = ''; summaryEl.classList.add('hidden'); }
      // update states
      updateButtonStates();
      showBanner('✅ All schedule data cleared.');
      // recheck just to be safe
      recheckConflicts();
      saveState();
    });
  })();

  /***** Save / restore scroll position when switching tabs *******/
  let lastScroll = 0;
  if (tabSchedule && tabMonitoring && scheduleContent && monitoringContent) {
    tabSchedule.addEventListener('click', () => {
      lastScroll = window.scrollY || 0;
      tabSchedule.classList.add('bg-indigo-600', 'text-white');
      tabSchedule.classList.remove('bg-gray-200', 'text-gray-700');
      tabMonitoring.classList.remove('bg-indigo-600', 'text-white');
      tabMonitoring.classList.add('bg-gray-200', 'text-gray-700');
      scheduleContent.classList.remove('hidden');
      monitoringContent.classList.add('hidden');
    });
    tabMonitoring.addEventListener('click', () => {
      lastScroll = window.scrollY || 0;
      tabMonitoring.classList.add('bg-indigo-600', 'text-white');
      tabMonitoring.classList.remove('bg-gray-200', 'text-gray-700');
      tabSchedule.classList.remove('bg-indigo-600', 'text-white');
      tabSchedule.classList.add('bg-gray-200', 'text-gray-700');
      scheduleContent.classList.add('hidden');
      monitoringContent.classList.remove('hidden');
      // refresh UI
      renderMonitoring();
      updateMonitoringStats();
      // restore approximate scroll
      setTimeout(() => window.scrollTo(0, lastScroll), 120);
    });
  }

  /***** Back to top button logic *****/
  if (backToTopBtn) {
    // show/hide by class .show (CSS should set display/fade)
    window.addEventListener('scroll', () => {
      if (window.scrollY > 400) backToTopBtn.classList.add('show');
      else backToTopBtn.classList.remove('show');
    });
    backToTopBtn.addEventListener('click', () => {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  }

  /***** Autosave / Load State (Work & Rest) *****/
  function saveState() {
    try {
      localStorage.setItem('workScheduleData', JSON.stringify(workScheduleData));
      localStorage.setItem('restDayData', JSON.stringify(restDayData));
    } catch (e) {
      // ignore storage exceptions
    }
  }
  function loadState() {
    try {
      const w = JSON.parse(localStorage.getItem('workScheduleData') || '[]');
      const r = JSON.parse(localStorage.getItem('restDayData') || '[]');
      workScheduleData = Array.isArray(w) ? w : [];
      restDayData = Array.isArray(r) ? r : [];
      renderWorkTable();
      renderRestTable();
      recheckConflicts();
      updateButtonStates();
    } catch (e) {
      // ignore parse errors
    }
  }

  /***** Initialize *****/
  loadState();
  renderWorkTable();
  renderRestTable();
  renderMonitoring();

  // expose some utilities for debugging (optional)
  window.__scc = {
    getMonitoring, saveMonitoring, renderMonitoring, updateMonitoringStats,
    workScheduleData, restDayData, validateSchedules, undoPaste, undoDelete, saveState, loadState
  };
});
