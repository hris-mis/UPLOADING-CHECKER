/* main.js - merged & fixed
   - Smart paste (work/rest)
   - Validation & HRIS file generation
   - Monitoring dashboard (animated progress)
   - Save/restore scroll position between tabs
   - Floating Add Branch + Back-to-Top
*/

document.addEventListener('DOMContentLoaded', () => {
  /***** State *****/
  let workScheduleData = [];
  let restDayData = [];
  let rejectedRows = [];
  const undoStack = { work: [], rest: [] };
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

  /***** Helpers *****/
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

  // detect header row and map column indices to our fields
  function detectHeaderAndMap(rows) {
    if (!rows || rows.length === 0) return { headerIndex: -1, dataRows: [], colMap: {} };
    const headerMap = {
      name: /name|employee\s*name|full\s*name|staff\s*name/,
      empNo: /emp.*no|employee.*no|employee|emp\s*id|id|emp\s*#|emp#|employee\s*number/,
      date: /date|work.*date|rest.*date|day\s*date|schedule.*date/,
      shift: /shift|shift\s*code|work\s*shift/,
      day: /day|day\s*of\s*week|dow/,
      position: /position|pos|job\s*title|role/
    };
    let bestIdx = -1, bestScore = -1;
    const maxHeaderRow = Math.min(6, rows.length);
    for (let i=0;i<maxHeaderRow;i++) {
      const row = rows[i].map(c => (c||'').toString().toLowerCase());
      let score = 0;
      for (const cell of row) {
        for (const key in headerMap) if (headerMap[key].test(cell)) score++;
      }
      if (score > bestScore) { bestScore = score; bestIdx = i; }
    }
    if (bestIdx === -1) bestIdx = 0;
    const headerRow = rows[bestIdx].map(c => (c||'').toString().toLowerCase());
    const colMap = {};
    headerRow.forEach((h, idx) => {
      for (const key in headerMap) {
        if (headerMap[key].test(h)) { colMap[key] = idx; break; }
      }
    });
    // sensible defaults if detection failed
    if (colMap.empNo === undefined) colMap.empNo = 1;
    if (colMap.name === undefined) colMap.name = 0;
    if (colMap.date === undefined) colMap.date = 2;
    if (colMap.shift === undefined) colMap.shift = 3;
    if (colMap.day === undefined) colMap.day = 4;
    if (colMap.position === undefined) colMap.position = 5;
    const dataRows = rows.slice(bestIdx + 1).filter(r => r.some(c => c && c.toString().trim() !== ''));
    return { headerIndex: bestIdx, dataRows, colMap };
  }

  function normalizeDate(dateStr) {
    if (!dateStr) return '';
    dateStr = (''+dateStr).trim();
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

  function updateButtonStates() {
    if (generateWorkFileBtn) generateWorkFileBtn.disabled = workScheduleData.length === 0;
    const hasConf = restDayData.some(r => r.conflicts && r.conflicts.length > 0);
    if (generateRestFileBtn) generateRestFileBtn.disabled = restDayData.length === 0 || hasConf;
  }

  /***** Paste handler (shared) *****/
  function handlePaste(e, type) {
    e.preventDefault();
    const clipboardData = e.clipboardData || window.clipboardData;
    const pastedData = clipboardData.getData('text/plain');
    if (!pastedData) return;
    const parsed = parseTabular(pastedData);
    const { headerIndex, dataRows, colMap } = detectHeaderAndMap(parsed);
    const cleaned = []; const rejected = [];

    dataRows.forEach((row, idx) => {
      const obj = {
        name: (row[colMap.name]||'').toString().trim(),
        empNo: (row[colMap.empNo]||'').toString().trim(),
        date: normalizeDate(row[colMap.date]||''),
        shift: (row[colMap.shift]||'').toString().trim(),
        day: (row[colMap.day]||'').toString().trim(),
        position: (row[colMap.position]||'').toString().trim()
      };
      const reasons = [];
      if (!obj.empNo || !isNumericStr(obj.empNo)) reasons.push('Missing or non-numeric Employee No');
      if (!obj.date || obj.date === '') reasons.push('Invalid or missing Date');
      if (reasons.length) rejected.push({ row: row.join(' | '), reasons });
      else cleaned.push(obj);
    });

    // snapshot for undo
    if (!undoStack[type]) undoStack[type] = [];
    undoStack[type].push({ work: JSON.parse(JSON.stringify(workScheduleData)), rest: JSON.parse(JSON.stringify(restDayData)) });

    if (type === 'work') {
      workScheduleData = cleaned;
      renderWorkTable();
      if (workInput) workInput.value = '';
      showBanner(`✅ ${cleaned.length} work schedule rows pasted. ${rejected.length ? rejected.length + ' rejected.' : ''}`);
    } else {
      restDayData = cleaned;
      validateSchedules();
      renderRestTable();
      if (restInput) restInput.value = '';
      showBanner(`✅ ${cleaned.length} rest day rows pasted. ${rejected.length ? rejected.length + ' rejected.' : ''}`);
    }
    if (rejected.length) showRejectedModal(rejected);
    updateButtonStates();
  }

  if (workInput) workInput.addEventListener('paste', (e) => handlePaste(e, 'work'));
  if (restInput) restInput.addEventListener('paste', (e) => handlePaste(e, 'rest'));

  /***** Row delete + undo delete *****/
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
    } else {
      const removed = restDayData.splice(idx,1)[0];
      deleteStack.rest.push({item: removed, idx});
      renderRestTable();
      showBanner('Row deleted. You can undo delete.');
    }
    updateButtonStates();
  });

  function undoDelete(type) {
    const stack = deleteStack[type];
    if (!stack || !stack.length) return showBanner('Nothing to undo.');
    const last = stack.pop();
    if (type === 'work') { workScheduleData.splice(last.idx,0,last.item); renderWorkTable(); }
    else { restDayData.splice(last.idx,0,last.item); renderRestTable(); }
    showBanner('Undo successful.');
    updateButtonStates();
  }
  function undoPaste(type) {
    const stack = undoStack[type];
    if (!stack || !stack.length) return showBanner('Nothing to undo.');
    const snap = stack.pop();
    workScheduleData = snap.work;
    restDayData = snap.rest;
    renderWorkTable();
    renderRestTable();
    showBanner('Undo paste restored previous data.');
    updateButtonStates();
  }

  /***** Rejected modal *****/
  function showRejectedModal(rejected) {
    const modal = document.getElementById('rejectedModal');
    if (!modal) return;
    const body = modal.querySelector('.modal-body');
    body.innerHTML = `<p>${rejected.length} rejected rows:</p>` + rejected.map(r => `<div style="padding:6px;border-bottom:1px solid #eee;"><strong>${escapeHtml(r.row)}</strong><div style="color:#b91c1c;margin-top:4px;">Reasons: ${r.reasons.join(', ')}</div></div>`).join('');
    modal.classList.remove('hidden');
    modal.style.display = 'block';
  }
  document.addEventListener('click', (e) => {
    if (e.target.matches('.modal-close') || e.target.matches('#rejectedModal .modal-overlay')) {
      const modal = document.getElementById('rejectedModal');
      if (!modal) return;
      modal.classList.add('hidden');
      modal.style.display = 'none';
    }
  });

  /***** HRIS File generation *****/
  function generateHrisFile(type) {
    const workBranchEl = document.getElementById('workBranchName');
    const restBranchEl = document.getElementById('restBranchName');
    const workBranch = workBranchEl ? workBranchEl.value.trim() : '';
    const restBranch = restBranchEl ? restBranchEl.value.trim() : '';

    if (type === 'work') {
      if (!workBranch) return showBanner('⚠️ Enter Work Branch Name.');
      if (workScheduleData.length === 0) return showBanner('⚠️ No Work Schedule data to generate.');
      const data = [['Employee Number','Work Date','Shift Code'], ...workScheduleData.map(r => [r.empNo, r.date, r.shift])];
      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'HRIS Upload');
      XLSX.writeFile(wb, `${workBranch}_WORK_SCHEDULE.xlsx`);
      showSuccess();
    } else {
      if (!restBranch) return showBanner('⚠️ Enter Rest Branch Name.');
      if (restDayData.length === 0) return showBanner('⚠️ No Rest Day data to generate.');
      if (restDayData.some(r => r.conflicts && r.conflicts.length > 0)) return showBanner('⚠️ Resolve conflicts before generating Rest Day file.');
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

  if (clearWorkBtn) clearWorkBtn.addEventListener('click', () => {
    workScheduleData = []; if (workInput) workInput.value = ''; renderWorkTable(); updateButtonStates(); showBanner('Work schedule cleared.');
  });
  if (clearRestBtn) clearRestBtn.addEventListener('click', () => {
    restDayData = []; if (restInput) restInput.value = ''; renderRestTable(); updateButtonStates(); showBanner('Rest day schedule cleared.');
  });

  // Ctrl/Cmd+Z to undo last paste (works across both types)
  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'z') {
      if (undoStack.work && undoStack.work.length) undoPaste('work');
      else if (undoStack.rest && undoStack.rest.length) undoPaste('rest');
    }
  });

  /***** Monitoring Data Handling *****/
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

  // render monitoring table with action icons
  function renderMonitoring() {
    const data = getMonitoring();
    const searchVal = (document.getElementById('monitorSearch') || {}).value || '';
    const showUnchecked = document.getElementById('filterUnchecked')?.checked;
    const filtered = data.filter((b) => {
      const matches = b.name.toLowerCase().includes(searchVal.toLowerCase());
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
        // map filtered index back to original index: we rely on filtered being subset, but to keep it simple
        // we'll update the first matching name occurrence if search/filter is active.
        const filteredNames = filtered.map(x => x.name);
        const targetName = filteredNames[index];
        const origIndex = d.findIndex(x => x.name === targetName);
        if (origIndex === -1) return;
        d[origIndex][field] = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
        saveMonitoring(d);
        updateMonitoringStats();
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

  /***** Clear & Export Buttons *****/
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
      const ws = XLSX.utils.json_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Branch Monitoring');
      const m = monthSelect ? monthSelect.value : '01';
      const y = yearSelect ? yearSelect.value : new Date().getFullYear();
      XLSX.writeFile(wb, `Monitoring_${m}_${y}.xlsx`);
      showSuccess();
    });
  }

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

  /***** Initialize *****/
  renderWorkTable();
  renderRestTable();
  renderMonitoring();

  // expose some utilities for debugging (optional)
  window.__scc = {
    getMonitoring, saveMonitoring, renderMonitoring, updateMonitoringStats,
    workScheduleData, restDayData, validateSchedules, undoPaste, undoDelete
  };
});
