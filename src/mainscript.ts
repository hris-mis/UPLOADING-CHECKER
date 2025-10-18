// Firebase imports — correct for TS
import { initializeApp } from "firebase/app";
import { getFirestore, doc, onSnapshot, setDoc } from "firebase/firestore";

const firebaseConfig = {
  apiKey: "AIzaSyDEGYeA0ere_txZPbwxMH5-BRflZqh_ef0",
  authDomain: "wikitehra.firebaseapp.com",
  projectId: "wikitehra",
  storageBucket: "wikitehra.firebasestorage.app",
  messagingSenderId: "761691537990",
  appId: "1:761691537990:web:70c47b4627350ade52c047"
};

const app = initializeApp(firebaseConfig);
const db = getFirestore(app);

declare const XLSX: any;

type RowObj = {
  name?: string;
  empNo?: string;
  date?: string;
  shift?: string;
  day?: string;
  position?: string;
  conflicts?: Array<{ type: string; reason: string }>;
  remarks?: string;
  checked?: boolean;
  uploaded?: boolean;
  uploadedBy?: string;
};

(() => {
  // Utility selectors with typed returns
  const $ = <T extends Element = Element>(sel: string): T | null => document.querySelector<T>(sel);
  const $$ = <T extends Element = Element>(sel: string): NodeListOf<T> => document.querySelectorAll<T>(sel);
  const pad2 = (v: string | number) => {
    const s = String(v);
    return s.length >= 2 ? s : '0' + s;
  };

  // Globals / state
  let workScheduleData: RowObj[] = [];
  let restDayData: RowObj[] = [];
  let rejectedRows: Array<{ row: string; reasons: string[] }> = [];
  const undoStack: Record<string, any[]> = { work: [], rest: [] };
  const redoStack: Record<string, any[]> = { work: [], rest: [] };
  const deleteStack: Record<string, any[]> = { work: [], rest: [] };
  const LEADERSHIP_POSITIONS = ['Branch Head', 'Site Supervisor', 'OIC'];
  let showUnchecked = false;
  let currentPercent = 0;

  // Handle show unchecked toggle
  const showUncheckedToggle = $('#showUncheckedToggle') as HTMLInputElement | null;
  if (showUncheckedToggle) {
    showUncheckedToggle.addEventListener('change', () => {
      showUnchecked = showUncheckedToggle.checked;
      renderMonitoring();
    });
  }

  // Element refs (may be null in some contexts)
  const workInput = $('#workScheduleInput') as HTMLTextAreaElement | null;
  const restInput = $('#restScheduleInput') as HTMLTextAreaElement | null;
  const workTableBody = $('#workTableBody') as HTMLElement | null;
  const restTableBody = $('#restTableBody') as HTMLElement | null;
  const summaryEl = $('#summary') as HTMLElement | null;
  const warningBanner = $('#warning-banner') as HTMLElement | null;
  const successMsg = $('#success-message') as HTMLElement | null;

  const generateWorkFileBtn = $('#generateWorkFile') as HTMLButtonElement | null;
  const generateRestFileBtn = $('#generateRestFile') as HTMLButtonElement | null;
  const clearWorkBtn = $('#clearWorkData') as HTMLButtonElement | null;
  const clearRestBtn = $('#clearRestData') as HTMLButtonElement | null;

  const tabSchedule = $('#tab-schedule') as HTMLElement | null;
  const tabMonitoring = $('#tab-monitoring') as HTMLElement | null;
  const scheduleContent = $('#tab-schedule-content') as HTMLElement | null;
  const monitoringContent = $('#tab-monitoring-content') as HTMLElement | null;

  // Monitoring refs
  const monitoringBody = $('#monitoringBody') as HTMLElement | null;
  const clearMonitoringBtn = $('#clearMonitoringBtn') as HTMLElement | null;
  const exportMonitoringBtn = $('#exportMonitoringBtn') as HTMLButtonElement | null;
  const monthSelect = $('#monthSelect') as HTMLSelectElement | null;
  const yearSelect = $('#yearSelect') as HTMLSelectElement | null;
  const totalBranchesEl = $('#totalBranches') as HTMLElement | null;
  const checkedBranchesEl = $('#checkedBranches') as HTMLElement | null;
  const uploadedBranchesEl = $('#uploadedBranches') as HTMLElement | null;
  const progressPercentEl = $('#progressPercent') as HTMLElement | null;
  const progressBar = $('#progressBar') as HTMLElement | null;
  const addBranchBtn = $('#addBranchBtn') as HTMLElement | null;
  const backToTopBtn = $('#backToTopBtn') as HTMLElement | null;

  /***** Helper UI functions *****/
  function showBanner(msg: string) {
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
  function escapeHtml(str: any) {
    return (str == null) ? '' : String(str)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

  /***** Parsers & detection *****/
  function parseTabular(text: string): string[][] {
    if (!text) return [];
    const rawLines = text.replace(/\r/g, '').split('\n');
    const lines = rawLines.map((l: string) => l.trim()).filter((l: string) => l && !/^(sheet|page|total|subtotal|page\s*\d+)/i.test(l));
    if (lines.length === 0) return [];
    const sample = lines.slice(0, 5).join('\n');
    let splitter: RegExp = /\t/;
    if (!/\t/.test(sample)) {
      if (/,/.test(sample)) splitter = /,/;
      else splitter = /\s{2,}/;
    }
    return lines.map((line: string) => line.split(splitter).map((c: string) => c.trim()));
  }

  function normalizeDate(dateStr: string | undefined): string {
    if (!dateStr) return '';
    let s = ('' + dateStr).trim();

    // keep weekday-like values
    const weekdays = ['mon', 'tue', 'wed', 'thu', 'fri', 'sat', 'sun'];
    if (weekdays.some(d => s.toLowerCase().startsWith(d))) return s;

    if (!isNaN(s as any) && Number(s) > 10000) {
      const excelEpoch = new Date(1899, 11, 30);
      const parsed = new Date(excelEpoch.getTime() + (Number(s) * 86400000));
      const mm = pad2(parsed.getMonth() + 1);
      const dd = pad2(parsed.getDate());
      const yy = String(parsed.getFullYear()).slice(2);
      return `${mm}/${dd}/${yy}`;
    }
    const m = s.match(/^(\d{1,2})[\/\.\-](\d{1,2})[\/\.\-](\d{2,4})$/);
    if (m) {
      let mm = pad2(m[1]), dd = pad2(m[2]), yy = m[3];
      if (yy.length === 4) yy = yy.slice(2);
      return `${mm}/${dd}/${yy}`;
    }
    const p = new Date(s);
    if (!isNaN(p.getTime())) {
      const mm = pad2(p.getMonth() + 1);
      const dd = pad2(p.getDate());
      const yy = String(p.getFullYear()).slice(2);
      return `${mm}/${dd}/${yy}`;
    }
    return s;
  }

  function dayNameFromDate(dateStr?: string) {
    if (!dateStr) return '';
    let s = dateStr.trim();

    const m = s.match(/^(\d{1,2})[\/\-\.](\d{1,2})[\/\-\.](\d{2})$/);
    if (m) s = `${m[1]}/${m[2]}/20${m[3]}`;

    const p = new Date(s);
    if (isNaN(p.getTime())) return '';
    return p.toLocaleDateString('en-US', { weekday: 'long' });
  }

  function detectBranchName(text?: string) {
    if (!text) return '';
    const m = text.match(/branch\s*[:\-]\s*(.+)/i);
    return m ? m[1].trim() : '';
  }

  function detectHeaderAndMap(rows: string[][]) {
    if (!rows || rows.length === 0) return { headerIndex: -1, dataRows: [] as string[][], colMap: {} as any };

    // helper detectors
    const isEmpNo = (v: string) => /^\d{3,6}$/.test(v);
    const isDateLike = (v: string) => /^\d{1,2}[\/\.\-]\d{1,2}[\/\.\-]\d{2,4}$/.test(v) || /^\d{5}$/.test(v) || /^\d{4}-\d{2}-\d{2}$/.test(v);
    const isShift = (v: string) => /am|pm|to|-|–|:/.test(v) && /\d/.test(v);
    const isDay = (v: string) => /^(mon|tue|wed|thu|fri|sat|sun)/i.test(v);
    const looksLikeName = (v: string) => /^[A-Za-z\s,.'-]{3,}$/.test(v) && v.split(' ').length >= 2;

    // if only one row: try to map by cell content
    if (rows.length === 1) {
      // keep original columns (do NOT filter empties) to preserve indexes
      const cells = rows[0].map((c: string | undefined) => (c?.trim() ?? ""));
      const colMap: any = {};
      cells.forEach((v: string, i: number) => {
        if (!v) return;
        if (isEmpNo(v)) colMap.empNo = i;
        else if (isDateLike(v)) colMap.date = i;
        else if (isShift(v)) colMap.shift = i;
        else if (isDay(v)) colMap.day = i;
        else if (looksLikeName(v)) colMap.name = i;
      });
      // sensible defaults (remain in-column order)
      if (colMap.name === undefined) colMap.name = 0;
      if (colMap.empNo === undefined) colMap.empNo = 1;
      if (colMap.date === undefined) colMap.date = 2;
      if (colMap.shift === undefined) colMap.shift = 3;
      if (colMap.day === undefined) colMap.day = 4;
      return { headerIndex: -1, dataRows: [cells], colMap };
    }

    // First, try to find an explicit header row (contains header keywords)
    let headerIndex = -1;
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i].map((c: string) => (c || '').toLowerCase());
      if (row.some((c: string) => c.includes('emp')) && row.some((c: string) => c.includes('date'))) { headerIndex = i; break; }
      if (row.some((c: string) => /name|fullname|employee|branch|position|shift|day/.test(c))) { headerIndex = i; break; }
    }

    // If no explicit header, try heuristic column-type detection across a sample of rows
    if (headerIndex === -1) {
      const sampleRows = rows.slice(0, Math.min(8, rows.length));
      const colCount = Math.max(...sampleRows.map(r => r.length));
      const scores: Array<{ num: number; date: number; shift: number; day: number; name: number }> = [];
      for (let c = 0; c < colCount; c++) {
        scores[c] = { num: 0, date: 0, shift: 0, day: 0, name: 0 };
        for (const r of sampleRows) {
          const cell = (r[c] || '').trim();
          if (!cell) continue;
          if (isEmpNo(cell)) scores[c].num++;
          if (isDateLike(cell)) scores[c].date++;
          if (isShift(cell)) scores[c].shift++;
          if (isDay(cell)) scores[c].day++;
          if (looksLikeName(cell)) scores[c].name++;
        }
      }
      // choose best candidates
      const pickBest = (key: keyof typeof scores[0]) => {
        let best = -1, bestScore = -1;
        for (let i = 0; i < scores.length; i++) {
          const s = (scores[i] as any)[key] || 0;
          if (s > bestScore) { bestScore = s; best = i; }
        }
        return best;
      };
      const mapping: any = {};
      const empIdx = pickBest('num');
      const dateIdx = pickBest('date');
      const nameIdx = pickBest('name');
      const shiftIdx = pickBest('shift');
      const dayIdx = pickBest('day');

      if (nameIdx !== -1) mapping.name = nameIdx;
      if (empIdx !== -1) mapping.empNo = empIdx;
      if (dateIdx !== -1) mapping.date = dateIdx;
      if (shiftIdx !== -1) mapping.shift = shiftIdx;
      if (dayIdx !== -1) mapping.day = dayIdx;

      // if mapping looks reasonable (found at least emp or date), treat as no header + auto-map
      if (mapping.empNo !== undefined || mapping.date !== undefined) {
        // sensible fallbacks
        if (mapping.name === undefined) mapping.name = 0;
        if (mapping.empNo === undefined) mapping.empNo = 1;
        if (mapping.date === undefined) mapping.date = 2;
        if (mapping.shift === undefined) mapping.shift = 3;
        if (mapping.day === undefined) mapping.day = 4;
        return { headerIndex: -1, dataRows: rows, colMap: mapping };
      }

      // if heuristics can't confidently map, do NOT assume the first row is header.
      // Instead, fallback to default positional mapping for all rows.
      const fallback: any = { name: 0, empNo: 1, date: 2, shift: 3, day: 4, position: 5 };
      return { headerIndex: -1, dataRows: rows, colMap: fallback };
    }

    // existing header parsing (when headerIndex is set)
    const detectInner = (headerRow: string[]) => {
      const normalize = (s: string) => s.replace(/[\s_\-\/\\\.]/g, '').toLowerCase();
      const h = headerRow.map(normalize);
      const mapping: any = {
        name: h.findIndex((x: string) => /name|fullname|employeename/.test(x)),
        empNo: h.findIndex((x: string) => /emp|employeenumber|idnum|id/.test(x)),
        date: h.findIndex((x: string) => /date|workdate|sched|schedule/.test(x)),
        shift: h.findIndex((x: string) => /shift|time|duty/.test(x)),
        day: h.findIndex((x: string) => /day|daytype|typeday/.test(x)),
        position: h.findIndex((x: string) => /position|title|role/.test(x))
      };
      if (mapping.name === -1) mapping.name = 0;
      if (mapping.empNo === -1) mapping.empNo = 1;
      if (mapping.date === -1) mapping.date = 2;
      if (mapping.shift === -1) mapping.shift = 3;
      if (mapping.day === -1) mapping.day = 4;
      if (mapping.position === -1) mapping.position = 5;
      return mapping;
    };
    const colMap = detectInner(rows[headerIndex]);
    const dataRows = rows.slice(headerIndex + 1);
    return { headerIndex, dataRows: dataRows.length ? dataRows : rows, colMap };
  }

  /***** Rendering tables *****/
  function renderWorkTable() {
    if (!workTableBody) return;
    workTableBody.innerHTML = workScheduleData.map((d: RowObj, i: number) => `
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
    restTableBody.innerHTML = restDayData.map((d: RowObj, i: number) => {
      const conflictHtml = (d.conflicts || []).map((c: { type: string; reason: string }) => `<div><strong>${escapeHtml(c.type)}:</strong> ${escapeHtml(c.reason)}</div>`).join('');
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

    const total = restDayData.length;
    const conflicts = restDayData.filter((r: RowObj) => r.conflicts && r.conflicts.length > 0).length;
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

  /***** Validation logic *****/
  function validateSchedules() {
    restDayData.forEach((r: RowObj) => (r.conflicts = []));
    const workMap = new Map<string, RowObj>(workScheduleData.map((w: RowObj) => [`${w.empNo}-${(w.date || '').trim()}`, w]));
    const workEmpSet = new Set<string>(workScheduleData.map((w: RowObj) => w.empNo || ''));
    const restByDate: Record<string, RowObj[]> = {};
    const weekendCount: Record<string, number> = {};
    const seen = new Set<string>();

    restDayData.forEach((rd: RowObj) => {
      if (!rd) return;
      const key = `${rd.empNo}-${(rd.date || '').trim()}`;
      if (!restByDate[rd.date || '']) restByDate[rd.date || ''] = [];
      (restByDate[rd.date || '']).push(rd);

      if (seen.has(key)) rd.conflicts!.push({ type: 'Duplicate Entry', reason: 'Duplicate rest day entry for same employee & date.' });
      else seen.add(key);

      if (!workEmpSet.has(rd.empNo || '')) rd.conflicts!.push({ type: 'Missing Employee', reason: 'Employee not found in Work Schedule data.' });

      const d = new Date(rd.date || '');
      if (isNaN(d.getTime())) rd.conflicts!.push({ type: 'Invalid Date Format', reason: 'Date format unrecognized.' });

      if (workMap.has(key)) rd.conflicts!.push({ type: 'Work Conflict', reason: 'Employee has a work schedule on same date.' });

      if (/saturday|sunday/i.test(rd.day || '')) {
        if (!isNaN(d.getTime())) {
          const monthYear = `${pad2(d.getMonth() + 1)}/${d.getFullYear()}`;
          const wkKey = `${rd.empNo}-${monthYear}`;
          weekendCount[wkKey] = (weekendCount[wkKey] || 0) + 1;
        }
      }
    });

    // leadership conflicts
    for (const date in restByDate) {
      const leaders = restByDate[date].filter((r: RowObj) => LEADERSHIP_POSITIONS.indexOf(r.position || '') !== -1);
      if (leaders.length > 1) leaders.forEach((l: RowObj) => l.conflicts!.push({ type: 'Leadership Conflict', reason: 'Multiple leaders have same rest day.' }));
    }
    // weekend limit enforcement
    for (const wkey in weekendCount) {
      if (weekendCount[wkey] > 2) {
        const [emp] = wkey.split('-');
        restDayData.filter((r: RowObj) => r.empNo === emp && /saturday|sunday/i.test(r.day || '')).forEach((r: RowObj) => {
          r.conflicts!.push({ type: 'Weekend Limit Exceeded', reason: `${weekendCount[wkey]} weekend rest days — maximum 2.` });
        });
      }
    }

    renderRestTable();
  }

  function recheckConflicts() {
    validateSchedules();
    updateButtonStates();
    saveState();
  }

  function updateButtonStates() {
    if (generateWorkFileBtn) generateWorkFileBtn.disabled = workScheduleData.length === 0;
    if (generateRestFileBtn) generateRestFileBtn.disabled = restDayData.length === 0;
  }

  /***** Monitoring helpers: local + shared Firestore document *****/
  const defaultMonitoring: RowObj[] = [
    { name: 'AASP ABREEZA', checked: false, uploaded: false, uploadedBy: '', remarks: '' },
    { name: 'AASP NES - ATLAS', checked: false, uploaded: false, uploadedBy: '', remarks: '' }
  ];

  function getMonitoring(): RowObj[] {
    const raw = localStorage.getItem('monitoringData');
    if (!raw) return [...defaultMonitoring];
    try { return JSON.parse(raw) as RowObj[]; } catch { return [...defaultMonitoring]; }
  }

  async function saveMonitoring(d: RowObj[]) {
    try {
      // Save locally first
      localStorage.setItem('monitoringData', JSON.stringify(d));
      // Save to Firestore shared document (best-effort)
      try {
        const sharedRef = doc(db, 'monitoring', 'sharedMonitoringData');
        await setDoc(sharedRef, { data: d, updatedAt: new Date() }, { merge: true });
      } catch (e) {
        console.warn('Firestore save failed (monitoring):', e);
      }
    } catch (error) {
      console.error('saveMonitoring error', error);
    }
  }

  function initMonitoringSync() {
    try {
      const sharedRef = doc(db, 'monitoring', 'sharedMonitoringData');
      onSnapshot(sharedRef, (snap) => {
        try {
          if (!snap.exists()) {
            // no shared doc yet — keep local
            return;
          }
          const payload = (snap.data() as any).data as RowObj[] | undefined;
          if (Array.isArray(payload)) {
            // overwrite local copy with shared data
            localStorage.setItem('monitoringData', JSON.stringify(payload));
            renderMonitoring();
          }
        } catch (e) {
          console.warn('monitoring snapshot parse error', e);
        }
      });
    } catch (e) {
      console.warn('initMonitoringSync skipped (Firestore)', e);
    }
  }

  /***** Paste handling (work/rest) *****/
  function handlePaste(e: ClipboardEvent, type: 'work' | 'rest') {
    e.preventDefault();
    const clipboardData = e.clipboardData || (window as any).clipboardData;
    const pastedData = clipboardData.getData('text/plain');
    if (!pastedData) return;

    const branchName = detectBranchName(pastedData);
    if (branchName) {
      const branchInput = type === 'work' ? $('#workBranchName') as HTMLInputElement | null : $('#restBranchName') as HTMLInputElement | null;
      if (branchInput && !branchInput.value) branchInput.value = branchName;
    }

    const parsed = parseTabular(pastedData);
    const { headerIndex, dataRows, colMap } = detectHeaderAndMap(parsed);
    const rows = dataRows.length ? dataRows : parsed;

    const cleaned: RowObj[] = [];
    const rejected: Array<{ row: string; reasons: string[] }> = [];

    rows.forEach((row: string[]) => {
      let name = '', emp = '', date = '', shift = '', day = '', position = '';
      if (colMap && row[colMap.empNo] !== undefined) {
        name = (row[colMap.name] || '').trim();
        emp = (row[colMap.empNo] || '').trim();
        date = normalizeDate(row[colMap.date] || '');
        shift = (row[colMap.shift] || '').trim();
        day = (row[colMap.day] || '').trim();
        position = (row[colMap.position] || '').trim();

        // If the detected 'day' cell actually contains a date, convert to weekday
        if (day && (/^[\d\/\.\-]{6,}$/.test(day) || /^\d{4}-\d{2}-\d{2}$/.test(day))) {
          const normalized = normalizeDate(day);
          const weekday = dayNameFromDate(normalized);
          if (weekday) day = weekday;
        }

        // ensure day from date when missing
        if (!day && date) day = dayNameFromDate(date);
      } else {
        const cells = row.map((c: string) => (c || '').trim());
        const empIdx = cells.findIndex((c: string) => /^\d+$/.test(c));
        const dateIdx = cells.findIndex((c: string) => ((!isNaN((c as any)) && Number(c) > 10000) || /[\/\.\-]/.test(c)));
        const nameIdx = cells.findIndex((c: string) => /^[A-Za-z\s,.'-]+$/.test(c) && c.split(' ').length >= 2);
        if (empIdx >= 0) emp = cells[empIdx];
        if (dateIdx >= 0) date = normalizeDate(cells[dateIdx]);
        if (nameIdx >= 0) name = cells[nameIdx];
        if (!day && date) day = dayNameFromDate(date);
      }

      if ((!emp || emp.length < 3) && row.some((c: string) => /^\d{3,}$/.test(c))) {
        const found = row.find((c: string) => /^\d{3,}$/.test(c));
        if (found) emp = String(found).trim();
      }
      if (!name && row.some((c: string) => /^[A-Za-z\s]+$/.test(String(c)) && String(c).split(' ').length >= 2)) {
        const found = row.find((c: string) => /^[A-Za-z\s]+$/.test(String(c)) && String(c).split(' ').length >= 2);
        if (found) name = String(found).trim();
      }
      if (!date && row.some((c: string) => /[\/\-\.]/.test(String(c)) || /^\d{5}$/.test(String(c)))) {
        const found = row.find((c: string) => /[\/\-\.]/.test(String(c)) || /^\d{5}$/.test(String(c)));
        if (found) date = normalizeDate(String(found));
      }
      if (!day && date) day = dayNameFromDate(date);

      const obj: RowObj = { name, empNo: emp ? String(emp).replace(/[^0-9]/g, '') : '', date, shift, day, position };
      const reasons: string[] = [];
      if (!obj.empNo || obj.empNo.length < 2 || obj.empNo.length > 6) reasons.push('Missing or invalid Employee No');
      if (reasons.length) rejected.push({ row: row.join(' | '), reasons });
      else cleaned.push(obj);
    });

    // Snapshot for undo
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

  if (workInput) workInput.addEventListener('paste', (e) => handlePaste(e as ClipboardEvent, 'work'));
  if (restInput) restInput.addEventListener('paste', (e) => handlePaste(e as ClipboardEvent, 'rest'));

  /***** Row deletion / undo/redo *****/
  document.addEventListener('click', (ev: MouseEvent) => {
    const target = ev.target as HTMLElement;
    const btn = (target.closest && (target.closest('.delete-row-btn') as HTMLElement | null));
    if (!btn) return;
    const type = btn.dataset.type;
    const idx = Number(btn.dataset.idx);
    if (type === 'work') {
      const removed = workScheduleData.splice(idx, 1)[0];
      deleteStack.work.push({ item: removed, idx });
      renderWorkTable();
      showBanner('Row deleted. You can undo delete.');
      recheckConflicts();
    } else {
      const removed = restDayData.splice(idx, 1)[0];
      deleteStack.rest.push({ item: removed, idx });
      renderRestTable();
      showBanner('Row deleted. You can undo delete.');
      recheckConflicts();
    }
    updateButtonStates();
    saveState();
  });

  function undoDelete(type: 'work' | 'rest') {
    const stack = deleteStack[type];
    if (!stack || !stack.length) return showBanner('Nothing to undo.');
    const last = stack.pop();
    if (type === 'work') { workScheduleData.splice(last.idx, 0, last.item); renderWorkTable(); }
    else { restDayData.splice(last.idx, 0, last.item); renderRestTable(); }
    showBanner('Undo successful.');
    recheckConflicts();
    updateButtonStates();
    saveState();
  }

  function undoPaste(type: 'work' | 'rest') {
    const stack = undoStack[type];
    if (!stack || !stack.length) return showBanner('Nothing to undo.');
    const snap = stack.pop();
    redoStack[type].push({ work: JSON.parse(JSON.stringify(workScheduleData)), rest: JSON.parse(JSON.stringify(restDayData)) });
    workScheduleData = snap.work;
    restDayData = snap.rest;
    renderWorkTable(); renderRestTable();
    showBanner('Undo paste restored previous data.');
    recheckConflicts();
    updateButtonStates();
    saveState();
  }

  function redoPaste(type: 'work' | 'rest') {
    const stack = redoStack[type];
    if (!stack || !stack.length) return showBanner('Nothing to redo.');
    const snap = stack.pop();
    undoStack[type].push({ work: JSON.parse(JSON.stringify(workScheduleData)), rest: JSON.parse(JSON.stringify(restDayData)) });
    workScheduleData = snap.work;
    restDayData = snap.rest;
    renderWorkTable(); renderRestTable();
    showBanner('Redo successful.');
    recheckConflicts();
    updateButtonStates();
    saveState();
  }

  /***** Rejected modal (summary) *****/
  function showRejectedModal(rejected: Array<{ row: string; reasons: string[] }>) {
    const modal = $('#rejectedModal') as HTMLElement | null;
    if (!modal) return;
    const body = modal.querySelector('.modal-body') as HTMLElement | null;
    if (!body) return;
    const informative = rejected.filter((r: { row: string; reasons: string[] }) => !/^(WORK\s*SCHEDULE|REST\s*DAY|TOTAL|SUMMARY|PAGE|PREPARED)/i.test(r.row));
    const msg = informative.length === 0
      ? `<p>✅ All rows have been processed successfully.<br>No critical issues detected.</p>`
      : `<p>⚙️ ${informative.length} rows were auto-corrected or skipped (decorative/non-critical):</p>` +
        informative.map((r: { row: string; reasons: string[] }) => `<div style="padding:6px;border-bottom:1px solid #eee;"><strong>${escapeHtml(r.row)}</strong><div style="color:#2563eb;margin-top:4px;">Notes: ${r.reasons.join(', ')}</div></div>`).join('');
    body.innerHTML = msg;
    const title = modal.querySelector('.modal-title');
    if (title) title.textContent = 'Smart Paste Summary';
    modal.classList.remove('hidden'); modal.style.display = 'block';
  }

  document.addEventListener('click', (e: Event) => {
    const target = e.target as HTMLElement;
    if (target.matches('.modal-close') || target.matches('#rejectedModal .modal-overlay')) {
      const modal = $('#rejectedModal') as HTMLElement | null;
      if (!modal) return;
      modal.classList.add('hidden'); modal.style.display = 'none';
    }
  });

  /***** HRIS generation *****/
  function generateHrisFile(type: 'work' | 'rest') {
    const workBranchEl = $('#workBranchName') as HTMLInputElement | null;
    const restBranchEl = $('#restBranchName') as HTMLInputElement | null;
    const workBranch = workBranchEl ? (workBranchEl.value || '') : '';
    const restBranch = restBranchEl ? (restBranchEl.value || '') : '';

    if (type === 'work') {
      if (!workBranch) return showBanner('⚠️ Enter Work Branch Name.');
      if (workScheduleData.length === 0) return showBanner('⚠️ No Work Schedule data to generate.');
      const cleanedData = workScheduleData.map((r: RowObj) => ({ empNo: r.empNo, date: r.date, shift: (r.shift || '').replace(/\s+/g, '').toUpperCase() }));
      const data = [['Employee Number', 'Work Date', 'Shift Code'], ...cleanedData.map((r: any) => [r.empNo, r.date, r.shift])];
      const ws = XLSX.utils.aoa_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'HRIS Upload');
      XLSX.writeFile(wb, `${workBranch}_WORK_SCHEDULE.xlsx`);
      showSuccess();
    } else {
      if (!restBranch) return showBanner('⚠️ Enter Rest Branch Name.');
      if (restDayData.length === 0) return showBanner('⚠️ No Rest Day data to generate.');
      if (restDayData.some((r: RowObj) => r.conflicts && r.conflicts.length > 0)) {
        showBanner('⚠️ Note: There are conflicts, but file generation will proceed.');
      }
      const data = restDayData.map((r: RowObj) => ({ 'Employee No': r.empNo, 'Rest Day Date': r.date }));
      const ws = XLSX.utils.json_to_sheet(data);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'HRIS Upload');
      XLSX.writeFile(wb, `${restBranch}_REST_DAY_UPLOAD.xlsx`);
      showSuccess();
    }
  }

  if (generateWorkFileBtn) generateWorkFileBtn.addEventListener('click', () => generateHrisFile('work'));
  if (generateRestFileBtn) generateRestFileBtn.addEventListener('click', () => generateHrisFile('rest'));

  /***** Clear per-section *****/
  if (clearWorkBtn) clearWorkBtn.addEventListener('click', () => {
    workScheduleData = []; if (workInput) workInput.value = ''; renderWorkTable(); updateButtonStates(); showBanner('Work schedule cleared.'); saveState();
  });
  if (clearRestBtn) clearRestBtn.addEventListener('click', () => {
    restDayData = []; if (restInput) restInput.value = ''; renderRestTable(); updateButtonStates(); showBanner('Rest day schedule cleared.'); saveState();
  });

  /***** Keyboard shortcuts *****/
  document.addEventListener('keydown', (e: KeyboardEvent) => {
    if ((e.ctrlKey || e.metaKey) && (String(e.key).toLowerCase() === 'z')) {
      if (undoStack.work && undoStack.work.length) undoPaste('work');
      else if (undoStack.rest && undoStack.rest.length) undoPaste('rest');
    }
    if ((e.ctrlKey || e.metaKey) && ((String(e.key).toLowerCase() === 'y') || (e.shiftKey && (String(e.key).toLowerCase() === 'z')))) {
      if (redoStack.work && redoStack.work.length) redoPaste('work');
      else if (redoStack.rest && redoStack.rest.length) redoPaste('rest');
    }
  });

  /***** Monitoring rendering & UI (uses getMonitoring/saveMonitoring) *****/
  function animateProgress(target: number) {
    const duration = 600;
    const start = performance.now();
    const from = currentPercent || 0;
    const diff = target - from;
    function frame(time: number) {
      const progress = Math.min((time - start) / duration, 1);
      const value = Math.round(from + diff * progress);
      if (progressPercentEl) progressPercentEl.textContent = `${value}%`;
      if (progressBar) progressBar.style.width = `${value}%`;
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
    const checked = data.filter((b: RowObj) => !!b.checked).length;
    const uploaded = data.filter((b: RowObj) => !!b.uploaded).length;
    if (totalBranchesEl) totalBranchesEl.textContent = String(total);
    if (checkedBranchesEl) checkedBranchesEl.textContent = String(checked);
    if (uploadedBranchesEl) uploadedBranchesEl.textContent = String(uploaded);
    const percent = total === 0 ? 0 : Math.round((uploaded / total) * 100);
    animateProgress(percent);
  }

  function renderMonitoring() {
    const data = getMonitoring();
    const searchVal = ($('#monitorSearch') as HTMLInputElement | null)?.value?.trim().toLowerCase() || '';
    const filtered = data.filter((b: RowObj) => {
      const matches = (b.name || '').toLowerCase().includes(searchVal);
      const passes = showUnchecked ? !b.checked : true;
      return matches && passes;
    });
    if (!monitoringBody) return;
    monitoringBody.innerHTML = filtered.map((b: RowObj, i: number) => `
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

    monitoringBody.querySelectorAll('input').forEach((inp: Element) => {
      const inputEl = inp as HTMLInputElement;
      inputEl.addEventListener('change', (ev: Event) => {
        const target = ev.target as HTMLInputElement;
        const index = Number(target.dataset.index);
        const field = target.dataset.field as keyof RowObj;
        const d = getMonitoring();
        const filteredNames = filtered.map((x: RowObj) => x.name);
        const targetName = filteredNames[index];
        const origIndex = d.findIndex((x: RowObj) => x.name === targetName);
        if (origIndex === -1) return;
        if (target.type === 'checkbox') d[origIndex][field] = target.checked as unknown as any;
        else d[origIndex][field] = target.value as any;
        saveMonitoring(d); updateMonitoringStats(); renderMonitoring();
      });
    });

    monitoringBody.querySelectorAll('.action-edit').forEach((btn: Element) => {
      btn.addEventListener('click', (ev: Event) => {
        const t = ev.currentTarget as HTMLElement;
        const i = Number(t.dataset.i);
        const d = getMonitoring();
        const filteredNames = filtered.map((x: RowObj) => x.name);
        const targetName = filteredNames[i];
        const origIndex = d.findIndex((x: RowObj) => x.name === targetName);
        if (origIndex === -1) return;
        const newName = prompt('Edit branch name:', d[origIndex].name);
        if (newName) { d[origIndex].name = newName; saveMonitoring(d); renderMonitoring(); showBanner('Branch updated.'); }
      });
    });

    monitoringBody.querySelectorAll('.action-delete').forEach((btn: Element) => {
      btn.addEventListener('click', (ev: Event) => {
        const t = ev.currentTarget as HTMLElement;
        const i = Number(t.dataset.i);
        const d = getMonitoring();
        const filteredNames = filtered.map((x: RowObj) => x.name);
        const targetName = filteredNames[i];
        const origIndex = d.findIndex((x: RowObj) => x.name === targetName);
        if (origIndex === -1) return;
        if (!confirm('Delete branch?')) return;
        d.splice(origIndex, 1); saveMonitoring(d); renderMonitoring(); showBanner('Branch deleted.');
      });
    });

    updateMonitoringStats();
  }

  const monitorSearchInput = $('#monitorSearch') as HTMLInputElement | null;
  const monitorFilterUnchecked = $('#filterUnchecked') as HTMLInputElement | null;
  if (monitorSearchInput) monitorSearchInput.addEventListener('input', () => renderMonitoring());
  if (monitorFilterUnchecked) monitorFilterUnchecked.addEventListener('change', () => { showUnchecked = !!monitorFilterUnchecked.checked; renderMonitoring(); });

  if (addBranchBtn) {
    addBranchBtn.addEventListener('click', () => {
      const name = prompt('Branch name:');
      if (!name) return;
      const d = getMonitoring();
      d.push({ name, checked: false, uploaded: false, uploadedBy: '', remarks: '' });
      saveMonitoring(d); renderMonitoring(); showBanner('Branch added.');
    });
  }

  if (clearMonitoringBtn) {
    clearMonitoringBtn.addEventListener('click', () => {
      if (!confirm('Clear all monitoring data?')) return;
      saveMonitoring([]); renderMonitoring(); showBanner('All monitoring data cleared.');
    });
  }

  if (exportMonitoringBtn) exportMonitoringBtn.addEventListener('click', () => {
   const data = getMonitoring();
   const month = monthSelect?.value || '';
   const year = yearSelect?.value || '';
   const percent = currentPercent || 0;
   const headerRows = [
     [`Monitoring Progress: ${percent}%`],
     [`Month: ${month} ${year}`],
     ['Branch Name', 'Checked', 'Uploaded', 'Uploaded By', 'Remarks']
   ];
   const rows = data.map((b: RowObj) => [
     b.name || '',
     b.checked ? 'Yes' : 'No',
     b.uploaded ? 'Yes' : 'No',
     b.uploadedBy || '',
     b.remarks || ''
   ]);
   const ws = XLSX.utils.aoa_to_sheet([...headerRows, ...rows]);
   const wb = XLSX.utils.book_new();
   XLSX.utils.book_append_sheet(wb, ws, 'Monitoring');
   XLSX.writeFile(wb, `Monitoring_${month}_${year}.xlsx`);
   showSuccess();
 });

  /***** Clear All insertion (schedule header) *****/
  (function insertClearAll() {
    const scheduleHeader = document.querySelector<HTMLElement>('#tab-schedule-content h1');
    if (!scheduleHeader) return;
    const btn = document.createElement('button');
    btn.id = 'clearAllBtn';
    btn.className = 'bg-blue-500 text-white px-3 py-1 rounded hover:bg-blue-600 shadow-sm transition';
    btn.textContent = '🧹 Clear All';
    btn.style.margin = '0.5rem auto 1rem';
    scheduleHeader.insertAdjacentElement('afterend', btn);
    btn.addEventListener('click', () => {
      if (!confirm('Are you sure you want to clear ALL schedule data (Work + Rest)?')) return;
      workScheduleData = []; restDayData = []; rejectedRows = [];
      if (workInput) workInput.value = '';
      if (restInput) restInput.value = '';
      const wbEl = $('#workBranchName') as HTMLInputElement | null;
      const rbEl = $('#restBranchName') as HTMLInputElement | null;
      if (wbEl) wbEl.value = ''; if (rbEl) rbEl.value = '';
      renderWorkTable(); renderRestTable(); hideBanner();
      if (summaryEl) { summaryEl.textContent = ''; summaryEl.classList.add('hidden'); }
      updateButtonStates(); showBanner('✅ All schedule data cleared.'); recheckConflicts(); saveState();
    });
  })();

  /***** Tab switching & scroll preservation *****/
  let lastScroll = 0;
  if (tabSchedule && tabMonitoring && scheduleContent && monitoringContent) {
    tabSchedule.addEventListener('click', () => {
      lastScroll = window.scrollY || 0;
      tabSchedule.classList.add('bg-indigo-600', 'text-white'); tabSchedule.classList.remove('bg-gray-200', 'text-gray-700');
      tabMonitoring.classList.remove('bg-indigo-600', 'text-white'); tabMonitoring.classList.add('bg-gray-200', 'text-gray-700');
      scheduleContent.classList.remove('hidden'); monitoringContent.classList.add('hidden');
      setTimeout(() => window.scrollTo(0, lastScroll), 120);
    });
    tabMonitoring.addEventListener('click', () => {
      lastScroll = window.scrollY || 0;
      tabMonitoring.classList.add('bg-indigo-600', 'text-white'); tabMonitoring.classList.remove('bg-gray-200', 'text-gray-700');
      tabSchedule.classList.remove('bg-indigo-600', 'text-white'); tabSchedule.classList.add('bg-gray-200', 'text-gray-700');
      scheduleContent.classList.add('hidden'); monitoringContent.classList.remove('hidden');
      renderMonitoring(); updateMonitoringStats();
      setTimeout(() => window.scrollTo(0, lastScroll), 120);
    });
  }

  /***** Back to top *****/
  if (backToTopBtn) {
    window.addEventListener('scroll', () => {
      if (window.scrollY > 400) backToTopBtn.classList.add('show'); else backToTopBtn.classList.remove('show');
    });
    backToTopBtn.addEventListener('click', () => window.scrollTo({ top: 0, behavior: 'smooth' }));
  }

  /***** Autosave / load state *****/
  function saveState() {
    try {
      localStorage.setItem('workScheduleData', JSON.stringify(workScheduleData));
      localStorage.setItem('restDayData', JSON.stringify(restDayData));
    } catch { /* ignore */ }
  }
  function loadState() {
    try {
      const w = JSON.parse(localStorage.getItem('workScheduleData') || '[]');
      const r = JSON.parse(localStorage.getItem('restDayData') || '[]');
      workScheduleData = Array.isArray(w) ? (w as RowObj[]) : [];
      restDayData = Array.isArray(r) ? (r as RowObj[]) : [];
      renderWorkTable(); renderRestTable(); recheckConflicts(); updateButtonStates();
    } catch { /* ignore parse errors */ }
  }

  /***** Initialization *****/
  loadState();
  renderWorkTable();
  renderRestTable();
  renderMonitoring();
  // start Firestore sync (best-effort, non-blocking)
  initMonitoringSync();

  // Expose utilities for debugging if needed
  (window as any).__scc = {
    getMonitoring, saveMonitoring, renderMonitoring, updateMonitoringStats,
    workScheduleData, restDayData, validateSchedules, undoPaste, undoDelete, saveState, loadState
  };
})();
