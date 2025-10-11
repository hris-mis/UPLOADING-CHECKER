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

  // 🧭 Save scroll position when switching tabs
let lastScroll = 0;


/***** Tabs *****/
tabSchedule.addEventListener('click', () => {
  lastScroll = window.scrollY; // save before switching
  tabSchedule.classList.add('bg-indigo-600', 'text-white');
  tabMonitoring.classList.remove('bg-indigo-600', 'text-white');
  tabMonitoring.classList.add('bg-gray-200', 'text-gray-700');
  scheduleContent.classList.remove('hidden');
  monitoringContent.classList.add('hidden');
});

tabMonitoring.addEventListener('click', () => {
  tabMonitoring.classList.add('bg-indigo-600', 'text-white');
  tabSchedule.classList.remove('bg-indigo-600', 'text-white');
  tabSchedule.classList.add('bg-gray-200', 'text-gray-700');
  scheduleContent.classList.add('hidden');
  monitoringContent.classList.remove('hidden');
  renderMonitoring();
  updateMonitoringStats();
  setTimeout(() => window.scrollTo(0, lastScroll), 100); // restore position
});


  /***** Helpers *****/
  function showBanner(msg) {
    warningBanner.textContent = msg;
    warningBanner.classList.remove('hidden');
    warningBanner.classList.add('opacity-100');
    setTimeout(() => {
      warningBanner.classList.remove('opacity-100');
      setTimeout(() => warningBanner.classList.add('hidden'), 400);
    }, 3000);
  }

  function showSuccess() {
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

  function saveMonitoring(d) {
    localStorage.setItem('monitoringData', JSON.stringify(d));
  }

function updateMonitoringStats() {
  const data = getMonitoring();
  const total = data.length;
  const checked = data.filter(b => b.checked).length;
  const uploaded = data.filter(b => b.uploaded).length;

  totalBranchesEl.textContent = total;
  checkedBranchesEl.textContent = checked;
  uploadedBranchesEl.textContent = uploaded;

  const percent = total === 0 ? 0 : Math.round((uploaded / total) * 100);
  animateProgress(percent);
}

// 🌀 Animate percent counting up
let currentPercent = 0;
function animateProgress(target) {
  const duration = 600;
  const start = performance.now();
  const from = currentPercent;
  const diff = target - from;

  function frame(time) {
    const progress = Math.min((time - start) / duration, 1);
    const value = Math.round(from + diff * progress);
    progressPercentEl.textContent = value + '%';
    progressBar.style.width = value + '%';
    if (value < 40)
      progressBar.style.background = 'linear-gradient(to right, #f43f5e, #fb7185)';
    else if (value < 80)
      progressBar.style.background = 'linear-gradient(to right, #fbbf24, #facc15)';
    else
      progressBar.style.background = 'linear-gradient(to right, #10b981, #34d399)';

    if (progress < 1) requestAnimationFrame(frame);
    else currentPercent = target;
  }
  requestAnimationFrame(frame);
}


  function renderMonitoring() {
    const data = getMonitoring();
    const searchVal = (document.getElementById('monitorSearch') || {}).value || '';
    const filtered = data.filter(b => b.name.toLowerCase().includes(searchVal.toLowerCase()));

    monitoringBody.innerHTML = filtered.map((b,i)=>`
      <tr class="hover:bg-gray-50 transition">
        <td class="text-left p-2">${escapeHtml(b.name)}</td>
        <td><input type="checkbox" data-index="${i}" data-field="checked" ${b.checked ? 'checked' : ''}></td>
        <td><input type="checkbox" data-index="${i}" data-field="uploaded" ${b.uploaded ? 'checked' : ''}></td>
        <td><input type="text" data-index="${i}" data-field="uploadedBy" value="${escapeHtml(b.uploadedBy)}" class="border rounded px-1 py-0.5 w-28"></td>
        <td><input type="text" data-index="${i}" data-field="remarks" value="${escapeHtml(b.remarks)}" class="border rounded px-1 py-0.5 w-28"></td>
        <td>
          <span class="cursor-pointer text-blue-500 action-edit" title="Edit" data-i="${i}">✏️</span>
          <span class="cursor-pointer text-red-500 action-delete" title="Delete" data-i="${i}">❌</span>
        </td>
      </tr>
    `).join('');

    monitoringBody.querySelectorAll('input').forEach(inp => {
      inp.addEventListener('change', (e)=>{
        const index = Number(e.target.dataset.index);
        const field = e.target.dataset.field;
        const d = getMonitoring();
        d[index][field] = e.target.type === 'checkbox' ? e.target.checked : e.target.value;
        saveMonitoring(d);
        updateMonitoringStats();
      });
    });

    monitoringBody.querySelectorAll('.action-edit').forEach(btn=>{
      btn.addEventListener('click', (e)=>{
        const i = Number(e.target.dataset.i);
        const d = getMonitoring();
        const newName = prompt('Edit branch name:', d[i].name);
        if (newName) { d[i].name = newName; saveMonitoring(d); renderMonitoring(); showBanner('Branch updated.'); }
      });
    });

    monitoringBody.querySelectorAll('.action-delete').forEach(btn=>{
      btn.addEventListener('click', (e)=>{
        const i = Number(e.target.dataset.i);
        if (!confirm('Delete branch?')) return;
        const d = getMonitoring();
        d.splice(i,1); saveMonitoring(d); renderMonitoring(); showBanner('Branch deleted.');
      });
    });

    updateMonitoringStats();
  }

  /***** Add Branch *****/
  const addBranchBtn = document.getElementById('addBranchBtn');
  if (addBranchBtn) {
    addBranchBtn.addEventListener('click', ()=>{
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
  clearMonitoringBtn?.addEventListener('click', ()=>{
    if (!confirm('Clear all monitoring data?')) return;
    saveMonitoring([]);
    renderMonitoring();
    showBanner('All monitoring data cleared.');
  });

  exportMonitoringBtn?.addEventListener('click', ()=>{
    const data = getMonitoring();
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Branch Monitoring');
    XLSX.writeFile(wb, `Monitoring_${monthSelect.value}_${yearSelect.value}.xlsx`);
    showSuccess();
  });
  // 🧭 Floating "Back to Top" Button Logic
const backToTopBtn = document.getElementById('backToTopBtn');
window.addEventListener('scroll', () => {
  if (window.scrollY > 400) backToTopBtn.classList.add('show');
  else backToTopBtn.classList.remove('show');
});


backToTopBtn.addEventListener('click', () => {
  window.scrollTo({ top: 0, behavior: 'smooth' });
});


  /***** Initialize *****/
  renderMonitoring();
  updateMonitoringStats();
});
