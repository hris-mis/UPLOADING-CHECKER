(function(){
  let workScheduleData = [];
  let restDayData = [];
  let undoStack = { work: [], rest: [] };
  let deleteStack = { work: [], rest: [] };
  const LEADERSHIP_POSITIONS = ['Branch Head','Site Supervisor','OIC'];

  const workInput = document.getElementById('workScheduleInput');
  const restInput = document.getElementById('restScheduleInput');
  const workTableBody = document.getElementById('workTableBody');
  const restTableBody = document.getElementById('restTableBody');
  const summaryEl = document.getElementById('summary');
  const warningBanner = document.getElementById('warning-banner');
  const successMsg = document.getElementById('success-message');
  const rejectedModal = document.getElementById('rejectedModal');
  const rejectedBody = document.getElementById('rejectedBody');

  const generateWorkFileBtn = document.getElementById('generateWorkFile');
  const generateRestFileBtn = document.getElementById('generateRestFile');

  // Monitoring
  const tabSchedule = document.getElementById('tab-schedule');
  const tabMonitoring = document.getElementById('tab-monitoring');
  const scheduleContent = document.getElementById('tab-schedule-content');
  const monitoringContent = document.getElementById('tab-monitoring-content');
  const monitoringBody = document.getElementById('monitoringBody');
  const monitorSearch = document.getElementById('monitorSearch');
  const addBranchBtn = document.getElementById('addBranchBtn');
  const exportMonitoringBtn = document.getElementById('exportMonitoringBtn');
  const totalBranchesEl = document.getElementById('totalBranches');
  const checkedBranchesEl = document.getElementById('checkedBranches');
  const uploadedBranchesEl = document.getElementById('uploadedBranches');
  const progressPercentEl = document.getElementById('progressPercent');

  function showBanner(msg){ warningBanner.textContent=msg; warningBanner.classList.remove('hidden'); setTimeout(()=>warningBanner.classList.add('hidden'),3000); }
  function showSuccess(){ successMsg.classList.remove('hidden'); setTimeout(()=>successMsg.classList.add('hidden'),2000); }

  function parseTabular(text){
    if(!text) return [];
    const lines = text.replace(/\r/g,'').split('\n').map(l=>l.trim()).filter(l=>l && !/^(sheet|page|total|subtotal)/i.test(l));
    if(lines.length===0) return [];
    const sample = lines.slice(0,5).join('\\n');
    let splitter = /\\t/;
    if(!/\\t/.test(sample)){ if(/,/.test(sample)) splitter = /,/; else splitter = /\\s{2,}/; }
    return lines.map(line=>line.split(splitter).map(c=>c.trim()));
  }

  function detectHeaderAndMap(rows){
    if(!rows||rows.length===0) return {headerIndex:-1,dataRows:[],colMap:{}};
    const headerMap = {
      name:/name|employee\\s*name|full\\s*name|staff\\s*name/,
      empNo:/emp.*no|employee.*no|emp\\s*id|emp#|id|employee\\s*number/,
      date:/date|work.*date|rest.*date|day\\s*date/,
      shift:/shift/,
      day:/day|dow/,
      position:/position|pos|role|title/
    };
    let bestIdx=0,bestScore=-1;
    const maxHeader = Math.min(6,rows.length);
    for(let i=0;i<maxHeader;i++){
      const row = rows[i].map(c=>(c||'').toLowerCase());
      let score=0;
      for(const cell of row){ for(const k in headerMap) if(headerMap[k].test(cell)) score++; }
      if(score>bestScore){ bestScore=score; bestIdx=i; }
    }
    const headerRow = rows[bestIdx].map(c=>(c||'').toLowerCase());
    const colMap = {};
    headerRow.forEach((h,idx)=>{ for(const key in headerMap){ if(headerMap[key].test(h)){ colMap[key]=idx; break; } } });
    if(colMap.empNo===undefined) colMap.empNo=1;
    if(colMap.name===undefined) colMap.name=0;
    if(colMap.date===undefined) colMap.date=2;
    if(colMap.shift===undefined) colMap.shift=3;
    if(colMap.day===undefined) colMap.day=4;
    if(colMap.position===undefined) colMap.position=5;
    const dataRows = rows.slice(bestIdx+1).filter(r=>r.some(c=>c&&c.toString().trim()!=='') );
    return {headerIndex:bestIdx,dataRows,colMap};
  }

  function normalizeDate(s){ if(!s) return ''; s=(''+s).trim(); if(!isNaN(s) && Number(s)>10000){ const excelEpoch=new Date(1899,11,30); const parsed=new Date(excelEpoch.getTime() + Number(s)*86400000); return `${String(parsed.getMonth()+1).padStart(2,'0')}/${String(parsed.getDate()).padStart(2,'0')}/${String(parsed.getFullYear()).slice(2)}`; } const m=s.match(/^(\\d{1,2})[\\/\\.\\-](\\d{1,2})[\\/\\.\\-](\\d{2,4})$/); if(m){ let mm=m[1].padStart(2,'0'), dd=m[2].padStart(2,'0'), yy=m[3]; if(yy.length===4) yy=yy.slice(2); return `${mm}/${dd}/${yy}`; } const p=new Date(s); if(!isNaN(p)){ return `${String(p.getMonth()+1).padStart(2,'0')}/${String(p.getDate()).padStart(2,'0')}/${String(p.getFullYear()).slice(2)}`; } return s; }

  function isNumericStr(s){ return /^\\d+$/.test((s||'').toString().trim()); }
  function escapeHtml(str){ return (str==null)?'':String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;'); }

  function renderWorkTable(){ workTableBody.innerHTML = workScheduleData.map((d,i)=>`<tr data-idx="${i}"><td class="table-fixed">${escapeHtml(d.name)}</td><td>${escapeHtml(d.empNo)}</td><td>${escapeHtml(d.date)}</td><td>${escapeHtml(d.shift)}</td><td>${escapeHtml(d.day)}</td><td>${escapeHtml(d.position)}</td><td><button class="delete-row-btn smallbtn" data-type="work" data-idx="${i}">❌</button></td></tr>`).join(''); }
  function renderRestTable(){ restTableBody.innerHTML = restDayData.map((d,i)=>{ const conflicts=(d.conflicts||[]).map(c=>`<div><strong>${escapeHtml(c.type)}:</strong> ${escapeHtml(c.reason)}</div>`).join(''); const rowClass=(d.conflicts&&d.conflicts.length>0)?'bg-yellow-50':''; return `<tr class="${rowClass}" data-idx="${i}"><td>${conflicts||''}</td><td>${escapeHtml(d.name)}</td><td>${escapeHtml(d.empNo)}</td><td>${escapeHtml(d.date)}</td><td>${escapeHtml(d.day)}</td><td>${escapeHtml(d.position)}</td><td><button class="delete-row-btn smallbtn" data-type="rest" data-idx="${i}">❌</button></td></tr>`; }).join(''); updateSummary(); }
  function updateSummary(){ const total=restDayData.length; const conflicts=restDayData.filter(r=>r.conflicts&&r.conflicts.length>0).length; if(total===0){ summaryEl.textContent=''; } else { summaryEl.textContent = conflicts===0?`✅ No conflicts detected for ${total} entries.`:`${conflicts} out of ${total} entries have conflicts detected.`; } updateButtonStates(); }
  function updateButtonStates(){ document.getElementById('generateWorkFile').disabled = workScheduleData.length===0; document.getElementById('generateRestFile').disabled = restDayData.length===0 || restDayData.some(r=>r.conflicts&&r.conflicts.length>0); }

  function validateSchedules(){ restDayData.forEach(r=>r.conflicts=[]); const workMap = new Map(workScheduleData.map(w=>[`${w.empNo}-${(w.date||'').trim()}`,w])); const workEmpSet=new Set(workScheduleData.map(w=>w.empNo)); const restByDate={}; const weekendCount={}; const seen=new Set(); restDayData.forEach(rd=>{ if(!rd) return; const key=`${rd.empNo}-${(rd.date||'').trim()}`; if(!restByDate[rd.date]) restByDate[rd.date]=[]; restByDate[rd.date].push(rd); if(seen.has(key)) rd.conflicts.push({type:'Duplicate Entry',reason:'Duplicate rest day entry for same employee & date.'}); else seen.add(key); if(!workEmpSet.has(rd.empNo)) rd.conflicts.push({type:'Missing Employee',reason:'Employee not found in Work Schedule data.'}); const d=new Date(rd.date); if(isNaN(d.getTime())) rd.conflicts.push({type:'Invalid Date Format',reason:'Date format unrecognized.'}); if(workMap.has(key)) rd.conflicts.push({type:'Work Conflict',reason:'Employee has a work schedule on same date.'}); if(/saturday|sunday/i.test(rd.day||'')){ if(!isNaN(d.getTime())){ const monthYear=`${String(d.getMonth()+1).padStart(2,'0')}/${d.getFullYear()}`; const wkKey=`${rd.empNo}-${monthYear}`; weekendCount[wkKey]=(weekendCount[wkKey]||0)+1; } } }); for(const date in restByDate){ const leaders = restByDate[date].filter(r=>LEADERSHIP_POSITIONS.includes(r.position)); if(leaders.length>1) leaders.forEach(l=>l.conflicts.push({type:'Leadership Conflict',reason:'Multiple leaders have same rest day.'})); } for(const wkey in weekendCount){ if(weekendCount[wkey]>2){ const emp=wkey.split('-')[0]; restDayData.filter(r=>r.empNo===emp && /saturday|sunday/i.test(r.day||'')).forEach(r=>r.conflicts.push({type:'Weekend Limit Exceeded',reason:`${weekendCount[wkey]} weekend rest days — maximum 2.`})); } } renderRestTable(); }

  function handlePaste(e,type){ e.preventDefault(); const pasted=(e.clipboardData||window.clipboardData).getData('text/plain'); if(!pasted) return; const parsed=parseTabular(pasted); const {headerIndex,dataRows,colMap}=detectHeaderAndMap(parsed); const cleaned=[]; const rejected=[]; dataRows.forEach(row=>{ const obj={ name: row[colMap.name]||'', empNo: row[colMap.empNo]||'', date: normalizeDate(row[colMap.date]||''), shift: row[colMap.shift]||'', day: row[colMap.day]||'', position: row[colMap.position]||'' }; const reasons=[]; if(!obj.empNo||!isNumericStr(obj.empNo)) reasons.push('Missing or non-numeric Employee No'); if(!obj.date||obj.date==='') reasons.push('Invalid or missing Date'); if(reasons.length) rejected.push({row:row.join(' | '),reasons}); else cleaned.push(obj); }); undoStack[type].push({work:JSON.parse(JSON.stringify(workScheduleData)),rest:JSON.parse(JSON.stringify(restDayData))}); if(type==='work'){ workScheduleData=cleaned; renderWorkTable(); workInput.value=''; showBanner(`${cleaned.length} work rows pasted. ${rejected.length?rejected.length+' rejected':''}`); } else { restDayData=cleaned; validateSchedules(); restInput.value=''; showBanner(`${cleaned.length} rest rows pasted. ${rejected.length?rejected.length+' rejected':''}`); } if(rejected.length) showRejectedModal(rejected); updateButtonStates(); }

  workInput.addEventListener('paste',(e)=>handlePaste(e,'work'));
  restInput.addEventListener('paste',(e)=>handlePaste(e,'rest'));

  document.addEventListener('click',(ev)=>{ const btn = ev.target.closest && ev.target.closest('.delete-row-btn'); if(btn){ const type=btn.dataset.type, idx=Number(btn.dataset.idx); if(type==='work'){ const rem = workScheduleData.splice(idx,1)[0]; deleteStack.work.push({item:rem,idx}); renderWorkTable(); showBanner('Row deleted. Undo available.'); } else { const rem = restDayData.splice(idx,1)[0]; deleteStack.rest.push({item:rem,idx}); renderRestTable(); showBanner('Row deleted. Undo available.'); } updateButtonStates(); } });

  function undoDelete(type){ const stack=deleteStack[type]; if(!stack||!stack.length) return showBanner('Nothing to undo'); const last=stack.pop(); if(type==='work'){ workScheduleData.splice(last.idx,0,last.item); renderWorkTable(); } else { restDayData.splice(last.idx,0,last.item); renderRestTable(); } showBanner('Undo delete done'); }
  function undoPaste(type){ const stack=undoStack[type]; if(!stack||!stack.length) return showBanner('Nothing to undo'); const snap=stack.pop(); workScheduleData=snap.work; restDayData=snap.rest; renderWorkTable(); renderRestTable(); showBanner('Undo paste restored previous data'); }
  window._undoDelete = undoDelete; window._undoPaste = undoPaste;

  function showRejectedModal(rejected){ rejectedBody.innerHTML = `<p>${rejected.length} rejected rows:</p>` + rejected.map(r=>`<div style="padding:6px;border-bottom:1px solid #eee;"><strong>${escapeHtml(r.row)}</strong><div style="color:#b91c1c;margin-top:4px;">Reasons: ${r.reasons.join(', ')}</div></div>`).join(''); rejectedModal.classList.remove('hidden'); rejectedModal.style.display='flex'; }
  document.addEventListener('click',(e)=>{ if(e.target.matches('.modal-close') || e.target.matches('#rejectedModal .modal-backdrop')){ rejectedModal.classList.add('hidden'); rejectedModal.style.display='none'; } });

  function generateHrisFile(type){ if(type==='work'){ const branch=document.getElementById('workBranchName').value.trim(); if(!branch) return showBanner('Enter Work Branch Name'); if(workScheduleData.length===0) return showBanner('No work data'); const data=[['Employee Number','Work Date','Shift Code'], ...workScheduleData.map(r=>[r.empNo,r.date,r.shift])]; const ws=XLSX.utils.aoa_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'HRIS Upload'); XLSX.writeFile(wb,`${branch}_WORK_SCHEDULE.xlsx`); showSuccess(); } else { const branch=document.getElementById('restBranchName').value.trim(); if(!branch) return showBanner('Enter Rest Branch Name'); if(restDayData.length===0) return showBanner('No rest data'); if(restDayData.some(r=>r.conflicts&&r.conflicts.length>0)) return showBanner('Resolve conflicts first'); const data=restDayData.map(r=>({'Employee No':r.empNo,'Rest Day Date':r.date})); const ws=XLSX.utils.json_to_sheet(data); const wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,'HRIS Upload'); XLSX.writeFile(wb,`${branch}_REST_DAY_UPLOAD.xlsx`); showSuccess(); } }
  document.getElementById('generateWorkFile').addEventListener('click',()=>generateHrisFile('work'));
  document.getElementById('generateRestFile').addEventListener('click',()=>generateHrisFile('rest'));

  const defaultMonitoring = [ {name:'AASP ABREEZA',checked:false,uploaded:false,uploadedBy:'',remarks:''}, {name:'AASP NES - ATLAS',checked:false,uploaded:false,uploadedBy:'',remarks:''} ];
  function getMonitoring(){ const d=localStorage.getItem('monitoringData'); if(d){ try{return JSON.parse(d);}catch(e){return [...defaultMonitoring];}} return [...defaultMonitoring]; }
  function saveMonitoring(d){ localStorage.setItem('monitoringData',JSON.stringify(d)); }

  function renderMonitoring(){ const data=getMonitoring(); const q=(monitorSearch.value||'').toLowerCase(); const filtered = data.filter(b=>b.name.toLowerCase().includes(q)); monitoringBody.innerHTML = filtered.map((b,i)=>`<tr><td class="p-2">${escapeHtml(b.name)}</td><td><input type="checkbox" data-i="${i}" data-field="checked" ${b.checked?'checked':''}></td><td><input type="checkbox" data-i="${i}" data-field="uploaded" ${b.uploaded?'checked':''}></td><td><input type="text" data-i="${i}" data-field="uploadedBy" value="${escapeHtml(b.uploadedBy)}" class="px-1 py-0.5"></td><td><input type="text" data-i="${i}" data-field="remarks" value="${escapeHtml(b.remarks)}" class="px-1 py-0.5"></td><td><button class="edit-branch smallbtn" data-i="${i}">✏️</button> <button class="delete-branch smallbtn" data-i="${i}">❌</button></td></tr>`).join(''); monitoringBody.querySelectorAll('input').forEach(inp=>{ inp.addEventListener('change',(e)=>{ const idx=Number(e.target.dataset.i); const field=e.target.dataset.field; const d=getMonitoring(); d[idx][field] = e.target.type==='checkbox'?e.target.checked:e.target.value; saveMonitoring(d); updateMonitoringStats(); }); }); monitoringBody.querySelectorAll('.edit-branch').forEach(btn=> btn.addEventListener('click',(e)=>{ const i=Number(e.target.dataset.i); const d=getMonitoring(); const name=prompt('Edit branch name:', d[i].name); if(name){ d[i].name=name; saveMonitoring(d); renderMonitoring(); showBanner('Branch updated'); } })); monitoringBody.querySelectorAll('.delete-branch').forEach(btn=> btn.addEventListener('click',(e)=>{ const i=Number(e.target.dataset.i); if(!confirm('Delete branch?')) return; const d=getMonitoring(); d.splice(i,1); saveMonitoring(d); renderMonitoring(); showBanner('Branch deleted'); })); updateMonitoringStats(); }
  function updateMonitoringStats(){ const d=getMonitoring(); totalBranchesEl.textContent=d.length; checkedBranchesEl.textContent=d.filter(x=>x.checked).length; uploadedBranchesEl.textContent=d.filter(x=>x.uploaded).length; const pct = d.length? Math.round((d.filter(x=>x.uploaded).length/d.length)*100):0; progressPercentEl.textContent=pct+'%'; }
  addBranchBtn.addEventListener('click',()=>{ const name=prompt('Branch name:'); if(!name) return; const d=getMonitoring(); d.push({name,checked:false,uploaded:false,uploadedBy:'',remarks:''}); saveMonitoring(d); renderMonitoring(); showBanner('Branch added'); });
  monitorSearch.addEventListener('input', renderMonitoring);

  function showTab(tab){ if(tab==='monitoring'){ scheduleContent.classList.add('hidden'); monitoringContent.classList.remove('hidden'); tabMonitoring.classList.add('bg-sky-600','text-white'); tabSchedule.classList.remove('bg-sky-600','text-white'); localStorage.setItem('lastTab','monitoring'); renderMonitoring(); } else { scheduleContent.classList.remove('hidden'); monitoringContent.classList.add('hidden'); tabSchedule.classList.add('bg-white'); tabMonitoring.classList.remove('bg-sky-600','text-white'); localStorage.setItem('lastTab','schedule'); } }
  tabSchedule.addEventListener('click',()=>showTab('schedule')); tabMonitoring.addEventListener('click',()=>showTab('monitoring'));
  const last = localStorage.getItem('lastTab')||'schedule'; showTab(last);

  document.addEventListener('click',(e)=>{ if(e.target.matches('.modal-close')){ rejectedModal.classList.add('hidden'); rejectedModal.style.display='none'; } });

  renderWorkTable(); renderRestTable(); updateButtonStates(); renderMonitoring();

  window._app = { undoDelete, undoPaste };

})();