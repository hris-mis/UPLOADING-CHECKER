// src/mainscript.ts
/* Main UI + logic for Schedule Checker & Monitoring */
/* Assumes firebase.ts exports: subscribeToData, subscribeToDoc, setSharedDoc */

declare const XLSX: any;

import { subscribeToData, subscribeToDoc, setSharedDoc } from "./firebase";

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

//
// ---------------------- Helpers (defined before use) ----------------------
//

function escapeHtml(str: string | undefined): string {
  if (str == null) return "";
  return String(str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function showBanner(msg: string) {
  const warningBanner = document.getElementById("warning-banner");
  if (!warningBanner) return;
  warningBanner.textContent = msg;
  warningBanner.classList.remove("hidden");
  warningBanner.classList.add("opacity-100");
  setTimeout(() => {
    warningBanner.classList.remove("opacity-100");
    setTimeout(() => warningBanner.classList.add("hidden"), 400);
  }, 3000);
}

function hideBanner() {
  const warningBanner = document.getElementById("warning-banner");
  if (!warningBanner) return;
  warningBanner.classList.add("hidden");
  warningBanner.classList.remove("opacity-100");
  warningBanner.textContent = "";
}

function showSuccess() {
  const successMsg = document.getElementById("success-message");
  if (!successMsg) return;
  successMsg.classList.remove("hidden");
  successMsg.style.animation = "fadeInOut 2.5s ease-in-out";
  setTimeout(() => {
    successMsg.classList.add("hidden");
    successMsg.style.animation = "";
  }, 2500);
}

//
// ---------------------- initApp (main) ----------------------
//
function initApp(): void {
  // small selector helpers
  const $ = <T extends Element = Element>(sel: string): T | null =>
    document.querySelector<T>(sel);
  const $$ = <T extends Element = Element>(sel: string): NodeListOf<T> =>
    document.querySelectorAll<T>(sel);

  const pad2 = (v: string | number) => {
    const s = String(v);
    return s.length >= 2 ? s : "0" + s;
  };

  // State
  let workScheduleData: RowObj[] = [];
  let restDayData: RowObj[] = [];
  let currentPercent = 0;
  let showUnchecked = false;

  // UI refs (guarded)
  const scheduleContent = $("#tab-schedule-content") as HTMLElement | null;
  const monitoringContent = $("#tab-monitoring-content") as HTMLElement | null;

  // ----- TAB SWITCHING (use IDs present in HTML) -----
  const tabScheduleBtn = document.getElementById("tab-schedule") as HTMLButtonElement | null;
  const tabMonitoringBtn = document.getElementById("tab-monitoring") as HTMLButtonElement | null;

  function activateTab(tab: "schedule" | "monitoring") {
    if (!tabScheduleBtn || !tabMonitoringBtn || !scheduleContent || !monitoringContent) return;
    if (tab === "schedule") {
      tabScheduleBtn.classList.add("bg-indigo-600", "text-white");
      tabScheduleBtn.classList.remove("bg-gray-200", "text-gray-700");
      tabMonitoringBtn.classList.remove("bg-indigo-600", "text-white");
      tabMonitoringBtn.classList.add("bg-gray-200", "text-gray-700");
      scheduleContent.classList.remove("hidden");
      monitoringContent.classList.add("hidden");
    } else {
      tabMonitoringBtn.classList.add("bg-indigo-600", "text-white");
      tabMonitoringBtn.classList.remove("bg-gray-200", "text-gray-700");
      tabScheduleBtn.classList.remove("bg-indigo-600", "text-white");
      tabScheduleBtn.classList.add("bg-gray-200", "text-gray-700");
      monitoringContent.classList.remove("hidden");
      scheduleContent.classList.add("hidden");
      // when opening monitoring ensure render and stats are up-to-date
      renderMonitoring();
      updateMonitoringStats();
    }
  }

  if (tabScheduleBtn && tabMonitoringBtn) {
    tabScheduleBtn.addEventListener("click", () => {
      console.log("🗓️ Schedule tab clicked");
      activateTab("schedule");
      // restore scroll if needed
    });
    tabMonitoringBtn.addEventListener("click", () => {
      console.log("📊 Monitoring tab clicked");
      activateTab("monitoring");
    });
  } else {
    console.warn("Tab elements missing - check HTML IDs");
  }

  // set default
  activateTab("schedule");

  // ----- Monitoring DOM refs -----
  const monitoringBody = $("#monitoringBody") as HTMLElement | null;
  const totalBranchesEl = $("#totalBranches") as HTMLElement | null;
  const checkedBranchesEl = $("#checkedBranches") as HTMLElement | null;
  const uploadedBranchesEl = $("#uploadedBranches") as HTMLElement | null;
  const progressBar = $("#progressBar") as HTMLElement | null;
  const progressPercentEl = $("#progressPercent") as HTMLElement | null;
  const monitorSearchInput = $("#monitorSearch") as HTMLInputElement | null;
  const monitorFilterUnchecked = $("#filterUnchecked") as HTMLInputElement | null;
  const clearMonitoringBtn = $("#clearMonitoringBtn") as HTMLButtonElement | null;
  const exportMonitoringBtn = $("#exportMonitoringBtn") as HTMLButtonElement | null;
  const addBranchBtn = $("#addBranchBtn") as HTMLElement | null;
  const monthSelect = $("#monthSelect") as HTMLSelectElement | null;
  const yearSelect = $("#yearSelect") as HTMLSelectElement | null;

  // ----- Local storage + shared Firestore helpers -----
  const defaultMonitoring: RowObj[] = [
    { name: "AASP ABREEZA", checked: false, uploaded: false, uploadedBy: "", remarks: "" },
    { name: "AASP NES - ATLAS", checked: false, uploaded: false, uploadedBy: "", remarks: "" },
  ];

  function getMonitoring(): RowObj[] {
    try {
      const raw = localStorage.getItem("monitoringData");
      if (!raw) return JSON.parse(JSON.stringify(defaultMonitoring));
      return JSON.parse(raw) as RowObj[];
    } catch {
      return JSON.parse(JSON.stringify(defaultMonitoring));
    }
  }

  async function saveMonitoring(d: RowObj[]) {
    try {
      localStorage.setItem("monitoringData", JSON.stringify(d));
      // attempt to save shared doc (best-effort)
      try {
        await setSharedDoc("monitoring", "sharedMonitoringData", { data: d, updatedAt: new Date() }, true);
      } catch (e) {
        console.warn("setSharedDoc failed:", e);
      }
    } catch (e) {
      console.error("saveMonitoring error", e);
    }
  }

  function animateProgress(target: number) {
    const duration = 600;
    const from = currentPercent || 0;
    const diff = target - from;
    const start = performance.now();
    function frame(now: number) {
      const t = Math.min((now - start) / duration, 1);
      const value = Math.round(from + diff * t);
      if (progressPercentEl) progressPercentEl.textContent = `${value}%`;
      if (progressBar) progressBar.style.width = `${value}%`;
      if (progressBar) {
        if (value < 40) progressBar.style.background = "linear-gradient(to right, #f43f5e, #fb7185)";
        else if (value < 80) progressBar.style.background = "linear-gradient(to right, #fbbf24, #facc15)";
        else progressBar.style.background = "linear-gradient(to right, #10b981, #34d399)";
      }
      if (t < 1) requestAnimationFrame(frame);
      else currentPercent = target;
    }
    requestAnimationFrame(frame);
  }

  function updateMonitoringStats() {
    const data = getMonitoring();
    const total = data.length;
    const checked = data.filter((b) => !!b.checked).length;
    const uploaded = data.filter((b) => !!b.uploaded).length;
    if (totalBranchesEl) totalBranchesEl.textContent = String(total);
    if (checkedBranchesEl) checkedBranchesEl.textContent = String(checked);
    if (uploadedBranchesEl) uploadedBranchesEl.textContent = String(uploaded);
    const percent = total === 0 ? 0 : Math.round((checked / total) * 100);
    animateProgress(percent);
  }

  function renderMonitoring() {
    const data = getMonitoring();
    const searchVal = (monitorSearchInput?.value || "").trim().toLowerCase();
    const filtered = data.filter((b) => {
      const matches = (b.name || "").toLowerCase().includes(searchVal);
      const passes = showUnchecked ? !b.checked : true;
      return matches && passes;
    });

    if (!monitoringBody) return;
    monitoringBody.innerHTML = filtered
      .map(
        (b, i) => `
      <tr class="hover:bg-gray-50 transition">
        <td class="text-left p-2">${escapeHtml(b.name || "")}</td>
        <td><input type="checkbox" data-index="${i}" data-field="checked" ${b.checked ? "checked" : ""}></td>
        <td><input type="checkbox" data-index="${i}" data-field="uploaded" ${b.uploaded ? "checked" : ""}></td>
        <td><input type="text" data-index="${i}" data-field="uploadedBy" value="${escapeHtml(
          b.uploadedBy || ""
        )}" class="border rounded px-1 py-0.5 w-28"></td>
        <td><input type="text" data-index="${i}" data-field="remarks" value="${escapeHtml(
          b.remarks || ""
        )}" class="border rounded px-1 py-0.5 w-28"></td>
        <td>
          <span class="action-edit" title="Edit" data-i="${i}" style="cursor:pointer;margin-right:8px;">✏️</span>
          <span class="action-delete" title="Delete" data-i="${i}" style="cursor:pointer;">❌</span>
        </td>
      </tr>`
      )
      .join("");

    // wire input change events (work with filtered -> map to original via name)
    monitoringBody.querySelectorAll("input").forEach((inp) => {
      const inputEl = inp as HTMLInputElement;
      inputEl.onchange = (ev: Event) => {
        const target = ev.target as HTMLInputElement;
        const index = Number(target.dataset.index);
        const field = target.dataset.field as keyof RowObj;
        const d = getMonitoring();
        const filteredNames = filtered.map((x) => x.name || "");
        const targetName = filteredNames[index];
        const origIndex = d.findIndex((x) => x.name === targetName);
        if (origIndex === -1) return;
        if (target.type === "checkbox") d[origIndex][field] = target.checked as unknown as any;
        else d[origIndex][field] = target.value as any;
        saveMonitoring(d);
        updateMonitoringStats();
        renderMonitoring();
      };
    });

    // edit/delete actions
    monitoringBody.querySelectorAll(".action-edit").forEach((btn) => {
      (btn as HTMLElement).onclick = (ev) => {
        const t = ev.currentTarget as HTMLElement;
        const i = Number(t.dataset.i);
        const d = getMonitoring();
        const filteredNames = filtered.map((x) => x.name || "");
        const targetName = filteredNames[i];
        const origIndex = d.findIndex((x) => x.name === targetName);
        if (origIndex === -1) return;
        const newName = prompt("Edit branch name:", d[origIndex].name || "");
        if (newName && newName.trim()) {
          d[origIndex].name = newName.trim();
          saveMonitoring(d);
          renderMonitoring();
          showBanner("Branch updated.");
        }
      };
    });

    monitoringBody.querySelectorAll(".action-delete").forEach((btn) => {
      (btn as HTMLElement).onclick = (ev) => {
        const t = ev.currentTarget as HTMLElement;
        const i = Number(t.dataset.i);
        const d = getMonitoring();
        const filteredNames = filtered.map((x) => x.name || "");
        const targetName = filteredNames[i];
        const origIndex = d.findIndex((x) => x.name === targetName);
        if (origIndex === -1) return;
        if (!confirm("Delete branch?")) return;
        d.splice(origIndex, 1);
        saveMonitoring(d);
        renderMonitoring();
        showBanner("Branch deleted.");
      };
    });

    updateMonitoringStats();
  }

  // ----- Firestore realtime: subscribe to 'branches' collection for progress cards -----
  subscribeToData("branches", (branches: any[]) => {
    try {
      if (!branches || branches.length === 0) {
        if (progressBar) progressBar.style.width = "0%";
        if (progressPercentEl) progressPercentEl.textContent = "0%";
        if (totalBranchesEl) totalBranchesEl.textContent = "0";
        if (checkedBranchesEl) checkedBranchesEl.textContent = "0";
        if (uploadedBranchesEl) uploadedBranchesEl.textContent = "0";
        return;
      }

      const checked = branches.filter((b: any) => !!b.checked).length;
      const uploaded = branches.filter((b: any) => !!b.uploaded).length;
      const percent = Math.round((checked / branches.length) * 100);

      if (totalBranchesEl) totalBranchesEl.textContent = String(branches.length);
      if (checkedBranchesEl) checkedBranchesEl.textContent = String(checked);
      if (uploadedBranchesEl) uploadedBranchesEl.textContent = String(uploaded);
      if (progressBar) progressBar.style.width = `${percent}%`;
      if (progressPercentEl) progressPercentEl.textContent = `${percent}%`;
      currentPercent = percent;
    } catch (e) {
      console.warn("branches snapshot handler error", e);
    }
  });

  // ----- Firestore realtime: subscribe to a shared monitoring document to sync across clients -----
  // (keeps localStorage in sync with shared doc)
  try {
    subscribeToDoc("monitoring", "sharedMonitoringData", (snap: any) => {
      try {
        // snap may be DocumentSnapshot or plain object depending on wrapper
        const exists = typeof snap.exists === "function" ? snap.exists() : !!snap.exists;
        if (!exists) return;
        const data = typeof snap.data === "function" ? snap.data() : snap;
        const payload = (data as any)?.data as RowObj[] | undefined;
        if (Array.isArray(payload)) {
          localStorage.setItem("monitoringData", JSON.stringify(payload));
          renderMonitoring();
        }
      } catch (e) {
        console.warn("monitoring doc parse error", e);
      }
    }, (err: any) => {
      console.warn("monitoring snapshot error", err);
    });
  } catch (e) {
    console.warn("subscribeToDoc skipped (Firestore)", e);
  }

  // ----- UI controls for monitoring -----
  if (monitorSearchInput) {
    monitorSearchInput.addEventListener("input", () => {
      renderMonitoring();
    });
  }
  if (monitorFilterUnchecked) {
    monitorFilterUnchecked.addEventListener("change", () => {
      showUnchecked = !!monitorFilterUnchecked.checked;
      renderMonitoring();
    });
  }
  if (addBranchBtn) {
    addBranchBtn.addEventListener("click", () => {
      const name = prompt("Branch name:");
      if (!name || !name.trim()) return;
      const d = getMonitoring();
      d.push({ name: name.trim(), checked: false, uploaded: false, uploadedBy: "", remarks: "" });
      saveMonitoring(d);
      renderMonitoring();
      showBanner("Branch added.");
    });
  }

  if (clearMonitoringBtn) {
    clearMonitoringBtn.addEventListener("click", () => {
      if (!confirm("Clear all monitoring data?")) return;
      saveMonitoring([]);
      renderMonitoring();
      showBanner("All monitoring data cleared.");
    });
  }

  if (exportMonitoringBtn) {
    exportMonitoringBtn.addEventListener("click", () => {
      const data = getMonitoring();
      const month = monthSelect?.value || "";
      const year = yearSelect?.value || "";
      const percent = currentPercent || 0;
      const headerRows = [
        [`Monitoring Progress: ${percent}%`],
        [`Month: ${month} ${year}`],
        ["Branch Name", "Checked", "Uploaded", "Uploaded By", "Remarks"],
      ];
      const rows = data.map((b) => [b.name || "", b.checked ? "Yes" : "No", b.uploaded ? "Yes" : "No", b.uploadedBy || "", b.remarks || ""]);
      const ws = XLSX.utils.aoa_to_sheet([...headerRows, ...rows]);
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "Monitoring");
      XLSX.writeFile(wb, `Monitoring_${month}_${year}.xlsx`);
      showSuccess();
    });
  }

  // Initial render/load from localStorage
  renderMonitoring();
  updateMonitoringStats();

  // Expose for debugging
  (window as any).__scc = { getMonitoring, saveMonitoring, renderMonitoring, updateMonitoringStats };
} // end initApp

// Script is loaded with `defer` in HTML. Run the app now.
initApp();
