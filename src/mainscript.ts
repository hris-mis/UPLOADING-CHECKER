declare const XLSX: any;

type Conflict = { type: string; reason: string };
type RowObj = { name: string; empNo: string; date: string; shift: string; day: string; position: string; conflicts?: Conflict[] };
type Snapshot = { work: RowObj[]; rest: RowObj[] };
type GroupKey = "operation" | "support";
type DataType = "work" | "rest";

const LEADERSHIP_POSITIONS = ["Branch Head", "Site Supervisor", "OIC"];
const clone = <T>(value: T): T => JSON.parse(JSON.stringify(value));
const $ = <T extends Element>(selector: string) => document.querySelector<T>(selector);
const pad2 = (value: string | number) => String(value).padStart(2, "0");

function escapeHtml(value: unknown): string {
  return String(value ?? "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}
function showBanner(message: string): void {
  const banner = $("#warning-banner");
  if (!banner) return;
  banner.textContent = message;
  banner.classList.remove("hidden");
  banner.classList.add("opacity-100");
  setTimeout(() => { banner.classList.remove("opacity-100"); setTimeout(() => banner.classList.add("hidden"), 400); }, 3000);
}
function showSuccess(): void {
  const message = $("#success-message") as HTMLElement | null;
  if (!message) return;
  message.classList.remove("hidden");
  message.style.animation = "fadeInOut 2.5s ease-in-out";
  setTimeout(() => { message.classList.add("hidden"); message.style.animation = ""; }, 2500);
}
function parseDate(value: string): Date | null {
  const match = value.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2}|\d{4})$/);
  if (!match) return null;
  const year = Number(match[3].length === 2 ? `20${match[3]}` : match[3]);
  const date = new Date(year, Number(match[1]) - 1, Number(match[2]));
  return date.getFullYear() === year && date.getMonth() === Number(match[1]) - 1 && date.getDate() === Number(match[2]) ? date : null;
}
function formatDate(date: Date): string { return `${pad2(date.getMonth() + 1)}/${pad2(date.getDate())}/${String(date.getFullYear()).slice(-2)}`; }
function normalizeDate(value: string): string {
  const raw = value.trim();
  if (/^\d{5}(?:\.\d+)?$/.test(raw)) return formatDate(new Date(1899, 11, 30 + Number(raw)));
  const parsed = parseDate(raw);
  if (parsed) return formatDate(parsed);
  const fallback = new Date(raw);
  return Number.isNaN(fallback.getTime()) ? raw : formatDate(fallback);
}
function dayName(value: string): string {
  const date = parseDate(value);
  return date ? date.toLocaleDateString("en-US", { weekday: "long" }) : "";
}
function parseTabular(text: string, html = ""): string[][] {
  if (html && /<table[\s>]/i.test(html)) {
    const documentFragment = new DOMParser().parseFromString(html, "text/html");
    const tableRows = Array.from(documentFragment.querySelectorAll("tr")).map(row =>
      Array.from(row.querySelectorAll("th, td")).map(cell => (cell.textContent || "").replace(/\u00a0/g, " ").trim())
    ).filter(row => row.some(Boolean));
    if (tableRows.length) return tableRows;
  }
  const lines = text.replace(/\r/g, "").replace(/\u00a0/g, " ").split("\n").map(line => line.trim()).filter(line => line && !/^(sheet|page|total|subtotal|summary|prepared|page\s*\d+)/i.test(line));
  const sample = lines.slice(0, 5).join("\n");
  const separator = /\t/.test(sample) ? /\t/ : /,/.test(sample) ? /,/ : /\s{2,}/;
  return lines.map(line => line.split(separator).map(cell => cell.trim()));
}
function mapImportedRows(importedRows: string[][], type: DataType): { rows: RowObj[]; rejected: string[] } {
  let rows = importedRows;
  if (!rows.length) return { rows: [], rejected: [] };
  const normalizedHeaders = rows[0].map(cell => cell.replace(/[\s_\-/.]/g, "").toLowerCase());
  const isHeader = normalizedHeaders.some(cell => /employee|emp|name/.test(cell)) && normalizedHeaders.some(cell => /date|schedule/.test(cell));
  const index = (pattern: RegExp, fallback: number) => { const found = normalizedHeaders.findIndex(cell => pattern.test(cell)); return found < 0 ? fallback : found; };
  let columns = isHeader ? {
    name: index(/name|fullname|employeename/, 0), empNo: index(/emp|employeenumber|idnum|^id$/, 1), date: index(/date|workdate|restdate|sched/, 2),
    shift: index(/shift|time|duty/, 3), day: index(/^day|daytype/, 4), position: index(/position|title|role/, 5)
  } : { name: 0, empNo: 1, date: 2, shift: 3, day: 4, position: 5 };
  if (isHeader) rows = rows.slice(1);
  else {
    const sample = rows.slice(0, 8);
    const score = (predicate: (cell: string) => boolean) => Array.from({ length: Math.max(...sample.map(row => row.length)) }, (_, column) => sample.filter(row => predicate((row[column] || "").trim())).length);
    const best = (scores: number[], fallback: number) => Math.max(...scores) > 0 ? scores.indexOf(Math.max(...scores)) : fallback;
    const employee = best(score(cell => /^\d{2,6}$/.test(cell.replace(/\D/g, ""))), 1);
    const date = best(score(cell => /^\d{5}(?:\.\d+)?$/.test(cell) || parseDate(cell) !== null), 2);
    const shift = best(score(cell => /^(?:[A-Z]{1,5}\d{0,4}|\d{1,2}:\d{2}\s*(?:AM|PM)?(?:\s*[-–]\s*\d{1,2}:\d{2}\s*(?:AM|PM)?)?)$/i.test(cell)), 3);
    const day = best(score(cell => /^(?:mon|tues?|wed(?:nes)?|thu(?:rs)?|fri|sat(?:ur)?|sun)(?:day)?$/i.test(cell)), 4);
    const name = best(score(cell => /^[A-Za-z][A-Za-z ,.'-]+$/.test(cell) && cell.includes(" ") && !/day$/i.test(cell)), 0);
    columns = { name, empNo: employee, date, shift, day, position: 5 };
  }
  const accepted: RowObj[] = [], rejected: string[] = [];
  for (const row of rows) {
    const empNo = (row[columns.empNo] || "").replace(/[^0-9]/g, "");
    const date = normalizeDate(row[columns.date] || "");
    if (empNo.length < 2 || empNo.length > 6) { rejected.push(row.join(" | ")); continue; }
    accepted.push({ name: row[columns.name] || "", empNo, date, shift: type === "work" ? row[columns.shift] || "" : "", day: row[columns.day] || dayName(date), position: row[columns.position] || "" });
  }
  return { rows: accepted, rejected };
}

class ScheduleGroup {
  work: RowObj[] = [];
  rest: RowObj[] = [];
  undo: Snapshot[] = [];
  redo: Snapshot[] = [];
  readonly prefix: string;
  constructor(readonly key: GroupKey, readonly allowWeekendRest: boolean) {
    this.prefix = key === "operation" ? "" : "support";
    const saved = localStorage.getItem(`scheduleChecker.${key}`);
    if (saved) { try { const state = JSON.parse(saved); this.work = state.work || []; this.rest = state.rest || []; } catch { /* ignore invalid older state */ } }
    this.bind(); this.validateAndRender();
  }
  id(name: string): string { return this.prefix ? `${this.prefix}${name[0].toUpperCase()}${name.slice(1)}` : name; }
  el<T extends HTMLElement>(name: string): T | null { return document.getElementById(this.id(name)) as T | null; }
  snapshot(): Snapshot { return clone({ work: this.work, rest: this.rest }); }
  checkpoint(): void { this.undo.push(this.snapshot()); this.redo = []; }
  save(): void { localStorage.setItem(`scheduleChecker.${this.key}`, JSON.stringify(this.snapshot())); }
  bind(): void {
    this.el<HTMLTextAreaElement>("workScheduleInput")?.addEventListener("paste", event => this.paste(event, "work"));
    this.el<HTMLTextAreaElement>("restScheduleInput")?.addEventListener("paste", event => this.paste(event, "rest"));
    this.el("generateWorkFile")?.addEventListener("click", () => this.generate("work"));
    this.el("generateRestFile")?.addEventListener("click", () => this.generate("rest"));
    this.el("clearWorkData")?.addEventListener("click", () => this.clear("work"));
    this.el("clearRestData")?.addEventListener("click", () => this.clear("rest"));
    this.el("workTableBody")?.addEventListener("click", event => this.tableClick(event));
    this.el("restTableBody")?.addEventListener("click", event => this.tableClick(event));
    this.el("workTableBody")?.addEventListener("dblclick", event => this.edit(event));
    this.el("restTableBody")?.addEventListener("dblclick", event => this.edit(event));
  }
  paste(event: ClipboardEvent, type: DataType): void {
    event.preventDefault();
    const text = event.clipboardData?.getData("text/plain") || "";
    const html = event.clipboardData?.getData("text/html") || "";
    const branch = text.match(/branch\s*[:\-]\s*(.+)/i)?.[1]?.trim();
    if (branch) { const input = this.el<HTMLInputElement>(type === "work" ? "workBranchName" : "restBranchName"); if (input && !input.value) input.value = branch; }
    const parsed = mapImportedRows(parseTabular(text, html), type);
    this.checkpoint(); this[type] = parsed.rows;
    const input = this.el<HTMLTextAreaElement>(type === "work" ? "workScheduleInput" : "restScheduleInput"); if (input) input.value = "";
    this.validateAndRender();
    showBanner(`✅ ${parsed.rows.length} ${type === "work" ? "work schedule" : "rest day"} rows pasted.${parsed.rejected.length ? ` ${parsed.rejected.length} rejected.` : ""}`);
    if (parsed.rejected.length) this.showRejected(parsed.rejected);
  }
  tableClick(event: Event): void {
    const button = (event.target as HTMLElement).closest<HTMLButtonElement>(".delete-row-btn");
    if (!button || button.dataset.group !== this.key) return;
    const type = button.dataset.type as DataType; this.checkpoint(); this[type].splice(Number(button.dataset.idx), 1); this.validateAndRender(); showBanner("Row deleted. Press Ctrl+Z to undo.");
  }
  edit(event: Event): void {
    const cell = (event.target as HTMLElement).closest<HTMLTableCellElement>("td[data-field]");
    if (!cell) return;
    const row = cell.closest<HTMLTableRowElement>("tr"); if (!row) return;
    const type = row.dataset.type as DataType, idx = Number(row.dataset.idx), field = cell.dataset.field as keyof RowObj;
    const value = prompt(`Edit ${field}:`, String(this[type][idx][field] || "")); if (value === null) return;
    this.checkpoint(); (this[type][idx][field] as string) = field === "date" ? normalizeDate(value) : value.trim(); this.validateAndRender(); showBanner("Schedule entry updated.");
  }
  clear(type: DataType): void { if (!this[type].length || confirm(`Clear all ${type} data?`)) { this.checkpoint(); this[type] = []; this.validateAndRender(); showBanner(`${type === "work" ? "Work schedule" : "Rest day schedule"} cleared.`); } }
  undoChange(): void { const previous = this.undo.pop(); if (!previous) return showBanner("Nothing to undo."); this.redo.push(this.snapshot()); this.work = previous.work; this.rest = previous.rest; this.validateAndRender(); showBanner("Undo successful."); }
  redoChange(): void { const next = this.redo.pop(); if (!next) return showBanner("Nothing to redo."); this.undo.push(this.snapshot()); this.work = next.work; this.rest = next.rest; this.validateAndRender(); showBanner("Redo successful."); }
  validateAndRender(): void {
    this.rest.forEach(row => row.conflicts = []);
    const workKeys = new Set(this.work.map(row => `${row.empNo}|${row.date}`)), employees = new Set(this.work.map(row => row.empNo));
    const seen = new Set<string>(), byDate = new Map<string, RowObj[]>(), weekendCounts = new Map<string, number>();
    for (const row of this.rest) {
      const key = `${row.empNo}|${row.date}`;
      if (seen.has(key)) row.conflicts!.push({ type: "Duplicate Entry", reason: "Duplicate rest day entry for same employee & date." }); else seen.add(key);
      if (!employees.has(row.empNo)) row.conflicts!.push({ type: "Missing Employee", reason: "Employee not found in Work Schedule data." });
      if (!parseDate(row.date)) row.conflicts!.push({ type: "Invalid Date Format", reason: "Date format unrecognized." });
      if (workKeys.has(key)) row.conflicts!.push({ type: "Work Conflict", reason: "Employee has a work schedule on same date." });
      const dateRows = byDate.get(row.date) || []; dateRows.push(row); byDate.set(row.date, dateRows);
      if (!this.allowWeekendRest && /saturday|sunday/i.test(row.day)) { const date = parseDate(row.date); if (date) { const monthKey = `${row.empNo}|${date.getFullYear()}-${date.getMonth()}`; weekendCounts.set(monthKey, (weekendCounts.get(monthKey) || 0) + 1); } }
    }
    byDate.forEach(rows => { const leaders = rows.filter(row => LEADERSHIP_POSITIONS.includes(row.position)); if (leaders.length > 1) leaders.forEach(row => row.conflicts!.push({ type: "Leadership Conflict", reason: "Multiple leaders have same rest day." })); });
    if (!this.allowWeekendRest) weekendCounts.forEach((count, key) => { if (count > 2) { const empNo = key.split("|")[0]; this.rest.filter(row => row.empNo === empNo && /saturday|sunday/i.test(row.day)).forEach(row => row.conflicts!.push({ type: "Weekend Limit Exceeded", reason: `${count} weekend rest days — maximum 2.` })); } });
    this.render(); this.save();
  }
  missingEntries(): Array<{ empNo: string; name: string; dates: string[] }> {
    const all = [...this.work, ...this.rest].filter(row => parseDate(row.date)); if (!all.length) return [];
    const times = all.map(row => parseDate(row.date)!.getTime()), start = Math.min(...times), end = Math.max(...times);
    const people = new Map<string, { name: string; dates: Set<string> }>();
    all.forEach(row => { const person = people.get(row.empNo) || { name: row.name, dates: new Set<string>() }; if (!person.name) person.name = row.name; person.dates.add(row.date); people.set(row.empNo, person); });
    const result: Array<{ empNo: string; name: string; dates: string[] }> = [];
    people.forEach((person, empNo) => { const dates: string[] = []; for (let time = start; time <= end; time += 86400000) { const date = formatDate(new Date(time)); if (!person.dates.has(date)) dates.push(date); } if (dates.length) result.push({ empNo, name: person.name, dates }); });
    return result;
  }
  render(): void {
    const workBody = this.el<HTMLTableSectionElement>("workTableBody"), restBody = this.el<HTMLTableSectionElement>("restTableBody");
    const cell = (value: string, field: string) => `<td data-field="${field}" title="Double-click to edit">${escapeHtml(value)}</td>`;
    if (workBody) workBody.innerHTML = this.work.map((row, idx) => `<tr data-type="work" data-idx="${idx}">${cell(row.name,"name")}${cell(row.empNo,"empNo")}${cell(row.date,"date")}${cell(row.shift,"shift")}${cell(row.day,"day")}${cell(row.position,"position")}<td><button class="delete-row-btn" data-group="${this.key}" data-type="work" data-idx="${idx}">❌</button></td></tr>`).join("");
    if (restBody) restBody.innerHTML = this.rest.map((row, idx) => { const conflicts = (row.conflicts || []).map(conflict => `<div><strong>${escapeHtml(conflict.type)}:</strong> ${escapeHtml(conflict.reason)}</div>`).join(""); return `<tr class="${conflicts ? "conflict-row" : ""}" data-type="rest" data-idx="${idx}"><td>${conflicts}</td>${cell(row.name,"name")}${cell(row.empNo,"empNo")}${cell(row.date,"date")}${cell(row.day,"day")}${cell(row.position,"position")}<td><button class="delete-row-btn" data-group="${this.key}" data-type="rest" data-idx="${idx}">❌</button></td></tr>`; }).join("");
    const summary = this.el<HTMLElement>("summary"), conflicts = this.rest.filter(row => row.conflicts?.length).length, missing = this.missingEntries();
    if (summary) { if (!this.work.length && !this.rest.length) { summary.classList.add("hidden"); summary.innerHTML = ""; } else { summary.classList.remove("hidden"); summary.innerHTML = `<strong>${conflicts ? `⚠️ ${conflicts} Rest Day entries have conflicts.` : "✅ No Rest Day conflicts detected."}</strong>${missing.length ? `<div class="mt-2 text-left"><strong>Missing Schedule Entry</strong>${missing.map(person => `<div>${escapeHtml(person.empNo)} — ${escapeHtml(person.name)}: ${person.dates.map(escapeHtml).join(", ")}</div>`).join("")}</div>` : `<div class="mt-2">✅ No Missing Schedule Entry.</div>`}`; } }
    const workGenerate = this.el<HTMLButtonElement>("generateWorkFile"), restGenerate = this.el<HTMLButtonElement>("generateRestFile"); if (workGenerate) workGenerate.disabled = !this.work.length; if (restGenerate) restGenerate.disabled = !this.rest.length;
  }
  generate(type: DataType): void {
    const branch = this.el<HTMLInputElement>(type === "work" ? "workBranchName" : "restBranchName")?.value.trim(); if (!branch) return showBanner(`⚠️ Enter ${type === "work" ? "Work" : "Rest"} Branch Name.`);
    if (!this[type].length) return showBanner(`⚠️ No ${type === "work" ? "Work Schedule" : "Rest Day"} data to generate.`);
    if (type === "rest" && this.rest.some(row => row.conflicts?.length)) showBanner("⚠️ Note: There are conflicts, but file generation will proceed.");
    const data = type === "work" ? [["Employee Number", "Work Date", "Shift Code"], ...this.work.map(row => [row.empNo, row.date, row.shift.replace(/\s+/g, "").toUpperCase()])] : [["Employee No", "Rest Day Date"], ...this.rest.map(row => [row.empNo, row.date])];
    const book = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(data), "HRIS Upload"); XLSX.writeFile(book, `${branch}_${type === "work" ? "WORK_SCHEDULE" : "REST_DAY_UPLOAD"}.xlsx`); showSuccess();
  }
  showRejected(rows: string[]): void { const modal = this.el<HTMLElement>("rejectedModal") || $("#rejectedModal") as HTMLElement | null; if (!modal) return; const body = modal.querySelector(".modal-body"); if (body) body.innerHTML = `<p>⚙️ ${rows.length} rows were skipped:</p>${rows.map(row => `<div class="p-2 border-b">${escapeHtml(row)}</div>`).join("")}`; modal.classList.remove("hidden"); modal.style.display = "block"; }
}

const operation = new ScheduleGroup("operation", false);
const support = new ScheduleGroup("support", false);
let active: GroupKey = "operation";
const scheduleContent = $("#tab-schedule-content"), supportContent = $("#tab-monitoring-content"), operationTab = $("#tab-schedule"), supportTab = $("#tab-monitoring");
function setTabState(tab: Element | null, selected: boolean): void {
  tab?.classList.toggle("bg-indigo-600", selected);
  tab?.classList.toggle("text-white", selected);
  tab?.classList.toggle("shadow-md", selected);
  tab?.classList.toggle("bg-gray-200", !selected);
  tab?.classList.toggle("text-gray-700", !selected);
}
function activate(group: GroupKey): void {
  active = group;
  scheduleContent?.classList.toggle("hidden", group !== "operation");
  supportContent?.classList.toggle("hidden", group !== "support");
  setTabState(operationTab, group === "operation");
  setTabState(supportTab, group === "support");
}
operationTab?.addEventListener("click", () => activate("operation")); supportTab?.addEventListener("click", () => activate("support")); activate("operation");
document.addEventListener("keydown", event => { if (!(event.ctrlKey || event.metaKey)) return; const group = active === "operation" ? operation : support; if (event.key.toLowerCase() === "z" && !event.shiftKey) { event.preventDefault(); group.undoChange(); } else if (event.key.toLowerCase() === "y" || (event.key.toLowerCase() === "z" && event.shiftKey)) { event.preventDefault(); group.redoChange(); } });
document.addEventListener("click", event => { const target = event.target as HTMLElement; if (target.matches(".modal-close, #rejectedModal .modal-overlay")) { const modal = $("#rejectedModal") as HTMLElement | null; if (modal) { modal.classList.add("hidden"); modal.style.display = "none"; } } });
const backToTop = $("#backToTopBtn") as HTMLElement | null; window.addEventListener("scroll", () => { if (backToTop) backToTop.style.display = window.scrollY > 300 ? "block" : "none"; }); backToTop?.addEventListener("click", () => window.scrollTo({ top: 0, behavior: "smooth" }));
