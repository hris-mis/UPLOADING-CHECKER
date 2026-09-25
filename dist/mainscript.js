"use strict";
const LEADERSHIP_POSITIONS = ["Branch Head", "Site Supervisor", "OIC"];
const clone = (value) => JSON.parse(JSON.stringify(value));
const $ = (selector) => document.querySelector(selector);
const pad2 = (value) => String(value).padStart(2, "0");
function escapeHtml(value) {
    return String(value !== null && value !== void 0 ? value : "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}
function showBanner(message) {
    const banner = $("#warning-banner");
    if (!banner)
        return;
    banner.textContent = message;
    banner.classList.remove("hidden");
    banner.classList.add("opacity-100");
    setTimeout(() => { banner.classList.remove("opacity-100"); setTimeout(() => banner.classList.add("hidden"), 400); }, 3000);
}
function showSuccess() {
    const message = $("#success-message");
    if (!message)
        return;
    message.classList.remove("hidden");
    message.style.animation = "fadeInOut 2.5s ease-in-out";
    setTimeout(() => { message.classList.add("hidden"); message.style.animation = ""; }, 2500);
}
function parseDate(value) {
    const match = value.match(/^(\d{1,2})[/.\-](\d{1,2})[/.\-](\d{2}|\d{4})$/);
    if (!match)
        return null;
    const year = Number(match[3].length === 2 ? `20${match[3]}` : match[3]);
    const date = new Date(year, Number(match[1]) - 1, Number(match[2]));
    return date.getFullYear() === year && date.getMonth() === Number(match[1]) - 1 && date.getDate() === Number(match[2]) ? date : null;
}
function formatDate(date) { return `${pad2(date.getMonth() + 1)}/${pad2(date.getDate())}/${String(date.getFullYear()).slice(-2)}`; }
function normalizeDate(value) {
    const raw = value.trim();
    if (/^\d{5}(?:\.\d+)?$/.test(raw))
        return formatDate(new Date(1899, 11, 30 + Number(raw)));
    const parsed = parseDate(raw);
    if (parsed)
        return formatDate(parsed);
    const fallback = new Date(raw);
    return Number.isNaN(fallback.getTime()) ? raw : formatDate(fallback);
}
function dayName(value) {
    const date = parseDate(value);
    return date ? date.toLocaleDateString("en-US", { weekday: "long" }) : "";
}
function parseTabular(text) {
    const lines = text.replace(/\r/g, "").split("\n").map(line => line.trim()).filter(line => line && !/^(sheet|page|total|subtotal|page\s*\d+)/i.test(line));
    const sample = lines.slice(0, 5).join("\n");
    const separator = /\t/.test(sample) ? /\t/ : /,/.test(sample) ? /,/ : /\s{2,}/;
    return lines.map(line => line.split(separator).map(cell => cell.trim()));
}
function mapRows(text, type) {
    let rows = parseTabular(text);
    if (!rows.length)
        return { rows: [], rejected: [] };
    const normalizedHeaders = rows[0].map(cell => cell.replace(/[\s_\-/.]/g, "").toLowerCase());
    const isHeader = normalizedHeaders.some(cell => /employee|emp|name/.test(cell)) && normalizedHeaders.some(cell => /date|schedule/.test(cell));
    const index = (pattern, fallback) => { const found = normalizedHeaders.findIndex(cell => pattern.test(cell)); return found < 0 ? fallback : found; };
    const columns = isHeader ? {
        name: index(/name|fullname|employeename/, 0), empNo: index(/emp|employeenumber|idnum|^id$/, 1), date: index(/date|workdate|restdate|sched/, 2),
        shift: index(/shift|time|duty/, 3), day: index(/^day|daytype/, 4), position: index(/position|title|role/, 5)
    } : { name: 0, empNo: 1, date: 2, shift: 3, day: 4, position: 5 };
    if (isHeader)
        rows = rows.slice(1);
    const accepted = [], rejected = [];
    for (const row of rows) {
        const empNo = (row[columns.empNo] || "").replace(/[^0-9]/g, "");
        const date = normalizeDate(row[columns.date] || "");
        if (empNo.length < 2 || empNo.length > 6 || !date) {
            rejected.push(row.join(" | "));
            continue;
        }
        accepted.push({ name: row[columns.name] || "", empNo, date, shift: type === "work" ? row[columns.shift] || "" : "", day: row[columns.day] || dayName(date), position: row[columns.position] || "" });
    }
    return { rows: accepted, rejected };
}
class ScheduleGroup {
    constructor(key, allowWeekendRest) {
        this.key = key;
        this.allowWeekendRest = allowWeekendRest;
        this.work = [];
        this.rest = [];
        this.undo = [];
        this.redo = [];
        this.prefix = key === "operation" ? "" : "support";
        const saved = localStorage.getItem(`scheduleChecker.${key}`);
        if (saved) {
            try {
                const state = JSON.parse(saved);
                this.work = state.work || [];
                this.rest = state.rest || [];
            }
            catch ( /* ignore invalid older state */_a) { /* ignore invalid older state */ }
        }
        this.bind();
        this.validateAndRender();
    }
    id(name) { return this.prefix ? `${this.prefix}${name[0].toUpperCase()}${name.slice(1)}` : name; }
    el(name) { return document.getElementById(this.id(name)); }
    snapshot() { return clone({ work: this.work, rest: this.rest }); }
    checkpoint() { this.undo.push(this.snapshot()); this.redo = []; }
    save() { localStorage.setItem(`scheduleChecker.${this.key}`, JSON.stringify(this.snapshot())); }
    bind() {
        var _a, _b, _c, _d, _e, _f, _g, _h, _j, _k;
        (_a = this.el("workScheduleInput")) === null || _a === void 0 ? void 0 : _a.addEventListener("paste", event => this.paste(event, "work"));
        (_b = this.el("restScheduleInput")) === null || _b === void 0 ? void 0 : _b.addEventListener("paste", event => this.paste(event, "rest"));
        (_c = this.el("generateWorkFile")) === null || _c === void 0 ? void 0 : _c.addEventListener("click", () => this.generate("work"));
        (_d = this.el("generateRestFile")) === null || _d === void 0 ? void 0 : _d.addEventListener("click", () => this.generate("rest"));
        (_e = this.el("clearWorkData")) === null || _e === void 0 ? void 0 : _e.addEventListener("click", () => this.clear("work"));
        (_f = this.el("clearRestData")) === null || _f === void 0 ? void 0 : _f.addEventListener("click", () => this.clear("rest"));
        (_g = this.el("workTableBody")) === null || _g === void 0 ? void 0 : _g.addEventListener("click", event => this.tableClick(event));
        (_h = this.el("restTableBody")) === null || _h === void 0 ? void 0 : _h.addEventListener("click", event => this.tableClick(event));
        (_j = this.el("workTableBody")) === null || _j === void 0 ? void 0 : _j.addEventListener("dblclick", event => this.edit(event));
        (_k = this.el("restTableBody")) === null || _k === void 0 ? void 0 : _k.addEventListener("dblclick", event => this.edit(event));
    }
    paste(event, type) {
        var _a, _b, _c;
        event.preventDefault();
        const text = ((_a = event.clipboardData) === null || _a === void 0 ? void 0 : _a.getData("text/plain")) || "";
        const branch = (_c = (_b = text.match(/branch\s*[:\-]\s*(.+)/i)) === null || _b === void 0 ? void 0 : _b[1]) === null || _c === void 0 ? void 0 : _c.trim();
        if (branch) {
            const input = this.el(type === "work" ? "workBranchName" : "restBranchName");
            if (input && !input.value)
                input.value = branch;
        }
        const parsed = mapRows(text, type);
        this.checkpoint();
        this[type] = parsed.rows;
        const input = this.el(type === "work" ? "workScheduleInput" : "restScheduleInput");
        if (input)
            input.value = "";
        this.validateAndRender();
        showBanner(`✅ ${parsed.rows.length} ${type === "work" ? "work schedule" : "rest day"} rows pasted.${parsed.rejected.length ? ` ${parsed.rejected.length} rejected.` : ""}`);
        if (parsed.rejected.length)
            this.showRejected(parsed.rejected);
    }
    tableClick(event) {
        const button = event.target.closest(".delete-row-btn");
        if (!button || button.dataset.group !== this.key)
            return;
        const type = button.dataset.type;
        this.checkpoint();
        this[type].splice(Number(button.dataset.idx), 1);
        this.validateAndRender();
        showBanner("Row deleted. Press Ctrl+Z to undo.");
    }
    edit(event) {
        const cell = event.target.closest("td[data-field]");
        if (!cell)
            return;
        const row = cell.closest("tr");
        if (!row)
            return;
        const type = row.dataset.type, idx = Number(row.dataset.idx), field = cell.dataset.field;
        const value = prompt(`Edit ${field}:`, String(this[type][idx][field] || ""));
        if (value === null)
            return;
        this.checkpoint();
        this[type][idx][field] = field === "date" ? normalizeDate(value) : value.trim();
        this.validateAndRender();
        showBanner("Schedule entry updated.");
    }
    clear(type) { if (!this[type].length || confirm(`Clear all ${type} data?`)) {
        this.checkpoint();
        this[type] = [];
        this.validateAndRender();
        showBanner(`${type === "work" ? "Work schedule" : "Rest day schedule"} cleared.`);
    } }
    undoChange() { const previous = this.undo.pop(); if (!previous)
        return showBanner("Nothing to undo."); this.redo.push(this.snapshot()); this.work = previous.work; this.rest = previous.rest; this.validateAndRender(); showBanner("Undo successful."); }
    redoChange() { const next = this.redo.pop(); if (!next)
        return showBanner("Nothing to redo."); this.undo.push(this.snapshot()); this.work = next.work; this.rest = next.rest; this.validateAndRender(); showBanner("Redo successful."); }
    validateAndRender() {
        this.rest.forEach(row => row.conflicts = []);
        const workKeys = new Set(this.work.map(row => `${row.empNo}|${row.date}`)), employees = new Set(this.work.map(row => row.empNo));
        const seen = new Set(), byDate = new Map(), weekendCounts = new Map();
        for (const row of this.rest) {
            const key = `${row.empNo}|${row.date}`;
            if (seen.has(key))
                row.conflicts.push({ type: "Duplicate Entry", reason: "Duplicate rest day entry for same employee & date." });
            else
                seen.add(key);
            if (!employees.has(row.empNo))
                row.conflicts.push({ type: "Missing Employee", reason: "Employee not found in Work Schedule data." });
            if (!parseDate(row.date))
                row.conflicts.push({ type: "Invalid Date Format", reason: "Date format unrecognized." });
            if (workKeys.has(key))
                row.conflicts.push({ type: "Work Conflict", reason: "Employee has a work schedule on same date." });
            const dateRows = byDate.get(row.date) || [];
            dateRows.push(row);
            byDate.set(row.date, dateRows);
            if (!this.allowWeekendRest && /saturday|sunday/i.test(row.day)) {
                const date = parseDate(row.date);
                if (date) {
                    const monthKey = `${row.empNo}|${date.getFullYear()}-${date.getMonth()}`;
                    weekendCounts.set(monthKey, (weekendCounts.get(monthKey) || 0) + 1);
                }
            }
        }
        byDate.forEach(rows => { const leaders = rows.filter(row => LEADERSHIP_POSITIONS.includes(row.position)); if (leaders.length > 1)
            leaders.forEach(row => row.conflicts.push({ type: "Leadership Conflict", reason: "Multiple leaders have same rest day." })); });
        if (!this.allowWeekendRest)
            weekendCounts.forEach((count, key) => { if (count > 2) {
                const empNo = key.split("|")[0];
                this.rest.filter(row => row.empNo === empNo && /saturday|sunday/i.test(row.day)).forEach(row => row.conflicts.push({ type: "Weekend Limit Exceeded", reason: `${count} weekend rest days — maximum 2.` }));
            } });
        this.render();
        this.save();
    }
    missingEntries() {
        const all = [...this.work, ...this.rest].filter(row => parseDate(row.date));
        if (!all.length)
            return [];
        const times = all.map(row => parseDate(row.date).getTime()), start = Math.min(...times), end = Math.max(...times);
        const people = new Map();
        all.forEach(row => { const person = people.get(row.empNo) || { name: row.name, dates: new Set() }; if (!person.name)
            person.name = row.name; person.dates.add(row.date); people.set(row.empNo, person); });
        const result = [];
        people.forEach((person, empNo) => { const dates = []; for (let time = start; time <= end; time += 86400000) {
            const date = formatDate(new Date(time));
            if (!person.dates.has(date))
                dates.push(date);
        } if (dates.length)
            result.push({ empNo, name: person.name, dates }); });
        return result;
    }
    render() {
        const workBody = this.el("workTableBody"), restBody = this.el("restTableBody");
        const cell = (value, field) => `<td data-field="${field}" title="Double-click to edit">${escapeHtml(value)}</td>`;
        if (workBody)
            workBody.innerHTML = this.work.map((row, idx) => `<tr data-type="work" data-idx="${idx}">${cell(row.name, "name")}${cell(row.empNo, "empNo")}${cell(row.date, "date")}${cell(row.shift, "shift")}${cell(row.day, "day")}${cell(row.position, "position")}<td><button class="delete-row-btn" data-group="${this.key}" data-type="work" data-idx="${idx}">❌</button></td></tr>`).join("");
        if (restBody)
            restBody.innerHTML = this.rest.map((row, idx) => { const conflicts = (row.conflicts || []).map(conflict => `<div><strong>${escapeHtml(conflict.type)}:</strong> ${escapeHtml(conflict.reason)}</div>`).join(""); return `<tr class="${conflicts ? "conflict-row" : ""}" data-type="rest" data-idx="${idx}"><td>${conflicts}</td>${cell(row.name, "name")}${cell(row.empNo, "empNo")}${cell(row.date, "date")}${cell(row.day, "day")}${cell(row.position, "position")}<td><button class="delete-row-btn" data-group="${this.key}" data-type="rest" data-idx="${idx}">❌</button></td></tr>`; }).join("");
        const summary = this.el("summary"), conflicts = this.rest.filter(row => { var _a; return (_a = row.conflicts) === null || _a === void 0 ? void 0 : _a.length; }).length, missing = this.missingEntries();
        if (summary) {
            if (!this.work.length && !this.rest.length) {
                summary.classList.add("hidden");
                summary.innerHTML = "";
            }
            else {
                summary.classList.remove("hidden");
                summary.innerHTML = `<strong>${conflicts ? `⚠️ ${conflicts} Rest Day entries have conflicts.` : "✅ No Rest Day conflicts detected."}</strong>${missing.length ? `<div class="mt-2 text-left"><strong>Missing Schedule Entry</strong>${missing.map(person => `<div>${escapeHtml(person.empNo)} — ${escapeHtml(person.name)}: ${person.dates.map(escapeHtml).join(", ")}</div>`).join("")}</div>` : `<div class="mt-2">✅ No Missing Schedule Entry.</div>`}`;
            }
        }
        const workGenerate = this.el("generateWorkFile"), restGenerate = this.el("generateRestFile");
        if (workGenerate)
            workGenerate.disabled = !this.work.length;
        if (restGenerate)
            restGenerate.disabled = !this.rest.length;
    }
    generate(type) {
        var _a;
        const branch = (_a = this.el(type === "work" ? "workBranchName" : "restBranchName")) === null || _a === void 0 ? void 0 : _a.value.trim();
        if (!branch)
            return showBanner(`⚠️ Enter ${type === "work" ? "Work" : "Rest"} Branch Name.`);
        if (!this[type].length)
            return showBanner(`⚠️ No ${type === "work" ? "Work Schedule" : "Rest Day"} data to generate.`);
        const data = type === "work" ? [["Employee Number", "Work Date", "Shift Code"], ...this.work.map(row => [row.empNo, row.date, row.shift.replace(/\s+/g, "").toUpperCase()])] : [["Employee No", "Rest Day Date"], ...this.rest.map(row => [row.empNo, row.date])];
        const book = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(book, XLSX.utils.aoa_to_sheet(data), "HRIS Upload");
        XLSX.writeFile(book, `${branch}_${type === "work" ? "WORK_SCHEDULE" : "REST_DAY_UPLOAD"}.xlsx`);
        showSuccess();
    }
    showRejected(rows) { const modal = this.el("rejectedModal") || $("#rejectedModal"); if (!modal)
        return; const body = modal.querySelector(".modal-body"); if (body)
        body.innerHTML = `<p>⚙️ ${rows.length} rows were skipped:</p>${rows.map(row => `<div class="p-2 border-b">${escapeHtml(row)}</div>`).join("")}`; modal.classList.remove("hidden"); modal.style.display = "block"; }
}
const operation = new ScheduleGroup("operation", false);
const support = new ScheduleGroup("support", true);
let active = "operation";
const scheduleContent = $("#tab-schedule-content"), supportContent = $("#tab-monitoring-content"), operationTab = $("#tab-schedule"), supportTab = $("#tab-monitoring");
function activate(group) { active = group; scheduleContent === null || scheduleContent === void 0 ? void 0 : scheduleContent.classList.toggle("hidden", group !== "operation"); supportContent === null || supportContent === void 0 ? void 0 : supportContent.classList.toggle("hidden", group !== "support"); operationTab === null || operationTab === void 0 ? void 0 : operationTab.classList.toggle("bg-indigo-600", group === "operation"); operationTab === null || operationTab === void 0 ? void 0 : operationTab.classList.toggle("text-white", group === "operation"); supportTab === null || supportTab === void 0 ? void 0 : supportTab.classList.toggle("bg-indigo-600", group === "support"); supportTab === null || supportTab === void 0 ? void 0 : supportTab.classList.toggle("text-white", group === "support"); }
operationTab === null || operationTab === void 0 ? void 0 : operationTab.addEventListener("click", () => activate("operation"));
supportTab === null || supportTab === void 0 ? void 0 : supportTab.addEventListener("click", () => activate("support"));
activate("operation");
document.addEventListener("keydown", event => { if (!(event.ctrlKey || event.metaKey))
    return; const group = active === "operation" ? operation : support; if (event.key.toLowerCase() === "z" && !event.shiftKey) {
    event.preventDefault();
    group.undoChange();
}
else if (event.key.toLowerCase() === "y" || (event.key.toLowerCase() === "z" && event.shiftKey)) {
    event.preventDefault();
    group.redoChange();
} });
document.addEventListener("click", event => { const target = event.target; if (target.matches(".modal-close, #rejectedModal .modal-overlay")) {
    const modal = $("#rejectedModal");
    if (modal) {
        modal.classList.add("hidden");
        modal.style.display = "none";
    }
} });
const backToTop = $("#backToTopBtn");
window.addEventListener("scroll", () => { if (backToTop)
    backToTop.style.display = window.scrollY > 300 ? "block" : "none"; });
backToTop === null || backToTop === void 0 ? void 0 : backToTop.addEventListener("click", () => window.scrollTo({ top: 0, behavior: "smooth" }));
