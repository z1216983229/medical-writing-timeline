import { BUILT_IN_HOLIDAYS, CSR_TEMPLATE } from "./holiday-data.js";
import {
  buildCalendar,
  buildExcelHtml,
  formatDisplayDate,
  normalizeDate,
  recalculateSteps,
} from "./timeline-core.js";

const STORAGE_KEY = "mwTimelineTool.v1";

const els = {
  newProjectBtn: document.querySelector("#newProjectBtn"),
  projectSearch: document.querySelector("#projectSearch"),
  projectList: document.querySelector("#projectList"),
  exportLibraryBtn: document.querySelector("#exportLibraryBtn"),
  importLibraryInput: document.querySelector("#importLibraryInput"),
  projectName: document.querySelector("#projectName"),
  templateSelect: document.querySelector("#templateSelect"),
  saveProjectBtn: document.querySelector("#saveProjectBtn"),
  exportExcelBtn: document.querySelector("#exportExcelBtn"),
  updateHolidaysBtn: document.querySelector("#updateHolidaysBtn"),
  deleteProjectBtn: document.querySelector("#deleteProjectBtn"),
  statusLine: document.querySelector("#statusLine"),
  timelineBody: document.querySelector("#timelineBody"),
  rowMenu: document.querySelector("#rowMenu"),
};

let state = loadState();

if (new URLSearchParams(location.search).get("reset") === "1") {
  localStorage.removeItem(STORAGE_KEY);
  history.replaceState(null, "", location.pathname);
  state = defaultState();
}

bindEvents();
render();

function defaultState() {
  const year = 2026;
  const project = createProject({
    name: "新项目 Timeline",
    templateName: CSR_TEMPLATE.name,
    startDate: `${year}-01-05`,
    holidayYear: year,
    steps: cloneSteps(CSR_TEMPLATE.steps),
  });
  return {
    activeProjectId: project.id,
    projects: [project],
    templates: [CSR_TEMPLATE],
    holidays: structuredClone(BUILT_IN_HOLIDAYS),
  };
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return defaultState();
    const parsed = JSON.parse(raw);
    if (!parsed.projects?.length) return defaultState();
    return {
      ...defaultState(),
      ...parsed,
      holidays: { ...structuredClone(BUILT_IN_HOLIDAYS), ...(parsed.holidays || {}) },
    };
  } catch {
    return defaultState();
  }
}

function saveState() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function bindEvents() {
  els.newProjectBtn.addEventListener("click", () => {
    const project = createProject({
      name: `项目 ${state.projects.length + 1}`,
      templateName: CSR_TEMPLATE.name,
      startDate: todayIso(),
      holidayYear: new Date().getFullYear(),
      steps: cloneSteps(CSR_TEMPLATE.steps),
    });
    state.projects.unshift(project);
    state.activeProjectId = project.id;
    saveAndRender("已新建项目");
  });

  els.projectSearch.addEventListener("input", renderProjectList);
  els.projectName.addEventListener("input", () => {
    const project = activeProject();
    if (!project) return;
    project.name = els.projectName.value.trim() || "未命名项目";
    project.updatedAt = new Date().toISOString();
    saveState();
    renderProjectList();
  });
  els.templateSelect.addEventListener("change", () => {
    const project = activeProject();
    const template = selectedTemplate();
    if (!project || !template) return;
    if (project.steps.length && !confirm("切换模板会替换当前 timeline。继续？")) {
      renderTemplates();
      return;
    }
    project.templateName = template.name;
    project.steps = cloneSteps(template.steps);
    project.startDate = project.startDate || todayIso();
    project.holidayYear = Number(project.startDate.slice(0, 4)) || new Date().getFullYear();
    saveAndRender("已套用模板");
  });

  els.saveProjectBtn.addEventListener("click", () => {
    const project = activeProject();
    if (!project) return;
    project.updatedAt = new Date().toISOString();
    saveState();
    renderProjectList();
    setStatus("已保存到左侧项目列表");
  });

  els.exportExcelBtn.addEventListener("click", () => {
    const project = activeProject();
    if (!project) return;
    const html = buildExcelHtml(project, calculatedSteps(project), holidayData(project.holidayYear).meta);
    download(`${safeFileName(project.name)}-timeline.xls`, html, "application/vnd.ms-excel;charset=utf-8");
    setStatus("已导出 Excel");
  });

  els.exportLibraryBtn.addEventListener("click", () => {
    download("medical-writing-timeline-library.json", JSON.stringify(state, null, 2), "application/json;charset=utf-8");
    setStatus("已导出项目库");
  });

  els.importLibraryInput.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;
    const text = await file.text();
    const imported = JSON.parse(text);
    if (!Array.isArray(imported.projects) || !Array.isArray(imported.templates)) {
      alert("项目库文件格式不正确");
      return;
    }
    state = {
      ...defaultState(),
      ...imported,
      holidays: { ...structuredClone(BUILT_IN_HOLIDAYS), ...(imported.holidays || {}) },
    };
    state.activeProjectId = state.activeProjectId || state.projects[0]?.id;
    els.importLibraryInput.value = "";
    saveAndRender("已导入项目库");
  });

  els.updateHolidaysBtn.addEventListener("click", updateHolidays);
  els.deleteProjectBtn.addEventListener("click", () => {
    const project = activeProject();
    if (!project) return;
    if (!confirm(`删除项目「${project.name}」？`)) return;
    state.projects = state.projects.filter((item) => item.id !== project.id);
    if (!state.projects.length) state.projects.push(createProject({
      name: "新项目 Timeline",
      templateName: CSR_TEMPLATE.name,
      startDate: todayIso(),
      holidayYear: new Date().getFullYear(),
      steps: cloneSteps(CSR_TEMPLATE.steps),
    }));
    state.activeProjectId = state.projects[0].id;
    saveAndRender("已删除项目");
  });
  els.rowMenu.addEventListener("click", handleMenuAction);
  document.addEventListener("click", hideRowMenu);
  document.addEventListener("keydown", (event) => {
    if (event.key === "Escape") hideRowMenu();
  });
  bindColumnResizers();
}

function render() {
  renderProjectList();
  renderProjectForm();
  renderTemplates();
  renderTimeline();
}

function renderProjectList() {
  const query = els.projectSearch.value.trim().toLowerCase();
  els.projectList.innerHTML = "";
  state.projects
    .filter((project) => project.name.toLowerCase().includes(query))
    .forEach((project) => {
      const button = document.createElement("button");
      button.className = `project-item${project.id === state.activeProjectId ? " active" : ""}`;
      button.innerHTML = `<strong>${escapeHtml(project.name)}</strong><span>${escapeHtml(project.templateName || "空白")} · ${formatDisplayDate(project.startDate)}</span>`;
      button.addEventListener("click", () => {
        state.activeProjectId = project.id;
        saveAndRender();
      });
      els.projectList.append(button);
    });
}

function renderProjectForm() {
  const project = activeProject();
  if (!project) return;
  els.projectName.value = project.name;
}

function renderTemplates() {
  const project = activeProject();
  els.templateSelect.innerHTML = state.templates
    .map((template) => `<option value="${template.id}">${escapeHtml(template.name)}</option>`)
    .join("");
  const matching = state.templates.find((template) => template.name === project?.templateName);
  els.templateSelect.value = matching?.id || state.templates[0]?.id || "";
}

function renderTimeline() {
  const project = activeProject();
  if (!project) return;
  const rows = calculatedSteps(project);
  els.timelineBody.innerHTML = "";
  rows.forEach((step, index) => {
    const tr = document.createElement("tr");
    tr.dataset.index = String(index);
    if (step.rowColor) tr.classList.add("row-highlight");
    tr.innerHTML = `
      <td data-col="order">${step.order}</td>
      <td data-col="scope" class="${cellClass(step, "scope")}">${textArea(index, "scope", step.scope)}</td>
      <td data-col="task" class="${cellClass(step, "task")}">${textArea(index, "task", step.task)}</td>
      <td data-col="owner" class="${cellClass(step, "owner")}">${textArea(index, "owner", step.owner)}</td>
      <td data-col="duration" class="${cellClass(step, "duration")}">${numberInput(index, "duration", step.duration)}</td>
      <td data-col="start" class="${cellClass(step, "start")}">${index === 0 ? dateInput(index, "projectStart", step.startDate) : calculatedDateCell(step.startDate)}</td>
      <td data-col="end" class="${cellClass(step, "end")}">${calculatedDateCell(step.endDate)}</td>
    `;
    els.timelineBody.append(tr);
  });

  els.timelineBody.querySelectorAll("input[data-field], textarea[data-field]").forEach((input) => {
    input.addEventListener("input", (event) => {
      const project = activeProject();
      const index = Number(event.target.dataset.index);
      const field = event.target.dataset.field;
      if (field === "projectStart") return;
      project.steps[index][field] = event.target.value;
      project.updatedAt = new Date().toISOString();
      saveState();
      if (field === "duration") renderTimeline();
    });
    input.addEventListener("change", () => {
      saveAndRender();
    });
  });

  els.timelineBody.querySelectorAll('input[data-date-field="projectStart"]').forEach((input) => {
    input.addEventListener("change", (event) => {
      const project = activeProject();
      if (!project) return;
      const index = Number(event.target.dataset.index);
      const field = event.target.dataset.dateField;
      const value = normalizeDate(event.target.value);
      project.startDate = value;
      project.holidayYear = Number(project.startDate.slice(0, 4)) || project.holidayYear;
      saveAndRender();
    });
  });

  els.timelineBody.querySelectorAll("td").forEach((cell) => {
    cell.addEventListener("contextmenu", showRowMenu);
  });
}

function handleRowAction(action, index, column) {
  const project = activeProject();
  if (!project) return;
  const step = project.steps[index];
  if (action === "up" && index > 0) swap(project.steps, index, index - 1);
  if (action === "down" && index < project.steps.length - 1) swap(project.steps, index, index + 1);
  if (action === "copy") project.steps.splice(index + 1, 0, { ...structuredClone(project.steps[index]), id: id() });
  if (action === "delete") project.steps.splice(index, 1);
  if (action === "insert-above") project.steps.splice(index, 0, newStep());
  if (action === "insert-below") project.steps.splice(index + 1, 0, newStep());
  if (action === "highlight-row" && step) step.rowColor = "highlight";
  if (action === "highlight-cell" && step && column && column !== "order") {
    step.cellColors = { ...(step.cellColors || {}), [column]: "highlight" };
  }
  if (action === "clear-highlight" && step) {
    delete step.rowColor;
    if (column && step.cellColors) delete step.cellColors[column];
  }
  if (action === "clear-cell" && step && isEditableColumn(column, index)) {
    clearEditableCell(project, step, index, column);
  }
  if (action === "cell-clear" && step?.cellColors && column) {
    delete step.cellColors[column];
  }
  saveAndRender();
}

function showRowMenu(event) {
  event.preventDefault();
  const row = event.currentTarget.closest("tr[data-index]");
  if (!row) return;
  const index = Number(row.dataset.index);
  const column = event.currentTarget.dataset.col || "";
  els.rowMenu.dataset.index = String(index);
  els.rowMenu.dataset.column = column;
  els.rowMenu.querySelector('[data-menu-action="clear-cell"]').disabled = !isEditableColumn(column, index);
  els.rowMenu.hidden = false;
  const left = Math.min(event.clientX, window.innerWidth - els.rowMenu.offsetWidth - 8);
  const top = Math.min(event.clientY, window.innerHeight - els.rowMenu.offsetHeight - 8);
  els.rowMenu.style.left = `${Math.max(8, left)}px`;
  els.rowMenu.style.top = `${Math.max(8, top)}px`;
}

function hideRowMenu() {
  els.rowMenu.hidden = true;
}

function handleMenuAction(event) {
  const button = event.target.closest("button[data-menu-action]");
  if (!button) return;
  event.stopPropagation();
  const index = Number(els.rowMenu.dataset.index);
  const column = els.rowMenu.dataset.column;
  const action = button.dataset.menuAction;
  hideRowMenu();
  handleRowAction(action, index, column);
}

async function updateHolidays() {
  const project = activeProject();
  if (!project) return;
  const year = Number(project.holidayYear);
  setStatus("正在更新节假日...");
  try {
    const response = await fetch(`https://api.jiejiariapi.com/v1/holidays/${year}`, { cache: "no-store" });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const data = await response.json();
    const offDays = [];
    const workDays = [];
    Object.values(data).forEach((item) => {
      if (!item?.date) return;
      const entry = { date: normalizeDate(item.date), name: item.name || "" };
      if (item.isOffDay) offDays.push(entry);
      if (!item.isOffDay && isWeekendDate(entry.date)) workDays.push(entry);
    });
    state.holidays[year] = {
      source: "api.jiejiariapi.com",
      updatedAt: new Date().toISOString(),
      offDays,
      workDays,
    };
    saveAndRender(`已更新 ${year} 节假日`);
  } catch (error) {
    setStatus(`更新失败，继续使用现有数据：${error.message}`);
  }
}

function createProject({ name, templateName, startDate, holidayYear, steps }) {
  return {
    id: id(),
    name,
    templateName,
    startDate,
    holidayYear,
    steps: cloneSteps(steps),
    notionPageId: "",
    lastSyncedAt: "",
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
}

function newStep() {
  return { id: id(), scope: "", task: "新步骤", owner: "", duration: 1, notes: "", manualStartDate: "", manualEndDate: "", rowColor: "", cellColors: {} };
}

function cloneSteps(steps) {
  return steps.map((step) => ({
    id: id(),
    scope: step.scope || "",
    task: step.task || "",
    owner: step.owner || "",
    duration: step.duration ?? "",
    notes: step.notes || "",
    manualStartDate: step.manualStartDate || "",
    manualEndDate: step.manualEndDate || "",
    rowColor: step.rowColor || "",
    cellColors: { ...(step.cellColors || {}) },
  }));
}

function activeProject() {
  return state.projects.find((project) => project.id === state.activeProjectId) || state.projects[0];
}

function selectedTemplate() {
  return state.templates.find((template) => template.id === els.templateSelect.value);
}

function calculatedSteps(project) {
  return recalculateSteps(project.steps, project.startDate, buildCalendar(holidayData(project.holidayYear)));
}

function holidayData(year) {
  return normalizeHolidayData(state.holidays[year] || BUILT_IN_HOLIDAYS[year] || { source: "空白数据", offDays: [], workDays: [] });
}

function ensureHolidayData(year) {
  const normalized = holidayData(year);
  state.holidays[year] = normalized;
  return state.holidays[year];
}

function normalizeHolidayData(data) {
  return {
    meta: { source: data.source || data.meta?.source || "内置数据", updatedAt: data.updatedAt || data.meta?.updatedAt || "" },
    source: data.source || data.meta?.source || "内置数据",
    updatedAt: data.updatedAt || data.meta?.updatedAt || "",
    offDays: (data.offDays || []).map((item) => ({ date: normalizeDate(item.date || item), name: item.name || "" })).filter((item) => item.date),
    workDays: (data.workDays || []).map((item) => ({ date: normalizeDate(item.date || item), name: item.name || "" })).filter((item) => item.date),
  };
}

function cellClass(step, field) {
  return step.cellColors?.[field] ? "cell-highlight" : "";
}

function textCell(value) {
  return `<div class="display-cell">${escapeHtml(value)}</div>`;
}

function textArea(index, field, value) {
  return `<textarea data-index="${index}" data-field="${field}" rows="2">${escapeHtml(value)}</textarea>`;
}

function numberInput(index, field, value) {
  return `<input data-index="${index}" data-field="${field}" type="number" min="1" step="1" value="${escapeAttr(value)}">`;
}

function dateInput(index, field, value) {
  return `<div class="date-cell"><input data-index="${index}" data-date-field="${field}" type="date" value="${escapeAttr(value)}"><span>${formatDisplayDate(value) || "年 / 月 / 日"}</span></div>`;
}

function calculatedDateCell(value) {
  return `<div class="display-cell calculated-date">${formatDisplayDate(value) || ""}</div>`;
}

function isEditableColumn(column, index) {
  return ["scope", "task", "owner", "duration"].includes(column) || (column === "start" && index === 0);
}

function clearEditableCell(project, step, index, column) {
  if (column === "scope" || column === "task" || column === "owner") step[column] = "";
  if (column === "duration") step.duration = "";
  if (column === "start") {
    if (index === 0) {
      project.startDate = "";
    } else {
      step.manualStartDate = "";
    }
  }
  if (column === "end") step.manualEndDate = "";
}

function bindColumnResizers() {
  const table = document.querySelector(".timeline-table");
  if (!table) return;
  const columns = [...table.querySelectorAll("col")];
  table.querySelectorAll(".col-resizer").forEach((resizer, index) => {
    resizer.addEventListener("mousedown", (event) => {
      event.preventDefault();
      event.stopPropagation();
      const startX = event.clientX;
      const startWidth = columns[index].getBoundingClientRect().width;
      const onMove = (moveEvent) => {
        const nextWidth = Math.max(64, startWidth + moveEvent.clientX - startX);
        columns[index].style.width = `${nextWidth}px`;
      };
      const onUp = () => {
        document.removeEventListener("mousemove", onMove);
        document.removeEventListener("mouseup", onUp);
      };
      document.addEventListener("mousemove", onMove);
      document.addEventListener("mouseup", onUp);
    });
  });
}

function saveAndRender(message = "") {
  const project = activeProject();
  if (project) project.updatedAt = new Date().toISOString();
  saveState();
  render();
  if (message) setStatus(message);
}

function setStatus(message) {
  els.statusLine.textContent = message;
  if (message) setTimeout(() => {
    if (els.statusLine.textContent === message) els.statusLine.textContent = "";
  }, 3500);
}

function download(filename, content, type) {
  const blob = new Blob([content], { type });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function todayIso() {
  const date = new Date();
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function id() {
  return crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function swap(list, a, b) {
  [list[a], list[b]] = [list[b], list[a]];
}

function safeFileName(name) {
  return String(name || "timeline").replace(/[\\/:*?"<>|]+/g, "-").trim() || "timeline";
}

function isWeekendDate(isoDate) {
  const day = new Date(`${isoDate}T00:00:00`).getDay();
  return day === 0 || day === 6;
}

function escapeHtml(value) {
  return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function escapeAttr(value) {
  return escapeHtml(value).replaceAll('"', "&quot;");
}
