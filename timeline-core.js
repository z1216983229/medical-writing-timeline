const ISO_DATE_RE = /^\d{4}-\d{2}-\d{2}$/;

export function buildCalendar(data = {}) {
  return {
    offDays: new Set((data.offDays || []).map((item) => normalizeDate(item.date || item))),
    workDays: new Set((data.workDays || []).map((item) => normalizeDate(item.date || item))),
    meta: data.meta || {},
  };
}

export function normalizeDate(value) {
  if (!value) return "";
  if (value instanceof Date) return toIsoDate(value);
  const text = String(value).trim();
  if (ISO_DATE_RE.test(text)) return text;
  const slashMatch = text.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
  if (slashMatch) {
    const [, year, month, day] = slashMatch;
    return `${year}-${month.padStart(2, "0")}-${day.padStart(2, "0")}`;
  }
  return "";
}

export function formatDisplayDate(value) {
  const iso = normalizeDate(value);
  return iso ? iso.replaceAll("-", "/") : "";
}

export function isWeekend(isoDate) {
  const day = fromIsoDate(isoDate).getDay();
  return day === 0 || day === 6;
}

export function isWorkday(isoDate, calendar = buildCalendar()) {
  const date = normalizeDate(isoDate);
  if (!date) return false;
  if (calendar.workDays.has(date)) return true;
  if (calendar.offDays.has(date)) return false;
  return !isWeekend(date);
}

export function nextWorkday(isoDate, calendar) {
  let cursor = addDays(normalizeDate(isoDate), 1);
  while (!isWorkday(cursor, calendar)) {
    cursor = addDays(cursor, 1);
  }
  return cursor;
}

export function firstWorkdayOnOrAfter(isoDate, calendar) {
  let cursor = normalizeDate(isoDate);
  while (cursor && !isWorkday(cursor, calendar)) {
    cursor = addDays(cursor, 1);
  }
  return cursor;
}

export function addWorkdaysInclusive(startDate, duration, calendar) {
  const count = Number(duration);
  let cursor = firstWorkdayOnOrAfter(startDate, calendar);
  if (!cursor || !Number.isFinite(count) || count < 1) return "";
  let remaining = Math.floor(count);
  while (remaining > 1) {
    cursor = nextWorkday(cursor, calendar);
    remaining -= 1;
  }
  return cursor;
}

export function recalculateSteps(steps, projectStartDate, calendar) {
  let cursor = firstWorkdayOnOrAfter(projectStartDate, calendar);
  return steps.map((step, index) => {
    const duration = Number(step.duration);
    const result = { ...step, order: index + 1, startDate: "", endDate: "", warning: "" };

    if (!cursor) {
      return result;
    }

    result.startDate = cursor;

    if (!Number.isFinite(duration) || duration < 1) {
      result.endDate = "";
      return result;
    }

    result.duration = Math.floor(duration);
    result.endDate = addWorkdaysInclusive(result.startDate, result.duration, calendar);
    cursor = nextWorkday(result.endDate, calendar);
    return result;
  });
}

export function toTsv(project, steps) {
  const rows = [
    ["序号", "任务范围", "任务名称", "任务负责人", "Duration 工作日", "计划开始日期", "计划结束日期"],
    ...steps.map((step) => [
      step.order,
      step.scope,
      step.task,
      step.owner,
      step.duration,
      formatDisplayDate(step.startDate),
      formatDisplayDate(step.endDate),
    ]),
  ];
  return rows.map((row) => row.map(escapeTsvCell).join("\t")).join("\n");
}

export function buildExcelHtml(project, steps, holidayMeta = {}) {
  const infoRows = [
    ["项目名称", project.name || ""],
    ["模板名称", project.templateName || ""],
    ["项目开始日期", formatDisplayDate(project.startDate)],
    ["节假日年份", project.holidayYear || ""],
    ["节假日版本", holidayMeta.source || "内置数据"],
    ["导出时间", new Date().toLocaleString("zh-CN")],
  ];
  const tableRows = [
    ["序号", "任务范围", "任务名称", "任务负责人", "Duration 工作日", "计划开始日期", "计划结束日期"],
    ...steps.map((step) => [
      step.order,
      step.scope,
      step.task,
      step.owner,
      step.duration,
      formatDisplayDate(step.startDate),
      formatDisplayDate(step.endDate),
    ]),
  ];
  const info = infoRows.map((row) => `<tr>${row.map(htmlCell).join("")}</tr>`).join("");
  const body = tableRows.map((row, index) => `<tr>${row.map((cell) => htmlCell(cell, index === 0)).join("")}</tr>`).join("");
  return `<!doctype html><html><head><meta charset="utf-8"><style>
    body{font-family:Arial,"Microsoft YaHei",sans-serif}
    table{border-collapse:collapse}
    td,th{border:1px solid #999;padding:6px 8px;white-space:nowrap}
    th{background:#e8eef7;font-weight:700}
    .info td:first-child{font-weight:700;background:#f5f5f5}
  </style></head><body><table class="info">${info}</table><br><table>${body}</table></body></html>`;
}

function escapeTsvCell(value) {
  return String(value ?? "").replace(/\t/g, " ").replace(/\r?\n/g, " ");
}

function htmlCell(value, header = false) {
  const tag = header ? "th" : "td";
  return `<${tag}>${String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")}</${tag}>`;
}

function fromIsoDate(isoDate) {
  return new Date(`${isoDate}T00:00:00`);
}

function toIsoDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function addDays(isoDate, count) {
  const date = fromIsoDate(isoDate);
  date.setDate(date.getDate() + count);
  return toIsoDate(date);
}
