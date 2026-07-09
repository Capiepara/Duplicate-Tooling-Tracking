let workbook = null;
let rawData = [];
let ganttRows = [];
let allWeeks = [];

const STAGES = [
  { name: "2D/3D Drawing", color: "#900C3F" },
  { name: "Tooling Spec Date", color: "#182B55" },
  { name: "Quotation", color: "#4871CC" },
  { name: "CAR", color: "#5F4E94" },
  { name: "POR", color: "#A291C7" },
  { name: "Tooling build", color: "#82CBEC" },
  { name: "FAI", color: "#D94F21" },
  { name: "QB Building", color: "#FEBD2B" },
  { name: "QB Testing complete date", color: "#9AAB4B" },
  { name: "PR building", color: "#D6568C" },
  { name: "SER fully release date", color: "#009999" },
];

const WEEK_WIDTH = 42;
const DAY_WIDTH = WEEK_WIDTH / 7;

const excelFile = document.getElementById("excelFile");
const fileName = document.getElementById("fileName");
const sheetSelect = document.getElementById("sheetSelect");
const loadSheetBtn = document.getElementById("loadSheetBtn");
const ganttGrid = document.getElementById("ganttGrid");
const summaryText = document.getElementById("summaryText");
const partSearch = document.getElementById("partSearch");
const stageFilter = document.getElementById("stageFilter");
const resetBtn = document.getElementById("resetBtn");

excelFile.addEventListener("change", handleFileImport);
loadSheetBtn.addEventListener("click", loadSelectedSheet);
partSearch.addEventListener("input", renderGantt);
stageFilter.addEventListener("change", renderGantt);
resetBtn.addEventListener("click", resetFilters);

function handleFileImport(event) {
  const file = event.target.files[0];
  if (!file) return;

  fileName.textContent = file.name;

  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);

    workbook = XLSX.read(data, {
      type: "array",
      cellDates: true,
    });

    sheetSelect.innerHTML = "";

    workbook.SheetNames.forEach((sheet) => {
      const option = document.createElement("option");
      option.value = sheet;
      option.textContent = sheet;
      sheetSelect.appendChild(option);
    });

    sheetSelect.disabled = false;
    loadSheetBtn.disabled = false;
  };

  reader.readAsArrayBuffer(file);
}

function loadSelectedSheet() {
  const sheetName = sheetSelect.value;
  const ws = workbook.Sheets[sheetName];

  const range = XLSX.utils.decode_range(ws["!ref"]);
  range.s.c = 0;
  range.e.c = 35;

  rawData = XLSX.utils.sheet_to_json(ws, {
    header: 1,
    range: XLSX.utils.encode_range(range),
    defval: "",
    raw: false,
  });

  buildStageFilter();
  parseGanttRows();
  renderGantt();
}

function buildStageFilter() {
  stageFilter.innerHTML = `<option value="ALL">All Stages</option>`;

  STAGES.forEach((stage) => {
    const option = document.createElement("option");
    option.value = stage.name;
    option.textContent = stage.name;
    stageFilter.appendChild(option);
  });
}

function parseGanttRows() {
  ganttRows = [];

  const headerRow = rawData[0] || [];
  const subHeaderRow = rawData[1] || [];

  const stageMap = getStageColumnMap(headerRow, subHeaderRow);

  for (let i = 2; i < rawData.length; i++) {
    const row = rawData[i];

    const partNo = row[0] || "";
    const type = row[1] || "";

    if (!partNo && !type) continue;
    if (!["Original", "Actual"].includes(type)) continue;

    const stages = [];

    STAGES.forEach((stage) => {
      const cols = stageMap[stage.name];
      if (!cols) return;

      const target = parseDate(row[cols.target]);
      const estimate = parseDate(row[cols.estimate]);
      const actual = parseDate(row[cols.actual]);

      const start = target;
      const end = actual || estimate || target;

      if (!start || !end) return;

      stages.push({
        name: stage.name,
        color: stage.color,
        start,
        end,
        target,
        estimate,
        actual,
      });
    });

    ganttRows.push({
      partNo,
      type,
      stages,
    });
  }

  allWeeks = generateWeeks(ganttRows);
}

function getStageColumnMap(headerRow, subHeaderRow) {
  const map = {};
  let currentStage = "";

  for (let c = 0; c < headerRow.length; c++) {
    if (headerRow[c]) {
      currentStage = normalizeStageName(headerRow[c]);
    }

    const sub = String(subHeaderRow[c] || "").trim();

    const stage = STAGES.find((s) => normalizeStageName(s.name) === currentStage);
    if (!stage) continue;

    if (!map[stage.name]) {
      map[stage.name] = {};
    }

    if (sub === "Target") map[stage.name].target = c;
    if (sub === "Estimate") map[stage.name].estimate = c;
    if (sub === "Actual") map[stage.name].actual = c;
  }

  return map;
}

function normalizeStageName(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function parseDate(value) {
  if (!value) return null;

  if (value instanceof Date && !isNaN(value)) return value;

  const text = String(value).trim();

  const parts = text.split(/[\/\-]/);
  if (parts.length >= 2) {
    const month = Number(parts[0]);
    const day = Number(parts[1]);
    const year = parts[2] ? Number(parts[2]) : 2026;

    if (!isNaN(month) && !isNaN(day)) {
      return new Date(year, month - 1, day);
    }
  }

  const parsed = new Date(text);
  if (!isNaN(parsed)) return parsed;

  return null;
}

function generateWeeks(rows) {
  let minDate = null;
  let maxDate = null;

  rows.forEach((row) => {
    row.stages.forEach((stage) => {
      if (!minDate || stage.start < minDate) minDate = stage.start;
      if (!maxDate || stage.end > maxDate) maxDate = stage.end;
    });
  });

  if (!minDate || !maxDate) return [];

  const start = getMonday(minDate);
  const end = getMonday(maxDate);

  const weeks = [];
  let cur = new Date(start);

  while (cur <= end) {
    weeks.push(new Date(cur));
    cur.setDate(cur.getDate() + 7);
  }

  return weeks;
}

function getMonday(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

function getWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;

  d.setUTCDate(d.getUTCDate() + 4 - dayNum);

  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));

  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

function renderGantt() {
  ganttGrid.innerHTML = "";

  const search = partSearch.value.trim().toLowerCase();
  const selectedStage = stageFilter.value;

  const filteredRows = ganttRows.filter((row) => {
    const matchPart = String(row.partNo).toLowerCase().includes(search);
    return matchPart;
  });

  renderHeader();

  filteredRows.forEach((row) => {
    renderRow(row, selectedStage);
  });

  summaryText.textContent = `${filteredRows.length} rows`;
}

function renderHeader() {
  const header = document.createElement("div");
  header.className = "gantt-header";

  header.innerHTML = `
    <div class="left-head">Part No.</div>
    <div class="left-head">Original / Actual</div>
    <div class="timeline-head">
      ${allWeeks.map((w) => `<div class="week-cell">WK ${getWeekNumber(w)}</div>`).join("")}
    </div>
  `;

  ganttGrid.appendChild(header);
}

function renderRow(row, selectedStage) {
  const rowDiv = document.createElement("div");
  rowDiv.className = "gantt-row";

  const partCell = document.createElement("div");
  partCell.className = "part-cell";
  partCell.textContent = row.partNo;

  const typeCell = document.createElement("div");
  typeCell.className = `type-cell ${row.type === "Original" ? "type-original" : "type-actual"}`;
  typeCell.textContent = row.type;

  const timeline = document.createElement("div");
  timeline.className = "timeline-cell";
  timeline.style.width = `${allWeeks.length * WEEK_WIDTH}px`;

  row.stages
    .filter((stage) => selectedStage === "ALL" || stage.name === selectedStage)
    .forEach((stage) => {
      const bar = document.createElement("div");
      bar.className = "stage-bar";
      bar.style.background = stage.color;

      const left = getDateOffset(stage.start);
      const width = Math.max(getDateDiff(stage.start, stage.end) * DAY_WIDTH, 8);

      bar.style.left = `${left}px`;
      bar.style.width = `${width}px`;

      bar.dataset.tip =
        `${stage.name}\n` +
        `Target: ${formatDate(stage.target)}\n` +
        `Estimate: ${formatDate(stage.estimate)}\n` +
        `Actual: ${formatDate(stage.actual)}\n` +
        `Using: ${stage.actual ? "Actual" : "Estimate"}`;

      timeline.appendChild(bar);
    });

  rowDiv.appendChild(partCell);
  rowDiv.appendChild(typeCell);
  rowDiv.appendChild(timeline);

  ganttGrid.appendChild(rowDiv);
}

function getDateOffset(date) {
  if (!allWeeks.length) return 0;

  const start = allWeeks[0];
  const diffDays = getDateDiff(start, date);

  return diffDays * DAY_WIDTH;
}

function getDateDiff(start, end) {
  const s = new Date(start);
  const e = new Date(end);

  s.setHours(0, 0, 0, 0);
  e.setHours(0, 0, 0, 0);

  return Math.max(1, Math.round((e - s) / 86400000) + 1);
}

function formatDate(date) {
  if (!date) return "";

  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");

  return `${mm}/${dd}`;
}

function resetFilters() {
  partSearch.value = "";
  stageFilter.value = "ALL";
  renderGantt();
}
