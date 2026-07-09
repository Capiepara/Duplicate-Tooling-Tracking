let workbook = null;
let rawData = [];
let ganttRows = [];
let allWeeks = [];

const WEEK_WIDTH = 44;
const DAY_WIDTH = WEEK_WIDTH / 7;

const STAGES = [
  { key: "2d", name: "2D/3D Drawing", color: "#900C3F", aliases: ["2d/3d drawing", "2d", "3d"] },
  { key: "toolingSpec", name: "Tooling Spec Date", color: "#182B55", aliases: ["tooling spec date", "tooling spec"] },
  { key: "quotation", name: "Quotation", color: "#4871CC", aliases: ["quotation"] },
  { key: "car", name: "CAR", color: "#5F4E94", aliases: ["car"] },
  { key: "por", name: "POR", color: "#A291C7", aliases: ["por"] },
  { key: "toolingBuild", name: "Tooling Build", color: "#82CBEC", aliases: ["tooling build", "tooling buid"] },
  { key: "fai", name: "FAI", color: "#D94F21", aliases: ["fai"] },
  { key: "qbBuilding", name: "QB Building", color: "#FEBD2B", aliases: ["qb building"] },
  { key: "qbTesting", name: "QB Testing Complete", color: "#9AAB4B", aliases: ["qb testing complete", "qb testing complete date"] },
  { key: "pr", name: "PR Building", color: "#D6568C", aliases: ["pr building"] },
  { key: "ser", name: "SER Fully Release", color: "#009999", aliases: ["ser fully release", "ser fully release date"] },
];

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
    option.value = stage.key;
    option.textContent = stage.name;
    stageFilter.appendChild(option);
  });
  renderLegend();
}

function parseGanttRows() {
  ganttRows = [];

  const headerInfo = findHeaderRows();
  if (!headerInfo) {
    alert("Không tìm thấy header stage. Kiểm tra file Excel.");
    return;
  }

  const { stageRowIndex, subRowIndex } = headerInfo;
  const stageMap = getStageColumnMap(stageRowIndex, subRowIndex);

  let currentPart = "";

  for (let r = subRowIndex + 1; r < rawData.length; r++) {
    const row = rawData[r];

    const partNo = clean(row[0]) || currentPart;
    const type = clean(row[1]);

    if (clean(row[0])) currentPart = clean(row[0]);

    if (!partNo) continue;
    if (type !== "Original" && type !== "Actual") continue;

    const stages = [];

    STAGES.forEach((stage) => {
      const cols = stageMap[stage.key];
      if (!cols) return;

      const target = parseDate(row[cols.target]);
      const estimate = parseDate(row[cols.estimate]);
      const actual = parseDate(row[cols.actual]);

      let start = target;
      let end = actual || estimate || target;

      if (!start || !end) return;

      if (end < start) {
        const temp = start;
        start = end;
        end = temp;
      }

      stages.push({
        key: stage.key,
        name: stage.name,
        color: stage.color,
        start,
        end,
        target,
        estimate,
        actual,
        using: actual ? "Actual" : "Estimate",
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

function findHeaderRows() {
  for (let r = 0; r < Math.min(rawData.length, 12); r++) {
    const rowText = rawData[r].map(clean).join(" ").toLowerCase();

    const hasStage = STAGES.some((s) =>
      s.aliases.some((a) => rowText.includes(a))
    );

    const nextRow = rawData[r + 1] || [];
    const nextText = nextRow.map(clean).join(" ").toLowerCase();

    const hasSubHeader =
      nextText.includes("target") ||
      nextText.includes("estimate") ||
      nextText.includes("actual");

    if (hasStage && hasSubHeader) {
      return {
        stageRowIndex: r,
        subRowIndex: r + 1,
      };
    }
  }

  return null;
}

function getStageColumnMap(stageRowIndex, subRowIndex) {
  const map = {};
  const stageRow = rawData[stageRowIndex];
  const subRow = rawData[subRowIndex];

  let currentStageKey = null;

  for (let c = 0; c < stageRow.length; c++) {
    const stageText = normalize(stageRow[c]);

    const matchedStage = STAGES.find((stage) =>
      stage.aliases.some((alias) => stageText.includes(alias))
    );

    if (matchedStage) {
      currentStageKey = matchedStage.key;
    }

    if (!currentStageKey) continue;

    const sub = normalize(subRow[c]);

    if (!map[currentStageKey]) {
      map[currentStageKey] = {};
    }

    if (sub.includes("target")) map[currentStageKey].target = c;
    if (sub.includes("estimate")) map[currentStageKey].estimate = c;
    if (sub.includes("actual")) map[currentStageKey].actual = c;
  }

  return map;
}

function renderGantt() {
  ganttGrid.innerHTML = "";

  const search = partSearch.value.trim().toLowerCase();
  const selectedStage = stageFilter.value;

  const filteredRows = ganttRows.filter((row) => {
    return String(row.partNo).toLowerCase().includes(search);
  });

  renderHeader();

  filteredRows.forEach((row) => {
    renderRow(row, selectedStage);
  });

  summaryText.textContent = `${filteredRows.length} rows`;
  updateKPI(filteredRows);
}

function renderHeader() {
  const header = document.createElement("div");
  header.className = "gantt-header";

  const timelineWidth = allWeeks.length * WEEK_WIDTH;

  header.innerHTML = `
    <div class="left-head">Part No.</div>
    <div class="left-head">Original<br>/ Actual</div>
    <div class="timeline-head" style="width:${timelineWidth}px">
      ${allWeeks.map((w) => `
        <div class="week-cell">
          <div>WK ${getWeekNumber(w)}</div>
          <small>${formatDate(w)}</small>
        </div>
      `).join("")}
    </div>
  `;

  ganttGrid.appendChild(header);
}

function renderRow(row, selectedStage) {
  const rowDiv = document.createElement("div");
  rowDiv.className = "gantt-row";

  const partCell = document.createElement("div");
  partCell.className = "part-cell";
  partCell.textContent = row.type === "Original" ? row.partNo : "";

  const typeCell = document.createElement("div");
  typeCell.className = `type-cell ${row.type === "Original" ? "type-original" : "type-actual"}`;
  typeCell.textContent = row.type;

  const timeline = document.createElement("div");
  timeline.className = "timeline-cell";
  timeline.style.width = `${allWeeks.length * WEEK_WIDTH}px`;

  row.stages
    .filter((stage) => selectedStage === "ALL" || stage.key === selectedStage)
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
        `Using: ${stage.using}`;

      timeline.appendChild(bar);
    });

  rowDiv.appendChild(partCell);
  rowDiv.appendChild(typeCell);
  rowDiv.appendChild(timeline);

  ganttGrid.appendChild(rowDiv);
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

function getDateOffset(date) {
  if (!allWeeks.length || !date) return 0;

  const start = allWeeks[0];
  const diffDays = getDateDiff(start, date) - 1;

  return diffDays * DAY_WIDTH;
}

function getDateDiff(start, end) {
  const s = new Date(start);
  const e = new Date(end);

  s.setHours(0, 0, 0, 0);
  e.setHours(0, 0, 0, 0);

  return Math.round((e - s) / 86400000) + 1;
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

function parseDate(value) {
  if (!value) return null;

  if (value instanceof Date && !isNaN(value)) return value;

  const text = String(value).trim();
  if (!text) return null;

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

function formatDate(date) {
  if (!date) return "";

  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");

  return `${mm}/${dd}`;
}

function normalize(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function clean(value) {
  return String(value || "").trim();
}

function resetFilters() {
  partSearch.value = "";
  stageFilter.value = "ALL";
  renderGantt();
}
function renderLegend() {
  const legendPanel = document.getElementById("legendPanel");
  if (!legendPanel) return;

  legendPanel.innerHTML = STAGES.map(stage => `
    <div class="legend-item">
      <span class="legend-color" style="background:${stage.color}"></span>
      <span>${stage.name}</span>
    </div>
  `).join("");
}

function updateKPI(filteredRows) {
  const kpiPart = document.getElementById("kpiPart");
  const kpiStage = document.getElementById("kpiStage");
  const kpiDate = document.getElementById("kpiDate");
  const kpiDelay = document.getElementById("kpiDelay");
  const kpiStageColor = document.getElementById("kpiStageColor");

  if (!filteredRows.length) {
    kpiPart.textContent = "-";
    kpiStage.textContent = "-";
    kpiDate.textContent = "-";
    kpiDelay.textContent = "-";
    kpiStageColor.style.background = "#d1d5db";
    return;
  }

  const search = partSearch.value.trim();

  if (!search) {
    kpiPart.textContent = "All";
    kpiStage.textContent = "Multiple";
    kpiDate.textContent = "-";
    kpiDelay.textContent = "-";
    kpiStageColor.style.background = "#d1d5db";
    return;
  }

  const actualRow = filteredRows.find(r => r.type === "Actual") || filteredRows[0];

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  let currentStage =
    actualRow.stages.find(s => today >= s.start && today <= s.end) ||
    actualRow.stages.find(s => today < s.start) ||
    actualRow.stages[actualRow.stages.length - 1];

  if (!currentStage) return;

  let delayText = "-";

  if (currentStage.actual && currentStage.estimate) {
    const delayDays = getDateDiff(currentStage.estimate, currentStage.actual) - 1;
    delayText = delayDays > 0 ? `+${delayDays} days` : "On time";
  }

  kpiPart.textContent = actualRow.partNo;
  kpiStage.textContent = currentStage.name;
  kpiDate.textContent = formatDate(currentStage.actual || currentStage.estimate || currentStage.end);
  kpiDelay.textContent = delayText;
  kpiStageColor.style.background = currentStage.color;
}
