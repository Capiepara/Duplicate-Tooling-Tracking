let workbook = null;
let currentData = [];

const excelFile = document.getElementById("excelFile");
const fileName = document.getElementById("fileName");
const sheetSelect = document.getElementById("sheetSelect");
const loadSheetBtn = document.getElementById("loadSheetBtn");
const previewTable = document.getElementById("previewTable");
const rowCount = document.getElementById("rowCount");

excelFile.addEventListener("change", handleFileImport);
loadSheetBtn.addEventListener("click", loadSelectedSheet);

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

    renderSheetOptions(workbook.SheetNames);
  };

  reader.readAsArrayBuffer(file);
}

function renderSheetOptions(sheetNames) {
  sheetSelect.innerHTML = "";

  sheetNames.forEach((sheetName) => {
    const option = document.createElement("option");
    option.value = sheetName;
    option.textContent = sheetName;
    sheetSelect.appendChild(option);
  });

  sheetSelect.disabled = false;
  loadSheetBtn.disabled = false;
}

function loadSelectedSheet() {
  const selectedSheet = sheetSelect.value;

  if (!workbook || !selectedSheet) return;

  const worksheet = workbook.Sheets[selectedSheet];

  // Chỉ đọc vùng A:AJ
  const range = XLSX.utils.decode_range(worksheet["!ref"]);
  range.s.c = 0;   // A
  range.e.c = 35;  // AJ

  const fixedRange = XLSX.utils.encode_range(range);

  currentData = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    range: fixedRange,
    defval: "",
    raw: false,
  });

  renderPreviewTable(currentData);
}

function renderPreviewTable(data) {
  const thead = previewTable.querySelector("thead");
  const tbody = previewTable.querySelector("tbody");

  thead.innerHTML = "";
  tbody.innerHTML = "";

  if (!data || data.length === 0) {
    rowCount.textContent = "0 rows";
    return;
  }

  const headerRow = data[0];
  const bodyRows = data.slice(1);

  const trHead = document.createElement("tr");

  headerRow.forEach((cell) => {
    const th = document.createElement("th");
    th.textContent = cell;
    trHead.appendChild(th);
  });

  thead.appendChild(trHead);

  bodyRows.forEach((row) => {
    const tr = document.createElement("tr");

    row.forEach((cell) => {
      const td = document.createElement("td");
      td.textContent = formatCell(cell);
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  });

  rowCount.textContent = `${bodyRows.length} rows`;
}

function formatCell(value) {
  if (value === null || value === undefined) return "";
  return value;
}
