// Excel -> JavaScript -> Website
// Reads an .xlsx file in the browser with the XLSX library and renders it as a table.

const fileInput = document.getElementById("fileInput");
const dropZone = document.getElementById("dropZone");
const statusEl = document.getElementById("status");
const tabsEl = document.getElementById("tabs");
const tableWrap = document.getElementById("tableWrap");
const tableEl = document.getElementById("table");

let workbook = null;

// --- File selection -------------------------------------------------------

fileInput.addEventListener("change", () => {
  if (fileInput.files.length > 0) {
    loadFile(fileInput.files[0]);
  }
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("dragover");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("dragover");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("dragover");
  if (event.dataTransfer.files.length > 0) {
    loadFile(event.dataTransfer.files[0]);
  }
});

// --- Reading the Excel file -----------------------------------------------

function loadFile(file) {
  if (!file.name.toLowerCase().endsWith(".xlsx")) {
    showError("Please choose an .xlsx file.");
    return;
  }

  const reader = new FileReader();

  reader.onload = (event) => {
    try {
      const data = new Uint8Array(event.target.result);
      workbook = XLSX.read(data, { type: "array" });
      renderTabs();
      showSheet(workbook.SheetNames[0]);
      setStatus("Loaded: " + file.name);
    } catch (error) {
      showError("This file could not be read. Is it a valid .xlsx file?");
    }
  };

  reader.onerror = () => showError("The file could not be opened.");

  setStatus("Reading " + file.name + " ...");
  reader.readAsArrayBuffer(file);
}

// --- Sheet tabs -----------------------------------------------------------

function renderTabs() {
  tabsEl.innerHTML = "";

  // Tabs are only useful when the workbook has more than one sheet.
  if (workbook.SheetNames.length < 2) return;

  workbook.SheetNames.forEach((name) => {
    const button = document.createElement("button");
    button.type = "button";
    button.textContent = name;
    button.addEventListener("click", () => showSheet(name));
    tabsEl.appendChild(button);
  });
}

function showSheet(name) {
  Array.from(tabsEl.children).forEach((button) => {
    button.classList.toggle("active", button.textContent === name);
  });
  renderTable(workbook.Sheets[name]);
}

// --- Building the table ---------------------------------------------------

function renderTable(sheet) {
  tableEl.innerHTML = "";

  if (!sheet || !sheet["!ref"]) {
    tableWrap.hidden = true;
    setStatus("This sheet is empty.");
    return;
  }

  // The used cell range of the sheet, e.g. A1:D20
  const range = XLSX.utils.decode_range(sheet["!ref"]);

  const thead = document.createElement("thead");
  const tbody = document.createElement("tbody");

  for (let row = range.s.r; row <= range.e.r; row++) {
    const tr = document.createElement("tr");
    const isHeader = row === range.s.r; // first row = column headings

    for (let col = range.s.c; col <= range.e.c; col++) {
      const address = XLSX.utils.encode_cell({ r: row, c: col });
      const cell = document.createElement(isHeader ? "th" : "td");
      cell.textContent = cellText(sheet[address]);
      tr.appendChild(cell);
    }

    (isHeader ? thead : tbody).appendChild(tr);
  }

  tableEl.appendChild(thead);
  tableEl.appendChild(tbody);
  tableWrap.hidden = false;
}

// Returns the cell's text as Excel displays it (formatted dates, numbers, ...).
function cellText(cell) {
  if (!cell) return "";
  if (cell.w !== undefined) return cell.w;
  if (cell.v !== undefined) return String(cell.v);
  return "";
}

// --- Status messages ------------------------------------------------------

function setStatus(message) {
  statusEl.textContent = message;
  statusEl.classList.remove("error");
}

function showError(message) {
  statusEl.textContent = message;
  statusEl.classList.add("error");
  tabsEl.innerHTML = "";
  tableWrap.hidden = true;
}
