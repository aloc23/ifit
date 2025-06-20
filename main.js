
let workbook, worksheetData = [], headerRow = [], weekMap = {};
document.getElementById("fileInput").addEventListener("change", handleFile);

function handleFile(event) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const wb = XLSX.read(data, { type: "array" });
    const sheetName = wb.SheetNames[0];
    const ws = wb.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
    worksheetData = json;
    headerRow = json[3];
    populateWeekDropdown(headerRow);
    renderTable();
  };
  reader.readAsArrayBuffer(event.target.files[0]);
}

function populateWeekDropdown(header) {
  const select = document.getElementById("weekSelect");
  select.innerHTML = "";
  header.forEach((text, i) => {
    if (typeof text === "string" && text.toLowerCase().includes("week")) {
      const option = document.createElement("option");
      option.value = i;
      option.text = text;
      select.appendChild(option);
    }
  });
}

function applyRepayment() {
  const col = parseInt(document.getElementById("weekSelect").value);
  const amount = parseFloat(document.getElementById("amountInput").value);
  const rowIdx = worksheetData.findIndex(r => r.some(c => typeof c === "string" && c.toLowerCase().includes("mayweather")));
  if (rowIdx === -1) return alert("Mayweather row not found.");
  worksheetData[rowIdx][col] = (worksheetData[rowIdx][col] || 0) - amount;
  renderTable();
}

function renderTable() {
  const table = document.createElement("table");
  worksheetData.forEach(row => {
    const tr = document.createElement("tr");
    row.forEach(cell => {
      const td = document.createElement("td");
      td.textContent = cell || "";
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
  document.getElementById("tableContainer").innerHTML = "";
  document.getElementById("tableContainer").appendChild(table);
}

function downloadCSV() {
  const csv = worksheetData.map(row => row.map(c => `"${c}"`).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "updated_cashflow.csv";
  a.click();
}
