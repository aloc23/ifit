
let worksheet = [], fileRaw;

document.getElementById("fileInput").addEventListener("change", handleFile);

function handleFile(evt) {
  const reader = new FileReader();
  reader.onload = e => {
    const data = new Uint8Array(e.target.result);
    const wb = XLSX.read(data, { type: "array" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws, { header: 1 });
    worksheet = json;
    fileRaw = ws;
    buildWeekDropdown();
    renderTable();
  };
  reader.readAsArrayBuffer(evt.target.files[0]);
}

function buildWeekDropdown() {
  const select = document.getElementById("weekSelect");
  select.innerHTML = "";
  const today = new Date();
  const dateRow = worksheet[2];
  const labelRow = worksheet[3];

  labelRow.forEach((label, i) => {
    try {
      const parsed = new Date(dateRow[i]);
      if (!isNaN(parsed) && parsed >= today && String(label).toLowerCase().includes("week")) {
        const option = document.createElement("option");
        option.value = i;
        option.text = `${label.trim()} (${parsed.getFullYear()})`;
        select.appendChild(option);
      }
    } catch (e) {}
  });
}

function applyRepayment() {
  const col = parseInt(document.getElementById("weekSelect").value);
  const amount = parseFloat(document.getElementById("amountInput").value);
  if (isNaN(col) || isNaN(amount)) return alert("Choose week and enter amount");
  const rowIndex = worksheet.findIndex(r => r.some(c => typeof c === "string" && c.toLowerCase().includes("mayweather")));
  if (rowIndex === -1) return alert("Mayweather repayment row not found.");
  worksheet[rowIndex][col] = (worksheet[rowIndex][col] || 0) - amount;
  renderTable();
}

function renderTable() {
  const table = document.createElement("table");
  worksheet.forEach(row => {
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
  const ws = XLSX.utils.aoa_to_sheet(worksheet);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Updated");
  XLSX.writeFile(wb, "cashflow_updated.xlsx");
}
