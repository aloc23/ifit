
let workbook, sheet, data = [];
let repaymentRows = [];
let remaining = 355000;
let chartInstance;

document.getElementById("fileUpload").addEventListener("change", e => {
  const file = e.target.files[0];
  const reader = new FileReader();
  reader.onload = function(evt) {
    const workbookData = XLSX.read(evt.target.result, { type: "binary" });
    workbook = workbookData;
    sheet = workbook.Sheets[workbook.SheetNames[0]];
    data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    populateWeekDropdowns();
    renderTable();
    updateChart();
  };
  reader.readAsBinaryString(file);
});

function populateWeekDropdowns() {
  const headers = data[3];
  const dropdowns = document.querySelectorAll(".week-dropdown");
  dropdowns.forEach(dropdown => {
    dropdown.innerHTML = "";
    headers.forEach((label, idx) => {
      if (label && label.toString().includes("Week")) {
        dropdown.innerHTML += `<option value="${idx}">${label}</option>`;
      }
    });
  });
}

function renderTable() {
  const preview = document.getElementById("tablePreview");
  preview.innerHTML = "";
  const table = document.createElement("table");
  data.forEach((row, i) => {
    const tr = document.createElement("tr");
    row.forEach((cell, j) => {
      const td = document.createElement(i === 0 ? "th" : "td");
      td.textContent = cell;
      if (i === 134 && repaymentRows.some(r => r.col == j)) {
        td.classList.add("highlight");
      }
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
  preview.appendChild(table);
}

function updateChart() {
  if (!data[135]) return;
  const labels = data[3].slice(5);
  const values = data[135].slice(5).map(v => parseFloat(v) || 0);
  const ctx = document.getElementById("forecastChart").getContext("2d");
  if (chartInstance) {
    chartInstance.data.labels = labels;
    chartInstance.data.datasets[0].data = values;
    chartInstance.update();
  } else {
    chartInstance = new Chart(ctx, {
      type: "line",
      data: {
        labels: labels,
        datasets: [{
          label: "Rolling Balance",
          data: values,
          borderColor: "green",
          fill: false
        }]
      }
    });
  }
}

document.getElementById("addRepaymentRow").addEventListener("click", () => {
  const row = document.createElement("div");
  row.className = "repayment-row";
  row.innerHTML = '<select class="week-dropdown"></select><input type="number" class="repayment-amount" placeholder="Amount €" />';
  document.getElementById("repaymentForm").appendChild(row);
  populateWeekDropdowns();
});

document.getElementById("applyRepayments").addEventListener("click", () => {
  repaymentRows = [];
  document.querySelectorAll(".repayment-row").forEach(r => {
    const col = parseInt(r.querySelector("select").value);
    const amount = parseFloat(r.querySelector("input").value);
    if (!isNaN(col) && !isNaN(amount)) {
      repaymentRows.push({ col, amount });
      if (!data[134][col]) data[134][col] = 0;
      data[134][col] = parseFloat(data[134][col]) + amount;
      remaining -= amount;
    }
  });
  document.getElementById("remainingBalance").innerHTML = `Remaining Mayweather Balance: <strong>€${remaining.toLocaleString()}</strong>`;
  renderTable();
  updateChart();
});

document.getElementById("toggleTable").addEventListener("click", () => {
  document.getElementById("tablePreview").classList.toggle("collapsed");
});

document.getElementById("exportFile").addEventListener("click", () => {
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Updated Sheet");
  XLSX.writeFile(wb, "updated_cashflow.xlsx");
});
