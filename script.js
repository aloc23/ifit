// script.js
let csvData = [];
let repaymentRows = [];
let forecastChart;
let remaining = 355000;

function parseCSV(file, callback) {
  const reader = new FileReader();
  reader.onload = () => {
    const lines = reader.result.split(/\r?\n/);
    csvData = lines.map(l => l.split(","));
    callback();
  };
  reader.readAsText(file);
}

function populateWeekDropdown() {
  const dropdowns = document.querySelectorAll(".week-dropdown");
  const headerRow = csvData[3]; // Week text (row 4 Excel = index 3)
  dropdowns.forEach(dropdown => {
    dropdown.innerHTML = "";
    headerRow.forEach((label, idx) => {
      if (label.match(/Week \d+/) && csvData[2][idx].match(/\d{4}/)) {
        const year = csvData[2][idx].match(/\d{4}/)[0];
        dropdown.innerHTML += `<option value="${idx}">Week ${label.match(/\d+/)[0]} (${year})</option>`;
      }
    });
  });
}

function renderTable() {
  const table = document.createElement("table");
  csvData.forEach((row, i) => {
    const tr = document.createElement("tr");
    row.forEach((cell, j) => {
      const td = document.createElement(i === 0 ? "th" : "td");
      td.textContent = cell;
      if (csvData[133]?.[0]?.includes("Mayweather") && i === 134 && repaymentRows.find(r => r.col === j)) {
        td.style.background = "#ffeeba";
      }
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
  document.getElementById("tablePreview").innerHTML = "";
  document.getElementById("tablePreview").appendChild(table);
}

function updateForecastChart() {
  const labels = csvData[3]?.slice(5) || [];
  const values = csvData[135]?.slice(5).map(v => parseFloat(v) || 0);
  if (!forecastChart) {
    forecastChart = new Chart(document.getElementById("forecastChart"), {
      type: "line",
      data: {
        labels,
        datasets: [{
          label: "Rolling Cash Balance",
          data: values,
          borderColor: "blue",
          tension: 0.3
        }]
      },
      options: { responsive: true }
    });
  } else {
    forecastChart.data.labels = labels;
    forecastChart.data.datasets[0].data = values;
    forecastChart.update();
  }
}

document.getElementById("csvUpload").addEventListener("change", e => {
  parseCSV(e.target.files[0], () => {
    populateWeekDropdown();
    renderTable();
    updateForecastChart();
  });
});

document.getElementById("addRepaymentRow").addEventListener("click", () => {
  const row = document.createElement("div");
  row.className = "repayment-row";
  row.innerHTML = `<select class="week-dropdown"></select><input type="number" class="repayment-amount" placeholder="Amount €" />`;
  document.getElementById("repaymentForm").insertBefore(row, document.getElementById("addRepaymentRow"));
  populateWeekDropdown();
});

document.getElementById("applyRepayments").addEventListener("click", () => {
  repaymentRows = [];
  document.querySelectorAll(".repayment-row").forEach(r => {
    const col = parseInt(r.querySelector("select").value);
    const amount = parseFloat(r.querySelector("input").value);
    if (!isNaN(col) && !isNaN(amount)) {
      repaymentRows.push({ col, amount });
      if (!csvData[134][col]) csvData[134][col] = 0;
      csvData[134][col] = parseFloat(csvData[134][col]) + amount;
      remaining -= amount;
    }
  });
  document.getElementById("remainingBalance").innerHTML = `Remaining Mayweather Balance: <strong>€${remaining.toLocaleString()}</strong>`;
  renderTable();
  updateForecastChart();
});

document.getElementById("toggleTable").addEventListener("click", () => {
  document.getElementById("tablePreview").classList.toggle("collapsed");
});

document.getElementById("exportCSV").addEventListener("click", () => {
  const csv = csvData.map(r => r.join(",")).join("\n");
  const blob = new Blob([csv], { type: 'text/csv' });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "updated_cashflow.csv";
  link.click();
});
