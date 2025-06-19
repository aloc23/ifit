
// JavaScript logic for the full interactive cashflow loan simulator

function updateDashboardSummary() {
  const table = document.querySelector("table");
  if (!table) return;

  const rows = Array.from(table.querySelectorAll("tr")).slice(1);
  const dates = header.slice(1);
  let inflowTotal = 0;
  let outflowTotal = 0;

  rows.forEach(row => {
    const cells = row.querySelectorAll("td");
    cells.forEach((cell, i) => {
      if (i === 0) return;
      const val = parseFloat(cell.textContent) || 0;
      if (val >= 0) inflowTotal += val;
      else outflowTotal += Math.abs(val);
    });
  });

  const net = inflowTotal - outflowTotal;
  document.getElementById("totalInflows").textContent = "€" + inflowTotal.toFixed(2);
  document.getElementById("totalOutflows").textContent = "€" + outflowTotal.toFixed(2);
  document.getElementById("netCashflow").textContent = "€" + net.toFixed(2);

  const repayment = parseFloat(document.getElementById("repaymentAmount").value || "0");
  const remainingLoan = 355000;
  const weeksToPayoff = repayment > 0 ? Math.ceil(remainingLoan / repayment) : "--";
  const monthsEstimate = repayment > 0 ? Math.ceil(weeksToPayoff / 4) : "--";
  document.getElementById("loanDuration").textContent = monthsEstimate !== "--" ? monthsEstimate + " mo" : "--";
}

function drawForecastChart() {
  const table = document.querySelector("table");
  if (!table) return;

  const rows = Array.from(table.querySelectorAll("tr")).slice(1);
  const dates = header.slice(1);
  let netByWeek = new Array(dates.length).fill(0);

  rows.forEach(row => {
    const cells = row.querySelectorAll("td");
    cells.forEach((cell, i) => {
      if (i === 0) return;
      const val = parseFloat(cell.textContent) || 0;
      netByWeek[i - 1] += val;
    });
  });

  const forecastWeeks = 8;
  const recentTrend = netByWeek.slice(-4).reduce((a, b) => a + b, 0) / 4;
  const futureDates = [...dates];
  const forecastValues = [...netByWeek];

  for (let i = 1; i <= forecastWeeks; i++) {
    forecastValues.push((forecastValues[forecastValues.length - 1] || 0) + recentTrend);
    futureDates.push("Future +" + i);
  }

  const ctx = document.getElementById("forecastChart").getContext("2d");
  if (window.forecastChart) window.forecastChart.destroy();
  window.forecastChart = new Chart(ctx, {
    type: "line",
    data: {
      labels: futureDates,
      datasets: [
        {
          label: "Projected Net Cashflow (Linear)",
          data: forecastValues,
          borderColor: "#6f42c1",
          borderDash: [],
          fill: false
        },
        {
          label: "Best Case",
          data: forecastValues.map(v => v * 1.1),
          borderColor: "#28a745",
          borderDash: [4, 4],
          fill: false
        },
        {
          label: "Worst Case",
          data: forecastValues.map(v => v * 0.9),
          borderColor: "#dc3545",
          borderDash: [4, 4],
          fill: false
        }
      ]
    },
    options: {
      responsive: true
    }
  });
}

const originalDrawCharts = drawCharts;
drawCharts = function() {
  originalDrawCharts();
  updateDashboardSummary();
  drawForecastChart();
};

let originalCSV = "";
let parsedData = [];
let header = [];

window.onload = () => {
  fetch("cashflow.csv")
    .then((res) => res.text())
    .then((csv) => {
      originalCSV = csv;
      renderTable(csv);
    });

  document.getElementById("csvUpload").addEventListener("change", (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();
    reader.onload = (evt) => {
      originalCSV = evt.target.result;
      renderTable(originalCSV);
    };
    reader.readAsText(file);
  });
};

function renderTable(csv) {
  const lines = csv.trim().split("\n");
  parsedData = lines.map((line) => line.split(","));
  header = parsedData[0];
  const container = document.getElementById("table-container");
  container.innerHTML = "";
  const table = document.createElement("table");
  parsedData.forEach((row, i) => {
    const tr = document.createElement("tr");
    row.forEach((cell, j) => {
      const td = document.createElement(i === 0 ? "th" : "td");
      td.textContent = cell.replace(/"/g, "");
      if (i > 0 && j > 0) td.contentEditable = true;
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
  container.appendChild(table);
  drawCharts();
}

function applyRepayment() {
  const repayment = parseFloat(document.getElementById("repaymentAmount").value);
  const selectedWeeks = Array.from(document.getElementById("weeksSelect").selectedOptions).map(o => parseInt(o.value));
  const selectedMonths = Array.from(document.getElementById("monthsSelect").selectedOptions).map(o => o.value);
  const cashThreshold = 0;

  if (!repayment || selectedWeeks.length === 0 || selectedMonths.length === 0) {
    return alert("Please enter a repayment, select week(s), and select month(s).");
  }

  const table = document.querySelector("table");
  const rows = Array.from(table.querySelectorAll("tr")).slice(1);

  rows.forEach((row) => {
    const cells = row.querySelectorAll("td");
    for (let j = 1; j < cells.length; j++) {
      const dateStr = header[j];
      const date = new Date(dateStr);
      const monthName = date.toLocaleString("default", { month: "long" });
      const weekNum = j % 4;

      if (selectedMonths.includes(monthName) && selectedWeeks.includes(weekNum)) {
        const currentValue = parseFloat(cells[j].textContent) || 0;
        if (currentValue - repayment >= cashThreshold) {
          cells[j].textContent = (currentValue - repayment).toFixed(2);
          cells[j].style.backgroundColor = "#ffeeba";
          cells[j].title = "Repayment applied";
        }
      }
    }
  });

  updateParsedDataFromTable();
  drawCharts();
}

function updateParsedDataFromTable() {
  const table = document.querySelector("table");
  if (!table) return;

  const rows = Array.from(table.querySelectorAll("tr"));
  parsedData = rows.map(row => {
    return Array.from(row.querySelectorAll("th, td")).map(cell => cell.textContent);
  });
}

function saveTable() {
  updateParsedDataFromTable();
  localStorage.setItem("cashflowData", JSON.stringify(parsedData));
  alert("Saved to browser storage!");
}

function resetTable() {
  if (confirm("Reset table to original?")) renderTable(originalCSV);
}

function exportCSV() {
  updateParsedDataFromTable();
  let csv = parsedData.map(row => row.map(cell => '"' + cell + '"').join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const a = document.createElement("a");
  a.href = URL.createObjectURL(blob);
  a.download = "updated_cashflow.csv";
  a.click();
}

function drawCharts() {
  const table = document.querySelector("table");
  if (!table) return;

  const rows = Array.from(table.querySelectorAll("tr")).slice(1);
  const dates = header.slice(1);
  const inflows = new Array(dates.length).fill(0);
  const outflows = new Array(dates.length).fill(0);

  rows.forEach(row => {
    const cells = row.querySelectorAll("td");
    cells.forEach((cell, i) => {
      if (i === 0) return;
      const val = parseFloat(cell.textContent) || 0;
      if (val >= 0) inflows[i - 1] += val;
      else outflows[i - 1] += Math.abs(val);
    });
  });

  const ctx1 = document.getElementById("balanceChart").getContext("2d");
  const ctx2 = document.getElementById("flowChart").getContext("2d");
  if (window.balanceChart) window.balanceChart.destroy();
  if (window.flowChart) window.flowChart.destroy();

  window.balanceChart = new Chart(ctx1, {
    type: "line",
    data: {
      labels: dates,
      datasets: [{
        label: "Net Cashflow (€)",
        data: inflows.map((val, i) => val - outflows[i]),
        borderColor: "#007bff",
        fill: false
      }]
    },
    options: {
      responsive: true
    }
  });

  window.flowChart = new Chart(ctx2, {
    type: "bar",
    data: {
      labels: dates,
      datasets: [
        {
          label: "Inflows (€)",
          data: inflows,
          backgroundColor: "#28a745"
        },
        {
          label: "Outflows (€)",
          data: outflows,
          backgroundColor: "#dc3545"
        }
      ]
    },
    options: {
      responsive: true
    }
  });
}
