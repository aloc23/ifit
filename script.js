let workbookData = [];
let chart;
let repayments = [];
let balanceRow = -1;
let incomeRow = -1;
let remainingStart = 0;

document.getElementById('fileInput').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = async (e) => {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    workbookData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    document.getElementById('fileName').textContent = file.name;

    locateKeyRows();
    initRepaymentRow();
    buildChart();
    buildTable();
  };
  reader.readAsArrayBuffer(file);
});

function locateKeyRows() {
  workbookData.forEach((row, idx) => {
    if (row.includes("Rolling Cash Balance")) balanceRow = idx;
    if (row.includes("Weekly income / cash position")) incomeRow = idx;
    if (row[0] && row[0].toString().includes("Bank Balance Overdraft")) {
      const firstVal = row.find(cell => typeof cell === 'number');
      if (firstVal) remainingStart = firstVal;
    }
  });
}

function initRepaymentRow() {
  const container = document.getElementById('repaymentContainer');
  container.innerHTML = '';
  addRepaymentRow();
}

function addRepaymentRow() {
  const row = document.createElement('div');
  const weekSelect = document.createElement('select');
  const amountInput = document.createElement('input');
  amountInput.type = 'number';
  amountInput.placeholder = "Amount €";

  const headers = workbookData[3]; // Week numbers
  const labels = workbookData[2];  // e.g. May 25, etc.

  if (!headers || !labels) return;

  for (let i = 1; i < headers.length; i++) {
    const weekLabel = `Week ${headers[i]} = ${labels[i] || ""}`;
    const opt = new Option(weekLabel, i);
    weekSelect.add(opt);
  }

  row.appendChild(weekSelect);
  row.appendChild(amountInput);
  document.getElementById('repaymentContainer').appendChild(row);
}

function applyRepayments() {
  repayments = [];
  const rows = document.getElementById('repaymentContainer').children;
  for (let row of rows) {
    const weekIdx = parseInt(row.children[0].value);
    const amount = parseFloat(row.children[1].value);
    if (!isNaN(weekIdx) && !isNaN(amount)) {
      repayments.push({ weekIdx, amount });
    }
  }
  buildChart();
  buildTable();
}

function buildChart() {
  if (balanceRow < 0 || incomeRow < 0) return;

  const balanceData = workbookData[balanceRow];
  const incomeData = workbookData[incomeRow];
  const labels = workbookData[2];

  const balances = [];
  const weeks = [];

  let totalRepaid = 0;
  const repaymentMap = {};

  repayments.forEach(r => {
    repaymentMap[r.weekIdx] = (repaymentMap[r.weekIdx] || 0) + r.amount;
    totalRepaid += r.amount;
  });

  let currentBalance = remainingStart;
  for (let i = 1; i < balanceData.length; i++) {
    const income = Number(incomeData[i]) || 0;
    const repayment = repaymentMap[i] || 0;

    currentBalance += income - repayment;
    balances.push(currentBalance);
    weeks.push(`Week ${workbookData[3][i]}`);
  }

  const ctx = document.getElementById("chartCanvas").getContext("2d");
  if (chart) chart.destroy();
  chart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: weeks,
      datasets: [{
        label: "Cash Balance Forecast",
        data: balances,
        borderColor: "blue",
        backgroundColor: "rgba(0, 0, 255, 0.1)",
        fill: true,
        tension: 0.2
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: true },
        tooltip: { mode: 'index' }
      },
      scales: {
        y: {
          beginAtZero: false
        }
      }
    }
  });

  const lowest = Math.min(...balances);
  const lowestIdx = balances.indexOf(lowest);
  document.getElementById("lowestWeek").textContent = `Week ${workbookData[3][lowestIdx + 1]} = ${workbookData[2][lowestIdx + 1]}`;
  document.getElementById("finalBalance").textContent = `€${balances[balances.length - 1].toLocaleString()}`;
  document.getElementById("totalRepaid").textContent = `€${totalRepaid.toLocaleString()}`;
  document.getElementById("remaining").textContent = `€${(remainingStart - totalRepaid).toLocaleString()}`;
}

function buildTable() {
  const table = document.getElementById("dataTable");
  table.innerHTML = "";
  workbookData.forEach(row => {
    const tr = document.createElement("tr");
    row.forEach(cell => {
      const td = document.createElement("td");
      td.textContent = cell;
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
}

document.getElementById("addRepaymentBtn").addEventListener("click", addRepaymentRow);
document.getElementById("applyRepaymentsBtn").addEventListener("click", applyRepayments);
document.getElementById("toggleTableBtn").addEventListener("click", () => {
  const wrap = document.getElementById("tableWrapper");
  wrap.style.display = wrap.style.display === "none" ? "block" : "none";
});
