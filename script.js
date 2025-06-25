let chartInstance;
let weekOptions = [];
let balanceData = [];

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const weekLabels = json[3].slice(2); // Row 4 = week labels
    weekOptions = weekLabels;

    updateDropdowns();

    const weeklyRow = json.find(row => row[1]?.toLowerCase().includes("weekly income"));
    const rollingRow = json.find(row => row[1]?.toLowerCase().includes("rolling cash balance"));

    if (!weeklyRow || !rollingRow) {
      alert("Spreadsheet missing income/balance rows.");
      return;
    }

    const weeklyCash = weeklyRow.slice(2).map(Number);
    const rollingBalance = rollingRow.slice(2).map(Number);

    balanceData = rollingBalance.map((value, i) => {
      return {
        week: weekLabels[i],
        income: weeklyCash[i] || 0,
        balance: value || 0,
      };
    });

    buildChart();
    updateSummary(0);
  };
  reader.readAsArrayBuffer(file);
}

function updateDropdowns() {
  const container = document.getElementById("repaymentContainer");
  container.innerHTML = '';
  addRepaymentRow();
}

function addRepaymentRow() {
  const container = document.getElementById("repaymentContainer");
  if (!container) return;

  const row = document.createElement("div");
  row.classList.add("repaymentRow");

  const select = document.createElement("select");
  weekOptions.forEach(week => {
    const option = document.createElement("option");
    option.value = week;
    option.textContent = week;
    select.appendChild(option);
  });

  const input = document.createElement("input");
  input.type = "number";
  input.placeholder = "Amount €";

  row.appendChild(select);
  row.appendChild(input);
  container.appendChild(row);
}

function applyRepayments() {
  const rows = document.querySelectorAll(".repaymentRow");
  const repayments = [];

  rows.forEach(row => {
    const week = row.querySelector("select").value;
    const amount = parseFloat(row.querySelector("input").value);
    if (!isNaN(amount)) {
      repayments.push({ week, amount });
    }
  });

  // Copy the original data
  const updatedData = balanceData.map(entry => ({ ...entry }));

  repayments.forEach(({ week, amount }) => {
    const index = updatedData.findIndex(e => e.week === week);
    if (index !== -1) {
      updatedData[index].income -= amount;
    }
  });

  // Recalculate rolling balances like the spreadsheet
  updatedData[0].balance = updatedData[0].income;
  for (let i = 1; i < updatedData.length; i++) {
    updatedData[i].balance = updatedData[i - 1].balance + updatedData[i].income;
  }

  // Rebuild chart with new balance values
  buildChart(updatedData);

  const totalRepaid = repayments.reduce((sum, r) => sum + r.amount, 0);
  const lowest = updatedData.reduce((min, curr) => curr.balance < min.balance ? curr : min, updatedData[0]);

  document.getElementById("lowestWeek").textContent = `Week ${lowest.week}`;
  updateSummary(totalRepaid);
}

function updateSummary(totalRepaid) {
  const remainingStart = 355000;
  document.getElementById("totalRepaid").textContent = `€${totalRepaid.toLocaleString()}`;
  document.getElementById("finalBalance").textContent = `€${(remainingStart - totalRepaid).toLocaleString()}`;
  document.getElementById("remaining").textContent = `€${(remainingStart - totalRepaid).toLocaleString()}`;
}

function buildChart(data = balanceData) {
  if (!data.length || data.some(d => isNaN(d.balance))) return;

  const labels = data.map(e => e.week);
  const values = data.map(e => e.balance);
  const ctx = document.getElementById("chartCanvas").getContext("2d");

  if (chartInstance) chartInstance.destroy();

  document.getElementById("chartCanvas").style.maxHeight = '400px';

  chartInstance = new Chart(ctx, {
    type: "line",
    data: {
      labels,
      datasets: [{
        label: "Rolling Cash Balance",
        data: values,
        borderColor: "blue",
        fill: true,
        tension: 0.3
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false
    }
  });
}

document.getElementById("fileInput").addEventListener("change", handleFile);
document.getElementById("addRowBtn").addEventListener("click", addRepaymentRow);
document.getElementById("applyBtn").addEventListener("click", applyRepayments);
