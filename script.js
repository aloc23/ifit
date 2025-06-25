let chart = null;
let weeks = [];
let balances = [];
let remaining = 0;

document.getElementById("fileInput").addEventListener("change", handleFileUpload);
document.getElementById("addRepaymentBtn").addEventListener("click", addRepaymentRow);
document.getElementById("applyRepaymentsBtn").addEventListener("click", applyRepayments);

function handleFileUpload(e) {
  // Fake initial data — replace with XLSX read logic
  weeks = Array.from({ length: 52 }, (_, i) => `Week ${i + 1}`);
  balances = weeks.map((_, i) => 355000 - i * 1000);
  remaining = 355000;

  updateSummary();
  renderChart();
  populateDropdowns();
}

function updateSummary() {
  document.getElementById("remaining").textContent = formatEUR(remaining);
  document.getElementById("finalBalance").textContent = formatEUR(balances[balances.length - 1]);
  const lowest = Math.min(...balances);
  const lowestIndex = balances.indexOf(lowest);
  document.getElementById("lowestWeek").textContent = `${weeks[lowestIndex]}`;
}

function renderChart() {
  const ctx = document.getElementById("chartCanvas").getContext("2d");
  if (chart) chart.destroy();

  chart = new Chart(ctx, {
    type: "line",
    data: {
      labels: weeks,
      datasets: [{
        label: "Cash Balance Forecast",
        data: balances,
        borderColor: "#007bff",
        backgroundColor: "rgba(0, 123, 255, 0.1)",
        fill: true,
        tension: 0.2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false
    }
  });
}

function addRepaymentRow() {
  const container = document.getElementById("repaymentInputs");
  const row = document.createElement("div");

  const select = document.createElement("select");
  weeks.forEach((week, idx) => {
    const opt = document.createElement("option");
    opt.value = idx;
    opt.textContent = week;
    select.appendChild(opt);
  });

  const input = document.createElement("input");
  input.type = "number";
  input.placeholder = "Amount €";

  row.appendChild(select);
  row.appendChild(input);
  container.appendChild(row);
}

function applyRepayments() {
  const rows = document.getElementById("repaymentInputs").querySelectorAll("div");
  let totalRepay = 0;
  const adjustments = new Array(balances.length).fill(0);

  rows.forEach(row => {
    const weekIndex = +row.querySelector("select").value;
    const amount = parseFloat(row.querySelector("input").value) || 0;
    adjustments[weekIndex] += amount;
    totalRepay += amount;
  });

  balances = balances.map((bal, i) => bal - adjustments[i]);
  remaining -= totalRepay;

  document.getElementById("totalRepaid").textContent = formatEUR(totalRepay);
  updateSummary();
  renderChart();
}

function populateDropdowns() {
  const selects = document.querySelectorAll("#repaymentInputs select");
  selects.forEach(select => {
    select.innerHTML = "";
    weeks.forEach((week, idx) => {
      const opt = document.createElement("option");
      opt.value = idx;
      opt.textContent = week;
      select.appendChild(opt);
    });
  });
}

function formatEUR(num) {
  return "€" + num.toLocaleString("en-IE");
}
