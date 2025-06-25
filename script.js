let chartInstance;
let weekOptions = [];
let balanceData = [];
let repaymentRows = [];

function handleFile(event) {
  const file = event.target.files[0];
  if (!file) return;

  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    // Get weeks from row 4
    const weekLabels = json[3].slice(2);
    weekOptions = weekLabels;

    // Display dropdowns
    updateDropdowns();

    // Find "Weekly income / cash position" and "Rolling Cash Balance"
    const weeklyRow = json.find(row => row[1] && typeof row[1] === 'string' && row[1].toLowerCase().includes("weekly income"));
    const rollingRow = json.find(row => row[1] && typeof row[1] === 'string' && row[1].toLowerCase().includes("rolling cash balance"));

    if (!weeklyRow || !rollingRow) {
      alert("Spreadsheet rows not detected");
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
  document.querySelectorAll(".repaymentRow").forEach(row => row.remove());

  addRepaymentRow();
}

function addRepaymentRow() {
  const container = document.createElement("div");
  container.classList.add("repaymentRow");

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

  container.appendChild(select);
  container.appendChild(input);
  document.body.insertBefore(container, document.getElementById("summary"));
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

  // Copy original data to preserve
  const updatedData = balanceData.map(entry => ({ ...entry }));

  let totalRepaid = 0;

  repayments.forEach(({ week, amount }) => {
    const index = updatedData.findIndex(e => e.week === week);
    if (index !== -1) {
      updatedData[index].balance -= amount;
      totalRepaid += amount;

      // Cascade the effect
      for (let i = index + 1; i < updatedData.length; i++) {
        updatedData[i].balance -= amount;
      }
    }
  });

  // Update chart
  buildChart(updatedData);

  // Determine lowest week
  const lowest = updatedData.reduce((min, curr) => curr.balance < min.balance ? curr : min, updatedData[0]);

  document.getElementById("lowestWeek").textContent = `Week ${lowest.week}`;
  updateSummary(totalRepaid);
}

function updateSummary(totalRepaid) {
  const remainingStart = balanceData[1] || 355000;
  document.getElementById("totalRepaid").textContent = `€${totalRepaid.toLocaleString()}`;
  document.getElementById("finalBalance").textContent = `€${(remainingStart - totalRepaid).toLocaleString()}`;
  document.getElementById("remaining").textContent = `€${(remainingStart - totalRepaid).toLocaleString()}`;
}

function buildChart(data = balanceData) {
  if (!data.length || data.some(d => isNaN(d.balance))) {
    console.warn("No valid data for chart");
    return;
  }

  const labels = data.map(e => e.week);
  const values = data.map(e => e.balance);

  const ctx = document.getElementById("chartCanvas").getContext("2d");

  if (chartInstance) {
    chartInstance.destroy();
  }

  document.getElementById("chartCanvas").style.maxHeight = '400px';

  chartInstance = new Chart(ctx, {
    type: "line",
    data: {
      labels: labels,
      datasets: [{
        label: "Cash Balance Forecast",
        data: values,
        borderColor: "blue",
        fill: true,
        pointRadius: 3
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
