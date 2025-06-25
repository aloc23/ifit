let weeklyLabels = [];
let weeklyIncome = [];
let rollingBalance = [];
let baseBalance = 0;
let chart;
let tableData = [];
let repayments = [];

document.getElementById("fileInput").addEventListener("change", handleFile);

function handleFile(e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (event) {
    const data = new Uint8Array(event.target.result);
    const workbook = XLSX.read(data, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const range = XLSX.utils.decode_range(sheet["!ref"]);

    weeklyLabels = [];
    weeklyIncome = [];
    rollingBalance = [];

    for (let c = 3; c <= range.e.c; c++) {
      const weekCell = sheet[XLSX.utils.encode_cell({ r: 3, c })];
      const incomeCell = sheet[XLSX.utils.encode_cell({ r: 270, c })];
      const balanceCell = sheet[XLSX.utils.encode_cell({ r: 271, c })];

      if (weekCell && weekCell.v) {
        weeklyLabels.push(weekCell.v);
        weeklyIncome.push(incomeCell ? Number(incomeCell.v) || 0 : 0);
        rollingBalance.push(balanceCell ? Number(balanceCell.v) || 0 : 0);
      }
    }

    baseBalance = rollingBalance[0] || 0;
    updateSummary(rollingBalance);
    drawChart(rollingBalance);
    populateTable(sheet, range);
    createRepaymentUI();
  };

  reader.readAsArrayBuffer(file);
}

function createRepaymentUI() {
  const container = document.getElementById("repaymentContainer");
  container.innerHTML = "";
  addRepaymentRow();
}

document.getElementById("addRow").addEventListener("click", addRepaymentRow);

function addRepaymentRow() {
  const container = document.getElementById("repaymentContainer");
  const div = document.createElement("div");

  const select = document.createElement("select");
  weeklyLabels.forEach((label, i) => {
    const opt = document.createElement("option");
    opt.value = i;
    opt.text = label;
    select.appendChild(opt);
  });

  const input = document.createElement("input");
  input.type = "number";
  input.placeholder = "Amount €";

  div.appendChild(select);
  div.appendChild(input);
  container.appendChild(div);
}

document.getElementById("applyRepayments").addEventListener("click", () => {
  repayments = [];
  const rows = document.getElementById("repaymentContainer").children;

  for (let row of rows) {
    const weekIndex = row.children[0].value;
    const amount = parseFloat(row.children[1].value) || 0;
    repayments.push({ weekIndex: parseInt(weekIndex), amount });
  }

  const adjusted = [...weeklyIncome];
  repayments.forEach(({ weekIndex, amount }) => {
    if (adjusted[weekIndex] !== undefined) {
      adjusted[weekIndex] -= amount;
    }
  });

  const recomputedBalance = [baseBalance];
  for (let i = 1; i < adjusted.length; i++) {
    recomputedBalance[i] = recomputedBalance[i - 1] + adjusted[i];
  }

  updateSummary(recomputedBalance);
  drawChart(recomputedBalance);
});

function updateSummary(balances) {
  const totalRepay = repayments.reduce((sum, r) => sum + r.amount, 0);
  const finalBal = balances[balances.length - 1];
  const minVal = Math.min(...balances);
  const minIndex = balances.indexOf(minVal);

  document.getElementById("totalRepaid").textContent = `€${totalRepay.toLocaleString()}`;
  document.getElementById("finalBalance").textContent = `€${finalBal.toLocaleString()}`;
  document.getElementById("lowestWeek").textContent = weeklyLabels[minIndex] || "–";
  document.getElementById("remaining").textContent = `€${(baseBalance - totalRepay).toLocaleString()}`;
}

function drawChart(balances) {
  const ctx = document.getElementById("chartCanvas").getContext("2d");
  if (chart) chart.destroy();
  chart = new Chart(ctx, {
    type: "line",
    data: {
      labels: weeklyLabels,
      datasets: [{
        label: "Cash Balance Forecast",
        data: balances,
        borderColor: "#007bff",
        backgroundColor: "rgba(0, 123, 255, 0.1)",
        pointRadius: 4,
        fill: true,
      }],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { position: "top" },
      },
      scales: {
        y: { beginAtZero: false },
      },
    },
  });
}

function populateTable(sheet, range) {
  const table = document.getElementById("tableContainer");
  const html = XLSX.utils.sheet_to_html(sheet);
  table.innerHTML = html;
}

document.getElementById("toggleTable").addEventListener("click", () => {
  const table = document.getElementById("tableContainer");
  table.style.display = table.style.display === "none" ? "block" : "none";
});
}
