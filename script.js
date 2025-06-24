let workbook, data = [], remaining = 355000, chart;
let weekStartCol = 5;
let repaymentRowIdx, cashPositionRowIdx, rollingBalanceRowIdx;

function loadXLSX(file) {
  const reader = new FileReader();
  reader.onload = function(e) {
    const wb = XLSX.read(e.target.result, { type: "binary" });
    workbook = wb;
    const sheet = wb.Sheets[wb.SheetNames[0]];
    data = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });

    detectSpecialRows();
    detectWeekStartCol();
    renderTable();
    addRepaymentRow(); // add initial row
    drawChart();
  };
  reader.readAsBinaryString(file);
}

function detectSpecialRows() {
  data.forEach((row, i) => {
    const label = String(row[0] || "").trim();
    if (label.includes("Repayment (Investment 1 and 2)")) repaymentRowIdx = i;
    if (label.includes("Weekly income / cash position")) cashPositionRowIdx = i;
    if (label.includes("Rolling cash balance")) rollingBalanceRowIdx = i;
  });
}

function detectWeekStartCol() {
  const headers = data[3] || [];
  for (let i = 0; i < headers.length; i++) {
    if (/Week/.test(headers[i])) {
      weekStartCol = i;
      break;
    }
  }
}

function addRepaymentRow() {
  const section = document.getElementById("repaymentSection");
  const headerRow = data[3] || [];
  const yearRow = data[2] || [];

  const container = document.createElement("div");
  container.className = "repayment-row";

  const select = document.createElement("select");
  for (let i = weekStartCol; i < headerRow.length; i++) {
    const week = headerRow[i];
    const year = yearRow[i];
    if (/Week/.test(week) && /\d{4}/.test(year)) {
      const option = document.createElement("option");
      option.value = i;
      option.text = `Week ${week} (${year})`;
      select.appendChild(option);
    }
  }

  const input = document.createElement("input");
  input.type = "number";
  input.placeholder = "Amount €";

  container.appendChild(select);
  container.appendChild(input);
  section.appendChild(container);
}

function applyRepayments() {
  remaining = 355000;

  // Reset repayment values before reapplying
  for (let i = weekStartCol; i < data[0].length; i++) {
    data[repaymentRowIdx][i] = 0;
  }

  document.querySelectorAll(".repayment-row").forEach(r => {
    const col = parseInt(r.querySelector("select").value);
    let val = parseFloat(r.querySelector("input").value);
    if (!isNaN(col) && !isNaN(val)) {
      val = -Math.abs(val);
      data[repaymentRowIdx][col] = (parseFloat(data[repaymentRowIdx][col]) || 0) + val;
      remaining -= val;
    }
  });

  // Recalculate Weekly income/cash position
  for (let i = weekStartCol; i < data[0].length; i++) {
    let total = 0;
    for (let r = 0; r < cashPositionRowIdx; r++) {
      const v = parseFloat(data[r][i]);
      if (!isNaN(v)) total += v;
    }
    data[cashPositionRowIdx][i] = total.toFixed(2);
  }

  // Recalculate rolling cash balance
  for (let i = weekStartCol; i < data[0].length; i++) {
    const prev = i === weekStartCol ? 0 : parseFloat(data[rollingBalanceRowIdx][i - 1]) || 0;
    const cash = parseFloat(data[cashPositionRowIdx][i]) || 0;
    data[rollingBalanceRowIdx][i] = (prev + cash).toFixed(2);
  }

  document.getElementById("remainingBalance").innerText = "Remaining: €" + remaining.toLocaleString();
  renderTable();
  drawChart();
}

function renderTable() {
  const container = document.getElementById("tablePreview");
  container.innerHTML = "<table></table>";
  const table = container.querySelector("table");

  data.forEach((row, i) => {
    const tr = document.createElement("tr");
    row.forEach((cell, j) => {
      const td = document.createElement(i === 0 ? "th" : "td");
      td.textContent = cell;
      if (i === repaymentRowIdx && cell !== 0) td.classList.add("highlight");
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
}

function drawChart() {
  const labels = data[3]?.slice(weekStartCol) || [];
  const values = data[rollingBalanceRowIdx]?.slice(weekStartCol).map(v => parseFloat(v) || 0);

  if (!chart) {
    chart = new Chart(document.getElementById("forecastChart"), {
      type: "line",
      data: {
        labels,
        datasets: [{
          label: "Cash Balance Forecast",
          data: values,
          borderColor: "#0077cc",
          backgroundColor: "rgba(0,119,204,0.1)",
          tension: 0.3
        }]
      },
      options: {
        responsive: true,
        plugins: {
          legend: { display: true }
        },
        scales: {
          y: { beginAtZero: true }
        }
      }
    });
  } else {
    chart.data.labels = labels;
    chart.data.datasets[0].data = values;
    chart.update();
  }
}

// Bind buttons
document.getElementById("addRepayment").addEventListener("click", addRepaymentRow);
document.getElementById("applyRepayments").addEventListener("click", applyRepayments);
document.getElementById("xlsxUpload").addEventListener("change", e => loadXLSX(e.target.files[0]));
document.getElementById("toggleTable").addEventListener("click", () => {
  document.getElementById("tablePreview").classList.toggle("collapsed");
});
document.getElementById("exportXLSX").addEventListener("click", () => {
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, "updated_cashflow.xlsx");
});
