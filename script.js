let workbook, data = [], repaymentRows = [], remaining = 355000, chart;

function loadXLSX(file) {
  const reader = new FileReader();
  reader.onload = function(e) {
    const wb = XLSX.read(e.target.result, { type: "binary" });
    workbook = wb;
    const sheet = wb.Sheets[wb.SheetNames[0]];
    data = XLSX.utils.sheet_to_json(sheet, { header: 1, blankrows: false });
    renderTable();
    populateDropdowns();
    drawChart();
  };
  reader.readAsBinaryString(file);
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
      if (repaymentRows.some(r => r.col === j && i === 134)) td.classList.add("highlight");
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
}

function populateDropdowns() {
  const section = document.getElementById("repaymentSection");
  const headerRow = data[3] || [];
  const yearRow = data[2] || [];
  const usedCols = [...section.querySelectorAll("select")].map(sel => parseInt(sel.value));

  for (let i = 5; i < headerRow.length; i++) {
    if (/Week/.test(headerRow[i]) && /\d{4}/.test(yearRow[i]) && !usedCols.includes(i)) {
      const container = document.createElement("div");
      container.className = "repayment-row";
      container.innerHTML = `
        <select data-col="${i}">
          <option value="${i}">Week ${headerRow[i]} (${yearRow[i]})</option>
        </select>
        <input type="number" placeholder="Amount €" />
      `;
      section.appendChild(container);
      break;
    }
  }
}

function updateRepayments() {
  remaining = 355000;
  repaymentRows = [];

  document.querySelectorAll(".repayment-row").forEach(r => {
    const col = parseInt(r.querySelector("select").value);
    const val = parseFloat(r.querySelector("input").value);
    if (!isNaN(col) && !isNaN(val)) {
      repaymentRows.push({ col, val });
      data[134] = data[134] || [];
      data[134][col] = (parseFloat(data[134][col]) || 0) + val;
      remaining -= val;
    }
  });

  data[135] = data[135] || [];
  for (let i = 5; i < data[3].length; i++) {
    const prev = parseFloat(data[135][i - 1] || 0);
    const inflow = parseFloat(data[134][i] || 0);
    data[135][i] = (prev + inflow).toFixed(2);
  }

  document.getElementById("remainingBalance").innerText = "Remaining: €" + remaining.toLocaleString();
  renderTable();
  drawChart();
}

function drawChart() {
  const labels = data[3]?.slice(5) || [];
  const values = data[135]?.slice(5).map(v => parseFloat(v) || 0);
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

document.getElementById("addRepayment").addEventListener("click", populateDropdowns);
document.getElementById("applyRepayments").addEventListener("click", updateRepayments);
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
