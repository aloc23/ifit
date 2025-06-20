let csvData = [], repaymentRows = [], chart, remaining = 355000;
document.getElementById('fileInput').addEventListener('change', e => {
  const reader = new FileReader();
  reader.onload = () => {
    csvData = reader.result.split(/\r?\n/).map(row => row.split(","));
    populateWeekDropdowns();
    renderTable();
    updateChart();
  };
  reader.readAsText(e.target.files[0]);
});
function populateWeekDropdowns() {
  document.querySelectorAll(".week-dropdown").forEach(drop => {
    drop.innerHTML = "";
    csvData[3].forEach((label, i) => {
      if (label.includes("Week") && csvData[2][i].match(/\d{4}/)) {
        drop.innerHTML += `<option value="\${i}">\${label}</option>`;
      }
    });
  });
}
document.getElementById("addRepaymentRow").addEventListener("click", () => {
  const row = document.createElement("div");
  row.className = "repayment-row";
  row.innerHTML = '<select class="week-dropdown"></select><input type="number" placeholder="Amount €">';
  document.getElementById("repaymentForm").appendChild(row);
  populateWeekDropdowns();
});
document.getElementById("applyRepayments").addEventListener("click", () => {
  repaymentRows = [];
  document.querySelectorAll(".repayment-row").forEach(r => {
    const col = parseInt(r.querySelector("select").value);
    const amt = parseFloat(r.querySelector("input").value);
    if (!isNaN(col) && !isNaN(amt)) {
      repaymentRows.push({col, amt});
      csvData[134][col] = (parseFloat(csvData[134][col]) || 0) + amt;
      remaining -= amt;
    }
  });
  document.getElementById("remainingBalance").innerHTML = `Remaining Mayweather Balance: <strong>€\${remaining.toLocaleString()}</strong>`;
  renderTable(); updateChart();
});
function renderTable() {
  const table = document.createElement("table");
  csvData.forEach((r, i) => {
    const tr = document.createElement("tr");
    r.forEach((c, j) => {
      const cell = document.createElement(i === 0 ? "th" : "td");
      cell.textContent = c;
      if (i === 134 && repaymentRows.some(row => row.col === j)) cell.classList.add("highlight");
      tr.appendChild(cell);
    });
    table.appendChild(tr);
  });
  document.getElementById("tablePreview").innerHTML = "";
  document.getElementById("tablePreview").appendChild(table);
}
function updateChart() {
  const labels = csvData[3]?.slice(5) || [];
  const values = csvData[135]?.slice(5).map(v => parseFloat(v) || 0);
  if (!chart) {
    chart = new Chart(document.getElementById("forecastChart"), {
      type: "line",
      data: { labels, datasets: [{ label: "Rolling Balance", data: values, borderColor: "green", fill: false }] }
    });
  } else {
    chart.data.labels = labels;
    chart.data.datasets[0].data = values;
    chart.update();
  }
}
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