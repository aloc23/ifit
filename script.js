let data = [], headers = [];
let originalBalance = 355000, remaining = originalBalance, chart;
let fullTableVisible = true;

const repaymentLabel = "Mayweather Investment Repayment (Investment 1 and 2)";
const incomeLabel = "Weekly income / cash position";
const balanceLabel = "Rolling cash balance";

document.getElementById("fileInput").addEventListener("change", handleFile);

function handleFile(e) {
  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(evt.target.result, { type: "binary" });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(ws, { header: 1 });
    headers = raw.shift();
    data = raw;
    renderTable();
    updateChart();
  };
  reader.readAsBinaryString(e.target.files[0]);
}

function addRepaymentRow() {
  if (!headers.length) {
    alert("Please upload a valid Excel file before adding a repayment.");
    return;
  }

  const container = document.getElementById("repayments");
  const div = document.createElement("div");
  div.className = "repayment-row";

  const sel = document.createElement("select");
  sel.className = "weekSelect";

  headers.forEach((h, i) => {
    if (typeof h === "string" && h.toLowerCase().includes("week")) {
      const opt = document.createElement("option");
      opt.value = i;
      opt.textContent = h;
      sel.appendChild(opt);
    }
  });

  const inp = document.createElement("input");
  inp.type = "number";
  inp.placeholder = "Amount €";

  div.appendChild(sel);
  div.appendChild(inp);
  container.appendChild(div);
}

  const sel = document.createElement("select");
  sel.className = "weekSelect";
  headers.forEach((h, i) => {
    if (typeof h === "string" && h.toLowerCase().includes("week")) {
      const opt = document.createElement("option");
      opt.value = i;
      opt.textContent = h;
      sel.appendChild(opt);
    }
  });

  const inp = document.createElement("input");
  inp.type = "number";
  inp.placeholder = "Amount €";

  div.appendChild(sel);
  div.appendChild(inp);
  container.appendChild(div);
}

function applyRepayments() {
  remaining = originalBalance;
  const rIdx = findOrCreateLabel(repaymentLabel);
  data[rIdx] = data[rIdx] || [];

  document.querySelectorAll(".repayment-row").forEach(r => {
    const col = +r.querySelector("select").value;
    let val = +r.querySelector("input").value;
    if (!isNaN(val)) {
      val = -Math.abs(val);
      data[rIdx][col] = (+(data[rIdx][col] || 0)) + val;
      remaining -= Math.abs(val);
    }
  });

  updateCashflow();
  renderTable();
}

function findOrCreateLabel(label) {
  let idx = data.findIndex(r => r[0] === label);
  if (idx < 0) {
    idx = data.push([label]) - 1;
  }
  return idx;
}

function updateCashflow() {
  const iIdx = findOrCreateLabel(incomeLabel);
  const bIdx = findOrCreateLabel(balanceLabel);
  data[bIdx] = [balanceLabel];

  for (let c = 1; c < headers.length; c++) {
    let sum = 0;
    for (let r = 0; r < iIdx; r++) {
      sum += +(data[r][c] || 0);
    }
    data[iIdx][c] = sum;
    const prev = +(data[bIdx][c - 1] || 0);
    data[bIdx][c] = prev + sum;
  }

  updateSummary();
  updateChart();
}

function updateSummary() {
  document.getElementById("remaining").textContent = `€${remaining.toLocaleString()}`;
  const bIdx = findOrCreateLabel(balanceLabel);
  const vals = data[bIdx].slice(1).map(n => +n);
  const min = Math.min(...vals);
  const final = vals[vals.length - 1];
  document.getElementById("totalRepaid").textContent = `€${(originalBalance - remaining).toLocaleString()}`;
  document.getElementById("finalBalance").textContent = `€${final.toLocaleString()}`;
  document.getElementById("minWeek").textContent = `${headers[vals.indexOf(min) + 1]}`;
}

function renderTable() {
  const container = document.getElementById("tableContainer");
  container.innerHTML = "";
  if (!fullTableVisible) return;
  const tbl = document.createElement("table");
  const hdr = tbl.createTHead().insertRow();
  headers.forEach(h => hdr.appendChild(Object.assign(document.createElement("th"), { textContent: h })));
  const body = tbl.createTBody();
  data.forEach(r => {
    const row = body.insertRow();
    headers.forEach((_, i) => {
      const td = row.insertCell();
      const val = r[i] ?? "";
      td.textContent = val;
      if (typeof val === "number" && val < 0) td.classList.add("highlight");
    });
  });
  container.appendChild(tbl);
}

function updateChart() {
  const bIdx = findOrCreateLabel(balanceLabel);
  const vals = data[bIdx].slice(1).map(n => +n);
  const labs = headers.slice(1);
  const ctx = document.getElementById("chartCanvas").getContext("2d");
  if (chart) chart.destroy();
  chart = new Chart(ctx, {
    type: "line",
    data: {
      labels: labs,
      datasets: [{
        label: "Cash Balance Forecast",
        data: vals,
        borderColor: "#0077cc",
        fill: false,
        tension: 0.2,
        pointRadius: 3,
        pointBackgroundColor: "#0077cc"
      }]
    },
    options: {
      responsive: true,
      scales: {
        y: {
          beginAtZero: false
        }
      }
    }
  });
}

function toggleFullTable() {
  fullTableVisible = !fullTableVisible;
  renderTable();
}

function exportFile() {
  const out = [headers, ...data];
  const ws = XLSX.utils.aoa_to_sheet(out);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, "CashflowUpdated.xlsx");
}
