let data = [];
let headers = [];
let originalBalance = 355000;
let remaining = originalBalance;
let chart;
let fullTableVisible = true;
const repaymentRowLabel = "Mayweather Investment Repayment (Investment 1 and 2)";
const cashPositionRowLabel = "Weekly income / cash position";
const rollingBalanceRowLabel = "Rolling cash balance";

document.getElementById("fileInput").addEventListener("change", handleFile, false);

function handleFile(e) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const workbook = XLSX.read(e.target.result, { type: "binary" });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const raw = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
    headers = raw[0];
    data = raw.slice(1);
    renderTable();
    populateDropdowns();
  };
  reader.readAsBinaryString(e.target.files[0]);
}

function populateDropdowns() {
  const dropdowns = document.querySelectorAll("select.weekSelect");
  dropdowns.forEach(dd => {
    dd.innerHTML = "";
    headers.forEach((header, index) => {
      if (header && typeof header === "string" && header.toLowerCase().includes("week")) {
        const opt = document.createElement("option");
        opt.value = index;
        opt.textContent = header;
        dd.appendChild(opt);
      }
    });
  });
}

function renderTable() {
  const container = document.getElementById("tableContainer");
  container.innerHTML = "";
  const table = document.createElement("table");
  const thead = table.createTHead();
  const row = thead.insertRow();
  headers.forEach(header => {
    const th = document.createElement("th");
    th.innerText = header;
    row.appendChild(th);
  });

  const tbody = table.createTBody();
  data.forEach(rowData => {
    const row = tbody.insertRow();
    headers.forEach((_, i) => {
      const cell = row.insertCell();
      const val = rowData[i] ?? "";
      cell.innerText = val;
      if (typeof val === "number" && val < 0) {
        cell.classList.add("highlight");
      }
    });
  });

  container.appendChild(table);
  updateChart();
  updateSummary();
}

function addRepayment() {
  const div = document.createElement("div");
  div.classList.add("repayment-row");

  const select = document.createElement("select");
  select.classList.add("weekSelect");

  const input = document.createElement("input");
  input.type = "number";

  div.appendChild(select);
  div.appendChild(input);
  document.getElementById("repayments").appendChild(div);

 function populateDropdown(targetSelect = null) {
  const options = headers.map((header, index) => {
    if (typeof header === "string" && header.toLowerCase().includes("week")) {
      return `<option value="${index}">${header}</option>`;
    }
    return '';
  }).join("");

  if (targetSelect) {
    targetSelect.innerHTML = options;
  } else {
    document.querySelectorAll("select.weekSelect").forEach(select => {
      select.innerHTML = options;
    });
  }
}

function applyRepayments() {
  const repaymentRowIdx = findOrCreateRow(repaymentRowLabel);
  data[repaymentRowIdx] = data[repaymentRowIdx] || [];

  document.querySelectorAll(".repayment-row").forEach(r => {
    const col = parseInt(r.querySelector("select").value);
    let val = parseFloat(r.querySelector("input").value);
    if (!isNaN(col) && !isNaN(val)) {
      val = -Math.abs(val);
      data[repaymentRowIdx][col] = (parseFloat(data[repaymentRowIdx][col]) || 0) + val;
      remaining -= Math.abs(val);
    }
  });

  updateCashFlow();
  renderTable();
}

function findOrCreateRow(label) {
  let idx = data.findIndex(r => r[0] === label);
  if (idx === -1) {
    idx = data.length;
    const row = [];
    row[0] = label;
    for (let i = 1; i < headers.length; i++) row[i] = 0;
    data.push(row);
  }
  return idx;
}

function updateCashFlow() {
  const posIdx = findOrCreateRow(cashPositionRowLabel);
  const rollIdx = findOrCreateRow(rollingBalanceRowLabel);
  data[rollIdx] = [rollingBalanceRowLabel];

  for (let i = 1; i < headers.length; i++) {
    const income = parseFloat(data[posIdx][i]) || 0;
    const prev = parseFloat(data[rollIdx][i - 1]) || 0;
    data[rollIdx][i] = prev + income;
  }
}

function updateSummary() {
  document.getElementById("remaining").innerText = `€${remaining.toLocaleString()}`;

  const rollIdx = findOrCreateRow(rollingBalanceRowLabel);
  const values = data[rollIdx].slice(1).map(x => parseFloat(x) || 0);
  const min = Math.min(...values);
  const final = values[values.length - 1];

  document.getElementById("totalRepaid").innerText = `€${(originalBalance - remaining).toLocaleString()}`;
  document.getElementById("finalBalance").innerText = `€${final.toLocaleString()}`;
  document.getElementById("minWeek").innerText = `${headers[values.indexOf(min) + 1]}`;
}

function updateChart() {
  const rollIdx = findOrCreateRow(rollingBalanceRowLabel);
  const values = data[rollIdx].slice(1).map(x => parseFloat(x) || 0);
  const labels = headers.slice(1);

  const ctx = document.getElementById("chartCanvas").getContext("2d");
  if (chart) chart.destroy();
  chart = new Chart(ctx, {
    type: "line",
    data: {
      labels: labels,
      datasets: [{
        label: "Cash Balance Forecast",
        data: values,
        fill: false,
        borderColor: "#0077cc",
        tension: 0.2,
        pointRadius: 4,
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
  document.getElementById("tableContainer").style.display = fullTableVisible ? "block" : "none";
}

function exportFile() {
  const ws = XLSX.utils.aoa_to_sheet([headers, ...data]);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, "CashflowForecast_Updated.xlsx");
}
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
    addRepaymentRow();
    updateChart();
  };
  reader.readAsBinaryString(e.target.files[0]);
}

function addRepaymentRow() {
  const container = document.getElementById("repayments");
  const div = document.createElement("div");
  div.className = "repayment-row";
  const sel = document.createElement("select");
  sel.className = "weekSelect";
  headers.forEach((h, i) => {
    if (typeof h === "string" && h.toLowerCase().includes("week")) {
      const opt = document.createElement("option");
      opt.value = i; opt.textContent = h;
      sel.appendChild(opt);
    }
  });
  const inp = document.createElement("input");
  inp.type = "number"; inp.placeholder = "Amount €";
  div.append(sel, inp);
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
    const prev = +(data[bIdx][c-1] || 0);
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
  document.getElementById("totalRepaid").textContent = `€${((originalBalance - remaining).toLocaleString())}`;
  document.getElementById("finalBalance").textContent = `€${final.toLocaleString()}`;
  document.getElementById("minWeek").textContent = `${headers[vals.indexOf(min) + 1]}`;
}

function renderTable() {
  const fullTbl = document.getElementById("tableContainer");
  fullTbl.innerHTML = "";
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
  fullTbl.appendChild(tbl);
}

function updateChart() {
  const bIdx = findOrCreateLabel(balanceLabel);
  const vals = data[bIdx].slice(1).map(n => +n);
  const labs = headers.slice(1);
  const ctx = document.getElementById("chartCanvas").getContext("2d");
  if (chart) chart.destroy();
  chart = new Chart(ctx, {
    type: "line",
    data: { labels: labs, datasets: [{ label: "Cash Balance Forecast", data: vals, borderColor: "#0077cc", tension: 0.2 }] },
    options: { scales: { y: { beginAtZero: false } } }
  });
}

function toggleFullTable() { fullTableVisible = !fullTableVisible; renderTable(); }
function exportFile() {
  const out = [headers, ...data];
  const ws = XLSX.utils.aoa_to_sheet(out);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, "CashflowUpdated.xlsx");
}
