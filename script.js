let weeklyLabels = [], weeklyIncome = [], rollingBalance = [];
let baseBalance = 0, repayments = [];
let chart, incomeRowIx, balanceRowIx, repayRowIx;

document.getElementById('fileInput').addEventListener('change', handleFile);
document.getElementById('addRepaymentBtn').addEventListener('click', addRepaymentRow);
document.getElementById('applyRepaymentsBtn').addEventListener('click', applyRepayments);
document.getElementById('toggleTableBtn').addEventListener('click', toggleTable);

function handleFile(e) {
  const reader = new FileReader();
  reader.onload = evt => {
    const wb = XLSX.read(evt.target.result, { type: 'array' });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    const arr = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    detectRows(arr);
    extractData(arr);
    buildTable(sheet);
    resetUI();
  };
  reader.readAsArrayBuffer(e.target.files[0]);
}

function detectRows(arr) {
  incomeRowIx = balanceRowIx = repayRowIx = -1;
  arr.forEach((row, i) => {
    const s = row.join(' ');
    if (/Weekly income \/ cash position/i.test(s)) incomeRowIx = i;
    if (/Rolling cash balance/i.test(s)) balanceRowIx = i;
    if (/Mayweather Investment Repayment/i.test(s)) repayRowIx = i;
  });
}

function extractData(arr) {
  weeklyLabels = arr[3].slice(1);
  weeklyIncome = arr[incomeRowIx].slice(1).map(v => +v || 0);
  rollingBalance = arr[balanceRowIx].slice(1).map(v => +v || 0);
  baseBalance = rollingBalance[0] || 0;
  renderAll();
}

function renderAll() {
  renderSummary(rollingBalance);
  renderChart(rollingBalance);
  document.getElementById('repaymentContainer').innerHTML = '';
  addRepaymentRow();
}

function addRepaymentRow() {
  const ctr = document.getElementById('repaymentContainer');
  const div = document.createElement('div');
  const sel = document.createElement('select');
  weeklyLabels.forEach((w, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = w;
    sel.append(opt);
  });
  const inp = document.createElement('input');
  inp.type = 'number';
  inp.placeholder = 'Amount €';
  div.append(sel, inp);
  ctr.append(div);
}

function applyRepayments() {
  repayments = [];
  Array.from(document.getElementById('repaymentContainer').children)
    .forEach(div => {
      const w = +div.children[0].value;
      const val = +div.children[1].value || 0;
      if (!isNaN(w) && !isNaN(val)) { repayments.push({ w, val }); }
    });

  let modIncome = [...weeklyIncome];
  repayments.forEach(({ w, val }) => {
    modIncome[w] = (modIncome[w] || 0) - val;
  });

  const newBal = [baseBalance];
  for (let i = 1; i < modIncome.length; i++) {
    newBal[i] = newBal[i - 1] + (modIncome[i] || 0);
  }

  if (newBal.some(n => isNaN(n))) {
    alert("Invalid balance calculation.");
    return;
  }

  updateSpreadsheetRows(repayments);
  renderSummary(newBal);
  renderChart(newBal);
}

function updateSpreadsheetRows(reps) {
  const tbl = document.querySelector('#tableWrapper table');
  if (!tbl) return;
  const rows = tbl.querySelectorAll('tr');
  const repayRow = rows[repayRowIx];
  const incomeRow = rows[incomeRowIx];

  reps.forEach(({ w, val }) => {
    const rc = repayRow?.children[w + 1];
    if (rc) {
      const old = parseFloat(rc.textContent.replace(/,/g, '')) || 0;
      rc.textContent = (old - val).toLocaleString();
      rc.style.backgroundColor = '#ffe6e6';
    }
    const ic = incomeRow?.children[w + 1];
    if (ic) {
      const old = parseFloat(ic.textContent.replace(/,/g, '')) || 0;
      ic.textContent = (old - val).toLocaleString();
      ic.style.backgroundColor = '#e6f7ff';
    }
  });
}

function renderSummary(bal) {
  const sum = repayments.reduce((a, v) => a + v.val, 0);
  const final = bal[bal.length - 1];
  const min = Math.min(...bal);
  const idx = bal.indexOf(min);

  document.getElementById('totalRepaid').textContent = `€${sum.toLocaleString()}`;
  document.getElementById('finalBalance').textContent = `€${final.toLocaleString()}`;
  document.getElementById('lowestWeek').textContent = weeklyLabels[idx];
  document.getElementById('remaining').textContent = `€${(baseBalance - sum).toLocaleString()}`;
}

function renderChart(bal) {
  const ctx = document.getElementById('chartCanvas').getContext('2d');
  if (chart) chart.destroy();
  chart = new Chart(ctx, {
    type: 'line',
    data: { labels: weeklyLabels, datasets: [{ label: 'Rolling Cash Balance', data: bal, borderColor: '#0288d1', fill: true, backgroundColor: 'rgba(2,136,209,0.2)' }] },
    options: { responsive: true, maintainAspectRatio: false, animation: false }
  });
}

function buildTable(sheet) {
  const html = XLSX.utils.sheet_to_html(sheet);
  document.getElementById('tableWrapper').innerHTML = html;
}

function toggleTable() {
  document.getElementById('tableWrapper').classList.toggle('hidden');
}

function resetUI() {
  repayments = [];
  document.getElementById('repaymentContainer').innerHTML = '';
  document.getElementById('totalRepaid').textContent = '€0';
  document.getElementById('finalBalance').textContent = '€0';
  document.getElementById('lowestWeek').textContent = '–';
  document.getElementById('remaining').textContent = '€0';
}
