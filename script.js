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
  baseBalance = rollingBalance[0];
  renderAll();
}

function renderAll() {
  renderSummary(rollingBalance);
  renderChart(rollingBalance);
  createRepaymentUI();
}

function createRepaymentUI() {
  const ctr = document.getElementById('repaymentContainer');
  ctr.innerHTML = '';
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
      repayments.push({ w, val });
    });

  let modIncome = [...weeklyIncome];
  repayments.forEach(({ w, val }) => modIncome[w] -= val);

  const newBal = [baseBalance];
  for (let i = 1; i < modIncome.length; i++)
    newBal[i] = newBal[i-1] + modIncome[i];

  updateSpreadsheetRows(repayments);
  renderSummary(newBal);
  renderChart(newBal);
}

function updateSpreadsheetRows(reps) {
  const tbl = document.querySelector('#tableWrapper table');
  if (!tbl) return;
  const row = tbl.querySelectorAll('tr')[repayRowIx];
  reps.forEach(({ w, val }) => {
    const cell = row.children[w+1];
    cell.textContent = (parseFloat(cell.textContent)||0) - val;
    cell.style.backgroundColor = '#ffe6e6';
  });
}

function renderSummary(bal) {
  const sum = repayments.reduce((a,v)=>a+v.val,0);
  const final = bal[bal.length-1];
  const min = Math.min(...bal);
  document.getElementById('totalRepaid').textContent = `€${sum.toLocaleString()}`;
  document.getElementById('finalBalance').textContent = `€${final.toLocaleString()}`;
  document.getElementById('lowestWeek').textContent = weeklyLabels[bal.indexOf(min)];
  document.getElementById('remaining').textContent = `€${(baseBalance - sum).toLocaleString()}`;
}

function renderChart(bal) {
  const ctx = document.getElementById('chartCanvas').getContext('2d');
  if (chart) chart.destroy();
  chart = new Chart(ctx, {
    type: 'line',
    data: { labels: weeklyLabels, datasets: [{
      label: 'Rolling Cash',
      data: bal,
      borderColor: '#0288d1',
      fill: true,
      backgroundColor: 'rgba(2,136,209,0.2)'
    }]},
    options: { responsive: true, maintainAspectRatio: false }
  });
}

function buildTable(sheet) {
  const html = XLSX.utils.sheet_to_html(sheet);
  const wr = document.getElementById('tableWrapper');
  wr.innerHTML = html;
}

function toggleTable() {
  document.getElementById('tableWrapper').classList.toggle('hidden');
}
