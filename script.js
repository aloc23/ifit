// --- Cashflow Forecast Tool: Ensure Weekly Income/Balance matches Spreadsheet precisely ---

let rawData = [];
let chart;
let weekOptions = [];
let chartType = 'line';
let showRepayments = true;
let lowestWeekCache = { value: null, index: null, label: null };

const LABEL_COL = 1; // Column B (index 1) for all label lookups

// DOM Elements
const fileInput = document.getElementById('fileInput');
const addRepaymentBtn = document.getElementById('addRepayment');
const clearRepaymentsBtn = document.getElementById('clearRepayments');
const savePlanBtn = document.getElementById('savePlan');
const loadPlanBtn = document.getElementById('loadPlan');
const exportExcelBtn = document.getElementById('exportExcel');
const exportPDFBtn = document.getElementById('exportPDF');
const exportPNGBtn = document.getElementById('exportPNG');
const repaymentInputs = document.getElementById('repaymentInputs');
const chartTypeSelect = document.getElementById('chartType');
const toggleRepaymentsChk = document.getElementById('toggleRepayments');
const resetZoomBtn = document.getElementById('resetZoom');

// Attach event listeners
fileInput.addEventListener('change', handleFile);
addRepaymentBtn.addEventListener('click', addRepaymentRow);
clearRepaymentsBtn.addEventListener('click', clearRepayments);
savePlanBtn.addEventListener('click', savePlan);
loadPlanBtn.addEventListener('click', loadPlan);
exportExcelBtn.addEventListener('click', exportToExcel);
exportPDFBtn.addEventListener('click', exportToPDF);
exportPNGBtn.addEventListener('click', exportChartPNG);
chartTypeSelect.addEventListener('change', function() {
  chartType = this.value;
  recalculateAndRender();
});
toggleRepaymentsChk.addEventListener('change', function() {
  showRepayments = this.checked;
  recalculateAndRender();
});
resetZoomBtn.addEventListener('click', function() {
  if (chart && chart.resetZoom) chart.resetZoom();
});

// --- Helper: Find the row index by label (case-insensitive, col B) ---
function findRowIndex(label) {
  label = label.trim().toLowerCase();
  let idx = rawData.findIndex(row =>
    row[LABEL_COL] && row[LABEL_COL].toString().trim().toLowerCase() === label
  );
  if (idx !== -1) return idx;
  idx = rawData.findIndex(row =>
    row[LABEL_COL] && row[LABEL_COL].toString().trim().toLowerCase().includes(label)
  );
  return idx;
}

function findRepaymentRowIndex() {
  return findRowIndex("Mayweather Investment Repayment (Investment 1 and 2)");
}

// --- Dynamically detect week columns using regex in header row ---
function extractWeekOptions(data) {
  const weeksRow = data[3] || [];
  weekOptions = [];
  // Find first column whose header matches "Week N = ..."
  let firstWeekCol = weeksRow.findIndex(cell => typeof cell === 'string' && /week\s*\d+\s*=/i.test(cell));
  if (firstWeekCol === -1) firstWeekCol = 5; // fallback
  for (let i = firstWeekCol; i < weeksRow.length; i++) {
    const label = typeof weeksRow[i] === 'string' ? weeksRow[i].trim() : '';
    if (label && /week\s*\d+\s*=/i.test(label)) {
      weekOptions.push({ index: i, label: label });
    }
  }
}

// --- Use the exact row range as in the spreadsheet for sums ---
// You may need to update startRow/endRow for your real sheet
const startRow = 5, endRow = 22; // <-- <--- ADJUST THIS TO MATCH YOUR SHEET! (use e.g. 5, 270 or as per your SUM formula)

// --- Weekly income matches SUM in spreadsheet for each week ---
function computeWeeklyIncomes() {
  return weekOptions.map(w => {
    let sum = 0;
    for (let r = startRow; r <= endRow; r++) {
      const val = parseFloat(rawData[r]?.[w.index] || 0);
      if (!isNaN(val)) sum += val;
    }
    return sum;
  });
}

// --- Repayment chart data (absolute value per week) ---
function getRepaymentsArr() {
  const repayRow = findRepaymentRowIndex();
  if (repayRow === -1) return weekOptions.map(() => 0);
  return weekOptions.map(w => {
    const val = parseFloat(rawData[repayRow][w.index] || 0);
    return Math.abs(val) || 0;
  });
}

// --- Write repayment (negative) into the correct cell in rawData ---
function setRepaymentForWeek(weekIdx, amount) {
  const repayRow = findRepaymentRowIndex();
  if (repayRow !== -1) {
    rawData[repayRow][weekIdx] = amount > 0 ? -Math.abs(amount) : amount;
  }
}

// --- UI Repayment: update spreadsheet and sum ---
function getRepaymentData() {
  let totalRepayment = 0;
  document.querySelectorAll('.repayment-row').forEach(row => {
    const weekIdx = parseInt(row.children[0].value);
    let amount = parseFloat(row.children[1].value) || 0;
    if (!isNaN(amount)) {
      setRepaymentForWeek(weekIdx, amount);
      totalRepayment += Math.abs(amount);
    }
  });
  const repaymentsArr = getRepaymentsArr();
  return { repaymentsArr, totalRepayment };
}

// --- Rolling cash: base + sum(week0) + sum(week1) + ... ---
function computeRollingCashArr(weeklyIncomes, baseValue) {
  let arr = [];
  let prev = baseValue;
  for (let i = 0; i < weeklyIncomes.length; i++) {
    const cur = prev + weeklyIncomes[i];
    arr.push(cur);
    prev = cur;
  }
  return arr;
}

// --- MAIN RECALC/RENDER ---
function recalculateAndRender() {
  if (weekOptions.length === 0 || rawData.length === 0) return;
  const { repaymentsArr, totalRepayment } = getRepaymentData();
  // Set this to your real opening balance (matches Excel)
  const baseValue = parseFloat(rawData[findRowIndex('Rolling cash balance')]?.[weekOptions[0].index - 1]) || 355000;
  const incomeArr = computeWeeklyIncomes();
  const rollingBalanceArr = computeRollingCashArr(incomeArr, baseValue);

  // Find lowest week
  let lowestWeek = { value: Infinity, index: null, label: '' };
  for (let i = 0; i < rollingBalanceArr.length; i++) {
    if (rollingBalanceArr[i] < lowestWeek.value) {
      lowestWeek.value = rollingBalanceArr[i];
      lowestWeek.index = i;
      lowestWeek.label = weekOptions[i].label;
    }
  }
  lowestWeekCache = lowestWeek;

  document.getElementById('remaining').textContent = `Remaining: €${(baseValue - totalRepayment).toLocaleString()}`;
  document.getElementById('totalRepaid').textContent = `Total Repaid: €${totalRepayment.toLocaleString()}`;
  document.getElementById('finalBalance').textContent = `Final Balance: €${(baseValue - totalRepayment).toLocaleString()}`;
  document.getElementById('lowestWeek').textContent = `Lowest Week: ${lowestWeek.label}`;

  renderChart(rollingBalanceArr, repaymentsArr, incomeArr);
  renderSpreadsheetSummary(incomeArr, rollingBalanceArr);
  renderTable(repaymentsArr, rollingBalanceArr, incomeArr);
}

// --- FILE HANDLING ---
function handleFile(event) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    rawData = json;
    extractWeekOptions(json);
    renderChart();
    renderTable();
    renderSpreadsheetSummary([], []);
    clearRepayments();
    recalculateAndRender();
  };
  reader.readAsArrayBuffer(event.target.files[0]);
}

function addRepaymentRow(weekIndex = null, amount = null) {
  const row = document.createElement('div');
  row.className = 'repayment-row';

  const weekSelect = document.createElement('select');
  weekOptions.forEach((week, idx) => {
    const option = document.createElement('option');
    option.value = week.index;
    option.textContent = week.label;
    weekSelect.appendChild(option);
  });
  if (weekIndex !== null) weekSelect.value = weekIndex;

  const amountInput = document.createElement('input');
  amountInput.type = 'number';
  amountInput.placeholder = 'Repayment €';
  if (amount !== null) amountInput.value = amount;

  weekSelect.addEventListener('change', recalculateAndRender);
  amountInput.addEventListener('input', recalculateAndRender);

  row.appendChild(weekSelect);
  row.appendChild(amountInput);

  repaymentInputs.appendChild(row);
}

function clearRepayments() {
  repaymentInputs.innerHTML = '';
  // Clear all repayments from spreadsheet row
  const repayRow = findRepaymentRowIndex();
  if (repayRow !== -1) {
    weekOptions.forEach(w => {
      rawData[repayRow][w.index] = '';
    });
  }
  recalculateAndRender();
}

// --- CHART ---
function renderChart(cashflowData = null, repaymentData = null, incomeData = null) {
  const ctx = document.getElementById('chartCanvas').getContext('2d');
  if (chart) chart.destroy();

  let datasets = [{
    label: 'Rolling Cash Balance',
    data: cashflowData ? cashflowData : Array(weekOptions.length).fill(355000),
    borderColor: '#0077cc',
    backgroundColor: chartType === 'bar' ? '#b3d7f6' : 'rgba(0, 119, 204, 0.09)',
    borderWidth: 2,
    fill: chartType === 'line' || chartType === 'radar',
    pointRadius: 4,
    tension: 0.2,
    yAxisID: 'y'
  }];

  if (showRepayments && repaymentData && repaymentData.some(r => r > 0) && chartType !== 'pie') {
    datasets.push({
      label: 'Repayments',
      data: repaymentData,
      type: 'bar',
      borderColor: '#ff9900',
      backgroundColor: 'rgba(255, 153, 0, 0.28)',
      borderWidth: 1,
      yAxisID: 'y2'
    });
  }

  chart = new Chart(ctx, {
    type: chartType,
    data: {
      labels: weekOptions.map(w => w.label),
      datasets: chartType === 'pie'
        ? [{
            label: 'Cash',
            data: cashflowData,
            backgroundColor: weekOptions.map((_, i) => i === lowestWeekCache.index ? '#ff8080' : '#0077cc')
          }]
        : datasets
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: true },
        annotation: {
          annotations: (chartType !== 'pie' && lowestWeekCache.index !== null) ? {
            lowestWeekLine: {
              type: 'line',
              xMin: lowestWeekCache.index,
              xMax: lowestWeekCache.index,
              borderColor: 'red',
              borderWidth: 2,
              label: {
                content: `Lowest (${lowestWeekCache.label})`,
                enabled: true,
                backgroundColor: 'red',
                color: 'white',
                position: 'start'
              }
            }
          } : {}
        },
        zoom: {
          pan: { enabled: true, mode: 'xy' },
          zoom: { wheel: { enabled: true }, pinch: { enabled: true }, mode: 'xy' }
        },
        tooltip: {
          callbacks: {
            title: function(context) {
              return context[0].label;
            },
            label: function(context) {
              let label = context.dataset.label || '';
              if (label) label += ': ';
              if (context.parsed.y !== undefined) {
                label += '€' + context.parsed.y.toLocaleString(undefined, {minimumFractionDigits: 0, maximumFractionDigits: 2});
              }
              return label;
            }
          }
        }
      },
      scales: chartType !== 'pie' ? {
        y: { beginAtZero: false, title: { display: true, text: 'Rolling Balance (€)' } },
        y2: {
          beginAtZero: true,
          position: 'right',
          title: { display: true, text: 'Repayments (€)' },
          grid: { drawOnChartArea: false },
          display: showRepayments
        }
      } : {}
    }
  });
}

// --- Summary Table: week labels, income, rolling balance ---
function renderSpreadsheetSummary(incomeArr, balanceArr) {
  const section = document.getElementById('spreadsheetSummarySection');
  section.innerHTML = "";
  if (!weekOptions.length) return;

  const div = document.createElement('div');
  div.className = 'spreadsheet-summary-scroll';
  const table = document.createElement('table');
  table.className = 'spreadsheet-summary-table';

  // Week label row
  const trWeeks = document.createElement('tr');
  trWeeks.className = 'week-label-row';
  trWeeks.appendChild(document.createElement('th'));
  weekOptions.forEach(w => {
    const th = document.createElement('th');
    th.textContent = w.label;
    th.className = 'sticky-week-label';
    trWeeks.appendChild(th);
  });
  table.appendChild(trWeeks);

  // Income row
  const trIncome = document.createElement('tr');
  trIncome.className = 'balance-row';
  const incomeLabel = document.createElement('td');
  incomeLabel.textContent = "Income";
  trIncome.appendChild(incomeLabel);
  incomeArr.forEach(val => {
    const td = document.createElement('td');
    td.textContent = "€" + Math.round(val);
    trIncome.appendChild(td);
  });
  table.appendChild(trIncome);

  // Balance row
  const trBal = document.createElement('tr');
  trBal.className = 'balance-row';
  const balLabel = document.createElement('td');
  balLabel.textContent = "Rolling Balance";
  trBal.appendChild(balLabel);
  balanceArr.forEach(val => {
    const td = document.createElement('td');
    td.textContent = "€" + Math.round(val);
    trBal.appendChild(td);
  });
  table.appendChild(trBal);

  div.appendChild(table);
  section.appendChild(div);
}

// --- Full Table: shows up to "Rolling cash balance" row, then computed income and balance rows ---
function renderTable(repaymentData = null, balanceArr = null, incomeArr = null) {
  const oldTable = document.getElementById('spreadsheetTable');
  if (oldTable) oldTable.innerHTML = "";
  if (rawData.length === 0) return;

  const table = document.createElement('table');
  table.style.width = '100%';
  table.style.borderCollapse = 'collapse';

  // Render up to "Rolling cash balance" row
  const rollingRowIdx = findRowIndex("Rolling cash balance");
  const tableEndIdx = rollingRowIdx > 0 ? rollingRowIdx : Math.min(rawData.length, 40);
  for (let rowIndex = 0; rowIndex <= tableEndIdx; rowIndex++) {
    const row = rawData[rowIndex];
    const tr = document.createElement('tr');
    row.forEach((cell, cellIndex) => {
      const td = document.createElement('td');
      td.style.border = '1px solid #ccc';
      td.style.padding = '4px 6px';
      // Repayment row highlight
      if (rowIndex === findRepaymentRowIndex() && weekOptions.some(w => w.index === cellIndex) && cell && cell !== '') {
        td.className = 'repayment-highlight';
      }
      td.textContent = cell || '';
      tr.appendChild(td);
    });
    table.appendChild(tr);
  }

  // Add a row for computed weekly income
  if (incomeArr) {
    const trIncome = document.createElement('tr');
    trIncome.className = 'balance-row';
    for (let i = 0; i < weekOptions[0].index; i++) trIncome.appendChild(document.createElement('td'));
    weekOptions.forEach((w, i) => {
      const td = document.createElement('td');
      td.textContent = `Income: €${Math.round(incomeArr[i])}`;
      trIncome.appendChild(td);
    });
    table.appendChild(trIncome);
  }

  // Add a row for computed rolling balance
  if (balanceArr) {
    const trBal = document.createElement('tr');
    trBal.className = 'balance-row';
    for (let i = 0; i < weekOptions[0].index; i++) trBal.appendChild(document.createElement('td'));
    weekOptions.forEach((w, i) => {
      const td = document.createElement('td');
      td.textContent = `Bal: €${Math.round(balanceArr[i])}`;
      trBal.appendChild(td);
    });
    table.appendChild(trBal);
  }

  document.getElementById('spreadsheetTable').appendChild(table);
}

// --- Export/Save/Load Functions (unchanged) ---
function exportToExcel() {
  const ws = XLSX.utils.aoa_to_sheet(rawData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, "repayment_data.xlsx");
}

function exportToPDF() {
  const container = document.querySelector('.container');
  html2canvas(container).then(canvas => {
    const imgData = canvas.toDataURL('image/png');
    const pdf = new window.jspdf.jsPDF('p', 'mm', 'a4');
    const pageWidth = pdf.internal.pageSize.getWidth();
    const pageHeight = pdf.internal.pageSize.getHeight();
    const ratio = Math.min(pageWidth / canvas.width, pageHeight / canvas.height);
    const imgWidth = canvas.width * ratio;
    const imgHeight = canvas.height * ratio;
    pdf.addImage(imgData, 'PNG', 0, 0, imgWidth, imgHeight);
    pdf.save("repayment_summary.pdf");
  });
}

function exportChartPNG() {
  const canvas = document.getElementById('chartCanvas');
  const link = document.createElement('a');
  link.href = canvas.toDataURL('image/png');
  link.download = 'cashflow_chart.png';
  link.click();
}

function savePlan() {
  const repayments = [];
  document.querySelectorAll('.repayment-row').forEach(row => {
    repayments.push({
      weekIndex: row.children[0].value,
      amount: row.children[1].value
    });
  });
  localStorage.setItem('repaymentPlan', JSON.stringify(repayments));
  alert('Repayment plan saved!');
}

function loadPlan() {
  const repayments = JSON.parse(localStorage.getItem('repaymentPlan') || '[]');
  repaymentInputs.innerHTML = '';
  repayments.forEach(rep => {
    addRepaymentRow(rep.weekIndex, rep.amount);
  });
  recalculateAndRender();
}

// --- UI/UX: Autofocus, Enter-to-add-row, Library loaded check ---
repaymentInputs.addEventListener('DOMNodeInserted', e => {
  if (e.target.classList && e.target.classList.contains('repayment-row')) {
    const input = e.target.querySelector('input[type="number"]');
    if (input) input.focus();
  }
});
repaymentInputs.addEventListener('keydown', e => {
  if (e.key === 'Enter' && e.target.tagName === 'INPUT') {
    const rows = repaymentInputs.querySelectorAll('.repayment-row');
    if (rows.length > 0 && e.target === rows[rows.length - 1].querySelector('input')) {
      addRepaymentRow();
    }
  }
});
if (!window.XLSX) {
  alert('This tool requires xlsx.js. Make sure the script is loaded.');
}
