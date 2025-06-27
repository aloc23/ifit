// ------- Cashflow Forecast Tool (Revised for Column B Labels and Excel-Logic) ---------

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

// ---------------- EVENT LISTENERS ----------------

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

// ----------------- FIND ROW HELPERS -----------------

function findRowIndex(label) {
  label = label.trim().toLowerCase();
  let idx = rawData.findIndex(row =>
    row[LABEL_COL] && row[LABEL_COL].toString().trim().toLowerCase() === label
  );
  if (idx !== -1) return idx;
  idx = rawData.findIndex(row =>
    row[LABEL_COL] && row[LABEL_COL].toString().trim().toLowerCase().includes(label)
  );
  if (idx !== -1) return idx;
  return -1;
}

function findRepaymentRowIndex() {
  return findRowIndex("Mayweather Investment Repayment (Investment 1 and 2)");
}
function findWeeklyIncomeRowIndex() {
  return findRowIndex("Weekly income / cash position");
}
function findRollingBalanceRowIndex() {
  return findRowIndex("Rolling cash balance");
}

// ----------- WEEK HEADER EXTRACTION (WEEKS IN COLUMNS F+) ------------

function extractWeekOptions(data) {
  const weeksRow = data[3] || [];
  weekOptions = [];
  for (let i = 5; i < weeksRow.length; i++) {
    const label = typeof weeksRow[i] === 'string' ? weeksRow[i].trim() : '';
    if (label) weekOptions.push({ index: i, label: label });
  }
}

// ----------- INCOME, REPAYMENT, ROLLING BALANCE CALCULATIONS ------------

// Sum for each week column over all relevant rows (Excel: =SUM(DH5:DH270))
function computeWeeklyIncome(weekIdx) {
  const startRow = 5;
  const endRow = 270;
  let sum = 0;
  for (let r = startRow; r <= endRow; r++) {
    const val = parseFloat(rawData[r]?.[weekIdx] || 0);
    if (!isNaN(val)) sum += val;
  }
  return sum;
}

// Get repayments for each week from the repayment row
function getRepaymentsArr() {
  const repaymentRow = findRepaymentRowIndex();
  if (repaymentRow === -1) return Array(weekOptions.length).fill(0);
  return weekOptions.map(w => {
    const val = parseFloat(rawData[repaymentRow][w.index] || 0);
    return isNaN(val) ? 0 : val;
  });
}

// Compute rolling cash balance as in Excel: previous balance + weekly income - repayment
function computeRollingCashArr(weekOptions, baseValue) {
  let rollingCashArr = [];
  let prevBalance = baseValue;
  const repayments = getRepaymentsArr();

  for (let i = 0; i < weekOptions.length; i++) {
    const weekIdx = weekOptions[i].index;
    const weeklyIncome = computeWeeklyIncome(weekIdx);
    const repayment = repayments[i] || 0;
    const thisBalance = prevBalance + weeklyIncome - repayment;
    rollingCashArr.push(thisBalance);
    prevBalance = thisBalance;
  }
  return rollingCashArr;
}

// Set repayment in spreadsheet's row for a specific week (UI to spreadsheet)
function setRepaymentForWeek(weekIdx, amount) {
  const repayRow = findRepaymentRowIndex();
  if (repayRow !== -1) {
    rawData[repayRow][weekIdx] = amount;
  }
}

// When user enters repayments, update spreadsheet and tally total
function getRepaymentData() {
  let totalRepayment = 0;
  document.querySelectorAll('.repayment-row').forEach(row => {
    const weekIdx = parseInt(row.children[0].value);
    const amount = parseFloat(row.children[1].value) || 0;
    if (!isNaN(amount)) {
      setRepaymentForWeek(weekIdx, amount);
      totalRepayment += amount;
    }
  });
  const repaymentArr = getRepaymentsArr();
  return { repaymentArr, totalRepayment };
}

// ------------------- MAIN RECALC/RENDER FUNCTION ----------------------

function recalculateAndRender() {
  if (weekOptions.length === 0 || rawData.length === 0) return;

  const { repaymentArr, totalRepayment } = getRepaymentData();
  const baseValue = 355000;

  // Income, rolling balance arrays
  const incomeArr = weekOptions.map(w => computeWeeklyIncome(w.index));
  const rollingBalanceArr = computeRollingCashArr(weekOptions, baseValue);

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
  document.getElementById('totalRepaid').textContent = `€${totalRepayment.toLocaleString()}`;
  document.getElementById('finalBalance').textContent = `€${rollingBalanceArr[rollingBalanceArr.length - 1].toLocaleString()}`;
  document.getElementById('lowestWeek').textContent = lowestWeek.label;

  renderChart(rollingBalanceArr, repaymentArr);
  renderTable(repaymentArr, rollingBalanceArr, incomeArr);
}

// ------------- FILE HANDLING, TABLE, CHART RENDERING -------------

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
  amountInput.placeholder = 'Amount €';
  if (amount !== null) amountInput.value = amount;

  weekSelect.addEventListener('change', recalculateAndRender);
  amountInput.addEventListener('input', recalculateAndRender);

  row.appendChild(weekSelect);
  row.appendChild(amountInput);

  repaymentInputs.appendChild(row);
}

function clearRepayments() {
  repaymentInputs.innerHTML = '';
  recalculateAndRender();
}

function renderChart(cashflowData = null, repaymentData = null) {
  const ctx = document.getElementById('chartCanvas').getContext('2d');
  if (chart) chart.destroy();

  let datasets = [{
    label: 'Rolling Cash Balance',
    data: cashflowData ? cashflowData : Array(weekOptions.length).fill(355000),
    borderColor: '#0077cc',
    backgroundColor: chartType === 'bar' ? '#b3d7f6' : 'rgba(0, 119, 204, 0.1)',
    borderWidth: 2,
    fill: chartType === 'line' || chartType === 'radar' ? true : false,
    pointRadius: 4,
    tension: 0.2
  }];

  if (showRepayments && repaymentData && repaymentData.some(r => r > 0) && chartType !== 'pie') {
    datasets.push({
      label: 'Repayments',
      data: repaymentData,
      type: 'bar',
      borderColor: '#ff9900',
      backgroundColor: 'rgba(255, 153, 0, 0.4)',
      borderWidth: 1,
      yAxisID: 'y2'
    });
  }

  let pieLabels = weekOptions.map(w => w.label);
  let pieData = cashflowData || Array(weekOptions.length).fill(355000);

  chart = new Chart(ctx, {
    type: chartType,
    data: {
      labels: pieLabels,
      datasets: chartType === 'pie'
        ? [{ label: 'Cash', data: pieData, backgroundColor: pieLabels.map((_,i) => i === lowestWeekCache.index ? '#ff8080' : '#0077cc') }]
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
            label: function(context) {
              let label = context.dataset.label || '';
              if (label) label += ': ';
              if (context.parsed.y !== undefined) {
                label += '€' + context.parsed.y.toLocaleString();
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

function renderTable(repaymentData = null, balanceArr = null, incomeArr = null) {
  const oldTable = document.getElementById('spreadsheetTable');
  if (oldTable) oldTable.innerHTML = "";
  if (rawData.length === 0) return;

  const table = document.createElement('table');
  table.style.width = '100%';
  table.style.borderCollapse = 'collapse';

  // Render spreadsheet up to and including the Rolling cash balance row
  const balanceRowIdx = findRollingBalanceRowIndex();

  for (let rowIndex = 0; rowIndex <= balanceRowIdx; rowIndex++) {
    const row = rawData[rowIndex];
    const tr = document.createElement('tr');
    row.forEach((cell, cellIndex) => {
      const td = document.createElement('td');
      td.style.border = '1px solid #ccc';
      td.style.padding = '4px 6px';

      // Editable: for all rows below header and not header itself
      if (rowIndex >= 4 && rowIndex !== balanceRowIdx) {
        td.className = 'editable-cell';
        const input = document.createElement('input');
        input.type = 'text';
        input.value = cell || '';
        input.addEventListener('input', (e) => {
          rawData[rowIndex][cellIndex] = e.target.value;
          recalculateAndRender();
        });
        td.appendChild(input);
      } else {
        td.textContent = cell || '';
      }
      // Highlight repayment columns (in the repayment row)
      if (rowIndex === findRepaymentRowIndex() && repaymentData && repaymentData[cellIndex-5] > 0) {
        td.classList.add('repayment-highlight');
      }
      tr.appendChild(td);
    });
    table.appendChild(tr);
  }

  // Add a row for computed weekly income
  if (incomeArr) {
    const trIncome = document.createElement('tr');
    trIncome.className = 'balance-row';
    for (let i = 0; i < 5; i++) trIncome.appendChild(document.createElement('td'));
    weekOptions.forEach((w, i) => {
      const td = document.createElement('td');
      td.textContent = incomeArr[i] !== 0 ? `Income: €${Math.round(incomeArr[i])}` : '';
      trIncome.appendChild(td);
    });
    table.appendChild(trIncome);
  }

  // Add a row for computed rolling balance
  if (balanceArr) {
    const trBal = document.createElement('tr');
    trBal.className = 'balance-row';
    for (let i = 0; i < 5; i++) trBal.appendChild(document.createElement('td'));
    weekOptions.forEach((w, i) => {
      const td = document.createElement('td');
      td.textContent = typeof balanceArr[i] === 'number' ? `Bal: €${Math.round(balanceArr[i])}` : '';
      trBal.appendChild(td);
    });
    table.appendChild(trBal);
  }

  document.getElementById('spreadsheetTable').appendChild(table);
}

// --------- Export/Save/Load Functions ----------

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
