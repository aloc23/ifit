// --- Cashflow Forecast Tool: Ensures Spreadsheet-Perfect Weekly Income and Rolling Cash Alignment ---

let rawData = [];
let chart;
let weekOptions = [];
let chartType = 'line';
let showRepayments = true;
let lowestWeekCache = { value: null, index: null, label: null };

// Filtering controls (for monthly/loan filtering if present)
let allMonths = [];
let allLoans = [];
let allMonthlyTotals = {};
let allPerLoanMonthly = {};

// Adjust these to match your sheet's actual structure!
const LABEL_COL = 1; // Column B (index 1) for all label lookups
const weeksHeaderRowIdx = 3; // Usually header is on row 4 (0-based index 3)
const startRow = 4;  // CH5 in Excel (row 5, 0-based index 4)
const endRow = 269;  // CH270 in Excel (row 270, 0-based index 269)

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

// Filtering controls (add these to your HTML if not present)
const loanSelect = document.getElementById('loanSelect');
const startMonth = document.getElementById('startMonth');
const endMonth = document.getElementById('endMonth');

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

// --- Helpers ---
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

// --- Detect week columns: Start from first "Week N =" column (e.g. CH) ---
function extractWeekOptions(data) {
  const weeksRow = data[weeksHeaderRowIdx] || [];
  weekOptions = [];
  let firstWeekCol = weeksRow.findIndex(cell => typeof cell === 'string' && /^Week\s*\d+\s*=/i.test(cell));
  if (firstWeekCol === -1) {
    alert("Could not find week columns. Check your spreadsheet headers!");
    return;
  }
  for (let i = firstWeekCol; i < weeksRow.length; i++) {
    const label = typeof weeksRow[i] === 'string' ? weeksRow[i].trim() : '';
    if (label && /^Week\s*\d+\s*=/i.test(label)) {
      weekOptions.push({ index: i, label: label });
    }
  }
}

function computeWeeklyIncomes() {
  // For each week column, sum from startRow to endRow
  let weeklyIncome = [];
  for (let w = 0; w < weekOptions.length; w++) {
    const weekCol = weekOptions[w].index;
    let sum = 0;
    for (let r = startRow; r <= endRow; r++) {
      const val = parseFloat(rawData[r]?.[weekCol] || 0);
      if (!isNaN(val)) sum += val;
    }
    weeklyIncome.push(sum);
  }
  return weeklyIncome;
}

// --- Repayments: get and set ---
function getRepaymentsArr() {
  const repayRow = findRepaymentRowIndex();
  if (repayRow === -1) return weekOptions.map(() => 0);
  return weekOptions.map(w => {
    const val = parseFloat(rawData[repayRow][w.index] || 0);
    return Math.abs(val) || 0;
  });
}
function setRepaymentForWeek(weekIdx, amount) {
  const repayRow = findRepaymentRowIndex();
  if (repayRow !== -1) {
    rawData[repayRow][weekIdx] = amount > 0 ? -Math.abs(amount) : amount;
  }
}
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

// --- Rolling Cash Balance: Spreadsheet logic ---
function computeRollingCashArr() {
  const rollingRowIdx = findRowIndex("Rolling cash balance");
  if (rollingRowIdx === -1) return weekOptions.map(() => 0);
  let rollingBalance = [];
  for (let w = 0; w < weekOptions.length; w++) {
    const weekCol = weekOptions[w].index;
    const prevRowVal = parseFloat(rawData[rollingRowIdx-1][weekCol] || 0);
    if (w === 0) {
      const prevColBal = parseFloat(rawData[rollingRowIdx][weekCol - 1]) || 0;
      rollingBalance.push(prevColBal + prevRowVal);
    } else {
      rollingBalance.push(rollingBalance[w-1] + prevRowVal);
    }
  }
  return rollingBalance;
}

function recalculateAndRender() {
  if (weekOptions.length === 0 || rawData.length === 0) return;
  const { repaymentsArr, totalRepayment } = getRepaymentData();
  const weeklyIncome = computeWeeklyIncomes();
  const rollingBalance = computeRollingCashArr();

  // Find lowest week
  let lowestWeek = { value: Infinity, index: null, label: '' };
  for (let i = 0; i < rollingBalance.length; i++) {
    if (rollingBalance[i] < lowestWeek.value) {
      lowestWeek.value = rollingBalance[i];
      lowestWeek.index = i;
      lowestWeek.label = weekOptions[i].label;
    }
  }
  lowestWeekCache = lowestWeek;

  // The "remaining" and "final balance" are now the last rolling balance
  document.getElementById('remaining').textContent = `Remaining: €${(rollingBalance[rollingBalance.length-1]||0).toLocaleString()}`;
  document.getElementById('totalRepaid').textContent = `Total Repaid: €${totalRepayment.toLocaleString()}`;
  document.getElementById('finalBalance').textContent = `Final Balance: €${(rollingBalance[rollingBalance.length-1]||0).toLocaleString()}`;
  document.getElementById('lowestWeek').textContent = `Lowest Week: ${lowestWeek.label}`;

  renderChart(rollingBalance, repaymentsArr, weeklyIncome);
  renderSpreadsheetSummary(weeklyIncome, rollingBalance);
  renderTable(repaymentsArr, rollingBalance, weeklyIncome);

  // Filtering controls: update summary/chart if present
  if (loanSelect && startMonth && endMonth) {
    const filtered = getFilteredMonthlyTotals();
    updateSummaryTable(filtered);
    updateRepaymentChart(filtered);
  }
}

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

    // If using monthly/loan filtering, attempt to parse for those as well
    parseRepaymentsByMonthFile(sheet);
  };
  reader.readAsArrayBuffer(event.target.files[0]);
}

// --- Filtering/Monthly/Lending Controls & Logic ---
function normalizeMonthLabel(label) {
  const months = ['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'];
  let m = label.match(/^([A-Za-z]+)\s*(\d{2,4})$/);
  if (m) {
    let month = months.indexOf(m[1].toLowerCase()) + 1;
    let year = parseInt(m[2], 10);
    if (year < 100) year += 2000;
    return `${year}-${month.toString().padStart(2,'0')}`;
  }
  m = label.match(/^(\d{4})[-\/](\d{1,2})/);
  if (m) {
    return `${m[1]}-${m[2].padStart(2,'0')}`;
  }
  return label;
}

// Parse repayments by month/account from a worksheet object (SheetJS)
function parseRepaymentsByMonthFile(sheet) {
  try {
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    if (!json.length) return;
    const headers = json[0];
    // Assume first column is label, rest are months
    const dateCols = headers.slice(1).map(normalizeMonthLabel);

    const monthlyTotals = {};
    const perLoanMonthly = {};

    for (let i = 1; i < json.length; ++i) {
      const row = json[i];
      if (!row || !row[0]) continue;
      const loan = String(row[0]).trim();
      perLoanMonthly[loan] = {};
      for (let j = 1; j < row.length; ++j) {
        const month = dateCols[j-1];
        const val = parseFloat(row[j]);
        if (!isNaN(val)) {
          perLoanMonthly[loan][month] = (perLoanMonthly[loan][month] || 0) + val;
          monthlyTotals[month] = (monthlyTotals[month] || 0) + val;
        }
      }
    }
    allMonthlyTotals = monthlyTotals;
    allPerLoanMonthly = perLoanMonthly;
    populateFilterControls(perLoanMonthly, monthlyTotals);
    setupFilterListeners();
    // Initial render
    const filtered = getFilteredMonthlyTotals();
    updateSummaryTable(filtered);
    updateRepaymentChart(filtered);
  } catch (err) {
    // Silently fail, don't block main tool
  }
}

// --- Filtering controls: populate and handle ---
function populateFilterControls(perLoanMonthly, monthlyTotals) {
  // Loans
  if (!loanSelect || !startMonth || !endMonth) return;
  allLoans = Object.keys(perLoanMonthly);
  loanSelect.innerHTML = '<option value="__all__">All Loans</option>';
  allLoans.forEach(loan => {
    const opt = document.createElement('option');
    opt.value = loan;
    opt.textContent = loan;
    loanSelect.appendChild(opt);
  });

  // Months (from all months in data)
  allMonths = Array.from(
    new Set(
      Object.values(perLoanMonthly).flatMap(m => Object.keys(m))
    )
  ).concat(Object.keys(monthlyTotals))
   .filter((v,i,a)=>a.indexOf(v)==i)
   .sort();

  startMonth.innerHTML = '';
  endMonth.innerHTML = '';
  allMonths.forEach(month => {
    const o1 = document.createElement('option');
    o1.value = month; o1.textContent = month;
    const o2 = o1.cloneNode(true);
    startMonth.appendChild(o1);
    endMonth.appendChild(o2);
  });
  // Set defaults
  startMonth.selectedIndex = 0;
  endMonth.selectedIndex = allMonths.length - 1;
}

function getFilteredMonthlyTotals() {
  if (!loanSelect || !startMonth || !endMonth) return {};
  const loan = loanSelect.value;
  const start = startMonth.value;
  const end = endMonth.value;
  let months = allMonths
    .filter(m => m >= start && m <= end);

  let filtered = {};
  if (loan === "__all__") {
    months.forEach(month => {
      filtered[month] = allMonthlyTotals[month] || 0;
    });
  } else {
    months.forEach(month => {
      filtered[month] = (allPerLoanMonthly[loan] && allPerLoanMonthly[loan][month]) || 0;
    });
  }
  return filtered;
}

function setupFilterListeners() {
  if (!loanSelect || !startMonth || !endMonth) return;
  ['loanSelect','startMonth','endMonth'].forEach(id => {
    document.getElementById(id).addEventListener('change', () => {
      const filtered = getFilteredMonthlyTotals();
      updateSummaryTable(filtered);
      updateRepaymentChart(filtered);
    });
  });
}

// --- Summary table and chart for filtered results ---
let repaymentChart;
function updateSummaryTable(monthlyTotals) {
  const tbody = document.getElementById('summaryTableBody');
  if (!tbody) return;
  tbody.innerHTML = '';
  const months = Object.keys(monthlyTotals).sort();
  for (const month of months) {
    const tr = document.createElement('tr');
    const tdMonth = document.createElement('td');
    tdMonth.textContent = month;
    const tdTotal = document.createElement('td');
    tdTotal.textContent = monthlyTotals[month].toLocaleString(undefined, {minimumFractionDigits: 2});
    tr.appendChild(tdMonth);
    tr.appendChild(tdTotal);
    tbody.appendChild(tr);
  }
}
function updateRepaymentChart(monthlyTotals) {
  const canvas = document.getElementById('repaymentChart');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const months = Object.keys(monthlyTotals).sort();
  const data = months.map(m => monthlyTotals[m]);
  if (repaymentChart) {
    repaymentChart.data.labels = months;
    repaymentChart.data.datasets[0].data = data;
    repaymentChart.update();
  } else {
    repaymentChart = new Chart(ctx, {
      type: 'bar',
      data: {
        labels: months,
        datasets: [{
          label: 'Total Repayments',
          data: data,
          backgroundColor: 'rgba(54, 162, 235, 0.6)'
        }]
      },
      options: {
        responsive: true,
        plugins: {
          legend: { display: false },
          title: { display: true, text: 'Monthly Repayment Totals' }
        },
        scales: {
          y: { beginAtZero: true }
        }
      }
    });
  }
}

// --- Repayment input rows and core logic ---
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
  const repayRow = findRepaymentRowIndex();
  if (repayRow !== -1) {
    weekOptions.forEach(w => {
      rawData[repayRow][w.index] = '';
    });
  }
  recalculateAndRender();
}

function renderChart(cashflowData = null, repaymentData = null, incomeData = null) {
  const ctx = document.getElementById('chartCanvas').getContext('2d');
  if (chart) chart.destroy();

  let datasets = [{
    label: 'Rolling Cash Balance',
    data: cashflowData ? cashflowData : weekOptions.map(() => 0),
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
      // Only render week columns and after; skip columns before first week column
      if (cellIndex < weekOptions[0].index) return;
      const td = document.createElement('td');
      td.style.border = '1px solid #ccc';
      td.style.padding = '4px 6px';
      if (rowIndex === findRepaymentRowIndex() && weekOptions.some(w => w.index === cellIndex) && cell && cell !== '') {
        td.className = 'repayment-highlight';
      }
      td.textContent = cell || '';
      tr.appendChild(td);
    });
    table.appendChild(tr);
  }

  // Add a row for weekly income
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

  // Add a row for rolling balance
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

// --- Export/Save/Load Functions ---
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

// --- UI/UX Improvements ---
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
