// Cashflow Forecast Tool -- Repayment Row Edit/Remove + Fix Old Repayment + Highlight Repayment Weeks

let rawData = [];
let chart;
let weekOptions = [];
let chartType = 'line';
let showRepayments = true;
let lowestWeekCache = { value: null, index: null, label: null };
let weekLabels = [];
let weekFilterRange = [0, 0];
let startingBalance = 0;
let loanOutstanding = 0;

// Sheet structure (adjust if your sheet is different)
const LABEL_COL = 1; // Column B (index 1) for label lookups
const weeksHeaderRowIdx = 3; // Header row (index 3, i.e. Excel row 4)
const startRow = 4;  // CH5 in Excel (row 5, 0-based index 4)
const endRow = 269;  // CH270 in Excel (row 270, 0-based index 269)
const firstWeekCol = 5; // Column F (index 5)

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
const weekFilterControls = document.getElementById('weekFilterControls');
const startWeekSelect = document.getElementById('startWeekSelect');
const endWeekSelect = document.getElementById('endWeekSelect');
const startingBalanceInput = document.getElementById('startingBalanceInput');
const loanOutstandingInput = document.getElementById('loanOutstandingInput');

startingBalanceInput.addEventListener('input', () => {
  startingBalance = parseFloat(startingBalanceInput.value) || 0;
  recalculateAndRender();
});
loanOutstandingInput.addEventListener('input', () => {
  loanOutstanding = parseFloat(loanOutstandingInput.value) || 0;
  recalculateAndRender();
});

fileInput.addEventListener('change', handleFile);
addRepaymentBtn.addEventListener('click', addRepaymentRow);
clearRepaymentsBtn.addEventListener('click', clearRepayments);
savePlanBtn.addEventListener('click', savePlan);
loadPlanBtn.addEventListener('click', loadPlan);
exportExcelBtn.addEventListener('click', exportToExcel);
exportPDFBtn.addEventListener('click', exportToPDF);
if (exportPNGBtn) exportPNGBtn.addEventListener('click', exportChartPNG);
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
  // Adjust as needed for your actual label!
  return findRowIndex("Mayweather Investment Repayment (Investment 1 and 2)");
}

function extractWeekOptions(data) {
  const weeksRow = data[weeksHeaderRowIdx] || [];
  weekOptions = [];
  weekLabels = [];
  for (let i = firstWeekCol; i < weeksRow.length; i++) {
    const label = typeof weeksRow[i] === 'string' ? weeksRow[i].trim() : '';
    if (label && /^Week\s*\d+/i.test(label)) {
      weekOptions.push({ index: i, label: label });
      weekLabels.push(label);
    }
  }
  weekFilterRange = [0, weekLabels.length-1];
}

function computeWeeklyIncomes(filtered = false) {
  let weeklyIncome = [];
  let sIdx = filtered ? weekFilterRange[0] : 0;
  let eIdx = filtered ? weekFilterRange[1] : weekOptions.length - 1;
  for (let w = sIdx; w <= eIdx; w++) {
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

function getRepaymentsArr(filtered = false) {
  const repayRow = findRepaymentRowIndex();
  if (repayRow === -1) return weekOptions.map(() => 0);
  let arr = weekOptions.map(w => {
    const val = parseFloat(rawData[repayRow][w.index] || 0);
    return Math.abs(val) || 0;
  });
  if (filtered) {
    return arr.slice(weekFilterRange[0], weekFilterRange[1]+1);
  }
  return arr;
}
function setRepaymentForWeek(weekIdx, amount) {
  const repayRow = findRepaymentRowIndex();
  if (repayRow !== -1) {
    rawData[repayRow][weekIdx] = amount > 0 ? -Math.abs(amount) : amount;
  }
}
function getRepaymentData(filtered = false) {
  let totalRepayment = 0;
  document.querySelectorAll('.repayment-row').forEach(row => {
    const weekIdx = parseInt(row.querySelector('select').value);
    let amount = parseFloat(row.querySelector('input[type="number"]').value) || 0;
    if (!isNaN(amount)) {
      setRepaymentForWeek(weekIdx, amount);
      totalRepayment += Math.abs(amount);
    }
  });
  const repaymentsArr = getRepaymentsArr(filtered);
  return { repaymentsArr, totalRepayment };
}

// Rolling Cash Balance: honors starting balance and prior rolling for filtered view
function computeRollingCashArr(filtered = false) {
  let sIdx = filtered ? weekFilterRange[0] : 0;
  let eIdx = filtered ? weekFilterRange[1] : weekOptions.length - 1;

  // Weekly sums for all weeks, from row 5 down
  let allWeeklySums = weekOptions.map(w => {
    let sum = 0;
    for (let r = startRow; r <= endRow; r++) {
      const val = parseFloat(rawData[r]?.[w.index] || 0);
      if (!isNaN(val)) sum += val;
    }
    return sum;
  });

  // 1. Find starting balance: prefer user entry, else rolling row prev week
  let rollingRowIdx = findRowIndex("Rolling cash balance");
  let prevBalance = Number.isFinite(startingBalance) && startingBalance !== 0 ? startingBalance : 0;
  if (rollingRowIdx !== -1 && sIdx > 0 && (!startingBalance || startingBalance === 0)) {
    let prevWeekCol = weekOptions[sIdx - 1].index;
    prevBalance = parseFloat(rawData[rollingRowIdx][prevWeekCol]) || 0;
  }

  let rollingBalance = [];
  for (let w = sIdx; w <= eIdx; w++) {
    if (w === sIdx) {
      rollingBalance.push(prevBalance + allWeeklySums[w]);
    } else {
      rollingBalance.push(rollingBalance[rollingBalance.length - 1] + allWeeklySums[w]);
    }
  }
  return rollingBalance;
}

function recalculateAndRender(filtered = false) {
  if (weekOptions.length === 0 || rawData.length === 0) return;
  const { repaymentsArr, totalRepayment } = getRepaymentData(filtered);
  const weeklyIncome = computeWeeklyIncomes(filtered);
  const rollingBalance = computeRollingCashArr(filtered);

  let offset = filtered ? weekFilterRange[0] : 0;
  let lowestWeek = { value: Infinity, index: null, label: '' };
  let negWeeks = [];
  for (let i = 0; i < rollingBalance.length; i++) {
    if (rollingBalance[i] < lowestWeek.value) {
      lowestWeek.value = rollingBalance[i];
      lowestWeek.index = i;
      lowestWeek.label = weekLabels[i+offset];
    }
    if (rollingBalance[i] < 0) {
      negWeeks.push(`${weekLabels[i + offset]}: €${Math.round(rollingBalance[i])}`);
    }
  }
  lowestWeekCache = lowestWeek;

  // Final Bank Balance = last rolling balance value
  const finalBankBalance = rollingBalance[rollingBalance.length-1] || 0;
  document.getElementById('finalBankBalance').textContent = `Final Bank Balance: €${finalBankBalance.toLocaleString()}`;
  document.getElementById('totalRepaid').textContent = `Total Repaid: €${totalRepayment.toLocaleString()}`;

  // Remaining = loanOutstanding - totalRepayment
  const remaining = (loanOutstanding || 0) - totalRepayment;
  document.getElementById('remaining').textContent = `Remaining: €${remaining.toLocaleString()}`;

  // Lowest week and negative weeks list
  let lowestLabel = `Lowest Week: ${lowestWeek.label}`;
  if (lowestWeek.value !== Infinity && lowestWeek.label) {
    lowestLabel += ` (€${Math.round(lowestWeek.value)})`;
  }
  document.getElementById('lowestWeek').childNodes[0].textContent = lowestLabel + " ";
  const negWeeksUl = document.getElementById('negWeeksUl');
  negWeeksUl.innerHTML = negWeeks.length ? negWeeks.map(w => `<li>${w}</li>`).join('') : '<li>None</li>';
  document.getElementById('negativeWeeksList').style.display = negWeeks.length ? "inline-block" : "none";

  renderChart(rollingBalance, repaymentsArr, weeklyIncome, filtered);
  renderSpreadsheetSummary(weeklyIncome, rollingBalance, filtered);
  renderTable(repaymentsArr, rollingBalance, weeklyIncome, filtered);
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
    setupWeekFilterControls();
    renderChart();
    renderTable();
    renderSpreadsheetSummary([], []);
    clearRepayments();
    recalculateAndRender();
  };
  reader.readAsArrayBuffer(event.target.files[0]);
}

function setupWeekFilterControls() {
  if (weekLabels.length === 0) {
    weekFilterControls.style.display = "none";
    return;
  }
  weekFilterControls.style.display = "";
  startWeekSelect.innerHTML = '';
  endWeekSelect.innerHTML = '';
  weekLabels.forEach((label, idx) => {
    let opt1 = document.createElement('option');
    opt1.value = idx;
    opt1.textContent = label;
    let opt2 = opt1.cloneNode(true);
    startWeekSelect.appendChild(opt1);
    endWeekSelect.appendChild(opt2);
  });
  startWeekSelect.selectedIndex = 0;
  endWeekSelect.selectedIndex = weekLabels.length - 1;
  weekFilterRange = [0, weekLabels.length-1];
  startWeekSelect.onchange = updateWeekFilter;
  endWeekSelect.onchange = updateWeekFilter;
}

function updateWeekFilter() {
  let sIdx = parseInt(startWeekSelect.value, 10);
  let eIdx = parseInt(endWeekSelect.value, 10);
  if (eIdx < sIdx) eIdx = sIdx;
  weekFilterRange = [sIdx, eIdx];
  updateRepaymentRowsForFilter();
  recalculateAndRender(true);
}

function updateRepaymentRowsForFilter() {
  document.querySelectorAll('.repayment-row').forEach(row => {
    const weekIdx = parseInt(row.querySelector('select').value);
    const weekArrIdx = weekOptions.findIndex(w => w.index == weekIdx);
    if (weekArrIdx < weekFilterRange[0] || weekArrIdx > weekFilterRange[1]) {
      row.style.display = "none";
    } else {
      row.style.display = "";
    }
  });
}

// --- Repayment Row With Edit/Remove Buttons and Fixes ---
function addRepaymentRow(weekIndex = null, amount = null) {
  const row = document.createElement('div');
  row.className = 'repayment-row';

  // Week select
  const weekSelect = document.createElement('select');
  weekOptions.forEach((week, idx) => {
    if (idx < weekFilterRange[0] || idx > weekFilterRange[1]) return;
    const option = document.createElement('option');
    option.value = week.index;
    option.textContent = week.label;
    weekSelect.appendChild(option);
  });
  if (weekIndex !== null) weekSelect.value = weekIndex;

  // Amount input
  const amountInput = document.createElement('input');
  amountInput.type = 'number';
  amountInput.placeholder = 'Repayment €';
  if (amount !== null) amountInput.value = amount;

  // Edit button
  const editBtn = document.createElement('button');
  editBtn.textContent = 'Edit';
  editBtn.type = 'button';
  editBtn.style.marginLeft = "6px";
  let isEditing = false;

  // Remove button
  const removeBtn = document.createElement('button');
  removeBtn.textContent = 'Remove';
  removeBtn.type = 'button';
  removeBtn.style.marginLeft = "6px";

  // Track previous week index for cleanup
  let repayRow = findRepaymentRowIndex();
  let previousWeekIndex = weekSelect.value;

  function saveEdit() {
    weekSelect.disabled = true;
    amountInput.disabled = true;
    editBtn.textContent = 'Edit';
    isEditing = false;
    // Clear old repayment if week changed
    if (repayRow !== -1 && previousWeekIndex !== weekSelect.value) {
      rawData[repayRow][previousWeekIndex] = '';
      previousWeekIndex = weekSelect.value;
    }
    recalculateAndRender(true);
  }

  editBtn.onclick = () => {
    isEditing = !isEditing;
    weekSelect.disabled = !isEditing;
    amountInput.disabled = !isEditing;
    editBtn.textContent = isEditing ? 'Save' : 'Edit';
    if (!isEditing) saveEdit();
  };

  removeBtn.onclick = () => {
    // Clear repayment in rawData
    repayRow = findRepaymentRowIndex();
    if (repayRow !== -1) {
      rawData[repayRow][weekSelect.value] = '';
    }
    row.remove();
    recalculateAndRender(true);
  };

  weekSelect.addEventListener('change', () => {
    if (!isEditing) saveEdit();
    recalculateAndRender(true);
  });
  amountInput.addEventListener('input', () => {
    if (!isEditing) saveEdit();
    recalculateAndRender(true);
  });

  // Default to "not editing" after creation
  weekSelect.disabled = true;
  amountInput.disabled = true;

  row.appendChild(weekSelect);
  row.appendChild(amountInput);
  row.appendChild(editBtn);
  row.appendChild(removeBtn);

  repaymentInputs.appendChild(row);
}

// --- End Repayment Row Edit/Remove ---

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

function renderChart(cashflowData = null, repaymentData = null, incomeData = null, filtered = false) {
  const canvas = document.getElementById('chartCanvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  if (chart) chart.destroy();

  let sIdx = filtered ? weekFilterRange[0] : 0;
  let eIdx = filtered ? weekFilterRange[1] : weekLabels.length - 1;
  let labels = weekLabels.slice(sIdx, eIdx+1);

  // Defensive: If labels or data are empty, don't render chart
  if (!labels.length || !cashflowData || cashflowData.length === 0) {
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    return;
  }

  let datasets = [{
    label: 'Rolling Cash Balance',
    data: cashflowData,
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
      labels: labels,
      datasets: chartType === 'pie'
        ? [{
            label: 'Cash',
            data: cashflowData,
            backgroundColor: labels.map((_, i) => i === lowestWeekCache.index ? '#ff8080' : '#0077cc')
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

// --- Highlight Repayment Weeks in Summary ---
function renderSpreadsheetSummary(incomeArr, balanceArr, filtered = false) {
  const section = document.getElementById('spreadsheetSummarySection');
  section.innerHTML = "";
  if (!weekOptions.length) return;
  let sIdx = filtered ? weekFilterRange[0] : 0;
  let eIdx = filtered ? weekFilterRange[1] : weekLabels.length - 1;

  const div = document.createElement('div');
  div.className = 'spreadsheet-summary-scroll';
  const table = document.createElement('table');
  table.className = 'spreadsheet-summary-table';

  const repayRowIdx = findRepaymentRowIndex();

  // Week label row (with repayment highlight)
  const trWeeks = document.createElement('tr');
  trWeeks.className = 'week-label-row';
  trWeeks.appendChild(document.createElement('th'));
  weekLabels.slice(sIdx, eIdx+1).forEach((w, idx) => {
    const th = document.createElement('th');
    th.textContent = w;
    th.className = 'sticky-week-label';
    const weekOptIdx = sIdx + idx;
    const weekCol = weekOptions[weekOptIdx].index;
    if (repayRowIdx !== -1 && rawData[repayRowIdx][weekCol] && parseFloat(rawData[repayRowIdx][weekCol]) !== 0) {
      th.classList.add('repayment-highlight');
    }
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
    td.textContent = "€" + (isNaN(val) ? '' : Math.round(val));
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
    td.textContent = "€" + (isNaN(val) ? '' : Math.round(val));
    trBal.appendChild(td);
  });
  table.appendChild(trBal);

  div.appendChild(table);
  section.appendChild(div);
}

// --- End Highlight Repayment Weeks ---

function renderTable(repaymentData = null, balanceArr = null, incomeArr = null, filtered = false) {
  const oldTable = document.getElementById('spreadsheetTable');
  if (oldTable) oldTable.innerHTML = "";
  if (rawData.length === 0) return;

  const table = document.createElement('table');
  table.style.width = '100%';
  table.style.borderCollapse = 'collapse';

  let sIdx = filtered ? weekFilterRange[0] : 0;
  let eIdx = filtered ? weekFilterRange[1] : weekLabels.length-1;

  // Render up to "Rolling cash balance" row
  const rollingRowIdx = findRowIndex("Rolling cash balance");
  const tableEndIdx = rollingRowIdx > 0 ? rollingRowIdx : Math.min(rawData.length, 40);
  for (let rowIndex = 0; rowIndex <= tableEndIdx; rowIndex++) {
    const row = rawData[rowIndex];
    const tr = document.createElement('tr');
    for (let i = sIdx+firstWeekCol; i <= eIdx+firstWeekCol; i++) {
      const td = document.createElement('td');
      td.style.border = '1px solid #ccc';
      td.style.padding = '4px 6px';
      if (rowIndex === findRepaymentRowIndex() && row[i] && row[i] !== '') {
        td.className = 'repayment-highlight';
      }
      td.textContent = row[i] || '';
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }

  // Add a row for weekly income
  if (incomeArr) {
    const trIncome = document.createElement('tr');
    trIncome.className = 'balance-row';
    for (let i = 0; i < incomeArr.length; i++) {
      const td = document.createElement('td');
      td.textContent = `Income: €${isNaN(incomeArr[i]) ? '' : Math.round(incomeArr[i])}`;
      trIncome.appendChild(td);
    }
    table.appendChild(trIncome);
  }

  // Add a row for rolling balance
  if (balanceArr) {
    const trBal = document.createElement('tr');
    trBal.className = 'balance-row';
    for (let i = 0; i < balanceArr.length; i++) {
      const td = document.createElement('td');
      td.textContent = `Bal: €${isNaN(balanceArr[i]) ? '' : Math.round(balanceArr[i])}`;
      trBal.appendChild(td);
    }
    table.appendChild(trBal);
  }

  document.getElementById('spreadsheetTable').appendChild(table);
}

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
      weekIndex: row.querySelector('select').value,
      amount: row.querySelector('input[type="number"]').value
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
