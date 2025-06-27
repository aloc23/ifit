let rawData = [];
let chart;
let weekOptions = [];
let chartType = 'line';
let showRepayments = true;
let lowestWeekCache = { value: null, index: null, label: null };

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
    setTimeout(recalculateAndRender, 100); // Ensure initial state is fresh
  };
  reader.readAsArrayBuffer(event.target.files[0]);
}

function extractWeekOptions(data) {
  const weeksRow = data[3] || [];
  weekOptions = weeksRow.map((weekLabel, i) => {
    const label = typeof weekLabel === 'string' ? weekLabel.trim() : '';
    if (label.toLowerCase().includes('week') || label) {
      return { index: i, label: label };
    }
    return null;
  }).filter(Boolean);
}

function addRepaymentRow(weekIndex = null, amount = null) {
  const row = document.createElement('div');
  row.className = 'repayment-row';

  const weekSelect = document.createElement('select');
  weekOptions.forEach(week => {
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

  // Instant update on change
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

function getRepaymentData() {
  let repaymentData = Array(weekOptions.length).fill(0);
  let totalRepayment = 0;
  document.querySelectorAll('.repayment-row').forEach(row => {
    const weekIndex = parseInt(row.children[0].value);
    const amount = parseFloat(row.children[1].value) || 0;
    if (!isNaN(amount)) {
      repaymentData[weekIndex] += amount;
      totalRepayment += amount;
    }
  });
  return { repaymentData, totalRepayment };
}

function recalculateAndRender() {
  if (weekOptions.length === 0 || rawData.length === 0) return;

  let base = 355000;
  const { repaymentData, totalRepayment } = getRepaymentData();

  let cashflow = [];
  let lowestWeek = { value: Infinity, index: null, label: '' };

  for (let i = 0; i < weekOptions.length; i++) {
    base -= repaymentData[i];
    cashflow.push(base);

    if (base < lowestWeek.value) {
      lowestWeek.value = base;
      lowestWeek.index = i;
      lowestWeek.label = weekOptions[i].label;
    }
  }

  lowestWeekCache = lowestWeek;

  document.getElementById('remaining').textContent = `Remaining: €${(355000 - totalRepayment).toLocaleString()}`;
  document.getElementById('totalRepaid').textContent = `€${totalRepayment.toLocaleString()}`;
  document.getElementById('finalBalance').textContent = `€${base.toLocaleString()}`;
  document.getElementById('lowestWeek').textContent = lowestWeek.label;

  renderChart(cashflow, repaymentData);
  renderTable(repaymentData);
}

function renderChart(cashflowData = null, repaymentData = null) {
  const ctx = document.getElementById('chartCanvas').getContext('2d');
  if (chart) chart.destroy();

  // Build dataset(s)
  let datasets = [{
    label: 'Cash Balance Forecast',
    data: cashflowData ? cashflowData : Array(weekOptions.length).fill(355000),
    borderColor: '#0077cc',
    backgroundColor: chartType === 'bar' ? '#b3d7f6' : 'rgba(0, 119, 204, 0.1)',
    borderWidth: 2,
    fill: chartType === 'line' || chartType === 'radar' ? true : false,
    pointRadius: 4,
    tension: 0.2
  }];

  // Optionally overlay repayments
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

  // Pie chart: show composition of final cash vs repaid
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
        // Annotation for lowest cash point
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
        y: { beginAtZero: false, title: { display: true, text: 'Cash (€)' } },
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

function renderTable(repaymentData = null) {
  // Remove old table if present
  const oldTable = document.getElementById('spreadsheetTable');
  if (oldTable) oldTable.innerHTML = "";

  if (rawData.length === 0) return;

  const table = document.createElement('table');
  table.style.width = '100%';
  table.style.borderCollapse = 'collapse';

  if (!repaymentData) {
    const tmp = getRepaymentData();
    repaymentData = tmp.repaymentData;
  }

  rawData.forEach((row, rowIndex) => {
    const tr = document.createElement('tr');
    row.forEach((cell, cellIndex) => {
      const td = document.createElement('td');
      td.style.border = '1px solid #ccc';
      td.style.padding = '4px 6px';

      // Make all cells editable except header row (first 4 rows typically header)
      if (rowIndex >= 4) {
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

      // Highlight cells in week columns with repayment
      if (rowIndex === 3 && repaymentData[cellIndex] > 0) {
        td.classList.add('repayment-highlight');
      }

      tr.appendChild(td);
    });
    table.appendChild(tr);
  });

  // Add a row to display repayment amounts per week (if any)
  if (weekOptions.length && repaymentData.some(val => val > 0)) {
    const repayTr = document.createElement('tr');
    rawData[3].forEach((_, i) => {
      const td = document.createElement('td');
      td.style.background = '#ffe9d2';
      td.style.fontWeight = 'bold';
      td.style.textAlign = 'center';
      td.textContent = repaymentData[i] > 0 ? `Repay: €${repaymentData[i]}` : '';
      repayTr.appendChild(td);
    });
    table.appendChild(repayTr);
  }

  document.getElementById('spreadsheetTable').appendChild(table);
}

// -------- Export/Save/Load Functions ----------

function exportToExcel() {
  // Export current table (rawData) as Excel
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
    // Draw image at full width, maintain aspect ratio
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