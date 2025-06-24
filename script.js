let rawData = [];
let chart;
let weekOptions = [];

document.getElementById('fileInput').addEventListener('change', handleFile);
document.getElementById('addRepayment').addEventListener('click', addRepaymentRow);
document.getElementById('applyRepayments').addEventListener('click', applyRepayments);

function handleFile(event) {
  const reader = new FileReader();
  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    rawData = json;
    extractWeekOptions(json);
    renderInitialChart();
  };
  reader.readAsArrayBuffer(event.target.files[0]);
}

function extractWeekOptions(data) {
  const weeksRow = data[3] || [];
  const labelsRow = data[2] || [];

  weekOptions = weeksRow.map((weekLabel, i) => {
    const label = typeof weekLabel === 'string' ? weekLabel.trim() : '';
    if (label.toLowerCase().includes('week')) {
      return { index: i, label: label };
    }
    return null;
  }).filter(Boolean);
}

function addRepaymentRow() {
  const container = document.getElementById('repaymentInputs');
  const row = document.createElement('div');
  row.className = 'repayment-row';

  const weekSelect = document.createElement('select');
  weekOptions.forEach(week => {
    const option = document.createElement('option');
    option.value = week.index;
    option.textContent = week.label;
    weekSelect.appendChild(option);
  });

  const amountInput = document.createElement('input');
  amountInput.type = 'number';
  amountInput.placeholder = 'Amount €';

  row.appendChild(weekSelect);
  row.appendChild(amountInput);
  container.appendChild(row);
}

function applyRepayments() {
  if (weekOptions.length === 0 || rawData.length === 0) return;

  let base = 355000;
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

  let cashflow = [];
  let lowestWeek = { value: Infinity, label: '' };

  for (let i = 0; i < weekOptions.length; i++) {
    base -= repaymentData[i];
    cashflow.push(base);

    if (base < lowestWeek.value) {
      lowestWeek.value = base;
      lowestWeek.label = weekOptions[i].label;
    }
  }

  document.getElementById('remaining').textContent = `Remaining: €${base.toLocaleString()}`;
  document.getElementById('totalRepaid').textContent = `€${totalRepayment.toLocaleString()}`;
  document.getElementById('finalBalance').textContent = `€${base.toLocaleString()}`;
  document.getElementById('lowestWeek').textContent = lowestWeek.label;

  updateChart(cashflow);
}

function renderInitialChart() {
  const ctx = document.getElementById('chartCanvas').getContext('2d');
  if (chart) chart.destroy();

  chart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: weekOptions.map(w => w.label),
      datasets: [{
        label: 'Cash Balance Forecast',
        data: Array(weekOptions.length).fill(355000),
        borderColor: '#0077cc',
        backgroundColor: 'rgba(0, 119, 204, 0.1)',
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: true
        }
      },
      scales: {
        y: {
          beginAtZero: false
        }
      }
    }
  });
}

function updateChart(data) {
  if (!chart) return;
  chart.data.datasets[0].data = data;
  chart.update();
}
