let workbook, sheetData, weekLabels = [], rollingCash = [], originalCash = [], chart;

document.getElementById('excelFile').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const data = await file.arrayBuffer();
  workbook = XLSX.read(data);
  sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
  initializeData();
  renderTable();
  renderChart();
});

function initializeData() {
  weekLabels = sheetData[3].slice(3);  // row 4, starting from column D
  const rollingRow = sheetData.find(row => row[0]?.toString().toLowerCase().includes('rolling cash balance'));
  rollingCash = rollingRow.slice(3).map(Number);
  originalCash = [...rollingCash];
}

document.getElementById('addRepayment').addEventListener('click', () => {
  const container = document.getElementById('repaymentInputs');
  const row = document.createElement('div');

  const select = document.createElement('select');
  weekLabels.forEach((label, i) => {
    const option = document.createElement('option');
    option.value = i;
    option.textContent = label;
    select.appendChild(option);
  });

  const input = document.createElement('input');
  input.type = 'number';
  input.placeholder = '€ Repayment';

  row.appendChild(select);
  row.appendChild(input);
  container.appendChild(row);
});

function applyRepayments() {
  // 1. Recompute rolling balances
  const rollingBalances = computeRollingBalancesWithRepayments();
  if (!rollingBalances || rollingBalances.length === 0) return;

  // 2. Recompute summaries
  const totalRepaid = repayments.reduce((a, b) => a + b.amount, 0);
  const remaining = 355000 - totalRepaid;
  const finalBalance = rollingBalances[rollingBalances.length - 1];
  const lowestWeekIndex = rollingBalances.indexOf(Math.min(...rollingBalances));
  const lowestWeekLabel = weekLabels[lowestWeekIndex] || '';

  // 3. Update UI
  document.getElementById('totalRepaid').innerText = formatEuro(totalRepaid);
  document.getElementById('remaining').innerText = formatEuro(remaining);
  document.getElementById('finalBalance').innerText = formatEuro(finalBalance);
  document.getElementById('lowestWeek').innerText = lowestWeekLabel;

  // 4. Update chart—but only if chart exists
  if (chart && chart.data && chart.data.datasets) {
    chart.data.datasets[0].data = rollingBalances;
    chart.update();
  }
}

  rows.forEach(row => {
    const week = parseInt(row.children[0].value);
    const amount = parseFloat(row.children[1].value) || 0;
    rollingCash[week] -= amount;
    totalRepaid += amount;
  });

  // Recalculate cumulative
  for (let i = 1; i < rollingCash.length; i++) {
    rollingCash[i] += rollingCash[i - 1] - originalCash[i - 1];
  }

  updateSummary(totalRepaid);
  updateChart();
});

function updateSummary(total) {
  document.getElementById('totalRepaid').textContent = `€${total.toLocaleString()}`;
  document.getElementById('remaining').textContent = `€${(355000 - total).toLocaleString()}`;
  document.getElementById('finalBalance').textContent = `€${rollingCash[rollingCash.length - 1].toLocaleString()}`;
  const min = Math.min(...rollingCash);
  document.getElementById('lowestWeek').textContent = weekLabels[rollingCash.indexOf(min)];
}

function renderChart() {
  const ctx = document.getElementById('chartCanvas').getContext('2d');
  chart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: weekLabels,
      datasets: [{
        label: 'Rolling Cash Balance',
        data: rollingCash,
        borderColor: 'blue',
        backgroundColor: 'rgba(0,0,255,0.1)',
        fill: true,
        tension: 0.3
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false
    }
  });
}

function updateChart() {
  chart.data.datasets[0].data = rollingCash;
  chart.update();
}

document.getElementById('toggleTable').addEventListener('click', () => {
  document.getElementById('tableContainer').classList.toggle('hidden');
});

function renderTable() {
  const container = document.getElementById('tableContainer');
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');

  const header = ['Label', ...weekLabels];
  const headerRow = document.createElement('tr');
  header.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const matchRows = sheetData.filter(r => typeof r[0] === 'string' && /(repayment|weekly income|rolling cash)/i.test(r[0]));

  matchRows.forEach(row => {
    const tr = document.createElement('tr');
    tr.appendChild(Object.assign(document.createElement('td'), { textContent: row[0] }));
    row.slice(3).forEach(c => {
      const td = document.createElement('td');
      td.textContent = c || '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  container.innerHTML = '';
  container.appendChild(table);
}
