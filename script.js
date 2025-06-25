let workbook, sheetData, weekLabels = [], rollingCash = [], originalCash = [], chart;

document.getElementById('excelFile').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const data = await file.arrayBuffer();
  workbook = XLSX.read(data);
  parseSheet();
  renderTable();
  setupChart();
});

function parseSheet() {
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  sheetData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

  weekLabels = sheetData[3].slice(3).map((w, i) => `Week ${i + 1} = ${w}`);
  const rollingRow = sheetData.find(r => r[0]?.toString().toLowerCase().includes('rolling cash balance'));
  rollingCash = rollingRow.slice(3).map(c => parseFloat(c) || 0);
  originalCash = [...rollingCash];
}

document.getElementById('addRepayment').addEventListener('click', () => {
  const container = document.getElementById('repaymentsContainer');
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
  input.placeholder = 'Amount €';

  row.appendChild(select);
  row.appendChild(input);
  container.appendChild(row);
});

document.getElementById('applyRepayments').addEventListener('click', () => {
  const repaymentRows = document.querySelectorAll('#repaymentsContainer > div');
  rollingCash = [...originalCash];
  let totalRepaid = 0;

  repaymentRows.forEach(row => {
    const weekIndex = parseInt(row.children[0].value);
    const amount = parseFloat(row.children[1].value) || 0;
    rollingCash[weekIndex] -= amount;
    totalRepaid += amount;
  });

  // Recalculate cumulative cash
  for (let i = 1; i < rollingCash.length; i++) {
    rollingCash[i] += rollingCash[i - 1] - originalCash[i - 1];
  }

  updateChart();
  updateSummary(totalRepaid);
});

function updateSummary(totalRepaid) {
  const remaining = 355000 - totalRepaid;
  const final = rollingCash[rollingCash.length - 1];
  const min = Math.min(...rollingCash);
  const minIndex = rollingCash.indexOf(min);

  document.getElementById('totalRepaid').textContent = `€${totalRepaid.toLocaleString()}`;
  document.getElementById('remainingBalance').textContent = `€${remaining.toLocaleString()}`;
  document.getElementById('finalBalance').textContent = `€${final.toLocaleString()}`;
  document.getElementById('lowestWeek').textContent = weekLabels[minIndex] || '–';
}

function renderTable() {
  const container = document.getElementById('tableContainer');
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');

  const headers = ['Row Label', ...weekLabels];
  const headerRow = document.createElement('tr');
  headers.forEach(h => {
    const th = document.createElement('th');
    th.textContent = h;
    headerRow.appendChild(th);
  });
  thead.appendChild(headerRow);

  const matchRows = sheetData.filter(row =>
    typeof row[0] === 'string' &&
    /(repayment|rolling cash balance|weekly income)/i.test(row[0])
  );

  matchRows.forEach(row => {
    const tr = document.createElement('tr');
    const label = document.createElement('td');
    label.textContent = row[0];
    tr.appendChild(label);

    row.slice(3).forEach(cell => {
      const td = document.createElement('td');
      td.textContent = cell || '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  table.appendChild(thead);
  table.appendChild(tbody);
  container.innerHTML = '';
  container.appendChild(table);
}

function setupChart() {
  const ctx = document.getElementById('chartCanvas').getContext('2d');
  chart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: weekLabels,
      datasets: [{
        label: 'Rolling Cash Balance',
        data: rollingCash,
        borderColor: 'blue',
        fill: true,
        backgroundColor: 'rgba(0, 0, 255, 0.1)',
        tension: 0.3,
        pointRadius: 3
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

document.getElementById('toggleFullTable').addEventListener('click', () => {
  const container = document.getElementById('tableContainer');
  container.classList.toggle('hidden');
});
