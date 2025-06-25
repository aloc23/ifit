let workbook, sheetData, weekLabels = [], rollingCash = [], originalCash = [], chart;
let repayments = [];

document.getElementById('excelFile').addEventListener('change', async (e) => {
  const file = e.target.files[0];
  const data = await file.arrayBuffer();
  workbook = XLSX.read(data);
  sheetData = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
  initializeData();
  renderTable();
  renderChart();
  populateWeekDropdowns();
});

function initializeData() {
  if (!sheetData || sheetData.length < 4) {
    console.error("❌ Excel data too short or not structured correctly.");
    return;
  }

  console.log("✅ Sheet loaded. Checking row 3 and 4:");
  console.log("Row 3:", sheetData[2]);
  console.log("Row 4 (week labels):", sheetData[3]);

  // Try to find a row with week labels — look for numeric values or "Week" strings
  const candidateRow = sheetData.find(row =>
    row.slice(3).some(cell => typeof cell === "string" && cell.toLowerCase().includes("week"))
  );

  if (!candidateRow) {
    console.error("❌ No row found with valid week labels.");
    return;
  }

  weekLabels = candidateRow.slice(3).map(label => label?.toString() || "");
  console.log("📅 Week labels set:", weekLabels);

  const rollingRow = sheetData.find(row =>
    row[0]?.toString().toLowerCase().includes("rolling cash balance")
  );

  if (!rollingRow) {
    console.error("❌ Could not find 'Rolling Cash Balance' row.");
    return;
  }

  rollingCash = rollingRow.slice(3).map(cell => parseFloat(cell) || 0);
  originalCash = [...rollingCash];
  console.log("💶 Loaded cash series:", rollingCash.length);
}

function populateWeekDropdowns() {
  document.getElementById('repaymentInputs').innerHTML = '';
  addRepaymentRow(); // start with one
}

function addRepaymentRow() {
  const container = document.getElementById('repaymentInputs');
  const row = document.createElement('div');

  const select = document.createElement('select');
  weekLabels.forEach((label, idx) => {
    const option = document.createElement('option');
    option.value = idx;
    option.text = `Week ${idx + 1} = ${label}`;
    select.appendChild(option);
  });

  const input = document.createElement('input');
  input.type = 'number';
  input.placeholder = '€ Repayment';

  row.appendChild(select);
  row.appendChild(input);
  container.appendChild(row);
}

document.getElementById('addRepayment').addEventListener('click', addRepaymentRow);

document.getElementById('applyRepayments').addEventListener('click', () => {
  if (!originalCash.length) return;

  const rows = document.querySelectorAll('#repaymentInputs > div');
  repayments = [];
  let cash = [...originalCash];

  rows.forEach(row => {
    const week = parseInt(row.children[0].value);
    const amount = parseFloat(row.children[1].value) || 0;
    if (!isNaN(amount) && week >= 0) {
      repayments.push({ week, amount });
      cash[week] = cash[week] - amount;
    }
  });

  for (let i = 1; i < cash.length; i++) {
    cash[i] = cash[i - 1] + (cash[i] - originalCash[i - 1]);
  }

  const totalRepaid = repayments.reduce((sum, r) => sum + r.amount, 0);
  const remaining = 355000 - totalRepaid;
  const finalBalance = cash[cash.length - 1] || 0;
  const lowest = Math.min(...cash);
  const lowestWeek = weekLabels[cash.indexOf(lowest)] || '–';

  document.getElementById('totalRepaid').textContent = `€${totalRepaid.toLocaleString()}`;
  document.getElementById('remaining').textContent = `€${remaining.toLocaleString()}`;
  document.getElementById('finalBalance').textContent = `€${finalBalance.toLocaleString()}`;
  document.getElementById('lowestWeek').textContent = lowestWeek;

  chart.data.datasets[0].data = cash;
  chart.update();
});

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

function renderTable() {
  const container = document.getElementById('tableContainer');
  const table = document.createElement('table');
  const thead = document.createElement('thead');
  const tbody = document.createElement('tbody');
  container.innerHTML = ''; // clear any old table

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

document.getElementById('toggleTable').addEventListener('click', () => {
  document.getElementById('tableContainer').classList.toggle('hidden');
  document.addEventListener('DOMContentLoaded', () => {
  if (sheetData?.length > 0) {
    initializeData();
    renderTable();
    renderChart();
    populateWeekDropdowns();
  }
});
});
