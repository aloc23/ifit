let weeklyIncome = [], baseBalance = 0, chart;
let repayRowIx = 0, incomeRowIx = 0, balanceRowIx = 0;
let repayments = [];

document.getElementById('fileInput').addEventListener('change', handleFile);
document.getElementById('addRepaymentBtn').addEventListener('click', addRepaymentRow);
document.getElementById('applyRepaymentsBtn').addEventListener('click', applyRepayments);
document.getElementById('toggleTableBtn').addEventListener('click', () => {
  const el = document.getElementById('tableWrapper');
  el.classList.toggle('hidden-scroll');
});

function handleFile(e) {
  const reader = new FileReader();
  reader.onload = function(e) {
    const workbook = XLSX.read(e.target.result, { type: 'binary' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const html = XLSX.utils.sheet_to_html(sheet);
    document.getElementById('tableWrapper').innerHTML = html;

    const rows = document.querySelectorAll('#tableWrapper table tr');
    const labelRow = Array.from(rows[3].cells).map(cell => cell.textContent);
    const incomeRow = Array.from(rows).find((row, i) => {
      const text = row.cells?.[0]?.textContent?.toLowerCase() || "";
      if (text.includes("weekly income")) incomeRowIx = i;
      if (text.includes("rolling cash")) balanceRowIx = i;
      if (text.includes("investment repayment")) repayRowIx = i;
      return false;
    });

    weeklyIncome = labelRow.map((_, i) => {
      const val = parseFloat(rows[incomeRowIx]?.cells[i]?.textContent.replace(/,/g, ''));
      return isNaN(val) ? 0 : val;
    });

    baseBalance = parseFloat(rows[balanceRowIx]?.cells[1]?.textContent.replace(/,/g, '')) || 0;
    renderChart([baseBalance]);
    renderSummary([baseBalance]);
  };
  reader.readAsBinaryString(e.target.files[0]);
}

function addRepaymentRow() {
  const div = document.createElement('div');
  const week = document.createElement('select');
  const val = document.createElement('input');
  val.type = 'number';
  val.placeholder = 'Amount €';

  for (let i = 0; i < weeklyIncome.length; i++) {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = `Week ${i + 1}`;
    week.appendChild(opt);
  }

  div.appendChild(week);
  div.appendChild(val);
  document.getElementById('repaymentContainer').appendChild(div);
}

function applyRepayments() {
  repayments = [];
  document.querySelectorAll('#repaymentContainer div').forEach(div => {
    const w = +div.children[0].value;
    const val = +div.children[1].value;
    if (!isNaN(w) && !isNaN(val)) repayments.push({ w, val });
  });

  const modified = [...weeklyIncome];
  repayments.forEach(({ w, val }) => {
    modified[w] -= val;
  });

  const bal = [baseBalance];
  for (let i = 1; i < modified.length; i++) {
    bal[i] = bal[i - 1] + (modified[i] || 0);
  }

  updateSpreadsheetRows(repayments);
  renderChart(bal);
  renderSummary(bal);
}

function updateSpreadsheetRows(reps) {
  const table = document.querySelector('#tableWrapper table');
  const rows = Array.from(table.querySelectorAll('tr'));
  reps.forEach(({ w, val }) => {
    const repayCell = rows[repayRowIx]?.cells[w + 1];
    const incomeCell = rows[incomeRowIx]?.cells[w + 1];
    if (repayCell) {
      const v = parseFloat(repayCell.textContent.replace(/,/g, '')) || 0;
      repayCell.textContent = (v - val).toLocaleString();
      repayCell.style.background = '#ffe6e6';
    }
    if (incomeCell) {
      const v = parseFloat(incomeCell.textContent.replace(/,/g, '')) || 0;
      incomeCell.textContent = (v - val).toLocaleString();
      incomeCell.style.background = '#e6f7ff';
    }
  });
}

function renderChart(bal) {
  const ctx = document.getElementById('chartCanvas').getContext('2d');
  if (chart) chart.destroy();
  chart = new Chart(ctx, {
    type: 'line',
    data: {
      labels: bal.map((_, i) => `Week ${i + 1}`),
      datasets: [{
        label: 'Cash Balance Forecast',
        data: bal,
        borderColor: 'blue',
        backgroundColor: 'rgba(0,0,255,0.05)',
        fill: true,
        tension: 0.3,
        pointRadius: 3
      }]
    },
    options: {
      responsive: true,
      scales: {
        y: { beginAtZero: false }
      }
    }
  });
}

function renderSummary(bal) {
  const total = repayments.reduce((acc, r) => acc + r.val, 0);
  document.getElementById('totalRepaid').textContent = `Total Repaid: €${total.toLocaleString()}`;
  document.getElementById('finalBalance').textContent = `Final Balance: €${bal[bal.length - 1].toLocaleString()}`;

  const minVal = Math.min(...bal);
  const minWeek = bal.indexOf(minVal) + 1;
  document.getElementById('lowestWeek').textContent = `Lowest Week: Week ${minWeek}`;
}
