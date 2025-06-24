let workbook, worksheet, data, chart;
let repayments = [];

document.getElementById('fileInput').addEventListener('change', handleFile);

function handleFile(event) {
    const reader = new FileReader();
    reader.onload = function (e) {
        const dataBinary = e.target.result;
        workbook = XLSX.read(dataBinary, { type: 'binary' });
        const sheetName = workbook.SheetNames[0];
        worksheet = workbook.Sheets[sheetName];

        const raw = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 3 }); // row 4
        const weeks = raw[0].slice(2); // week labels start from col index 2
        data = raw.slice(1); // actual data begins from row 5

        populateWeekSelect(weeks);
        renderChart(weeks, data);
        updateSummary(weeks, data);
    };
    reader.readAsBinaryString(event.target.files[0]);
}

function populateWeekSelect(weeks) {
    const container = document.getElementById('repaymentContainer');
    container.innerHTML = ''; // clear old

    addRepaymentRow(weeks);
}

function addRepaymentRow(weeks) {
    const container = document.getElementById('repaymentContainer');

    const rowDiv = document.createElement('div');
    rowDiv.className = 'repayment-row';

    const select = document.createElement('select');
    weeks.forEach((week, i) => {
        const option = document.createElement('option');
        option.value = i;
        option.textContent = week;
        select.appendChild(option);
    });

    const input = document.createElement('input');
    input.type = 'number';
    input.placeholder = 'Amount €';

    rowDiv.appendChild(select);
    rowDiv.appendChild(input);
    container.appendChild(rowDiv);
}

function applyRepayments() {
    repayments = [];
    const rows = document.querySelectorAll('.repayment-row');

    rows.forEach(row => {
        const select = row.querySelector('select');
        const input = row.querySelector('input');
        const weekIndex = parseInt(select.value);
        const amount = parseFloat(input.value);

        if (!isNaN(amount)) {
            repayments.push({ weekIndex, amount });
        }
    });

    updateSummary();
    renderChart();
}

function updateSummary(weeks, sheetData) {
    const totals = sheetData.map(row => {
        return row.slice(2).reduce((a, b) => a + (parseFloat(b) || 0), 0);
    });

    const baseTotal = totals.reduce((a, b) => a + b, 0);
    const totalRepayments = repayments.reduce((a, r) => a + r.amount, 0);

    const finalBalance = baseTotal - totalRepayments;
    const lowestWeek = getLowestWeek(totals);

    document.getElementById('totalRepaid').textContent = `€${totalRepayments.toLocaleString()}`;
    document.getElementById('finalBalance').textContent = `€${finalBalance.toLocaleString()}`;
    document.getElementById('lowestWeek').textContent = lowestWeek;
    document.getElementById('remaining').textContent = `Remaining: €${finalBalance.toLocaleString()}`;
}

function getLowestWeek(totals) {
    let min = Math.min(...totals);
    let index = totals.indexOf(min);
    return `Week ${index + 1}: Cashflow Forecast`;
}

function renderChart(weeks, sheetData) {
    if (!weeks) weeks = Array.from(document.querySelector('.repayment-row select').options).map(o => o.textContent);
    if (!sheetData) sheetData = data;

    const totals = sheetData.map(row => {
        return row.slice(2).reduce((a, b) => a + (parseFloat(b) || 0), 0);
    });

    repayments.forEach(rep => {
        if (rep.weekIndex < totals.length) {
            totals[rep.weekIndex] -= rep.amount;
        }
    });

    const ctx = document.getElementById('chartCanvas').getContext('2d');
    if (chart) chart.destroy();

    chart = new Chart(ctx, {
        type: 'line',
        data: {
            labels: weeks,
            datasets: [{
                label: 'Cash Balance Forecast',
                data: totals,
                fill: false,
                borderColor: 'blue',
                tension: 0.1,
                pointRadius: 4
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false
        }
    });
}

// Event wiring
document.getElementById('addRepaymentBtn').addEventListener('click', () => {
    const weeks = Array.from(document.querySelector('.repayment-row select').options).map(o => o.textContent);
    addRepaymentRow(weeks);
});

document.getElementById('applyRepaymentsBtn').addEventListener('click', applyRepayments);
