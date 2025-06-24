let weekLabels = [];
let repayments = [];

function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return alert("No file selected.");

    const reader = new FileReader();
    reader.onload = function (e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });

        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

        if (!json || json.length === 0) {
            alert("Error reading Excel file.");
            return;
        }

        // Example: Extract week labels from first row
        const headerRow = json[0];
        weekLabels = headerRow.slice(2).filter(Boolean); // Adjust start index if needed

        console.log("Extracted week labels:", weekLabels);
        if (weekLabels.length === 0) {
            alert("No week labels found in the file. Check column headers.");
        }

        // Update week dropdown in all rows
        populateWeekDropdowns();

        // Trigger rendering
        renderForecast(json);
    };
    reader.readAsArrayBuffer(file);
}

function populateWeekDropdowns() {
    document.querySelectorAll('.repayment-row select').forEach(select => {
        select.innerHTML = '';
        weekLabels.forEach(week => {
            const option = document.createElement('option');
            option.value = week;
            option.textContent = week;
            select.appendChild(option);
        });
    });
}

function addRepaymentRow() {
    const container = document.querySelector('.repayment-container');
    const row = document.createElement('div');
    row.className = 'repayment-row';

    const select = document.createElement('select');
    weekLabels.forEach(week => {
        const option = document.createElement('option');
        option.value = week;
        option.textContent = week;
        select.appendChild(option);
    });

    const input = document.createElement('input');
    input.type = 'number';
    input.placeholder = 'Amount €';

    row.appendChild(select);
    row.appendChild(input);
    container.appendChild(row);
}

function applyRepayments() {
    repayments = [];
    const rows = document.querySelectorAll('.repayment-row');
    rows.forEach(row => {
        const week = row.querySelector('select')?.value;
        const amount = parseFloat(row.querySelector('input')?.value || 0);
        if (week && amount > 0) {
            repayments.push({ week, amount });
        }
    });

    updateRepaymentSummary();
    renderForecast(); // Re-render with repayments
}

function updateRepaymentSummary() {
    const total = repayments.reduce((sum, r) => sum + r.amount, 0);
    document.getElementById('totalRepaid').textContent = `€${total.toLocaleString()}`;
}

function renderForecast(parsedData = null) {
    const ctx = document.getElementById('chartCanvas').getContext('2d');
    const labels = weekLabels;
    const balances = Array(labels.length).fill(0);

    repayments.forEach(r => {
        const idx = labels.indexOf(r.week);
        if (idx >= 0) balances[idx] -= r.amount;
    });

    new Chart(ctx, {
        type: 'line',
        data: {
            labels: labels,
            datasets: [{
                label: "Cash Balance Forecast",
                data: balances,
                borderColor: '#007bff',
                backgroundColor: 'rgba(0, 123, 255, 0.1)',
                fill: true,
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: true }
            },
            scales: {
                y: { beginAtZero: true }
            }
        }
    });

    const finalBalance = balances.reduce((a, b) => a + b, 0);
    document.getElementById('finalBalance').textContent = `€${finalBalance.toLocaleString()}`;
}

document.getElementById('fileInput').addEventListener('change', handleFile);
document.getElementById('addRepaymentBtn').addEventListener('click', addRepaymentRow);
document.getElementById('applyRepaymentsBtn').addEventListener('click', applyRepayments);
