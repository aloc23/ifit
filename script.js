
document.getElementById('file-input').addEventListener('change', handleFile, false);

function handleFile(e) {
  const file = e.target.files[0];
  const reader = new FileReader();
  reader.onload = function(e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    displayTable(json);
    updateChart(json);
  };
  reader.readAsArrayBuffer(file);
}

function displayTable(data) {
  const container = document.getElementById('excel-table');
  container.innerHTML = '';
  const table = document.createElement('table');
  data.forEach(row => {
    const tr = document.createElement('tr');
    row.forEach(cell => {
      const td = document.createElement('td');
      td.textContent = cell || '';
      tr.appendChild(td);
    });
    table.appendChild(tr);
  });
  container.appendChild(table);

  document.getElementById('toggle-table').addEventListener('click', () => {
    container.style.display = container.style.display === 'none' ? 'block' : 'none';
  });
}

function updateChart(data) {
  const ctx = document.getElementById('forecastChart').getContext('2d');
  if (!data[4]) return;

  const labels = data[3].slice(2); // Week labels
  const values = data[4].slice(2).map(v => parseFloat((v + '').replace(/[^0-9.-]+/g,"")) || 0);

  new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Mayweather Repayment Forecast',
        data: values,
        backgroundColor: 'blue'
      }]
    }
  });
}
