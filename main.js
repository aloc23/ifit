let repaymentEntries = [];
let maxLoanRepayment = 355000;

function addRepaymentEntry() {
  const amount = parseFloat(document.getElementById("entryAmount").value);
  const month = document.getElementById("entryMonth").value;
  const week = document.getElementById("entryWeek").value;

  if (!amount || amount <= 0 || !month || week === "") {
    alert("Please enter a valid amount, month, and week.");
    return;
  }

  const currentTotal = repaymentEntries.reduce((sum, e) => sum + e.amount, 0);
  if (currentTotal + amount > maxLoanRepayment) {
    alert("Total repayment exceeds €355,000 limit.");
    return;
  }

  repaymentEntries.push({ amount, month, week });
  renderRepaymentList();
}

function renderRepaymentList() {
  const table = document.getElementById("repaymentList");
  table.innerHTML = "";

  repaymentEntries.forEach((entry, i) => {
    const row = document.createElement("tr");
    row.innerHTML = `
      <td>€${entry.amount.toLocaleString()}</td>
      <td>${entry.month}</td>
      <td>Week ${parseInt(entry.week) + 1}</td>
      <td><button onclick="removeRepaymentEntry(${i})">🗑️</button></td>
    `;
    table.appendChild(row);
  });
}

function removeRepaymentEntry(index) {
  repaymentEntries.splice(index, 1);
  renderRepaymentList();
}

function applyRepayment() {
  const table = document.querySelector("#table-container table");
  if (!table) {
    alert("Please upload or load the spreadsheet first.");
    return;
  }

  const headers = Array.from(table.querySelectorAll("thead th")).map(th => th.textContent.trim());
  const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
  ];

  const rows = Array.from(table.querySelectorAll("tbody tr"));
  rows.forEach(row => {
    const cells = row.querySelectorAll("td");

    repaymentEntries.forEach(({ amount, month, week }) => {
      for (let col = 1; col < headers.length; col++) {
        const dateStr = headers[col];
        const date = new Date(dateStr);
        if (isNaN(date)) continue;

        const headerMonth = monthNames[date.getMonth()];
        const weekOfMonth = Math.floor((date.getDate() - 1) / 7); // 0-indexed weeks

        if (month === headerMonth && parseInt(week) === weekOfMonth) {
          const cell = cells[col];
          const currentValue = parseFloat(cell.textContent.replace(/[€,]/g, "")) || 0;
          const newValue = currentValue - amount;
          cell.textContent = newValue.toFixed(2);
          cell.style.backgroundColor = "#fff3cd";
          cell.title = `Repayment applied: €${amount.toFixed(2)}`;
        }
      }
    });
  });
}