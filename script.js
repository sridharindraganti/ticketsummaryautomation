document.getElementById('processBtn').addEventListener('click', async function() {
  const fileInput = document.getElementById('excelFiles');
  const output = document.getElementById('output');
  output.innerHTML = '';

  if (fileInput.files.length === 0) {
    alert('Please upload an Excel file.');
    return;
  }

  const file = fileInput.files[0];

  // Helper: read Excel file and return rows
  const readExcel = (file) => {
    return new Promise((resolve) => {
      const reader = new FileReader();
      reader.onload = (event) => {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const sheet = workbook.Sheets[workbook.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet);
        resolve(rows);
      };
      reader.readAsArrayBuffer(file);
    });
  };

  // Read uploaded file
  const rows = await readExcel(file);
  const table = createSummaryTable(rows, `Summary for ${file.name}`);
  output.appendChild(table);
});

// Helper: create summary table from rows
function createSummaryTable(rows, title) {
  const statusTypes = [
    'Scheduled', 'Successful', 'Unsuccessful', 'In Process', 'Others', 
    'Cancelled', 'DTE', 'QS', 'Pending', 'In CA by Friday', 
    'In CA Weekend', 'Rescheduled', 'SAP Queue', 'Future', 'Escalated'
  ];

  const summary = {};
  rows.forEach(row => {
    const region = row['Region'];
    const status = row['Status'];
    if (!summary[region]) {
      summary[region] = { Total: 0 };
      statusTypes.forEach(st => summary[region][st] = 0);
    }
    summary[region].Total++;
    if (summary[region][status] !== undefined) {
      summary[region][status]++;
    } else {
      summary[region]['Others']++; // fallback for unknown statuses
    }
  });

  const container = document.createElement('div');
  container.innerHTML = `<h3>${title}</h3>`;
  const table = document.createElement('table');
  table.border = 1;

  // Create table header
  let headerRow = '<tr><th>Region</th><th>Total Tickets</th>';
  statusTypes.forEach(st => headerRow += `<th>${st}</th>`);
  headerRow += '</tr>';
  table.innerHTML = headerRow;

  // Fill data
  Object.keys(summary).forEach(region => {
    const s = summary[region];
    let rowHTML = `<tr><td>${region}</td><td>${s.Total}</td>`;
    statusTypes.forEach(st => rowHTML += `<td>${s[st]}</td>`);
    rowHTML += '</tr>';
    table.innerHTML += rowHTML;
  });

  container.appendChild(table);
  return container;
}

// Delete button logic
document.getElementById('deleteBtn').addEventListener('click', function () {
  const confirmDelete = confirm("Are you sure you want to delete the uploaded Excel file?");
  if (confirmDelete) {
    document.getElementById('excelFiles').value = '';
    document.getElementById('output').innerHTML = '';
    alert('The uploaded Excel file has been deleted. You can now upload a new one.');
  } else {
    alert('Deletion cancelled.');
  }
});

// Download button logic
document.getElementById('downloadBtn').addEventListener('click', function () {
  const tables = document.querySelectorAll('#output table');
  if (tables.length === 0) {
    alert("No summary to download!");
    return;
  }

  const wb = XLSX.utils.book_new();
  tables.forEach((table, index) => {
    const ws = XLSX.utils.table_to_sheet(table);
    XLSX.utils.book_append_sheet(wb, ws, `Summary`);
  });

  XLSX.writeFile(wb, 'Ticket_Summary.xlsx');
});

// Theme toggle logic
document.getElementById('themeToggleBtn').addEventListener('click', function () {
  document.body.classList.toggle('dark-mode');

  const currentTheme = document.body.classList.contains('dark-mode') ? 'Dark' : 'Light';
  this.textContent = `Switch to ${currentTheme === 'Dark' ? 'Light' : 'Dark'} Mode`;
});
