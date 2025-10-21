document.getElementById('processBtn').addEventListener('click', async function() {
  const files = document.getElementById('excelFiles').files;
  const mode = document.getElementById('modeSelect').value;
  const output = document.getElementById('output');
  output.innerHTML = '';

  if (files.length === 0) {
    alert('Please upload at least one Excel file.');
    return;
  }

  // Helper function: read file and return rows
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

  // Read all uploaded files
  const allData = [];
  for (const file of files) {
    const rows = await readExcel(file);
    allData.push({ name: file.name, rows });
  }

  if (mode === 'combined') {
    // Combine all sheets’ data
    const combinedRows = allData.flatMap(item => item.rows);
    const table = createSummaryTable(combinedRows, 'Combined Summary');
    output.appendChild(table);
  } else {
    // Create a separate table for each file
    allData.forEach(item => {
      const table = createSummaryTable(item.rows, `Summary for ${item.name}`);
      output.appendChild(table);
    });
  }
});

// Helper: create summary table from rows
function createSummaryTable(rows, title) {
  const summary = {};
  rows.forEach(row => {
    const region = row['Region'];
    const status = row['Status'];
    if (!summary[region]) {
      summary[region] = { Total: 0, Completed: 0, Scheduled: 0, 'Not Completed': 0, Escalated: 0 };
    }
    summary[region].Total++;
    if (summary[region][status] !== undefined) {
      summary[region][status]++;
    }
  });

  const container = document.createElement('div');
  container.innerHTML = `<h3>${title}</h3>`;
  const table = document.createElement('table');
  table.border = 1;
  table.innerHTML = `
    <tr>
      <th>Region</th>
      <th>Total Tickets</th>
      <th>Completed</th>
      <th>Scheduled</th>
      <th>Not Completed</th>
      <th>Escalated</th>
    </tr>
  `;
  Object.keys(summary).forEach(region => {
    const s = summary[region];
    table.innerHTML += `
      <tr>
        <td>${region}</td>
        <td>${s.Total}</td>
        <td>${s.Completed}</td>
        <td>${s.Scheduled}</td>
        <td>${s['Not Completed']}</td>
        <td>${s.Escalated}</td>
      </tr>
    `;
  });

  container.appendChild(table);
  return container;
}

// Delete button logic
document.getElementById('deleteBtn').addEventListener('click', function () {
  const confirmDelete = confirm("Are you sure you want to delete all uploaded Excel files?");
  
  if (confirmDelete) {
    document.getElementById('excelFiles').value = '';  // Correct input ID
    document.getElementById('output').innerHTML = '';  // Correct container ID
    alert('All uploaded Excel files have been deleted. You can now upload new files.');
  } else {
    alert('Deletion cancelled.');
  }
});

document.getElementById('downloadBtn').addEventListener('click', function () {
  const tables = document.querySelectorAll('#output table'); // get all tables in output
  const wb = XLSX.utils.book_new();

  tables.forEach((table, index) => {
    const ws = XLSX.utils.table_to_sheet(table);
    XLSX.utils.book_append_sheet(wb, ws, `Sheet${index + 1}`);
  });

  XLSX.writeFile(wb, 'Ticket_Summary.xlsx'); // downloads file
});
