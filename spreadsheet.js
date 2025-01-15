let currentSheetIndex = 0;
const sheets = [];

function createNewSheet(name = `Sheet ${sheets.length + 1}`) {
    // Create a new sheet
    const table = document.createElement('table');
    table.id = `sheet-${sheets.length}`;
    table.classList.add('bg-white', 'rounded', 'shadow-lg');
    table.innerHTML = `
        <thead>
            <tr>
                <th>#</th>
                <th contenteditable="true">A</th>
                <th contenteditable="true">B</th>
                <th contenteditable="true">C</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td>1</td>
                <td contenteditable="true"></td>
                <td contenteditable="true"></td>
                <td contenteditable="true"></td>
            </tr>
        </tbody>
    `;
    sheets.push({ name, table });
    updateSheetTabs();
    switchToSheet(sheets.length - 1);

    // Add double-click functionality to rename columns
    const headerCells = table.querySelectorAll('th');
    headerCells.forEach((header, index) => {
        header.addEventListener('dblclick', () => renameColumn(index));
    });
}

function updateSheetTabs() {
    // Update the sheet tabs
    const tabs = document.getElementById('sheet-tabs');
    tabs.innerHTML = sheets
        .map(
            (sheet, index) => `
        <button
            class="px-4 py-2 rounded shadow ${
                index === currentSheetIndex ? 'bg-blue-500 text-white' : 'bg-gray-300 text-gray-800'
            }"
            onclick="switchToSheet(${index})"
        >
            ${sheet.name}
        </button>
    `
        )
        .join('');
}

function switchToSheet(index) {
    // Switch to the selected sheet
    const container = document.getElementById('spreadsheet-container');
    container.innerHTML = '';
    container.appendChild(sheets[index].table);
    currentSheetIndex = index;
    updateSheetTabs();
}

function renameColumn(colIndex) {
    const table = sheets[currentSheetIndex].table;
    const headerCell = table.querySelectorAll('th')[colIndex];
    
    // Create an input field for renaming
    const input = document.createElement('input');
    input.type = 'text';
    input.value = headerCell.textContent;
    input.classList.add('border', 'p-1', 'rounded');
    
    // Replace the column header with input field
    headerCell.innerHTML = '';
    headerCell.appendChild(input);
    
    // Focus the input field and listen for blur (when the user clicks away or presses Enter)
    input.focus();

    input.addEventListener('blur', () => {
        headerCell.textContent = input.value;
    });

    input.addEventListener('keypress', (e) => {
        if (e.key === 'Enter') {
            headerCell.textContent = input.value;
        }
    });
}

document.getElementById('add-sheet').addEventListener('click', () => {
    createNewSheet();
});

document.getElementById('add-row').addEventListener('click', () => {
    const table = sheets[currentSheetIndex].table;
    const tbody = table.querySelector('tbody');
    const newRow = document.createElement('tr');
    const cols = table.querySelector('thead tr').children.length;
    const rows = tbody.children.length;
    newRow.innerHTML = `<td>${rows + 1}</td>` + Array.from({ length: cols - 1 })
        .map(() => `<td contenteditable="true"></td>`)
        .join('');
    tbody.appendChild(newRow);
});

document.getElementById('add-col').addEventListener('click', () => {
    const table = sheets[currentSheetIndex].table;
    const headerRow = table.querySelector('thead tr');
    const colIndex = headerRow.children.length;
    const colLetter = String.fromCharCode(64 + colIndex);
    const newHeader = document.createElement('th');
    newHeader.textContent = colLetter;
    headerRow.appendChild(newHeader);
    table.querySelectorAll('tbody tr').forEach(row => {
        const newCell = document.createElement('td');
        newCell.contentEditable = true;
        row.appendChild(newCell);
    });
});

document.getElementById('export-xlsx').addEventListener('click', () => {
    const wb = XLSX.utils.book_new();
    sheets.forEach(sheet => {
        const rows = Array.from(sheet.table.querySelectorAll('tr')).map(row =>
            Array.from(row.children).map(cell => cell.textContent)
        );
        const ws = XLSX.utils.aoa_to_sheet(rows);
        XLSX.utils.book_append_sheet(wb, ws, sheet.name);
    });
    XLSX.writeFile(wb, 'spreadsheet.xlsx');
});

// Initialize the first sheet
createNewSheet('Sheet 1');