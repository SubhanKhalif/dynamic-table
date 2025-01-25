document.addEventListener("DOMContentLoaded", function () {
    const spreadsheetContainer = document.getElementById("spreadsheet-container");
    spreadsheetContainer.style.position = "relative";
    spreadsheetContainer.style.height = "500px";
    spreadsheetContainer.style.overflow = "hidden";

    const undoStack = [];
    const redoStack = [];
    const maxStackSize = 100;

    function saveState() {
        const table = document.getElementById("spreadsheet");
        const state = {
            data: Array.from(table.querySelector("tbody").rows).map(row =>
                Array.from(row.cells).slice(1).map(cell => cell.textContent)
            ),
            headers: Array.from(table.querySelector("thead").rows[1].cells)
                .slice(1)
                .map(header => header.textContent),
            cellLockStates: JSON.parse(localStorage.getItem("cellLockStates")) || {}
        };

        undoStack.push(JSON.stringify(state));
        if (undoStack.length > maxStackSize) {
            undoStack.shift();
        }
        redoStack.length = 0;
    }

    const tooltip = document.createElement("div");
    tooltip.style.position = "absolute";
    tooltip.style.background = "#f9f9f9";
    tooltip.style.border = "1px solid #ccc";
    tooltip.style.padding = "5px";
    tooltip.style.borderRadius = "4px";
    tooltip.style.boxShadow = "0 2px 5px rgba(0,0,0,0.2)";
    tooltip.style.display = "none";
    tooltip.style.fontSize = "12px";
    tooltip.style.pointerEvents = "none";
    document.body.appendChild(tooltip);

    const contextMenu = document.createElement("div");
    contextMenu.style.position = "absolute";
    contextMenu.style.background = "#fff";
    contextMenu.style.border = "1px solid #ccc";
    contextMenu.style.padding = "5px";
    contextMenu.style.borderRadius = "4px";
    contextMenu.style.boxShadow = "0 2px 5px rgba(0,0,0,0.2)";
    contextMenu.style.display = "none";
    contextMenu.style.zIndex = "1000";
    document.body.appendChild(contextMenu);

    document.addEventListener('paste', (e) => {
        e.preventDefault();
        saveState();

        const table = document.getElementById("spreadsheet");
        const activeElement = document.activeElement;
        
        if (!activeElement || !activeElement.closest('td, th')) return;
        
        const cell = activeElement.closest('td, th');
        const row = cell.parentElement;
        const tbody = table.querySelector('tbody');
        
        const clipboardData = e.clipboardData.getData('text');
        const rows = clipboardData.split('\n').filter(row => row.trim());
        const maxCols = Math.max(...rows.map(row => row.split('\t').length));
        
        let startRow = Array.from(tbody.children).indexOf(row);
        let startCol = Array.from(row.children).indexOf(cell);

        const currentColCount = table.querySelector("thead tr").cells.length - 1;
        const neededCols = startCol + maxCols - currentColCount;
        
        for(let i = 0; i < neededCols; i++) {
            document.getElementById("add-col").click();
        }
        
        rows.forEach((rowData, rowIndex) => {
            const cells = rowData.split('\t');
            let currentRow = tbody.children[startRow + rowIndex];
            
            if (!currentRow) {
                currentRow = addNewRow();
            }
            
            cells.forEach((cellData, colIndex) => {
                const currentCell = currentRow.children[startCol + colIndex];
                if (currentCell) {
                    currentCell.textContent = cellData.trim();
                    if (cellData.trim() !== '') {
                        currentCell.contentEditable = false;
                        const cellKey = `${startRow + rowIndex}-${startCol + colIndex - 1}`;
                        const cellLockStates = JSON.parse(localStorage.getItem("cellLockStates")) || {};
                        cellLockStates[cellKey] = true;
                        localStorage.setItem("cellLockStates", JSON.stringify(cellLockStates));
                    }
                }
            });
        });
        
        saveData();
    });

    function initializeSpreadsheet() {
        const savedData = JSON.parse(localStorage.getItem("Tan")) || [[]];
        const cellLockStates = JSON.parse(localStorage.getItem("cellLockStates")) || {};
        const savedHeaders = JSON.parse(localStorage.getItem("columnHeaders")) || [];
        renderSpreadsheet(savedData, cellLockStates, savedHeaders);

        const undoButton = document.createElement("button");
        undoButton.id = "undu";
        undoButton.textContent = "Undu";
        undoButton.onclick = undo;
        document.getElementById("spreadsheet-controls").appendChild(undoButton);

        const redoButton = document.createElement("button");
        redoButton.id = "redu";
        redoButton.textContent = "Redu";
        redoButton.onclick = redo;
        document.getElementById("spreadsheet-controls").appendChild(redoButton);
    }

    function undo() {
        if (undoStack.length > 1) {
            redoStack.push(undoStack.pop());
            const previousState = JSON.parse(undoStack[undoStack.length - 1]);
            
            localStorage.setItem("Tan", JSON.stringify(previousState.data));
            localStorage.setItem("columnHeaders", JSON.stringify(previousState.headers));
            localStorage.setItem("cellLockStates", JSON.stringify(previousState.cellLockStates));
            
            renderSpreadsheet(previousState.data, previousState.cellLockStates, previousState.headers);
        }
    }

    function redo() {
        if (redoStack.length > 0) {
            const nextState = JSON.parse(redoStack.pop());
            undoStack.push(JSON.stringify(nextState));
            
            localStorage.setItem("Tan", JSON.stringify(nextState.data));
            localStorage.setItem("columnHeaders", JSON.stringify(nextState.headers));
            localStorage.setItem("cellLockStates", JSON.stringify(nextState.cellLockStates));
            
            renderSpreadsheet(nextState.data, nextState.cellLockStates, nextState.headers);
        }
    }

    document.addEventListener('keydown', (e) => {
        if (e.ctrlKey || e.metaKey) {
            if (e.key === 'z') {
                e.preventDefault();
                if (e.shiftKey) {
                    redo();
                } else {
                    undo();
                }
            } else if (e.key === 'y') {
                e.preventDefault();
                redo();
            }
        }
    });

    function renderSpreadsheet(data, cellLockStates, savedHeaders) {
        spreadsheetContainer.innerHTML = "";

        const wrapper = document.createElement("div");
        wrapper.style.position = "relative";
        wrapper.style.height = "100%";

        const table = document.createElement("table");
        table.id = "spreadsheet";
        table.style.position = "relative";

        const rowCount = data.length || 1;
        const colCount = data[0]?.length || 1;

        const headerContainer = document.createElement("div");
        headerContainer.style.position = "sticky";
        headerContainer.style.top = "0";
        headerContainer.style.zIndex = "2";
        headerContainer.style.backgroundColor = "#fff";

        const thead = document.createElement("thead");
        
        const headerRow1 = document.createElement("tr");
        const topLeftCorner1 = document.createElement("th");
        topLeftCorner1.style.backgroundColor = "#e0e0e0";
        topLeftCorner1.style.position = "sticky";
        topLeftCorner1.style.left = "0";
        topLeftCorner1.style.zIndex = "3";
        headerRow1.appendChild(topLeftCorner1);

        for (let i = 0; i < colCount; i++) {
            const th = document.createElement("th");
            th.textContent = i + 1;
            th.style.backgroundColor = "#e0e0e0";
            th.style.position = "sticky";
            th.style.top = "0";
            th.style.zIndex = "2";
            th.style.fontWeight = "bold";
            headerRow1.appendChild(th);
        }
        thead.appendChild(headerRow1);
        
        const headerRow2 = document.createElement("tr");
        const topLeftCorner2 = document.createElement("th");
        topLeftCorner2.style.backgroundColor = "#e0e0e0";
        topLeftCorner2.style.position = "sticky";
        topLeftCorner2.style.left = "0";
        topLeftCorner2.style.zIndex = "3";
        headerRow2.appendChild(topLeftCorner2);

        for (let i = 0; i < colCount; i++) {
            const th = document.createElement("th");
            th.contentEditable = true;
            th.textContent = savedHeaders[i] || '';
            th.style.backgroundColor = "#e0e0e0";
            th.style.position = "sticky";
            th.style.top = "0";
            th.style.zIndex = "2";
            th.style.fontWeight = "bold";

            th.addEventListener('input', () => {
                saveState();
                const headers = Array.from(headerRow2.children)
                    .slice(1)
                    .map(header => header.textContent);
                localStorage.setItem("columnHeaders", JSON.stringify(headers));
            });

            headerRow2.appendChild(th);
        }
        thead.appendChild(headerRow2);
        
        table.appendChild(thead);

        const scrollContainer = document.createElement("div");
        scrollContainer.style.height = "calc(100% - 80px)";
        scrollContainer.style.overflowY = "auto";
        scrollContainer.style.overflowX = "auto";

        const tbody = document.createElement("tbody");
        data.forEach((row, rowIndex) => {
            const tr = document.createElement("tr");

            const rowHeader = document.createElement("th");
            rowHeader.textContent = rowIndex + 1;
            rowHeader.style.backgroundColor = "#e0e0e0";
            rowHeader.style.position = "sticky";
            rowHeader.style.left = "0";
            rowHeader.style.zIndex = "1";
            tr.appendChild(rowHeader);

            row.forEach((cell, colIndex) => {
                const td = document.createElement("td");
                const cellKey = `${rowIndex}-${colIndex}`;
                td.contentEditable = !cell;
                td.textContent = cell;

                td.addEventListener("mouseover", (e) => {
                    tooltip.style.display = "block";
                    tooltip.style.left = `${e.pageX + 10}px`;
                    tooltip.style.top = `${e.pageY + 10}px`;
                    tooltip.textContent = `Row: ${rowIndex + 1}, Column: ${colIndex + 1}`;
                });

                td.addEventListener("mouseout", () => {
                    tooltip.style.display = "none";
                });

                let previousContent = td.textContent;
                td.addEventListener("blur", () => {
                    if (td.textContent.trim() !== previousContent) {
                        saveState();
                        td.contentEditable = false;
                        cellLockStates[cellKey] = true;
                        localStorage.setItem("cellLockStates", JSON.stringify(cellLockStates));
                        saveData();
                    }
                });

                td.addEventListener("contextmenu", (e) => {
                    e.preventDefault();
                    contextMenu.innerHTML = "";
                    const unlockButton = document.createElement("button");
                    unlockButton.textContent = "Unlock Cell";
                    unlockButton.style.padding = "5px 10px";
                    unlockButton.style.cursor = "pointer";
                    unlockButton.onclick = () => {
                        saveState();
                        td.contentEditable = true;
                        const existingData = td.textContent;
                        cellLockStates[cellKey] = false;
                        localStorage.setItem("cellLockStates", JSON.stringify(cellLockStates));
                        td.focus();
                        contextMenu.style.display = "none";
                    };
                    contextMenu.appendChild(unlockButton);
                    contextMenu.style.display = "block";
                    contextMenu.style.left = `${e.pageX}px`;
                    contextMenu.style.top = `${e.pageY}px`;
                });

                tr.appendChild(td);
            });
            tbody.appendChild(tr);
        });
        table.appendChild(tbody);

        scrollContainer.appendChild(table);
        wrapper.appendChild(headerContainer);
        wrapper.appendChild(scrollContainer);
        spreadsheetContainer.appendChild(wrapper);
        
        updateRowAndColumnCount();

        document.addEventListener("click", (e) => {
            if (!e.target.closest("#spreadsheet")) {
                contextMenu.style.display = "none";
            }
        });

        if (undoStack.length === 0) {
            saveState();
        }
    }

    function addNewRow() {
        saveState();
        const table = document.getElementById("spreadsheet");
        const tbody = table.querySelector("tbody");
        const newRow = tbody.insertRow();
        const colCount = table.querySelector("thead tr").cells.length - 1;
        const cellLockStates = JSON.parse(localStorage.getItem("cellLockStates")) || {};

        const rowHeader = document.createElement("th");
        rowHeader.textContent = tbody.rows.length;
        rowHeader.style.backgroundColor = "#e0e0e0";
        rowHeader.style.position = "sticky";
        rowHeader.style.left = "0";
        rowHeader.style.zIndex = "1";
        newRow.appendChild(rowHeader);

        for (let i = 0; i < colCount; i++) {
            const newCell = newRow.insertCell();
            const cellKey = `${tbody.rows.length-1}-${i}`;
            newCell.contentEditable = true;

            newCell.addEventListener("mouseover", (e) => {
                tooltip.style.display = "block";
                tooltip.style.left = `${e.pageX + 10}px`;
                tooltip.style.top = `${e.pageY + 10}px`;
                tooltip.textContent = `Row: ${newRow.rowIndex}, Column: ${i + 1}`;
            });

            newCell.addEventListener("mouseout", () => {
                tooltip.style.display = "none";
            });

            let previousContent = newCell.textContent;
            newCell.addEventListener("blur", () => {
                if (newCell.textContent.trim() !== previousContent) {
                    saveState();
                    newCell.contentEditable = false;
                    cellLockStates[cellKey] = true;
                    localStorage.setItem("cellLockStates", JSON.stringify(cellLockStates));
                    saveData();
                }
            });

            newCell.addEventListener("contextmenu", (e) => {
                e.preventDefault();
                contextMenu.innerHTML = "";
                const unlockButton = document.createElement("button");
                unlockButton.textContent = "Unlock Cell";
                unlockButton.style.padding = "5px 10px";
                unlockButton.style.cursor = "pointer";
                unlockButton.onclick = () => {
                    saveState();
                    newCell.contentEditable = true;
                    const existingData = newCell.textContent;
                    cellLockStates[cellKey] = false;
                    localStorage.setItem("cellLockStates", JSON.stringify(cellLockStates));
                    newCell.focus();
                    contextMenu.style.display = "none";
                };
                contextMenu.appendChild(unlockButton);
                contextMenu.style.display = "block";
                contextMenu.style.left = `${e.pageX}px`;
                contextMenu.style.top = `${e.pageY}px`;
            });
        }
        saveData();
        return newRow;
    }

    function saveData() {
        const table = document.getElementById("spreadsheet");
        const data = Array.from(table.querySelector("tbody").rows).map((row) =>
            Array.from(row.cells).slice(1).map((cell) => cell.textContent)
        );

        // Save data to localStorage with the key "ZP"
        localStorage.setItem("Tan", JSON.stringify(data));

        const headerRow = table.querySelector("thead").rows[1];
        const headers = Array.from(headerRow.cells)
            .slice(1)
            .map(header => header.textContent);
        localStorage.setItem("columnHeaders", JSON.stringify(headers));

        updateRowAndColumnCount();
    }

    function updateRowAndColumnCount() {
        const table = document.getElementById("spreadsheet");
        const rowCount = table.querySelector("tbody").rows.length || 0;
        const colCount = table.querySelector("thead tr").cells.length - 1 || 0;

        console.log(`Row Count: ${rowCount}, Column Count: ${colCount}`);
        updateHeaders();
    }

    function updateHeaders() {
        const table = document.getElementById("spreadsheet");

        const headerRows = table.querySelectorAll("thead tr");
        Array.from(headerRows[0].cells).forEach((cell, index) => {
            if (index > 0) {
                cell.textContent = index;
                cell.style.backgroundColor = "#e0e0e0";
                cell.style.fontWeight = "bold";
            }
        });

        Array.from(table.querySelector("tbody").rows).forEach((row, index) => {
            const rowHeader = row.querySelector("th");
            if (rowHeader) {
                rowHeader.textContent = index + 1;
                rowHeader.style.backgroundColor = "#e0e0e0";
            }
        });
    }

    document.getElementById("add-row").addEventListener("click", addNewRow);

    document.getElementById("add-col").addEventListener("click", () => {
        saveState();
        const table = document.getElementById("spreadsheet");
        const colCount = table.querySelector("thead tr").cells.length;
        const cellLockStates = JSON.parse(localStorage.getItem("cellLockStates")) || {};

        const headerRows = table.querySelectorAll("thead tr");
        
        const newColHeader1 = document.createElement("th");
        newColHeader1.textContent = colCount;
        newColHeader1.style.backgroundColor = "#e0e0e0";
        newColHeader1.style.position = "sticky";
        newColHeader1.style.top = "0";
        newColHeader1.style.zIndex = "2";
        newColHeader1.style.fontWeight = "bold";
        headerRows[0].appendChild(newColHeader1);

        const newColHeader2 = document.createElement("th");
        newColHeader2.contentEditable = true;
        newColHeader2.style.backgroundColor = "#e0e0e0";
        newColHeader2.style.position = "sticky";
        newColHeader2.style.top = "0";
        newColHeader2.style.zIndex = "2";
        newColHeader2.style.fontWeight = "bold";

        newColHeader2.addEventListener('input', () => {
            saveState();
            const headers = Array.from(headerRows[1].children)
                .slice(1)
                .map(header => header.textContent);
            localStorage.setItem("columnHeaders", JSON.stringify(headers));
        });

        headerRows[1].appendChild(newColHeader2);

        Array.from(table.querySelector("tbody").rows).forEach((row, rowIndex) => {
            const newCell = row.insertCell();
            const cellKey = `${rowIndex}-${colCount-1}`;
            newCell.contentEditable = true;

            newCell.addEventListener("mouseover", (e) => {
                tooltip.style.display = "block";
                tooltip.style.left = `${e.pageX + 10}px`;
                tooltip.style.top = `${e.pageY + 10}px`;
                tooltip.textContent = `Row: ${rowIndex + 1}, Column: ${colCount}`;
            });

            newCell.addEventListener("mouseout", () => {
                tooltip.style.display = "none";
            });

            let previousContent = newCell.textContent;
            newCell.addEventListener("blur", () => {
                if (newCell.textContent.trim() !== previousContent) {
                    saveState();
                    newCell.contentEditable = false;
                    cellLockStates[cellKey] = true;
                    localStorage.setItem("cellLockStates", JSON.stringify(cellLockStates));
                    saveData();
                }
            });

            newCell.addEventListener("contextmenu", (e) => {
                e.preventDefault();
                contextMenu.innerHTML = "";
                const unlockButton = document.createElement("button");
                unlockButton.textContent = "Unlock Cell";
                unlockButton.style.padding = "5px 10px";
                unlockButton.style.cursor = "pointer";
                unlockButton.onclick = () => {
                    saveState();
                    newCell.contentEditable = true;
                    const existingData = newCell.textContent;
                    cellLockStates[cellKey] = false;
                    localStorage.setItem("cellLockStates", JSON.stringify(cellLockStates));
                    newCell.focus();
                    contextMenu.style.display = "none";
                };
                contextMenu.appendChild(unlockButton);
                contextMenu.style.display = "block";
                contextMenu.style.left = `${e.pageX}px`;
                contextMenu.style.top = `${e.pageY}px`;
            });
        });
        saveData();
    });

    document.getElementById("delete-row").addEventListener("click", () => {
        const rowIndexes = prompt("Enter row numbers to delete (comma separated, starting from 1):");
        if (rowIndexes) {
            saveState();
            const table = document.getElementById("spreadsheet");
            const tbody = table.querySelector("tbody");

            const indexes = rowIndexes.split(',').map(num => parseInt(num.trim())).sort((a, b) => b - a);
            let invalidIndexes = [];

            indexes.forEach(index => {
                if (index > 0 && tbody.rows[index - 1]) {
                    tbody.deleteRow(index - 1);
                } else {
                    invalidIndexes.push(index);
                }
            });

            if (invalidIndexes.length > 0) {
                alert(`Invalid row numbers: ${invalidIndexes.join(', ')}`);
            }
            saveData();
        }
    });

    document.getElementById("delete-column").addEventListener("click", () => {
        const colIndexes = prompt("Enter column numbers to delete (comma separated, starting from 1):");
        if (colIndexes) {
            saveState();
            const table = document.getElementById("spreadsheet");

            const indexes = colIndexes.split(',').map(num => parseInt(num.trim())).sort((a, b) => b - a);
            let invalidIndexes = [];

            indexes.forEach(index => {
                if (index > 0) {
                    const headerRows = table.querySelectorAll("thead tr");
                    if (headerRows[0].cells[index]) {
                        headerRows[0].deleteCell(index);
                        headerRows[1].deleteCell(index);

                        Array.from(table.querySelector("tbody").rows).forEach((row) => {
                            if (row.cells[index]) row.deleteCell(index);
                        });
                    } else {
                        invalidIndexes.push(index);
                    }
                } else {
                    invalidIndexes.push(index);
                }
            });

            if (invalidIndexes.length > 0) {
                alert(`Invalid column numbers: ${invalidIndexes.join(', ')}`);
            }
            saveData();
        }
    });

    document.getElementById("export-xlsx").addEventListener("click", () => {
        const table = document.getElementById("spreadsheet");
        const data = Array.from(table.querySelector("tbody").rows).map((row) =>
            Array.from(row.cells).slice(1).map((cell) => cell.textContent)
        );
        const ws = XLSX.utils.aoa_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
        XLSX.writeFile(wb, "pwd.xlsx");
    });

    initializeSpreadsheet();
});
