let chart = null;
let history = [];
let historyIndex = -1;

document.addEventListener('DOMContentLoaded', () => {
    const spreadsheet = document.getElementById('spreadsheet');
    const rows = 20;
    const cols = 10;

    for (let i = 0; i < rows; i++) {
        for (let j = 0; j < cols; j++) {
            const cell = document.createElement('div');
            cell.contentEditable = true;
            cell.classList.add('cell');
            cell.dataset.row = i;
            cell.dataset.col = j;

            // Data validation
            cell.addEventListener('input', (e) => {
                const value = e.target.textContent;
                if (isNaN(value)) {
                    e.target.style.backgroundColor = '#ffcccc'; // Highlight invalid input
                } else {
                    e.target.style.backgroundColor = ''; // Reset background
                }
                saveState();
            });

            // Formula evaluation
            cell.addEventListener('input', (e) => {
                const value = e.target.textContent;
                if (value.startsWith('=')) {
                    e.target.textContent = evaluateFormula(value);
                }
                saveState();
            });

            spreadsheet.appendChild(cell);
        }
    }
});

function saveState() {
    const state = [];
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
        state.push(cell.textContent);
    });
    history = history.slice(0, historyIndex + 1);
    history.push(state);
    historyIndex++;
}

function undo() {
    if (historyIndex > 0) {
        historyIndex--;
        const state = history[historyIndex];
        const cells = document.querySelectorAll('.cell');
        cells.forEach((cell, index) => {
            cell.textContent = state[index];
        });
    }
}

function redo() {
    if (historyIndex < history.length - 1) {
        historyIndex++;
        const state = history[historyIndex];
        const cells = document.querySelectorAll('.cell');
        cells.forEach((cell, index) => {
            cell.textContent = state[index];
        });
    }
}

function getCellValue(row, col) {
    const cell = document.querySelector(`.cell[data-row="${row}"][data-col="${col}"]`);
    return cell ? cell.textContent : '';
}

function setCellValue(row, col, value) {
    const cell = document.querySelector(`.cell[data-row="${row}"][data-col="${col}"]`);
    if (cell) {
        cell.textContent = value;
    }
}

function sum() {
    let total = 0;
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
        const value = parseFloat(cell.textContent);
        if (!isNaN(value)) {
            total += value;
        }
    });
    alert(`Sum: ${total}`);
}

function average() {
    let total = 0;
    let count = 0;
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
        const value = parseFloat(cell.textContent);
        if (!isNaN(value)) {
            total += value;
            count++;
        }
    });
    alert(`Average: ${count > 0 ? total / count : 0}`);
}

function max() {
    let maxVal = -Infinity;
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
        const value = parseFloat(cell.textContent);
        if (!isNaN(value) && value > maxVal) {
            maxVal = value;
        }
    });
    alert(`Max: ${maxVal}`);
}

function min() {
    let minVal = Infinity;
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
        const value = parseFloat(cell.textContent);
        if (!isNaN(value) && value < minVal) {
            minVal = value;
        }
    });
    alert(`Min: ${minVal}`);
}

function count() {
    let count = 0;
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
        const value = parseFloat(cell.textContent);
        if (!isNaN(value)) {
            count++;
        }
    });
    alert(`Count: ${count}`);
}

function trim() {
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
        cell.textContent = cell.textContent.trim();
    });
    saveState();
}

function toUpper() {
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
        cell.textContent = cell.textContent.toUpperCase();
    });
    saveState();
}

function toLower() {
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
        cell.textContent = cell.textContent.toLowerCase();
    });
    saveState();
}

function removeDuplicates() {
    const values = new Set();
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
        if (values.has(cell.textContent)) {
            cell.textContent = '';
        } else {
            values.add(cell.textContent);
        }
    });
    saveState();
}

function findAndReplace() {
    const find = prompt("Enter text to find:");
    const replace = prompt("Enter text to replace with:");
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
        cell.textContent = cell.textContent.replace(new RegExp(find, 'g'), replace);
    });
    saveState();
}

function createChart(type) {
    const labels = [];
    const data = [];
    const cells = document.querySelectorAll('.cell');

    // Extract data from the first column (A1, A2, A3, ...)
    for (let i = 0; i < 10; i++) { // Use first 10 rows for simplicity
        const cell = document.querySelector(`.cell[data-row="${i}"][data-col="0"]`);
        if (cell) {
            labels.push(`Row ${i + 1}`);
            data.push(parseFloat(cell.textContent) || 0);
        }
    }

    // Destroy existing chart if it exists
    if (chart) {
        chart.destroy();
    }

    // Create new chart
    const ctx = document.getElementById('chart').getContext('2d');
    chart = new Chart(ctx, {
        type: type,
        data: {
            labels: labels,
            datasets: [{
                label: 'Column A Data',
                data: data,
                backgroundColor: 'rgba(75, 192, 192, 0.2)',
                borderColor: 'rgba(75, 192, 192, 1)',
                borderWidth: 1
            }]
        },
        options: {
            scales: {
                y: {
                    beginAtZero: true
                }
            }
        }
    });
}

function save() {
    const data = [];
    const cells = document.querySelectorAll('.cell');
    cells.forEach(cell => {
        data.push(cell.textContent);
    });
    localStorage.setItem('spreadsheetData', JSON.stringify(data));
    alert('Spreadsheet saved!');
}

function load() {
    const data = JSON.parse(localStorage.getItem('spreadsheetData'));
    if (data) {
        const cells = document.querySelectorAll('.cell');
        cells.forEach((cell, index) => {
            cell.textContent = data[index];
        });
        alert('Spreadsheet loaded!');
    } else {
        alert('No saved data found!');
    }
}

function exportCSV() {
    let csvContent = "data:text/csv;charset=utf-8,";
    const cells = document.querySelectorAll('.cell');
    cells.forEach((cell, index) => {
        csvContent += cell.textContent + (index % 10 === 9 ? "\n" : ",");
    });
    const encodedUri = encodeURI(csvContent);
    const link = document.createElement("a");
    link.setAttribute("href", encodedUri);
    link.setAttribute("download", "spreadsheet.csv");
    document.body.appendChild(link);
    link.click();
}

function importCSV() {
    const fileInput = document.getElementById('csvFileInput');
    fileInput.click();
    fileInput.onchange = (e) => {
        const file = e.target.files[0];
        const reader = new FileReader();
        reader.onload = (e) => {
            const text = e.target.result;
            const rows = text.split("\n");
            const cells = document.querySelectorAll('.cell');
            rows.forEach((row, i) => {
                const cols = row.split(",");
                cols.forEach((col, j) => {
                    const index = i * 10 + j;
                    if (cells[index]) {
                        cells[index].textContent = col;
                    }
                });
            });
        };
        reader.readAsText(file);
    };
}

function evaluateFormula(formula) {
    if (formula.startsWith('=')) {
        formula = formula.slice(1); // Remove the '=' sign
        const cells = formula.split(/[\+\-\*\/]/); // Split by operators
        const operators = formula.split(/[0-9A-Z\$]+/).filter(op => op); // Extract operators

        let result = parseCellValue(cells[0]); // Parse the first cell value
        for (let i = 1; i < cells.length; i++) {
            const operator = operators[i - 1];
            const value = parseCellValue(cells[i]);
            switch (operator) {
                case '+': result += value; break;
                case '-': result -= value; break;
                case '*': result *= value; break;
                case '/': result /= value; break;
                default: break;
            }
        }
        return result;
    }
    return formula; // Return the original value if not a formula
}

function parseCellValue(cellRef) {
    const col = cellRef.replace(/\$/, '').charCodeAt(0) - 65; // Convert A-Z to 0-25
    const row = parseInt(cellRef.replace(/\$/, '').slice(1)) - 1; // Convert 1-based to 0-based
    const cell = document.querySelector(`.cell[data-row="${row}"][data-col="${col}"]`);
    return cell ? parseFloat(cell.textContent) || 0 : 0;
}

// Add Row
function addRow() {
    const spreadsheet = document.getElementById('spreadsheet');
    const cols = spreadsheet.children.length > 0 ? spreadsheet.children[0].children.length : 10;
    const newRow = document.createElement('div');
    newRow.style.display = 'contents';

    for (let j = 0; j < cols; j++) {
        const cell = document.createElement('div');
        cell.contentEditable = true;
        cell.classList.add('cell');
        cell.dataset.row = spreadsheet.children.length;
        cell.dataset.col = j;

        cell.addEventListener('input', (e) => {
            const value = e.target.textContent;
            if (isNaN(value)) {
                e.target.style.backgroundColor = '#ffcccc'; // Highlight invalid input
            } else {
                e.target.style.backgroundColor = ''; // Reset background
            }
            saveState();
        });

        cell.addEventListener('input', (e) => {
            const value = e.target.textContent;
            if (value.startsWith('=')) {
                e.target.textContent = evaluateFormula(value);
            }
            saveState();
        });

        newRow.appendChild(cell);
    }

    spreadsheet.appendChild(newRow);
    saveState();
}

// Add Column
function addColumn() {
    const spreadsheet = document.getElementById('spreadsheet');
    const rows = spreadsheet.children.length;

    for (let i = 0; i < rows; i++) {
        const cell = document.createElement('div');
        cell.contentEditable = true;
        cell.classList.add('cell');
        cell.dataset.row = i;
        cell.dataset.col = spreadsheet.children[i].children.length;

        cell.addEventListener('input', (e) => {
            const value = e.target.textContent;
            if (isNaN(value)) {
                e.target.style.backgroundColor = '#ffcccc'; // Highlight invalid input
            } else {
                e.target.style.backgroundColor = ''; // Reset background
            }
            saveState();
        });

        cell.addEventListener('input', (e) => {
            const value = e.target.textContent;
            if (value.startsWith('=')) {
                e.target.textContent = evaluateFormula(value);
            }
            saveState();
        });

        spreadsheet.children[i].appendChild(cell);
    }
    saveState();
}

// Delete Row
function deleteRow() {
    const spreadsheet = document.getElementById('spreadsheet');
    if (spreadsheet.children.length > 1) {
        spreadsheet.removeChild(spreadsheet.lastChild);
        saveState();
    } else {
        alert("Cannot delete the last row.");
    }
}

// Delete Column
function deleteColumn() {
    const spreadsheet = document.getElementById('spreadsheet');
    const rows = spreadsheet.children.length;

    if (spreadsheet.children[0].children.length > 1) {
        for (let i = 0; i < rows; i++) {
            spreadsheet.children[i].removeChild(spreadsheet.children[i].lastChild);
        }
        saveState();
    } else {
        alert("Cannot delete the last column.");
    }
}

// Toggle Bold
function toggleBold() {
    const selectedCells = document.querySelectorAll('.cell:focus');
    selectedCells.forEach(cell => {
        cell.style.fontWeight = cell.style.fontWeight === 'bold' ? 'normal' : 'bold';
    });
    saveState();
}

// Toggle Italic
function toggleItalic() {
    const selectedCells = document.querySelectorAll('.cell:focus');
    selectedCells.forEach(cell => {
        cell.style.fontStyle = cell.style.fontStyle === 'italic' ? 'normal' : 'italic';
    });
    saveState();
}

document.addEventListener('DOMContentLoaded', () => {
    const spreadsheet = document.getElementById('spreadsheet');
    const rows = 20, cols = 10;
    
    for (let i = 0; i < rows; i++) {
        const rowDiv = document.createElement('div');
        rowDiv.classList.add('row');
        for (let j = 0; j < cols; j++) {
            const cell = createCell(i, j);
            rowDiv.appendChild(cell);
        }
        spreadsheet.appendChild(rowDiv);
    }
});

function createCell(row, col) {
    const cell = document.createElement('div');
    cell.contentEditable = true;
    cell.classList.add('cell');
    cell.dataset.row = row;
    cell.dataset.col = col;

    cell.addEventListener('input', (e) => {
        const value = e.target.textContent;
        e.target.style.backgroundColor = isNaN(value) && !value.startsWith('=') ? '#ffcccc' : '';
        if (value.startsWith('=')) {
            e.target.textContent = evaluateFormula(value);
        }
        saveState();
    });

    return cell;
}

// Improved Evaluate Formula Function
function evaluateFormula(formula) {
    if (!formula.startsWith('=')) return formula;
    try {
        const expression = formula.slice(1).replace(/[A-Z]+\d+/g, match => {
            const value = parseCellValue(match);
            return isNaN(value) ? '0' : value; 
        });
        return eval(expression);
    } catch {
        return 'ERROR';
    }
}

function parseCellValue(ref) {
    const col = ref.charCodeAt(0) - 65; // Convert 'A' to 0, 'B' to 1...
    const row = parseInt(ref.slice(1)) - 1; // Convert 1-based to 0-based index
    const cell = document.querySelector(`.cell[data-row="${row}"][data-col="${col}"]`);
    return cell ? parseFloat(cell.textContent) || 0 : 0;
}

// Save & Load Enhancements
function save() {
    const data = [];
    document.querySelectorAll('.row').forEach(row => {
        const rowData = [];
        row.querySelectorAll('.cell').forEach(cell => rowData.push(cell.textContent));
        data.push(rowData);
    });
    localStorage.setItem('spreadsheetData', JSON.stringify(data));
    alert('Spreadsheet saved!');
}

function load() {
    const data = JSON.parse(localStorage.getItem('spreadsheetData'));
    if (data) {
        document.querySelectorAll('.row').forEach((row, i) => {
            row.querySelectorAll('.cell').forEach((cell, j) => {
                cell.textContent = data[i] && data[i][j] !== undefined ? data[i][j] : '';
            });
        });
        alert('Spreadsheet loaded!');
    } else {
        alert('No saved data found!');
    }
}
