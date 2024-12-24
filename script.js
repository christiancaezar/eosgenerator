function addEmptyRow() {
    let tbody = document.querySelector('#spreadsheet tbody');
    let newRow = tbody.insertRow();  // Insert a new row at the end of the table
    for (let i = 0; i < 5; i++) {  // Create 5 cells (one for each column)
        let cell = newRow.insertCell(i);
        cell.contentEditable = true;  // Make each cell editable
    }
}

document.getElementById('spreadsheet').addEventListener('paste', function(event) {
event.preventDefault();  // Prevent the default paste behavior

let clipboardData = event.clipboardData || window.clipboardData;
let pastedData = clipboardData.getData('text/plain');  // Get the text data from the clipboard

let rows = pastedData.split('\n');  // Split data into rows (Excel rows are separated by newlines)

let tbody = this.querySelector('tbody');
let currentRows = tbody.rows.length;

// Find the first empty row in the table
let firstEmptyRowIndex = -1;
for (let i = 0; i < currentRows; i++) {
    if (tbody.rows[i].cells[0].innerText === '') {
    firstEmptyRowIndex = i;
    break;
    }
}

// If no empty row exists, add a new one at the bottom
if (firstEmptyRowIndex === -1) {
    firstEmptyRowIndex = currentRows;
    addEmptyRow();  // Ensure there's at least one empty row
}

// Add more rows if necessary
if (rows.length + firstEmptyRowIndex > currentRows) {
    for (let i = currentRows; i < rows.length + firstEmptyRowIndex; i++) {
    let newRow = tbody.insertRow();
    for (let j = 0; j < 5; j++) {
        let cell = newRow.insertCell(j);
        cell.contentEditable = true;
    }
    }
}

// Fill the table with the pasted data starting from the first empty row
rows.forEach((row, rowIndex) => {
    let cols = row.split('\t');  // Split each row into columns (Excel cells are separated by tabs)
    let tableRow = tbody.rows[firstEmptyRowIndex + rowIndex];
    
    if (tableRow) {
    cols.forEach((col, colIndex) => {
        let cell = tableRow.cells[colIndex];
        if (cell) {
        cell.innerText = col;  // Place the clipboard data in the table cell
        }
    });
    }
});

// Always add an empty row after pasting the data
addEmptyRow();
});



