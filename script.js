function convertSerialToDate(serial) {
    // Converte um número serial do Excel para uma data JavaScript
    const baseDate = new Date(1899, 11, 30); // Data base do Excel (30/12/1899)
    const date = new Date(baseDate.getTime() + serial * 24 * 60 * 60 * 1000);
    return date.toLocaleDateString(); // Formata a data para o formato local
}

function importData() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
            populateTable(jsonData);
        };
        reader.readAsArrayBuffer(file);
    }
}

function populateTable(data) {
    const tableBody = document.querySelector('#reportTable tbody');
    tableBody.innerHTML = '';
    data.forEach((row, index) => {
        if (index === 0) return; // Ignora a primeira linha (cabeçalho)
        const tr = document.createElement('tr');
        row.forEach((cell, cellIndex) => {
            const td = document.createElement('td');
            if (cellIndex === 0) {
                // Converte a primeira coluna (data)
                td.textContent = convertSerialToDate(cell);
            } else {
                td.textContent = cell;
            }
            tr.appendChild(td);
        });
        const checklistTd = document.createElement('td');
        const checklistInput = document.createElement('input');
        checklistInput.type = 'checkbox';
        checklistTd.appendChild(checklistInput);
        tr.appendChild(checklistTd);

        const actionsTd = document.createElement('td');
        const editButton = document.createElement('button');
        editButton.textContent = 'Editar';
        editButton.onclick = () => editRow(tr);
        const deleteButton = document.createElement('button');
        deleteButton.textContent = 'Excluir';
        deleteButton.onclick = () => deleteRow(tr);
        actionsTd.appendChild(editButton);
        actionsTd.appendChild(deleteButton);
        tr.appendChild(actionsTd);

        tableBody.appendChild(tr);
    });
}

function editRow(row) {
    const cells = row.cells;
    for (let i = 0; i < cells.length - 2; i++) {
        const cell = cells[i];
        const input = document.createElement('input');
        input.type = 'text';
        input.value = cell.textContent;
        cell.textContent = '';
        cell.appendChild(input);
    }
    const saveButton = document.createElement('button');
    saveButton.textContent = 'Salvar';
    saveButton.onclick = () => saveRow(row);
    cells[cells.length - 1].appendChild(saveButton);
}

function saveRow(row) {
    const cells = row.cells;
    for (let i = 0; i < cells.length - 2; i++) {
        const cell = cells[i];
        const input = cell.querySelector('input');
        cell.textContent = input.value;
    }
    const saveButton = cells[cells.length - 1].querySelector('button');
    cells[cells.length - 1].removeChild(saveButton);
}

function deleteRow(row) {
    row.remove();
}
