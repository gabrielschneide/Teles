function convertSerialToDate(serial) {
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
            populateImportedTable(jsonData);
        };
        reader.readAsArrayBuffer(file);
    }
}

function populateImportedTable(data) {
    const tableBody = document.querySelector('#importedTable tbody');
    tableBody.innerHTML = '';
    data.forEach((row, index) => {
        if (index === 0) return; // Ignora a primeira linha (cabeçalho)
        const tr = document.createElement('tr');
        row.forEach((cell, cellIndex) => {
            const td = document.createElement('td');
            if (cellIndex === 0) {
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

function addManualData() {
    const form = document.getElementById('manualForm');
    const data = new FormData(form);

    const tableBody = document.querySelector('#manualTable tbody');
    const tr = document.createElement('tr');

    const date = new Date(data.get('data'));
    const dateTd = document.createElement('td');
    dateTd.textContent = date.toLocaleDateString();
    tr.appendChild(dateTd);

    const tipoTd = document.createElement('td');
    tipoTd.textContent = data.get('tipo');
    tr.appendChild(tipoTd);

    const ordemTd = document.createElement('td');
    ordemTd.textContent = data.get('ordem');
    tr.appendChild(ordemTd);

    const atividadeTd = document.createElement('td');
    atividadeTd.textContent = data.get('atividade');
    tr.appendChild(atividadeTd);

    const wiseTd = document.createElement('td');
    wiseTd.textContent = data.get('wise');
    tr.appendChild(wiseTd);

    const mantenedorTd = document.createElement('td');
    mantenedorTd.textContent = data.get('mantenedor');
    tr.appendChild(mantenedorTd);

    const observacaoTd = document.createElement('td');
    observacaoTd.textContent = data.get('observacao');
    tr.appendChild(observacaoTd);

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

    form.reset();
}

function exportToPDF() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF('portrait', 'mm', 'a4');

    // Adicionar título
    doc.setFontSize(24);
    doc.text("Relatório de Atividades", 105, 20, null, null, 'center');

    // Adicionar data de exportação
    const exportDate = new Date().toLocaleDateString();
    doc.setFontSize(12);
    doc.text(`Data de Exportação: ${exportDate}`, 170, 30, null, null, 'right');

    // Adicionar cabeçalho da tabela
    doc.setFontSize(12);
    doc.text("Data", 20, 40);
    doc.text("Tipo", 50, 40);
    doc.text("Ordem", 80, 40);
    doc.text("Atividade", 110, 40);
    doc.text("Wise", 160, 40);
    doc.text("Mantenedor", 190, 40);
    doc.text("Observação", 230, 40);

    // Adicionar dados da tabela
    const table = document.getElementById('manualTable');
    const rows = table.querySelectorAll('tbody tr');
    let y = 50;

    rows.forEach((row) => {
        const cells = row.querySelectorAll('td');
        doc.text(cells[0].textContent, 20, y);
        doc.text(cells[1].textContent, 50, y);
        doc.text(cells[2].textContent, 80, y);
        doc.text(cells[3].textContent, 110, y);
        doc.text(cells[4].textContent, 160, y);
        doc.text(cells[5].textContent, 190, y);
        doc.text(cells[6].textContent, 230, y, { maxWidth: 150 });
        y += 10;

        // Verificar se precisa adicionar uma nova página
        if (y > 270) {
            doc.addPage();
            y = 20;
            // Adicionar cabeçalho da tabela na nova página
            doc.text("Data", 20, y);
            doc.text("Tipo", 50, y);
            doc.text("Ordem", 80, y);
            doc.text("Atividade", 110, y);
            doc.text("Wise", 160, y);
            doc.text("Mantenedor", 190, y);
            doc.text("Observação", 230, y);
            y += 10;
        }
    });

    doc.save('relatorio.pdf');
}
