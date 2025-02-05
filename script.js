// Configurações globais
const PDF_CONFIG = {
    pageSize: 'A4',
    orientation: 'portrait',
    margin: 20,
    headerHeight: 40,
    lineHeight: 8,
    fontSizes: {
        title: 18,
        subtitle: 14,
        header: 12,
        body: 10
    },
    colors: {
        primary: '#2c3e50',
        secondary: '#3498db',
        text: '#34495e',
        border: '#ecf0f1'
    }
};

// Função principal de exportação
function exportToPDF() {
    const { jsPDF } = window.jspdf;
    const doc = new jsPDF(PDF_CONFIG.orientation, 'mm', PDF_CONFIG.pageSize);

    // Configurações iniciais
    let currentY = PDF_CONFIG.margin;
    const pageWidth = doc.internal.pageSize.getWidth();
    const contentWidth = pageWidth - (PDF_CONFIG.margin * 2);

    // Adicionar cabeçalho
    currentY = addHeader(doc, currentY, pageWidth);

    // Adicionar informações do relatório
    currentY = addReportInfo(doc, currentY);

    // Adicionar tabela de dados
    currentY = addDataTable(doc, currentY, contentWidth);

    // Adicionar rodapé
    addFooter(doc, pageWidth);

    // Salvar o PDF
    doc.save(`relatorio_${new Date().toISOString().slice(0,10)}.pdf`);
}

// Funções auxiliares
function addHeader(doc, currentY, pageWidth) {
    doc.setFontSize(PDF_CONFIG.fontSizes.title);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(PDF_CONFIG.colors.primary);

    // Título principal
    doc.text("Relatório Técnico de Atividades", pageWidth / 2, currentY, { align: 'center' });
    currentY += 10;

    // Subtítulo
    doc.setFontSize(PDF_CONFIG.fontSizes.subtitle);
    doc.setTextColor(PDF_CONFIG.colors.text);
    doc.text("Relatório gerado automaticamente pelo sistema", pageWidth / 2, currentY, { align: 'center' });

    return currentY + 15;
}

function addReportInfo(doc, currentY) {
    doc.setFontSize(PDF_CONFIG.fontSizes.body);
    doc.setFont('helvetica', 'normal');

    // Informações do relatório
    const reportDate = new Date().toLocaleDateString('pt-BR', {
        day: '2-digit',
        month: 'long',
        year: 'numeric'
    });

    const reportInfo = [
        `Data de emissão: ${reportDate}`,
        `Total de registros: ${document.querySelectorAll('#manualTable tbody tr').length}`
    ];

    reportInfo.forEach((info, index) => {
        doc.text(info, PDF_CONFIG.margin, currentY + (index * PDF_CONFIG.lineHeight));
    });

    return currentY + (reportInfo.length * PDF_CONFIG.lineHeight) + 10;
}

function addDataTable(doc, currentY, contentWidth) {
    const table = document.getElementById('manualTable');
    const headers = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent);
    const rows = table.querySelectorAll('tbody tr');

    // Configurações da tabela
    const columnCount = headers.length;
    const columnWidth = contentWidth / columnCount;
    const rowHeight = PDF_CONFIG.lineHeight * 1.5;

    // Estilo do cabeçalho
    doc.setFontSize(PDF_CONFIG.fontSizes.header);
    doc.setFont('helvetica', 'bold');
    doc.setFillColor(PDF_CONFIG.colors.primary);
    doc.setTextColor('#ffffff');

    // Desenhar cabeçalho
    headers.forEach((header, index) => {
        doc.rect(
            PDF_CONFIG.margin + (index * columnWidth),
            currentY,
            columnWidth,
            rowHeight,
            'F'
        );
        doc.text(
            header,
            PDF_CONFIG.margin + (index * columnWidth) + (columnWidth / 2),
            currentY + rowHeight / 2,
            { align: 'center' }
        );
    });

    currentY += rowHeight;

    // Estilo do conteúdo
    doc.setFontSize(PDF_CONFIG.fontSizes.body);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(PDF_CONFIG.colors.text);

    // Desenhar linhas
    rows.forEach(row => {
        const cells = Array.from(row.querySelectorAll('td'));

        // Verificar se precisa de nova página
        if (currentY + rowHeight > doc.internal.pageSize.getHeight() - PDF_CONFIG.margin) {
            doc.addPage();
            currentY = PDF_CONFIG.margin;

            // Redesenhar cabeçalho na nova página
            headers.forEach((header, index) => {
                doc.rect(
                    PDF_CONFIG.margin + (index * columnWidth),
                    currentY,
                    columnWidth,
                    rowHeight,
                    'F'
                );
                doc.text(
                    header,
                    PDF_CONFIG.margin + (index * columnWidth) + (columnWidth / 2),
                    currentY + rowHeight / 2,
                    { align: 'center' }
                );
            });
            currentY += rowHeight;
        }

        cells.forEach((cell, index) => {
            if (index < columnCount - 1) { // Ignorar a coluna de ações
                doc.text(
                    cell.textContent,
                    PDF_CONFIG.margin + (index * columnWidth) + (columnWidth / 2),
                    currentY + rowHeight / 2,
                    { align: 'center', maxWidth: columnWidth - 4 }
                );

                // Bordas da célula
                doc.rect(
                    PDF_CONFIG.margin + (index * columnWidth),
                    currentY,
                    columnWidth,
                    rowHeight
                );
            }
        });

        currentY += rowHeight;
    });

    return currentY + 10;
}

function addFooter(doc, pageWidth) {
    doc.setFontSize(PDF_CONFIG.fontSizes.body - 2);
    doc.setFont('helvetica', 'italic');
    doc.setTextColor(PDF_CONFIG.colors.text);

    const footerText = `Gerado automaticamente pelo Sistema de Gestão de Relatórios - Página ${doc.internal.getNumberOfPages()}`;

    doc.text(
        footerText,
        pageWidth / 2,
        doc.internal.pageSize.getHeight() - PDF_CONFIG.margin + 5,
        { align: 'center' }
    );

    // Linha do rodapé
    doc.setLineWidth(0.2);
    doc.line(
        PDF_CONFIG.margin,
        doc.internal.pageSize.getHeight() - PDF_CONFIG.margin,
        pageWidth - PDF_CONFIG.margin,
        doc.internal.pageSize.getHeight() - PDF_CONFIG.margin
    );
}

function convertSerialToDate(serial) {
    const baseDate = new Date(1899, 11, 30); // Data base do Excel (30/12/1899)
    const date = new Date(baseDate.getTime() + serial * 24 * 60 * 60 * 1000);
    return date.toLocaleDateString(); // Formata a data para o formato local
}

function importData() {
    const fileInput = document.getElementById('fileInput');
    const file = fileInput.files[0];
    if (!file) {
        alert('Por favor, selecione um arquivo.');
        return;
    }

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
