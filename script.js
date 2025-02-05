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

function addHeader(doc, currentY, pageWidth) {
    doc.setFontSize(PDF_CONFIG.fontSizes.title);
    doc.setFont('helvetica', 'bold');
    doc.setTextColor(PDF_CONFIG.colors.primary);
    doc.text("Relatório Técnico de Atividades", pageWidth / 2, currentY, { align: 'center' });
    currentY += 10;

    doc.setFontSize(PDF_CONFIG.fontSizes.subtitle);
    doc.setTextColor(PDF_CONFIG.colors.text);
    doc.text("Relatório gerado automaticamente pelo sistema", pageWidth / 2, currentY, { align: 'center' });
    
    return currentY + 15;
}

function addReportInfo(doc, currentY) {
    doc.setFontSize(PDF_CONFIG.fontSizes.body);
    doc.setFont('helvetica', 'normal');

    const reportDate = new Date().toLocaleDateString('pt-BR', { day: '2-digit', month: 'long', year: 'numeric' });
    const reportInfo = [`Data de emissão: ${reportDate}`];

    reportInfo.forEach((info, index) => {
        doc.text(info, PDF_CONFIG.margin, currentY + (index * PDF_CONFIG.lineHeight));
    });

    return currentY + (reportInfo.length * PDF_CONFIG.lineHeight) + 10;
}

function addDataTable(doc, currentY, contentWidth) {
    const table = document.getElementById('manualTable');
    const headers = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent);
    const rows = table.querySelectorAll('tbody tr');

    const columnCount = headers.length;
    const columnWidth = contentWidth / columnCount;
    const rowHeight = PDF_CONFIG.lineHeight * 1.5;

    doc.setFontSize(PDF_CONFIG.fontSizes.header);
    doc.setFont('helvetica', 'bold');
    doc.setFillColor(PDF_CONFIG.colors.primary);
    doc.setTextColor('#ffffff');

    headers.forEach((header, index) => {
        doc.rect(PDF_CONFIG.margin + (index * columnWidth), currentY, columnWidth, rowHeight, 'F');
        doc.text(header, PDF_CONFIG.margin + (index * columnWidth) + columnWidth / 2, currentY + rowHeight / 2 + 2, { align: 'center' });
    });

    currentY += rowHeight;
    doc.setFontSize(PDF_CONFIG.fontSizes.body);
    doc.setFont('helvetica', 'normal');
    doc.setTextColor(PDF_CONFIG.colors.text);

    rows.forEach(row => {
        const cells = Array.from(row.querySelectorAll('td'));
        if (currentY + rowHeight > doc.internal.pageSize.getHeight() - PDF_CONFIG.margin) {
            doc.addPage();
            currentY = PDF_CONFIG.margin;
            headers.forEach((header, index) => {
                doc.rect(PDF_CONFIG.margin + (index * columnWidth), currentY, columnWidth, rowHeight, 'F');
                doc.text(header, PDF_CONFIG.margin + (index * columnWidth) + columnWidth / 2, currentY + rowHeight / 2 + 2, { align: 'center' });
            });
            currentY += rowHeight;
        }

        cells.forEach((cell, index) => {
            doc.text(cell.textContent, PDF_CONFIG.margin + (index * columnWidth) + 2, currentY + rowHeight / 2 + 2, { maxWidth: columnWidth - 4 });
            doc.rect(PDF_CONFIG.margin + (index * columnWidth), currentY, columnWidth, rowHeight);
        });

        currentY += rowHeight;
    });

    return currentY + 10;
}

function addFooter(doc, pageWidth) {
    doc.setFontSize(PDF_CONFIG.fontSizes.body - 2);
    doc.setFont('helvetica', 'italic');
    doc.setTextColor(PDF_CONFIG.colors.text);

    const footerText = `Gerado automaticamente pelo Sistema - Página ${doc.internal.getNumberOfPages()}`;
    doc.text(footerText, pageWidth / 2, doc.internal.pageSize.getHeight() - PDF_CONFIG.margin + 5, { align: 'center' });
    doc.setLineWidth(0.2);
    doc.line(PDF_CONFIG.margin, doc.internal.pageSize.getHeight() - PDF_CONFIG.margin, pageWidth - PDF_CONFIG.margin, doc.internal.pageSize.getHeight() - PDF_CONFIG.margin);
}
