/**
 * Office Viewer - 主入口文件
 * 处理文件上传和初始化各个预览器
 */

document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const previewContainer = document.getElementById('preview-container');
    const tableWrapper = document.getElementById('table-wrapper');
    const controls = document.getElementById('controls');
    const sheetTabs = document.getElementById('sheet-tabs');
    const uploadWrapper = document.getElementById('upload-wrapper');
    const fileNameDisplay = document.getElementById('file-name');
    const cellInfo = document.getElementById('cell-info');
    const cellInfoBar = document.getElementById('cell-info-bar');
    const cellPosition = document.getElementById('cell-position');
    const cellInfoContent = document.getElementById('cell-info-content');

    // 初始化预览器
    const excelViewer = new ExcelViewer({
        tableWrapper,
        controls,
        sheetTabs,
        previewContainer,
        uploadWrapper,
        cellInfo,
        cellInfoBar,
        cellPosition,
        cellInfoContent
    });

    const pdfViewer = new PDFViewer({
        previewContainer,
        uploadWrapper,
        controls
    });

    const docxViewer = new DOCXViewer({
        previewContainer,
        uploadWrapper,
        controls
    });

    // ==========================================
    // 文件处理
    // ==========================================

    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('dragover'); });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('dragover'));
    dropZone.addEventListener('drop', e => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
        if (e.dataTransfer.files.length) handleFile(e.dataTransfer.files[0]);
    });
    fileInput.addEventListener('change', e => {
        if (e.target.files.length) handleFile(e.target.files[0]);
    });

    async function handleFile(file) {
        fileNameDisplay.textContent = file.name;
        
        const fileExtension = file.name.split('.').pop().toLowerCase();
        
        try {
            if (fileExtension === 'xlsx' || fileExtension === 'xls') {
                pdfViewer.destroy();
                docxViewer.destroy();
                const tableWrapper = previewContainer.querySelector('.table-wrapper');
                const pdfContainer = previewContainer.querySelector('.pdf-viewer-container');
                const docxContainer = previewContainer.querySelector('.docx-viewer-container');
                const cellInfoBar = previewContainer.querySelector('.cell-info-bar');
                if (sheetTabs) sheetTabs.style.display = '';
                if (tableWrapper) tableWrapper.style.display = '';
                if (pdfContainer) pdfContainer.style.display = 'none';
                if (docxContainer) docxContainer.style.display = 'none';
                if (cellInfoBar) cellInfoBar.style.display = '';
                await excelViewer.loadFile(file);
            } else if (fileExtension === 'pdf') {
                docxViewer.destroy();
                if (sheetTabs) sheetTabs.style.display = 'none';
                await pdfViewer.loadFile(file);
            } else if (fileExtension === 'docx') {
                pdfViewer.destroy();
                if (sheetTabs) sheetTabs.style.display = 'none';
                const tableWrapper = previewContainer.querySelector('.table-wrapper');
                const pdfContainer = previewContainer.querySelector('.pdf-viewer-container');
                const cellInfoBar = previewContainer.querySelector('.cell-info-bar');
                if (tableWrapper) tableWrapper.style.display = 'none';
                if (pdfContainer) pdfContainer.style.display = 'none';
                if (cellInfoBar) cellInfoBar.style.display = 'none';
                await docxViewer.loadFile(file);
            } else {
                alert('暂不支持 ' + fileExtension.toUpperCase() + ' 格式');
            }
        } catch (err) {
            console.error('处理文件失败:', err);
            alert('无法处理文件: ' + err.message);
        }
    }
});
