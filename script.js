document.addEventListener('DOMContentLoaded', () => {
    const dropZone = document.getElementById('drop-zone');
    const fileInput = document.getElementById('file-input');
    const uploadStatus = document.getElementById('upload-status');
    const exportBtns = document.querySelectorAll('.export-btn');
    const exportStatus = document.getElementById('export-status');
    
    // UI steps
    const stepEditor = document.getElementById('step-editor');
    const stepExport = document.getElementById('step-export');
    
    // Table Elements
    const tableHead = document.getElementById('table-head');
    const tableBody = document.getElementById('table-body');
    
    // Data State
    let currentData = [];
    let originalData = [];

    // --- UPLOAD LOGIC ---
    dropZone.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', handleFiles);

    // Styling drag/drop
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.add('dragover'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => dropZone.classList.remove('dragover'), false);
    });

    dropZone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        handleFiles({ target: { files: files }});
    });

    // Real File Processing with SheetJS
    function handleFiles(e) {
        let files = e.target.files;
        if (files.length > 0) {
            let file = files[0];
            uploadStatus.style.color = '#10b981';
            uploadStatus.innerHTML = `⏳ Идет парсинг файла: <strong>${file.name}</strong>...`;
            
            const reader = new FileReader();
            
            reader.onload = function(evt) {
                try {
                    const data = evt.target.result;
                    // Load the workbook
                    const workbook = XLSX.read(data, {type: 'binary'});
                    
                    // Access the first sheet
                    const firstSheetName = workbook.SheetNames[0];
                    const worksheet = workbook.Sheets[firstSheetName];
                    
                    // Convert sheet to json
                    const json = XLSX.utils.sheet_to_json(worksheet, {defval: ""});
                    
                    if (json.length === 0) {
                        throw new Error("Файл пуст или имеет несовместимую структуру.");
                    }
                    
                    originalData = JSON.parse(JSON.stringify(json));
                    currentData = JSON.parse(JSON.stringify(json));
                    
                    renderTable(currentData);
                    
                    uploadStatus.innerHTML = `✅ Данные из файла загружены! Распознано строк: ${currentData.length}.`;
                    
                    // Show next UI states
                    stepEditor.style.display = 'block';
                    stepExport.style.display = 'block';
                    
                    setTimeout(() => {
                        stepEditor.scrollIntoView({behavior: 'smooth', block: 'start'});
                    }, 100);

                } catch (error) {
                    uploadStatus.style.color = '#ef4444';
                    uploadStatus.innerHTML = `❌ Ошибка при чтении файла: ${error.message}`;
                }
            };
            
            reader.onerror = function() {
                uploadStatus.style.color = '#ef4444';
                uploadStatus.innerHTML = `❌ Системная ошибка при загрузке.`;
            };
            
            reader.readAsBinaryString(file);
        }
    }

    // --- TABLE RENDER LOGIC ---
    function renderTable(dataArray) {
        tableHead.innerHTML = '';
        tableBody.innerHTML = '';
        
        if (dataArray.length === 0) return;
        
        const headers = Object.keys(dataArray[0]);
        
        // Add headers
        headers.forEach(h => {
            let th = document.createElement('th');
            th.textContent = h;
            tableHead.appendChild(th);
        });
        
        // Only render max 50 rows to keep DOM fast
        const displayData = dataArray.slice(0, 50);
        
        displayData.forEach(row => {
            let tr = document.createElement('tr');
            headers.forEach(h => {
                let td = document.createElement('td');
                td.textContent = row[h];
                tr.appendChild(td);
            });
            tableBody.appendChild(tr);
        });
    }

    // --- TRANSFORMATION TOOLS ---
    document.getElementById('tool-uppercase').addEventListener('click', () => {
        if(currentData.length === 0) return;
        currentData = currentData.map(row => {
            let newRow = {};
            for(let key in row) {
                newRow[key] = typeof row[key] === 'string' ? row[key].toUpperCase() : row[key];
            }
            return newRow;
        });
        renderTable(currentData);
    });

    document.getElementById('tool-add-col').addEventListener('click', () => {
        if(currentData.length === 0) return;
        const now = new Date().toLocaleString('ru-RU');
        currentData = currentData.map(row => {
            let newRow = {...row};
            newRow['Process_Date'] = now; // adding a stamp to each row
            return newRow;
        });
        renderTable(currentData);
    });

    document.getElementById('tool-reset').addEventListener('click', () => {
        if(originalData.length === 0) return;
        currentData = JSON.parse(JSON.stringify(originalData));
        renderTable(currentData);
    });

    // --- EXPORT LOGIC ---
    exportBtns.forEach(btn => {
        btn.addEventListener('click', (e) => {
            if(currentData.length === 0) {
                alert("Нет данных для экспорта. Пожалуйста, загрузите файл на первом этапе.");
                return;
            }

            const type = e.currentTarget.dataset.type; // 'excel' or 'csv'
            const icon = type === 'excel' ? '📊' : '📝';
            const ext = type === 'excel' ? '.xlsx' : '.csv';
            
            exportStatus.style.color = '#3b82f6';
            exportStatus.innerHTML = `⏳ Генерация ${ext} файла из данных мода...`;
            
            setTimeout(() => {
                try {
                    // Turn our parsed JSON back to a SheetJS object
                    const worksheet = XLSX.utils.json_to_sheet(currentData);
                    const newWorkbook = XLSX.utils.book_new();
                    XLSX.utils.book_append_sheet(newWorkbook, worksheet, "Processed Data");
                    
                    // Actually trigger a browser download! 
                    XLSX.writeFile(newWorkbook, `Datasync_Result${ext}`);
                    
                    exportStatus.style.color = '#10b981';
                    exportStatus.innerHTML = `✅ ${icon} Файл Datasync_Result${ext} успешно скачан на ваше устройство!`;
                    
                    setTimeout(() => {
                        exportStatus.innerHTML = '';
                    }, 5000);
                } catch(err) {
                    exportStatus.style.color = '#ef4444';
                    exportStatus.innerHTML = `❌ Произошла ошибка экспорта: ${err.message}`;
                }
            }, 600); // Small visual delay
        });
    });
});
