let workbook = null;
let fileName = '';
let selectedSheet = null;

// DOM elements
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileNameSpan = document.getElementById('fileName');
const sheetSelector = document.getElementById('sheetSelector');
const sheetSelect = document.getElementById('sheetSelect');
const generateBtn = document.getElementById('generateBtn');
const preview = document.getElementById('preview');
const previewContent = document.getElementById('previewContent');
const status = document.getElementById('status');

// Upload area click
uploadArea.addEventListener('click', () => {
    fileInput.click();
});

// Drag and drop
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

// File input change
fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

// Sheet selection change
sheetSelect.addEventListener('change', (e) => {
    if (e.target.value) {
        selectedSheet = e.target.value;
        generateBtn.classList.add('show');
        preview.classList.remove('show');
        status.classList.remove('show');
    } else {
        generateBtn.classList.remove('show');
    }
});

// Generate button click
generateBtn.addEventListener('click', generateSQL);

// Handle file upload
function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        showStatus('请上传有效的Excel文件（.xlsx 或 .xls）', 'error');
        return;
    }

    fileName = file.name.replace(/\.(xlsx|xls)$/, '');
    fileNameSpan.textContent = file.name;
    fileInfo.classList.add('show');

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            workbook = XLSX.read(data, { type: 'array' });

            // Populate sheet selector
            sheetSelect.innerHTML = '<option value="">请选择...</option>';
            workbook.SheetNames.forEach(name => {
                const option = document.createElement('option');
                option.value = name;
                option.textContent = name;
                sheetSelect.appendChild(option);
            });

            sheetSelector.classList.add('show');
            showStatus('文件加载成功！请选择Sheet表单。', 'success');
        } catch (error) {
            showStatus('文件解析失败：' + error.message, 'error');
        }
    };
    reader.readAsArrayBuffer(file);
}

// Parse INSERT SQL to extract table name and columns
function parseInsertSQL(sql) {
    // Match INSERT INTO table_name (col1, col2, ...) VALUES (val1, val2, ...)
    const pattern = /INSERT\s+INTO\s+(\S+)\s*\(([^)]+)\)\s*VALUES\s*\(([^)]+)\)/i;
    const match = sql.match(pattern);

    if (!match) {
        return null;
    }

    const tableName = match[1].trim();
    const columns = match[2].split(',').map(col => col.trim());
    const values = match[3].split(',').map(val => val.trim());

    return { tableName, columns, values };
}

// Generate SQL
function generateSQL() {
    if (!workbook || !selectedSheet) {
        showStatus('请先选择文件和Sheet表单', 'error');
        return;
    }

    try {
        const worksheet = workbook.Sheets[selectedSheet];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

        if (jsonData.length < 4) {
            showStatus('Sheet表单数据不足：至少需要4行（SQL示例、表头描述、表头、数据）', 'error');
            return;
        }

        // First row: SQL example
        const sqlExample = jsonData[0][0];
        if (!sqlExample || typeof sqlExample !== 'string') {
            showStatus('第一行必须包含INSERT SQL示例', 'error');
            return;
        }

        // Parse the SQL example
        const parsedSQL = parseInsertSQL(sqlExample);
        if (!parsedSQL) {
            showStatus('无法解析SQL语句，请确保格式为：INSERT INTO table (col1, col2) VALUES (val1, val2)', 'error');
            return;
        }

        // Second row: Header descriptions (ignored)
        // Third row: Headers
        const headers = jsonData[2];
        if (!headers || headers.length === 0) {
            showStatus('第三行必须包含表头', 'error');
            return;
        }

        // Clean headers
        const cleanHeaders = headers.map(h => h ? h.toString().trim() : '').filter(h => h);

        if (cleanHeaders.length === 0) {
            showStatus('表头不能为空', 'error');
            return;
        }

        // Map headers to column indices
        const headerMap = {};
        headers.forEach((header, index) => {
            if (header) {
                const cleanHeader = header.toString().trim();
                headerMap[cleanHeader] = index;
            }
        });

        // Match SQL columns with Excel headers
        const columnMapping = [];
        for (let i = 0; i < parsedSQL.columns.length; i++) {
            const sqlColumn = parsedSQL.columns[i];
            const sqlValue = parsedSQL.values[i];

            // Try exact match first
            if (headerMap.hasOwnProperty(sqlColumn)) {
                columnMapping.push({ sqlColumn, excelIndex: headerMap[sqlColumn], originalValue: null });
            } else {
                // Try case-insensitive match
                const lowerSqlColumn = sqlColumn.toLowerCase();
                const matchedHeader = Object.keys(headerMap).find(
                    h => h.toLowerCase() === lowerSqlColumn
                );

                if (matchedHeader) {
                    columnMapping.push({ sqlColumn, excelIndex: headerMap[matchedHeader], originalValue: null });
                } else {
                    // Field not found in headers, keep original value
                    columnMapping.push({ sqlColumn, excelIndex: null, originalValue: sqlValue });
                }
            }
        }

        // Generate SQL statements
        const sqlStatements = [];
        for (let i = 3; i < jsonData.length; i++) {
            const row = jsonData[i];

            // Skip empty rows
            if (!row || row.every(cell => !cell && cell !== 0)) {
                continue;
            }

            // Build VALUES clause
            const values = columnMapping.map(mapping => {
                // If field not found in headers, use original value from SQL example
                if (mapping.excelIndex === null) {
                    return mapping.originalValue;
                }

                let value = row[mapping.excelIndex];

                // Handle different data types
                if (value === null || value === undefined || value === '') {
                    return 'NULL';
                } else if (typeof value === 'number') {
                    return value.toString();
                } else {
                    // Escape single quotes in strings
                    value = value.toString().replace(/'/g, "''");
                    return `'${value}'`;
                }
            });

            // Construct SQL statement
            const sql = `INSERT INTO ${parsedSQL.tableName} (${parsedSQL.columns.join(', ')}) VALUES (${values.join(', ')});`;
            sqlStatements.push(sql);
        }

        if (sqlStatements.length === 0) {
            showStatus('没有生成任何SQL语句（可能所有数据行都是空的）', 'error');
            return;
        }

        // Show preview
        const previewText = sqlStatements.slice(0, 10).join('\n\n');
        previewContent.textContent = previewText;
        preview.classList.add('show');

        // Generate file name
        const today = new Date();
        const dateStr = today.getFullYear() +
                       String(today.getMonth() + 1).padStart(2, '0') +
                       String(today.getDate()).padStart(2, '0') + 
                       String(today.getHours()) + 
                       String(today.getMinutes()) + 
                       String(today.getMilliseconds());
        const outputFileName = `${fileName}_${selectedSheet}_${dateStr}.sql`;

        // Download file
        const sqlContent = sqlStatements.join('\n\n');
        downloadFile(sqlContent, outputFileName);

        showStatus(`成功生成 ${sqlStatements.length} 条SQL语句！文件已下载。`, 'success');
    } catch (error) {
        showStatus('生成SQL时出错：' + error.message, 'error');
        console.error(error);
    }
}

// Download file
function downloadFile(content, filename) {
    const blob = new Blob([content], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// Show status message
function showStatus(message, type) {
    status.textContent = message;
    status.className = 'status show ' + type;
    setTimeout(() => {
        status.classList.remove('show');
    }, 5000);
}