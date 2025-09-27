// Global variables
let uploadedData = null;
let columnMappings = {};
let fixedValues = {}; // New: store fixed values for columns
let valueMappings = {};
let sourceColumns = [];
let validationResults = null;
let currentConfigName = null;

// Valid options from the POD system (exact column order from template)
const requiredColumns = [
    { key: 'ISBN', label: 'ISBN', required: true, description: '13-digit ISBN number', order: 1 },
    { key: 'Title', label: 'Title', required: true, description: 'Book title (max 58 characters)', order: 2 },
    { key: 'Trim Height', label: 'Trim Height', required: true, description: 'Height in mm', order: 3 },
    { key: 'Trim Width', label: 'Trim Width', required: true, description: 'Width in mm', order: 4 },
    { key: 'Paper Type', label: 'Paper Type', required: true, description: 'Paper specification', order: 5, hasValueMapping: true },
    { key: 'Binding Style', label: 'Binding Style', required: true, description: 'Limp or Cased', order: 6, hasValueMapping: true },
    { key: 'Page Extent', label: 'Page Extent', required: true, description: 'Number of pages', order: 7 },
    { key: 'Lamination', label: 'Lamination', required: true, description: 'Gloss, Matt, or None', order: 8, hasValueMapping: true },
    { key: 'Plate Section 1', label: 'Plate Section 1', required: false, description: 'Optional: Insert after p[page]-[pages]pp-[paper type]', order: 9 },
    { key: 'Plate Section 2', label: 'Plate Section 2', required: false, description: 'Optional: Insert after p[page]-[pages]pp-[paper type]', order: 10 }
];

// Valid values from the template
const validOptions = {
    'Paper Type': [
        'Amber Preprint 80 gsm',
        'Woodfree 80 gsm',
        'Munken Print Cream 70 gsm',
        'Munken Print Cream 80 gsm',
        'Navigator 80 gsm',
        'LetsGo Silk 90 gsm',
        'Matt 115 gsm',
        'Holmen Book Cream 60 gsm',
        'Enso 70 gsm',
        'Holmen Bulky 52 gsm',
        'Holmen Book 55 gsm',
        'Holmen Cream 65 gsm',
        'Holmen Book 52 gsm',
        'Premium Mono 90 gsm',
        'Premium Colour 90 gsm'
    ],
    'Binding Style': ['Limp', 'Cased'],
    'Lamination': ['Gloss', 'Matt', 'None']
};

// Initialize value mappings
function initializeValueMappings() {
    valueMappings = {
        'Paper Type': {},
        'Binding Style': {},
        'Lamination': {}
    };
}

// Initialize the app
document.addEventListener('DOMContentLoaded', function() {
    setupFileInput();
    initializeValueMappings();
});

function setupFileInput() {
    document.getElementById('fileInput').addEventListener('change', function(e) {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0]);
        }
    });
}

async function handleFile(file) {
    if (!file.name.match(/\.(xls|xlsx|xlsm)$/i)) {
        showAlert('Please upload an Excel file (.xls, .xlsx, or .xlsm)', 'danger');
        return;
    }

    try {
        showAlert('Processing file...', 'info');
        const buffer = await readFileAsArrayBuffer(file);
        const workbook = XLSX.read(new Uint8Array(buffer), { type: 'array' });
        
        // Get first worksheet
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
        
        if (jsonData.length === 0) {
            showAlert('The Excel file appears to be empty or has no data rows.', 'warning');
            return;
        }

        uploadedData = jsonData;
        sourceColumns = Object.keys(jsonData[0]);
        
        displayFileInfo(file, jsonData.length, sourceColumns.length);
        createMappingInterface();
        showAlert(`Successfully loaded ${jsonData.length} rows with ${sourceColumns.length} columns`, 'success');

    } catch (error) {
        console.error('Error processing file:', error);
        showAlert('Error processing Excel file. Please check the file format.', 'danger');
    }
}

function readFileAsArrayBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = e => resolve(e.target.result);
        reader.onerror = e => reject(e);
        reader.readAsArrayBuffer(file);
    });
}

function displayFileInfo(file, rowCount, columnCount) {
    const fileInfoSection = document.getElementById('fileInfoSection');
    const fileInfoContent = document.getElementById('fileInfoContent');
    
    fileInfoContent.innerHTML = `
        <div class="row">
            <div class="col-md-6">
                <h6>File Details</h6>
                <p class="mb-1"><strong>Name:</strong> ${file.name}</p>
                <p class="mb-1"><strong>Size:</strong> ${(file.size / 1024 / 1024).toFixed(2)} MB</p>
                <p class="mb-1"><strong>Type:</strong> ${file.type}</p>
            </div>
            <div class="col-md-6">
                <h6>Data Summary</h6>
                <p class="mb-1"><strong>Rows:</strong> ${rowCount}</p>
                <p class="mb-1"><strong>Columns:</strong> ${columnCount}</p>
                <p class="mb-1"><strong>Column Names:</strong> ${sourceColumns.join(', ')}</p>
            </div>
        </div>
    `;
    
    fileInfoSection.classList.remove('d-none');
}

function createMappingInterface() {
    const mappingContainer = document.getElementById('mappingContainer');
    mappingContainer.innerHTML = '';

    // Sort columns by order for display
    const sortedColumns = [...requiredColumns].sort((a, b) => a.order - b.order);

    sortedColumns.forEach(column => {
        const mappingRow = document.createElement('div');
        mappingRow.className = `mapping-row ${column.required ? 'required-field' : 'optional-field'}`;
        
        mappingRow.innerHTML = `
            <div class="row align-items-center">
                <div class="col-md-3">
                    <label class="form-label mb-0">
                        <strong>${column.label}</strong>
                        ${column.required ? '<span class="badge bg-danger info-badge">Required</span>' : '<span class="badge bg-success info-badge">Optional</span>'}
                        ${column.hasValueMapping ? '<span class="badge bg-warning info-badge">Mappable</span>' : ''}
                    </label>
                    <small class="text-muted d-block">${column.description}</small>
                </div>
                <div class="col-md-4">
                    <select class="form-select" id="mapping_${column.key}" onchange="updateMapping('${column.key}')">
                        <option value="">-- Select source column --</option>
                        ${sourceColumns.map(col => `<option value="${col}">${col}</option>`).join('')}
                        <option value="__FIXED__">Use fixed value...</option>
                    </select>
                    <input type="text" class="form-control mt-1 fixed-value-input d-none" 
                           id="fixed_${column.key}" 
                           placeholder="Enter fixed value"
                           onchange="updateFixedValue('${column.key}', this.value)">
                </div>
                <div class="col-md-5">
                    <div class="column-preview" id="preview_${column.key}">
                        Select a column to see preview data
                    </div>
                </div>
            </div>
        `;
        
        mappingContainer.appendChild(mappingRow);
    });

    // Auto-map columns based on similarity
    autoMapColumns();
    
    document.getElementById('mappingSection').classList.remove('d-none');
    document.getElementById('exportSection').classList.remove('d-none');
    updateExportButtonState();
    updateConfigSection();
}

function autoMapColumns() {
    requiredColumns.forEach(column => {
        const bestMatch = findBestColumnMatch(column.key, column.label);
        if (bestMatch) {
            const selectElement = document.getElementById(`mapping_${column.key}`);
            selectElement.value = bestMatch;
            updateMapping(column.key);
        }
    });
}

function findBestColumnMatch(key, label) {
    const searchTerms = [key.toLowerCase(), label.toLowerCase()];
    
    for (const term of searchTerms) {
        // Exact match
        const exactMatch = sourceColumns.find(col => col.toLowerCase() === term);
        if (exactMatch) return exactMatch;
        
        // Partial match
        const partialMatch = sourceColumns.find(col => 
            col.toLowerCase().includes(term) || term.includes(col.toLowerCase())
        );
        if (partialMatch) return partialMatch;
    }
    
    return null;
}

function updateMapping(columnKey) {
    const selectElement = document.getElementById(`mapping_${columnKey}`);
    const selectedColumn = selectElement.value;
    const previewElement = document.getElementById(`preview_${columnKey}`);
    const fixedInputElement = document.getElementById(`fixed_${columnKey}`);
    
    if (selectedColumn === '__FIXED__') {
        // Show fixed value input
        fixedInputElement.classList.remove('d-none');
        previewElement.innerHTML = '<span class="text-warning">Enter a fixed value</span>';
        delete columnMappings[columnKey];
    } else if (selectedColumn) {
        // Hide fixed value input
        fixedInputElement.classList.add('d-none');
        delete fixedValues[columnKey];
        
        columnMappings[columnKey] = selectedColumn;
        
        // Show preview data
        const sampleValues = uploadedData.slice(0, 5).map(row => row[selectedColumn]);
        previewElement.innerHTML = `Sample: ${sampleValues.join(', ')}`;
    } else {
        // No selection
        fixedInputElement.classList.add('d-none');
        delete columnMappings[columnKey];
        delete fixedValues[columnKey];
        previewElement.innerHTML = 'Select a column to see preview data';
    }
    
    updatePreviewData();
    updateExportButtonState();
    
    // Clear validation when mappings change
    validationResults = null;
    document.getElementById('validationSection').classList.add('d-none');
    document.getElementById('validationStatus').textContent = 'Not validated';
    
    // Update value mappings
    updateValueMappings();
}

function updateFixedValue(columnKey, value) {
    if (value && value.trim()) {
        fixedValues[columnKey] = value.trim();
        const previewElement = document.getElementById(`preview_${columnKey}`);
        previewElement.innerHTML = `<span class="badge bg-warning">Fixed: ${value.trim()}</span>`;
    } else {
        delete fixedValues[columnKey];
        const previewElement = document.getElementById(`preview_${columnKey}`);
        previewElement.innerHTML = '<span class="text-warning">Enter a fixed value</span>';
    }
    
    updatePreviewData();
    updateExportButtonState();
    updateConfigSection();
}

function updatePreviewData() {
    const previewData = document.getElementById('previewData');
    
    if (Object.keys(columnMappings).length === 0 && Object.keys(fixedValues).length === 0) {
        previewData.innerHTML = '<p class="text-muted">Map some columns to see preview data</p>';
        return;
    }

    const previewRows = uploadedData.slice(0, 5);
    let tableHTML = '<table class="table table-sm table-bordered"><thead><tr>';
    
    // Headers in correct order
    const sortedColumns = [...requiredColumns].sort((a, b) => a.order - b.order);
    sortedColumns.forEach(column => {
        if (columnMappings[column.key] || fixedValues[column.key]) {
            tableHTML += `<th>${column.label}</th>`;
        }
    });
    tableHTML += '</tr></thead><tbody>';
    
    // Data rows
    previewRows.forEach(row => {
        tableHTML += '<tr>';
        sortedColumns.forEach(column => {
            if (fixedValues[column.key]) {
                tableHTML += `<td><span class="badge bg-warning">${fixedValues[column.key]}</span></td>`;
            } else if (columnMappings[column.key]) {
                const value = row[columnMappings[column.key]] || '';
                tableHTML += `<td>${value}</td>`;
            }
        });
        tableHTML += '</tr>';
    });
    
    tableHTML += '</tbody></table>';
    previewData.innerHTML = tableHTML;
}

function updateValueMappings() {
    // Check if we have any columns that need value mapping
    const columnsWithValueMapping = requiredColumns.filter(col => 
        col.hasValueMapping && columnMappings[col.key]
    );
    
    if (columnsWithValueMapping.length === 0) {
        document.getElementById('valueMappingSection').classList.add('d-none');
        return;
    }
    
    const valueMappingContent = document.getElementById('valueMappingContent');
    let html = '';
    
    columnsWithValueMapping.forEach(column => {
        const sourceColumn = columnMappings[column.key];
        const uniqueValues = [...new Set(uploadedData.map(row => row[sourceColumn]).filter(val => val))];
        const validValues = validOptions[column.key];
        
        html += `
            <div class="value-mapping-section">
                <h6><i class="bi bi-arrow-right me-2"></i>${column.label} Value Mapping</h6>
                <p class="text-muted">Map your values to standard POD options:</p>
                
                <div class="row">
                    <div class="col-md-6">
                        <strong>Your Values (${uniqueValues.length})</strong>
                        <div class="mt-2">
                            ${uniqueValues.map(value => {
                                const escapedValue = value.replace(/'/g, "\\'");
                                return `
                                <div class="mapping-option">
                                    <div class="input-group input-group-sm">
                                        <span class="input-group-text">${value}</span>
                                        <select class="form-select" onchange="updateValueMapping('${column.key}', '${escapedValue}', this.value)">
                                            <option value="">-- Map to standard value --</option>
                                            ${validValues.map(validValue => `
                                                <option value="${validValue}" ${valueMappings[column.key][value] === validValue ? 'selected' : ''}>
                                                    ${validValue}
                                                </option>
                                            `).join('')}
                                        </select>
                                    </div>
                                </div>
                            `}).join('')}
                        </div>
                    </div>
                    <div class="col-md-6">
                        <strong>Standard POD Values</strong>
                        <div class="mt-2">
                            ${validValues.map(value => `
                                <span class="badge bg-success me-1 mb-1">${value}</span>
                            `).join('')}
                        </div>
                        
                        <div class="mt-3">
                            <button class="btn btn-outline-primary btn-sm" onclick="autoMapValues('${column.key}')">
                                <i class="bi bi-magic me-1"></i>Auto-map similar values
                            </button>
                            <button class="btn btn-outline-secondary btn-sm" onclick="clearValueMappings('${column.key}')">
                                <i class="bi bi-x-lg me-1"></i>Clear mappings
                            </button>
                        </div>
                    </div>
                </div>
            </div>
        `;
    });
    
    valueMappingContent.innerHTML = html;
    document.getElementById('valueMappingSection').classList.remove('d-none');
    updateValueMappingStats();
}

function updateValueMapping(columnKey, sourceValue, targetValue) {
    if (!valueMappings[columnKey]) {
        valueMappings[columnKey] = {};
    }
    
    if (targetValue) {
        valueMappings[columnKey][sourceValue] = targetValue;
    } else {
        delete valueMappings[columnKey][sourceValue];
    }
    
    updateValueMappingStats();
}

function autoMapValues(columnKey) {
    const sourceColumn = columnMappings[columnKey];
    const uniqueValues = [...new Set(uploadedData.map(row => row[sourceColumn]).filter(val => val))];
    const validValues = validOptions[columnKey];
    
    uniqueValues.forEach(sourceValue => {
        const normalizedSource = sourceValue.toLowerCase().trim();
        
        // Try exact match first
        const exactMatch = validValues.find(valid => valid.toLowerCase() === normalizedSource);
        if (exactMatch) {
            updateValueMapping(columnKey, sourceValue, exactMatch);
            return;
        }
        
        // Try partial match
        const partialMatch = validValues.find(valid => 
            valid.toLowerCase().includes(normalizedSource) || 
            normalizedSource.includes(valid.toLowerCase())
        );
        if (partialMatch) {
            updateValueMapping(columnKey, sourceValue, partialMatch);
        }
    });
    
    // Refresh the interface
    updateValueMappings();
}

function clearValueMappings(columnKey) {
    valueMappings[columnKey] = {};
    updateValueMappings();
}

function updateValueMappingStats() {
    const totalMappings = Object.values(valueMappings).reduce((sum, mappings) => 
        sum + Object.keys(mappings).length, 0
    );
    document.getElementById('valueMappingCount').textContent = totalMappings;
}

function updateExportButtonState() {
    const requiredMapped = requiredColumns.filter(col => 
        col.required && (columnMappings[col.key] || fixedValues[col.key])
    ).length;
    const totalRequired = requiredColumns.filter(col => col.required).length;
    const totalMapped = Object.keys(columnMappings).length + Object.keys(fixedValues).length;
    const fixedCount = Object.keys(fixedValues).length;
    
    document.getElementById('totalRows').textContent = uploadedData ? uploadedData.length : 0;
    document.getElementById('mappedColumns').textContent = totalMapped;
    document.getElementById('requiredMapped').textContent = `${requiredMapped}/${totalRequired}`;
    document.getElementById('fixedValueCount').textContent = fixedCount;
    
    const exportBtn = document.getElementById('exportBtn');
    const validateBtn = document.getElementById('validateBtn');
    
    const allRequiredMapped = requiredMapped === totalRequired;
    validateBtn.disabled = !allRequiredMapped;
    exportBtn.disabled = !allRequiredMapped;
    
    if (allRequiredMapped) {
        exportBtn.classList.remove('btn-outline-success');
        exportBtn.classList.add('btn-success');
    } else {
        exportBtn.classList.remove('btn-success');
        exportBtn.classList.add('btn-outline-success');
    }
}

function applyValueMapping(value, columnKey) {
    if (valueMappings[columnKey] && valueMappings[columnKey][value]) {
        return valueMappings[columnKey][value];
    }
    return value;
}

function validateMappedData() {
    if (!uploadedData || (Object.keys(columnMappings).length === 0 && Object.keys(fixedValues).length === 0)) {
        showAlert('No data to validate', 'warning');
        return;
    }

    const results = {
        totalRows: uploadedData.length,
        validRows: 0,
        errors: [],
        warnings: [],
        valueMappingsApplied: 0
    };

    uploadedData.forEach((row, index) => {
        let hasErrors = false;
        const rowNumber = index + 1;

        // Validate required fields
        requiredColumns.filter(col => col.required).forEach(column => {
            let value = null;
            
            // Get value from fixed value or column mapping
            if (fixedValues[column.key]) {
                value = fixedValues[column.key];
            } else if (columnMappings[column.key]) {
                value = row[columnMappings[column.key]];
                
                // Apply value mapping if available
                if (column.hasValueMapping) {
                    const mappedValue = applyValueMapping(value, column.key);
                    if (mappedValue !== value) {
                        results.valueMappingsApplied++;
                    }
                    value = mappedValue;
                }
            }
            
            if (!value || value.toString().trim() === '') {
                results.errors.push(`Row ${rowNumber}: ${column.label} is required but empty`);
                hasErrors = true;
            }

            // Specific validations
            if (column.key === 'ISBN' && value) {
                const isbn = value.toString().replace(/[-\s]/g, '');
                if (isbn.length !== 13 || !/^\d{13}$/.test(isbn)) {
                    results.errors.push(`Row ${rowNumber}: Invalid ISBN format (must be 13 digits)`);
                    hasErrors = true;
                }
            }

            if (column.key === 'Title' && value && value.toString().length > 58) {
                results.errors.push(`Row ${rowNumber}: Title exceeds 58 characters (${value.toString().length} chars): "${value}"`);
                hasErrors = true;
            }

            if (column.key === 'Paper Type' && value && !validOptions['Paper Type'].includes(value)) {
                results.errors.push(`Row ${rowNumber}: Invalid paper type "${value}"`);
                hasErrors = true;
            }

            if (column.key === 'Binding Style' && value && !validOptions['Binding Style'].includes(value)) {
                results.errors.push(`Row ${rowNumber}: Invalid binding style "${value}" (must be Limp or Cased)`);
                hasErrors = true;
            }

            if (column.key === 'Lamination' && value && !validOptions['Lamination'].includes(value)) {
                results.errors.push(`Row ${rowNumber}: Invalid lamination "${value}" (must be Gloss, Matt, or None)`);
                hasErrors = true;
            }

            if ((column.key === 'Trim Height' || column.key === 'Trim Width' || column.key === 'Page Extent') && value) {
                const numValue = parseFloat(value);
                if (isNaN(numValue) || numValue <= 0) {
                    results.errors.push(`Row ${rowNumber}: ${column.label} must be a positive number`);
                    hasErrors = true;
                }
            }
        });

        if (!hasErrors) {
            results.validRows++;
        }
    });

    validationResults = results;
    displayValidationResults(results);
    updateExportButtonValidation();
}

function displayValidationResults(results) {
    const validationSection = document.getElementById('validationSection');
    const validationContent = document.getElementById('validationContent');
    
    let statusClass = 'success';
    let statusText = 'All data valid';
    
    if (results.errors.length > 0) {
        statusClass = 'danger';
        statusText = `${results.errors.length} errors found`;
    } else if (results.warnings.length > 0) {
        statusClass = 'warning';
        statusText = `${results.warnings.length} warnings found`;
    }

    let html = `
        <div class="alert alert-${statusClass}">
            <h6><i class="bi bi-${statusClass === 'success' ? 'check-circle' : statusClass === 'warning' ? 'exclamation-triangle' : 'x-circle'}"></i> Validation Results</h6>
            <p class="mb-0"><strong>Status:</strong> ${statusText}</p>
            <p class="mb-0"><strong>Valid rows:</strong> ${results.validRows} / ${results.totalRows}</p>
            ${results.valueMappingsApplied > 0 ? `<p class="mb-0"><strong>Value mappings applied:</strong> ${results.valueMappingsApplied}</p>` : ''}
        </div>
    `;

    if (results.errors.length > 0) {
        html += `
            <div class="mt-3">
                <h6 class="text-danger">Errors (${results.errors.length})</h6>
                <div class="alert alert-danger">
                    ${results.errors.slice(0, 20).map(error => `<div class="validation-error">• ${error}</div>`).join('')}
                    ${results.errors.length > 20 ? `<div class="text-muted">... and ${results.errors.length - 20} more errors</div>` : ''}
                </div>
            </div>
        `;
    }

    if (results.warnings.length > 0) {
        html += `
            <div class="mt-3">
                <h6 class="text-warning">Warnings (${results.warnings.length})</h6>
                <div class="alert alert-warning">
                    ${results.warnings.slice(0, 10).map(warning => `<div class="validation-error">• ${warning}</div>`).join('')}
                    ${results.warnings.length > 10 ? `<div class="text-muted">... and ${results.warnings.length - 10} more warnings</div>` : ''}
                </div>
            </div>
        `;
    }

    validationContent.innerHTML = html;
    validationSection.classList.remove('d-none');
    
    document.getElementById('validationStatus').textContent = statusText;
    document.getElementById('validationStatus').className = `validation-${statusClass}`;
}

function updateExportButtonValidation() {
    const exportBtn = document.getElementById('exportBtn');
    if (validationResults && validationResults.errors.length > 0) {
        exportBtn.classList.remove('btn-success');
        exportBtn.classList.add('btn-warning');
        exportBtn.innerHTML = '<i class="bi bi-exclamation-triangle me-2"></i>Export with Errors';
    } else {
        exportBtn.classList.remove('btn-warning');
        exportBtn.classList.add('btn-success');
        exportBtn.innerHTML = '<i class="bi bi-file-earmark-arrow-down me-2"></i>Export Remapped Excel File';
    }
}

function exportRemappedFile() {
    if (!uploadedData || (Object.keys(columnMappings).length === 0 && Object.keys(fixedValues).length === 0)) {
        showAlert('No data to export', 'warning');
        return;
    }

    try {
        const includeHeaders = document.getElementById('includeHeaders').checked;
        const validateData = document.getElementById('validateData').checked;
        const applyValueMappingsOption = document.getElementById('applyValueMappings').checked;
        
        // Create remapped data in the exact column order
        const remappedData = [];
        
        // Sort columns by order for export
        const sortedColumns = [...requiredColumns].sort((a, b) => a.order - b.order);
        
        if (includeHeaders) {
            const headers = sortedColumns.map(col => col.label);
            remappedData.push(headers);
        }
        
        let invalidRows = 0;
        let appliedMappings = 0;
        
        uploadedData.forEach((row, index) => {
            const newRow = sortedColumns.map(column => {
                let value = '';
                
                // Check for fixed value first
                if (fixedValues[column.key]) {
                    value = fixedValues[column.key];
                } else if (columnMappings[column.key]) {
                    value = row[columnMappings[column.key]] || '';
                    
                    // Apply value mappings if enabled
                    if (applyValueMappingsOption && column.hasValueMapping) {
                        const mappedValue = applyValueMapping(value, column.key);
                        if (mappedValue !== value) {
                            appliedMappings++;
                        }
                        value = mappedValue;
                    }
                }
                
                // Basic validation if enabled
                if (validateData && column.required && !value) {
                    invalidRows++;
                }
                
                return value;
            });
            
            remappedData.push(newRow);
        });
        
        // Create workbook with proper ISBN formatting
        const worksheet = XLSX.utils.aoa_to_sheet(remappedData);
        
        // Format ISBN column as number with no decimals
        const isbnColumnIndex = sortedColumns.findIndex(col => col.key === 'ISBN');
        if (isbnColumnIndex !== -1) {
            const startRow = includeHeaders ? 1 : 0;
            const endRow = remappedData.length - 1;
            
            for (let i = startRow; i <= endRow; i++) {
                const cellAddress = XLSX.utils.encode_cell({ r: i, c: isbnColumnIndex });
                if (worksheet[cellAddress]) {
                    const isbn = worksheet[cellAddress].v;
                    if (isbn) {
                        // Convert to number and set format
                        worksheet[cellAddress].t = 'n'; // number type
                        worksheet[cellAddress].v = parseFloat(isbn.toString().replace(/[-\s]/g, ''));
                        worksheet[cellAddress].z = '0'; // number format with no decimals
                    }
                }
            }
        }
        
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'POD_Template_Data');
        
        // Generate filename
        const timestamp = new Date().toISOString().replace(/[:.]/g, '_').replace('T', '_').replace('Z', '');
        const filename = `POD_Remapped_${timestamp}.xlsx`;
        
        // Download file
        XLSX.writeFile(workbook, filename);
        
        let message = `Successfully exported ${remappedData.length - (includeHeaders ? 1 : 0)} rows`;
        if (appliedMappings > 0) {
            message += ` (${appliedMappings} value mappings applied)`;
        }
        if (validateData && invalidRows > 0) {
            message += ` (${invalidRows} rows may have missing required data)`;
        }
        
        showAlert(message, 'success');
        
    } catch (error) {
        console.error('Error exporting file:', error);
        showAlert('Error exporting file', 'danger');
    }
}

function downloadTemplate() {
    // Create template with exact column order and sample data
    const templateData = [
        ['ISBN', 'Title', 'Trim Height', 'Trim Width', 'Paper Type', 'Binding Style', 'Page Extent', 'Lamination'],
        ['9781234567890', 'Sample Book Title - Max 58 Characters', '198', '129', 'Amber Preprint 80 gsm', 'Limp', '320', 'Gloss']
    ];
    
    const worksheet = XLSX.utils.aoa_to_sheet(templateData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'POD_Template');
    
    XLSX.writeFile(workbook, 'POD_Template_Format.xlsx');
    showAlert('Template downloaded successfully', 'success');
}

function resetApp() {
    uploadedData = null;
    columnMappings = {};
    fixedValues = {};
    sourceColumns = [];
    validationResults = null;
    initializeValueMappings();
    
    document.getElementById('fileInput').value = '';
    document.getElementById('fileInfoSection').classList.add('d-none');
    document.getElementById('mappingSection').classList.add('d-none');
    document.getElementById('valueMappingSection').classList.add('d-none');
    document.getElementById('validationSection').classList.add('d-none');
    document.getElementById('exportSection').classList.add('d-none');
    
    showAlert('App reset successfully', 'info');
}

function showAlert(message, type = 'info', duration = 4000) {
    const alertBox = document.getElementById('alertBox');
    const alertMessage = document.getElementById('alertMessage');
    
    alertBox.className = `alert alert-${type} fade show`;
    alertMessage.textContent = message;
    alertBox.classList.remove('d-none');
    
    if (window.alertTimeout) {
        clearTimeout(window.alertTimeout);
    }
    
    window.alertTimeout = setTimeout(() => {
        alertBox.classList.add('fade-out');
        setTimeout(() => {
            alertBox.classList.add('d-none');
            alertBox.classList.remove('fade-out');
        }, 500);
    }, duration);
}

// Configuration Management Functions
function updateConfigSection() {
    const hasColumnMappings = Object.keys(columnMappings).length > 0;
    const hasFixedValues = Object.keys(fixedValues).length > 0;
    const hasValueMappings = Object.values(valueMappings).some(mapping => Object.keys(mapping).length > 0);
    
    const configSection = document.getElementById('configSection');
    const saveConfigBtn = document.getElementById('saveConfigBtn');
    const saveConfigBtn2 = document.getElementById('saveConfigBtn2');
    
    if (hasColumnMappings || hasFixedValues) {
        configSection.classList.remove('d-none');
        saveConfigBtn.disabled = false;
        saveConfigBtn2.disabled = false;
        updateConfigSummary();
    } else {
        configSection.classList.add('d-none');
        saveConfigBtn.disabled = true;
        saveConfigBtn2.disabled = true;
    }
}

function updateConfigSummary() {
    const summaryElement = document.getElementById('currentConfigSummary');
    const columnCount = Object.keys(columnMappings).length;
    const fixedCount = Object.keys(fixedValues).length;
    const valueCount = Object.values(valueMappings).reduce((sum, mapping) => sum + Object.keys(mapping).length, 0);
    
    const requiredMapped = requiredColumns.filter(col => 
        col.required && (columnMappings[col.key] || fixedValues[col.key])
    ).length;
    const totalRequired = requiredColumns.filter(col => col.required).length;
    
    let statusBadge = '';
    if (requiredMapped === totalRequired) {
        statusBadge = '<span class="badge bg-success">Complete</span>';
    } else {
        statusBadge = '<span class="badge bg-warning">Incomplete</span>';
    }
    
    summaryElement.innerHTML = `
        <div class="small">
            <p class="mb-1"><strong>Status:</strong> ${statusBadge}</p>
            <p class="mb-1"><strong>Column mappings:</strong> ${columnCount} (${requiredMapped}/${totalRequired} required)</p>
            <p class="mb-1"><strong>Fixed values:</strong> ${fixedCount}</p>
            <p class="mb-1"><strong>Value mappings:</strong> ${valueCount}</p>
            ${currentConfigName ? `<p class="mb-0"><strong>Based on:</strong> ${currentConfigName}</p>` : ''}
        </div>
    `;
}

function saveConfig() {
    if (Object.keys(columnMappings).length === 0 && Object.keys(fixedValues).length === 0) {
        showAlert('No column mappings or fixed values to save', 'warning');
        return;
    }
    
    // Update modal with current stats
    document.getElementById('modalColumnCount').textContent = Object.keys(columnMappings).length;
    document.getElementById('modalValueCount').textContent = Object.values(valueMappings).reduce((sum, mapping) => sum + Object.keys(mapping).length, 0);
    document.getElementById('modalFixedCount').textContent = Object.keys(fixedValues).length;
    
    // Generate default name
    const timestamp = new Date().toLocaleDateString();
    document.getElementById('configName').value = `POD Mapping Config - ${timestamp}`;
    
    // Show modal
    const modal = new bootstrap.Modal(document.getElementById('saveConfigModal'));
    modal.show();
}

function downloadConfig() {
    const configName = document.getElementById('configName').value.trim();
    const configDescription = document.getElementById('configDescription').value.trim();
    
    if (!configName) {
        showAlert('Please enter a configuration name', 'warning');
        return;
    }
    
    const config = {
        name: configName,
        description: configDescription,
        version: '1.1',
        createdAt: new Date().toISOString(),
        columnMappings: columnMappings,
        fixedValues: fixedValues,
        valueMappings: valueMappings,
        exportSettings: {
            includeHeaders: document.getElementById('includeHeaders')?.checked || true,
            validateData: document.getElementById('validateData')?.checked || true,
            applyValueMappings: document.getElementById('applyValueMappings')?.checked || true
        },
        requiredColumns: requiredColumns.map(col => ({ key: col.key, label: col.label, required: col.required }))
    };
    
    try {
        const blob = new Blob([JSON.stringify(config, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        
        // Clean filename
        const cleanName = configName.replace(/[^a-z0-9]/gi, '_').toLowerCase();
        const filename = `pod_mapping_config_${cleanName}.json`;
        
        link.href = url;
        link.download = filename;
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        URL.revokeObjectURL(url);
        
        currentConfigName = configName;
        updateConfigSummary();
        
        // Close modal
        const modal = bootstrap.Modal.getInstance(document.getElementById('saveConfigModal'));
        modal.hide();
        
        showAlert(`Configuration "${configName}" saved successfully`, 'success');
        
    } catch (error) {
        console.error('Error saving configuration:', error);
        showAlert('Error saving configuration', 'danger');
    }
}

function loadConfig() {
    // Create hidden file input for config loading
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.json';
    input.style.display = 'none';
    
    input.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;
        
        const reader = new FileReader();
        reader.onload = function(event) {
            try {
                const config = JSON.parse(event.target.result);
                applyConfig(config);
            } catch (error) {
                console.error('Error loading configuration:', error);
                showAlert('Error loading configuration file: Invalid format', 'danger');
            }
        };
        reader.readAsText(file);
        
        // Remove the input element
        document.body.removeChild(input);
    });
    
    document.body.appendChild(input);
    input.click();
}

function applyConfig(config) {
    try {
        // Validate config structure
        if (!config.columnMappings && !config.fixedValues) {
            throw new Error('Invalid configuration format');
        }
        
        // Apply column mappings
        columnMappings = { ...config.columnMappings } || {};
        
        // Apply fixed values
        fixedValues = { ...config.fixedValues } || {};
        
        // Apply value mappings
        valueMappings = { ...config.valueMappings } || {};
        
        // Apply export settings if available
        if (config.exportSettings) {
            const includeHeadersEl = document.getElementById('includeHeaders');
            const validateDataEl = document.getElementById('validateData');
            const applyValueMappingsEl = document.getElementById('applyValueMappings');
            
            if (includeHeadersEl) includeHeadersEl.checked = config.exportSettings.includeHeaders !== false;
            if (validateDataEl) validateDataEl.checked = config.exportSettings.validateData !== false;
            if (applyValueMappingsEl) applyValueMappingsEl.checked = config.exportSettings.applyValueMappings !== false;
        }
        
        currentConfigName = config.name || 'Loaded Configuration';
        
        // Update UI if data is loaded
        if (uploadedData && sourceColumns.length > 0) {
            updateMappingSelections();
            updateValueMappings();
            updatePreviewData();
        }
        
        updateExportButtonState();
        updateConfigSection();
        
        showAlert(`Configuration "${currentConfigName}" loaded successfully`, 'success');
        
    } catch (error) {
        console.error('Error applying configuration:', error);
        showAlert('Error applying configuration: ' + error.message, 'danger');
    }
}

function updateMappingSelections() {
    // Update column mapping dropdowns
    Object.keys(columnMappings).forEach(columnKey => {
        const selectElement = document.getElementById(`mapping_${columnKey}`);
        if (selectElement && sourceColumns.includes(columnMappings[columnKey])) {
            selectElement.value = columnMappings[columnKey];
            
            // Update preview
            const previewElement = document.getElementById(`preview_${columnKey}`);
            if (previewElement && uploadedData) {
                const sampleValues = uploadedData.slice(0, 5).map(row => row[columnMappings[columnKey]]);
                previewElement.innerHTML = `Sample: ${sampleValues.join(', ')}`;
            }
        } else if (selectElement) {
            // Column not found in current data
            selectElement.value = '';
            const previewElement = document.getElementById(`preview_${columnKey}`);
            if (previewElement) {
                previewElement.innerHTML = '<span class="text-warning">Mapped column not found in current file</span>';
            }
        }
    });
    
    // Update fixed values
    Object.keys(fixedValues).forEach(columnKey => {
        const selectElement = document.getElementById(`mapping_${columnKey}`);
        const fixedInputElement = document.getElementById(`fixed_${columnKey}`);
        const previewElement = document.getElementById(`preview_${columnKey}`);
        
        if (selectElement) {
            selectElement.value = '__FIXED__';
            if (fixedInputElement) {
                fixedInputElement.classList.remove('d-none');
                fixedInputElement.value = fixedValues[columnKey];
            }
            if (previewElement) {
                previewElement.innerHTML = `<span class="badge bg-warning">Fixed: ${fixedValues[columnKey]}</span>`;
            }
        }
    });
}