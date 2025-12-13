/**
 * VBA to Python Converter - Frontend JavaScript
 */

// DOM Elements
const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const uploadStatus = document.getElementById('uploadStatus');
const modulesSection = document.getElementById('modulesSection');
const modulesList = document.getElementById('modulesList');
const optionsSection = document.getElementById('optionsSection');
const comparisonSection = document.getElementById('comparisonSection');
const convertAllBtn = document.getElementById('convertAllBtn');
const convertPastedBtn = document.getElementById('convertPastedBtn');
const vbaInput = document.getElementById('vbaInput');
const vbaCode = document.getElementById('vbaCode');
const pythonCode = document.getElementById('pythonCode');
const vbaModuleName = document.getElementById('vbaModuleName');
const conversionStatus = document.getElementById('conversionStatus');
const notesSection = document.getElementById('notesSection');
const conversionNotes = document.getElementById('conversionNotes');
const copyPythonBtn = document.getElementById('copyPythonBtn');
const downloadPythonBtn = document.getElementById('downloadPythonBtn');
const loadingOverlay = document.getElementById('loadingOverlay');
const loadingMessage = document.getElementById('loadingMessage');
const targetLibrary = document.getElementById('targetLibrary');

// State
let extractedModules = [];
let currentModuleName = 'converted_module';
let currentPythonCode = '';

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    setupDragAndDrop();
    setupEventListeners();
});

// Drag and Drop Setup
function setupDragAndDrop() {
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, preventDefaults, false);
    });

    ['dragenter', 'dragover'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.add('dragover');
        });
    });

    ['dragleave', 'drop'].forEach(eventName => {
        dropZone.addEventListener(eventName, () => {
            dropZone.classList.remove('dragover');
        });
    });

    dropZone.addEventListener('drop', handleDrop);
    dropZone.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', handleFileSelect);
}

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

function handleDrop(e) {
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        uploadFile(files[0]);
    }
}

function handleFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        uploadFile(files[0]);
    }
}

// Event Listeners
function setupEventListeners() {
    convertAllBtn.addEventListener('click', convertAllModules);
    convertPastedBtn.addEventListener('click', convertPastedVBA);
    copyPythonBtn.addEventListener('click', copyPythonToClipboard);
    downloadPythonBtn.addEventListener('click', downloadPythonFile);
    
    // Formula extraction
    const extractFormulasBtn = document.getElementById('extractFormulasBtn');
    const formulaFileInput = document.getElementById('formulaFileInput');
    if (extractFormulasBtn) {
        extractFormulasBtn.addEventListener('click', () => formulaFileInput.click());
        formulaFileInput.addEventListener('change', handleFormulaFileSelect);
    }
    
    // Data export
    const exportDataBtn = document.getElementById('exportDataBtn');
    const dataFileInput = document.getElementById('dataFileInput');
    if (exportDataBtn) {
        exportDataBtn.addEventListener('click', () => dataFileInput.click());
        dataFileInput.addEventListener('change', handleDataFileSelect);
    }
    
    // Workbook analysis
    const analyzeWorkbookBtn = document.getElementById('analyzeWorkbookBtn');
    const analysisFileInput = document.getElementById('analysisFileInput');
    if (analyzeWorkbookBtn) {
        analyzeWorkbookBtn.addEventListener('click', () => analysisFileInput.click());
        analysisFileInput.addEventListener('change', handleAnalysisFileSelect);
    }
    
    // Copy buttons for new sections
    const copyDataCodeBtn = document.getElementById('copyDataCodeBtn');
    const copyAnalysisCodeBtn = document.getElementById('copyAnalysisCodeBtn');
    if (copyDataCodeBtn) copyDataCodeBtn.addEventListener('click', () => copyCode('dataCode'));
    if (copyAnalysisCodeBtn) copyAnalysisCodeBtn.addEventListener('click', () => copyCode('analysisCode'));
}

// File Upload
async function uploadFile(file) {
    const allowedExtensions = ['xlsm', 'xls', 'xlsb', 'xla', 'xlam'];
    const extension = file.name.split('.').pop().toLowerCase();

    if (!allowedExtensions.includes(extension)) {
        showStatus('error', `Invalid file type. Allowed: ${allowedExtensions.join(', ')}`);
        return;
    }

    showLoading('Extracting VBA code...');

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();
        hideLoading();

        if (data.error) {
            showStatus('error', data.error);
            return;
        }

        if (data.warning) {
            showStatus('warning', data.warning);
            return;
        }

        extractedModules = data.modules;
        showStatus('success', `Successfully extracted ${data.modules.length} VBA module(s) from ${data.filename}`);
        displayModules(data.modules);

    } catch (error) {
        hideLoading();
        showStatus('error', `Upload failed: ${error.message}`);
    }
}

// Display Extracted Modules
function displayModules(modules) {
    modulesSection.classList.remove('hidden');
    optionsSection.classList.remove('hidden');
    modulesList.innerHTML = '';

    modules.forEach((module, index) => {
        const card = document.createElement('div');
        card.className = 'module-card';
        card.innerHTML = `
            <div class="module-info">
                <h4>${escapeHtml(module.name)}</h4>
                <span class="module-type">${escapeHtml(module.type)}</span>
            </div>
            <div class="module-actions">
                <button class="btn btn-secondary btn-small" onclick="viewModule(${index})">👁️ View</button>
                <button class="btn btn-primary btn-small" onclick="convertModule(${index})">🔄 Convert</button>
            </div>
        `;
        modulesList.appendChild(card);
    });
}

// View Module
function viewModule(index) {
    const module = extractedModules[index];
    currentModuleName = module.name;

    comparisonSection.classList.remove('hidden');
    vbaModuleName.textContent = module.name;
    vbaCode.textContent = module.code;
    pythonCode.textContent = '// Click "Convert" to generate Python code';
    conversionStatus.textContent = '';
    conversionStatus.className = 'status-badge';
    notesSection.classList.add('hidden');

    // Re-highlight code
    Prism.highlightElement(vbaCode);
    
    // Scroll to comparison section
    comparisonSection.scrollIntoView({ behavior: 'smooth' });
}

// Convert Single Module
async function convertModule(index) {
    const module = extractedModules[index];
    await convertVBACode(module.code, module.name);
}

// Convert All Modules
async function convertAllModules() {
    if (extractedModules.length === 0) {
        showStatus('warning', 'No modules to convert');
        return;
    }

    showLoading('Converting all modules...');

    try {
        const response = await fetch('/api/convert-all', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                modules: extractedModules,
                target_library: targetLibrary.value
            })
        });

        const data = await response.json();
        hideLoading();

        if (data.error) {
            showStatus('error', data.error);
            return;
        }

        // Show first converted module
        if (data.converted_modules && data.converted_modules.length > 0) {
            const first = data.converted_modules[0];
            displayConversion(first.original_code, first.python_code, first.name, first.conversion_notes);
            showStatus('success', `Successfully converted ${data.converted_modules.length} module(s)`);
        }

    } catch (error) {
        hideLoading();
        showStatus('error', `Conversion failed: ${error.message}`);
    }
}

// Convert Pasted VBA
async function convertPastedVBA() {
    const vbaCodeText = vbaInput.value.trim();
    
    if (!vbaCodeText) {
        showStatus('warning', 'Please paste some VBA code first');
        return;
    }

    await convertVBACode(vbaCodeText, 'pasted_module');
}

// Core Conversion Function
async function convertVBACode(vbaCodeText, moduleName) {
    showLoading('Converting VBA to Python...');
    
    comparisonSection.classList.remove('hidden');
    vbaModuleName.textContent = moduleName;
    vbaCode.textContent = vbaCodeText;
    conversionStatus.textContent = 'Converting...';
    conversionStatus.className = 'status-badge converting';
    pythonCode.textContent = '// Converting...';
    
    Prism.highlightElement(vbaCode);

    try {
        const response = await fetch('/api/convert', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                vba_code: vbaCodeText,
                module_name: moduleName,
                target_library: targetLibrary.value
            })
        });

        const data = await response.json();
        hideLoading();

        if (data.error) {
            conversionStatus.textContent = 'Error';
            conversionStatus.className = 'status-badge error';
            pythonCode.textContent = `# Error: ${data.error}`;
            showStatus('error', data.error);
            return;
        }

        displayConversion(vbaCodeText, data.python_code, moduleName, data.conversion_notes);
        showStatus('success', 'VBA code successfully converted to Python!');

    } catch (error) {
        hideLoading();
        conversionStatus.textContent = 'Error';
        conversionStatus.className = 'status-badge error';
        pythonCode.textContent = `# Error: ${error.message}`;
        showStatus('error', `Conversion failed: ${error.message}`);
    }
}

// Display Conversion Result
function displayConversion(originalVba, convertedPython, moduleName, notes) {
    currentModuleName = moduleName;
    currentPythonCode = convertedPython;

    comparisonSection.classList.remove('hidden');
    vbaModuleName.textContent = moduleName;
    vbaCode.textContent = originalVba;
    pythonCode.textContent = convertedPython;
    conversionStatus.textContent = 'Converted';
    conversionStatus.className = 'status-badge success';

    // Re-highlight code
    Prism.highlightElement(vbaCode);
    Prism.highlightElement(pythonCode);

    // Display notes if any
    if (notes && notes.length > 0) {
        notesSection.classList.remove('hidden');
        conversionNotes.innerHTML = notes.map(note => `<li>${escapeHtml(note)}</li>`).join('');
    } else {
        notesSection.classList.add('hidden');
    }

    // Scroll to comparison section
    comparisonSection.scrollIntoView({ behavior: 'smooth' });
}

// Copy to Clipboard
async function copyPythonToClipboard() {
    if (!currentPythonCode) {
        showStatus('warning', 'No Python code to copy');
        return;
    }

    try {
        await navigator.clipboard.writeText(currentPythonCode);
        copyPythonBtn.textContent = '✅ Copied!';
        setTimeout(() => {
            copyPythonBtn.textContent = '📋 Copy Python';
        }, 2000);
    } catch (error) {
        showStatus('error', 'Failed to copy to clipboard');
    }
}

// Download Python File
function downloadPythonFile() {
    if (!currentPythonCode) {
        showStatus('warning', 'No Python code to download');
        return;
    }

    const blob = new Blob([currentPythonCode], { type: 'text/x-python' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${currentModuleName.replace(/[^a-zA-Z0-9_]/g, '_')}.py`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

// UI Helpers
function showStatus(type, message) {
    uploadStatus.className = `status-message ${type}`;
    uploadStatus.textContent = message;
    uploadStatus.classList.remove('hidden');
}

function showLoading(message) {
    loadingMessage.textContent = message;
    loadingOverlay.classList.remove('hidden');
}

function hideLoading() {
    loadingOverlay.classList.add('hidden');
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

// ============================================================================
// FORMULA EXTRACTION
// ============================================================================

function handleFormulaFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        extractFormulas(files[0]);
    }
}

async function extractFormulas(file) {
    showLoading('Extracting formulas from Excel file...');

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('/api/extract-formulas', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();
        hideLoading();

        if (data.error) {
            showStatus('error', data.error);
            return;
        }

        displayFormulaResults(data);
        showStatus('success', `Extracted ${data.formulas.length} formulas from ${data.filename}`);

    } catch (error) {
        hideLoading();
        showStatus('error', `Formula extraction failed: ${error.message}`);
    }
}

function displayFormulaResults(data) {
    const formulaResults = document.getElementById('formulaResults');
    const formulaStats = document.getElementById('formulaStats');
    const formulaList = document.getElementById('formulaList');
    
    formulaResults.classList.remove('hidden');
    
    // Display statistics
    const stats = data.statistics;
    formulaStats.innerHTML = `
        <div class="stats-grid">
            <div class="stat-card">
                <h4>${stats.total_formulas}</h4>
                <p>Total Formulas</p>
            </div>
            <div class="stat-card">
                <h4>${stats.sheets_with_formulas}</h4>
                <p>Sheets</p>
            </div>
            <div class="stat-card">
                <h4>${stats.unique_functions_used}</h4>
                <p>Unique Functions</p>
            </div>
        </div>
        <div class="common-functions">
            <h4>Most Common Functions:</h4>
            <ul>
                ${stats.most_common_functions.slice(0, 5).map(([func, count]) => 
                    `<li>${func}: ${count} times</li>`
                ).join('')}
            </ul>
        </div>
    `;
    
    // Display formula list
    formulaList.innerHTML = '';
    data.formulas.forEach((formula, index) => {
        const card = document.createElement('div');
        card.className = 'formula-card';
        card.innerHTML = `
            <div class="formula-info">
                <strong>${formula.sheet_name} - ${formula.cell_address}</strong>
                <p class="formula-text">${escapeHtml(formula.formula)}</p>
                <div class="formula-functions">
                    ${formula.functions.map(f => `<span class="badge">${f}</span>`).join('')}
                </div>
            </div>
            <button class="btn btn-primary btn-small" onclick="convertFormula('${escapeHtml(formula.formula)}', '${formula.cell_address}', '${formula.sheet_name}')">
                🔄 Convert to Python
            </button>
        `;
        formulaList.appendChild(card);
    });
}

async function convertFormula(formula, cellAddress, sheetName) {
    showLoading('Converting formula to Python...');

    try {
        const response = await fetch('/api/convert-formula', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                formula: formula,
                cell_address: cellAddress,
                sheet_name: sheetName
            })
        });

        const data = await response.json();
        hideLoading();

        if (data.error) {
            showStatus('error', data.error);
            return;
        }

        // Display in comparison section
        displayConversion(
            `Excel Formula in ${sheetName}!${cellAddress}:\n${formula}`,
            data.python_code,
            `formula_${cellAddress}`,
            data.conversion_notes
        );
        showStatus('success', 'Formula converted to Python!');

    } catch (error) {
        hideLoading();
        showStatus('error', `Formula conversion failed: ${error.message}`);
    }
}

// ============================================================================
// DATA EXPORT
// ============================================================================

function handleDataFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        exportDataToCode(files[0]);
    }
}

async function exportDataToCode(file) {
    showLoading('Exporting Excel data to Python...');

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('/api/export-data', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();
        hideLoading();

        if (data.error) {
            showStatus('error', data.error);
            return;
        }

        displayDataResults(data);
        showStatus('success', `Exported data from ${data.filename}`);

    } catch (error) {
        hideLoading();
        showStatus('error', `Data export failed: ${error.message}`);
    }
}

function displayDataResults(data) {
    const dataResults = document.getElementById('dataResults');
    const dataCode = document.getElementById('dataCode');
    const dataMetadata = document.getElementById('dataMetadata');
    
    dataResults.classList.remove('hidden');
    
    // Display generated code
    dataCode.textContent = data.python_code;
    Prism.highlightElement(dataCode);
    
    // Display metadata
    const meta = data.metadata;
    dataMetadata.innerHTML = `
        <h4>Data Summary</h4>
        <div class="stats-grid">
            <div class="stat-card">
                <h4>${meta.total_sheets}</h4>
                <p>Sheets</p>
            </div>
            <div class="stat-card">
                <h4>${meta.total_rows}</h4>
                <p>Total Rows</p>
            </div>
            <div class="stat-card">
                <h4>${meta.total_columns}</h4>
                <p>Total Columns</p>
            </div>
        </div>
        <div class="sheets-list">
            <h4>Sheets:</h4>
            <ul>
                ${meta.sheets.map(sheet => 
                    `<li><strong>${sheet.name}</strong>: ${sheet.rows} rows × ${sheet.columns} columns</li>`
                ).join('')}
            </ul>
        </div>
    `;
}

// ============================================================================
// COMPLETE WORKBOOK ANALYSIS
// ============================================================================

function handleAnalysisFileSelect(e) {
    const files = e.target.files;
    if (files.length > 0) {
        analyzeCompleteWorkbook(files[0]);
    }
}

async function analyzeCompleteWorkbook(file) {
    showLoading('Analyzing complete workbook (VBA + Formulas + Data)...');

    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('/api/analyze-workbook', {
            method: 'POST',
            body: formData
        });

        const data = await response.json();
        hideLoading();

        if (data.error) {
            showStatus('error', data.error);
            return;
        }

        displayAnalysisResults(data);
        showStatus('success', `Complete analysis of ${data.filename} finished!`);

    } catch (error) {
        hideLoading();
        showStatus('error', `Workbook analysis failed: ${error.message}`);
    }
}

function displayAnalysisResults(data) {
    const analysisResults = document.getElementById('analysisResults');
    const analysisSummary = document.getElementById('analysisSummary');
    const analysisCode = document.getElementById('analysisCode');
    const analysisReport = document.getElementById('analysisReport');
    
    analysisResults.classList.remove('hidden');
    
    // Display summary
    analysisSummary.innerHTML = `
        <div class="stats-grid">
            <div class="stat-card ${data.has_vba ? 'highlight' : ''}">
                <h4>${data.vba_modules_count}</h4>
                <p>VBA Modules</p>
            </div>
            <div class="stat-card ${data.has_formulas ? 'highlight' : ''}">
                <h4>${data.formulas_count}</h4>
                <p>Formulas</p>
            </div>
            <div class="stat-card">
                <h4>${data.sheets_count}</h4>
                <p>Data Sheets</p>
            </div>
        </div>
    `;
    
    // Display generated Python script
    analysisCode.textContent = data.python_script;
    Prism.highlightElement(analysisCode);
    
    // Display detailed report
    analysisReport.textContent = data.report;
    
    // Scroll to results
    analysisResults.scrollIntoView({ behavior: 'smooth' });
}

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

async function copyCode(elementId) {
    const element = document.getElementById(elementId);
    if (!element) return;
    
    const code = element.textContent;
    if (!code) {
        showStatus('warning', 'No code to copy');
        return;
    }

    try {
        await navigator.clipboard.writeText(code);
        showStatus('success', 'Code copied to clipboard!');
    } catch (error) {
        showStatus('error', 'Failed to copy to clipboard');
    }
}