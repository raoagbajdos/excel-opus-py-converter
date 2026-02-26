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
let lastUploadedFile = null;  // Store last file for reuse across sections
let batchConvertedModules = [];  // Store batch results for ZIP download
let conversionHistory = JSON.parse(localStorage.getItem('conversionHistory') || '[]');
let currentBatchIndex = 0;  // Track active module in batch navigator

// Helper: extract error message from FastAPI / generic response
function getErrorMessage(data) {
    return data.detail || data.error || data.message || 'Unknown error';
}

// Create ARIA live region for screen reader announcements
function createLiveRegion() {
    const region = document.createElement('div');
    region.id = 'srAnnouncements';
    region.setAttribute('aria-live', 'polite');
    region.setAttribute('aria-atomic', 'true');
    region.classList.add('sr-only');
    document.body.appendChild(region);
    return region;
}

// Announce message to screen readers via live region
function announceToSR(message, priority = 'polite') {
    const region = document.getElementById('srAnnouncements');
    if (!region) return;
    region.setAttribute('aria-live', priority);
    region.textContent = '';
    // Small timeout to ensure DOM update triggers announcement
    setTimeout(() => { region.textContent = message; }, 50);
}

// Initialize
document.addEventListener('DOMContentLoaded', () => {
    createLiveRegion();
    initTheme();
    setupDragAndDrop();
    setupEventListeners();
    setupKeyboardShortcuts();
    setupCollapsibleSections();
    setupResizablePanels();
    displayHistory();
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
    // Keyboard support: Enter/Space to trigger file browse on dropzone
    dropZone.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' || e.key === ' ') {
            e.preventDefault();
            fileInput.click();
        }
    });
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
        extractFormulasBtn.addEventListener('click', () => {
            if (lastUploadedFile) {
                extractFormulas(lastUploadedFile);
            } else {
                formulaFileInput.click();
            }
        });
        formulaFileInput.addEventListener('change', handleFormulaFileSelect);
    }
    
    // Data export
    const exportDataBtn = document.getElementById('exportDataBtn');
    const dataFileInput = document.getElementById('dataFileInput');
    if (exportDataBtn) {
        exportDataBtn.addEventListener('click', () => {
            if (lastUploadedFile) {
                exportDataToCode(lastUploadedFile);
            } else {
                dataFileInput.click();
            }
        });
        dataFileInput.addEventListener('change', handleDataFileSelect);
    }
    
    // Workbook analysis
    const analyzeWorkbookBtn = document.getElementById('analyzeWorkbookBtn');
    const analysisFileInput = document.getElementById('analysisFileInput');
    if (analyzeWorkbookBtn) {
        analyzeWorkbookBtn.addEventListener('click', () => {
            if (lastUploadedFile) {
                analyzeCompleteWorkbook(lastUploadedFile);
            } else {
                analysisFileInput.click();
            }
        });
        analysisFileInput.addEventListener('change', handleAnalysisFileSelect);
    }
    
    // Copy buttons for new sections
    const copyDataCodeBtn = document.getElementById('copyDataCodeBtn');
    const copyAnalysisCodeBtn = document.getElementById('copyAnalysisCodeBtn');
    if (copyDataCodeBtn) copyDataCodeBtn.addEventListener('click', () => copyCode('dataCode'));
    if (copyAnalysisCodeBtn) copyAnalysisCodeBtn.addEventListener('click', () => copyCode('analysisCode'));

    // Theme toggle
    const themeToggle = document.getElementById('themeToggle');
    if (themeToggle) themeToggle.addEventListener('click', toggleTheme);

    // Conversion history
    const clearHistoryBtn = document.getElementById('clearHistoryBtn');
    const toggleHistoryBtn = document.getElementById('toggleHistoryBtn');
    if (clearHistoryBtn) clearHistoryBtn.addEventListener('click', clearHistory);
    if (toggleHistoryBtn) toggleHistoryBtn.addEventListener('click', toggleHistoryPanel);

    // Download All as ZIP
    const downloadAllZipBtn = document.getElementById('downloadAllZipBtn');
    if (downloadAllZipBtn) downloadAllZipBtn.addEventListener('click', downloadAllAsZip);

    // Copy VBA code
    const copyVbaBtn = document.getElementById('copyVbaBtn');
    if (copyVbaBtn) copyVbaBtn.addEventListener('click', copyVBAToClipboard);

    // Export / Import history
    const exportHistoryBtn = document.getElementById('exportHistoryBtn');
    const importHistoryInput = document.getElementById('importHistoryInput');
    if (exportHistoryBtn) exportHistoryBtn.addEventListener('click', exportHistory);
    if (importHistoryInput) importHistoryInput.addEventListener('change', importHistory);

    // Module navigator tabs
    const prevModuleBtn = document.getElementById('prevModuleBtn');
    const nextModuleBtn = document.getElementById('nextModuleBtn');
    if (prevModuleBtn) prevModuleBtn.addEventListener('click', () => navigateBatchModule(currentBatchIndex - 1));
    if (nextModuleBtn) nextModuleBtn.addEventListener('click', () => navigateBatchModule(currentBatchIndex + 1));

    // Syntax validation (live as user types)
    const vbaTextarea = document.getElementById('vbaInput');
    if (vbaTextarea) {
        vbaTextarea.addEventListener('input', debounce(validateVBASyntax, 300));
        vbaTextarea.addEventListener('paste', () => setTimeout(validateVBASyntax, 50));
    }

    // History search & filter
    const historySearch = document.getElementById('historySearch');
    const historyEngineFilter = document.getElementById('historyEngineFilter');
    const historyStatusFilter = document.getElementById('historyStatusFilter');
    if (historySearch) historySearch.addEventListener('input', debounce(displayHistory, 200));
    if (historyEngineFilter) historyEngineFilter.addEventListener('change', displayHistory);
    if (historyStatusFilter) historyStatusFilter.addEventListener('change', displayHistory);

    // Diff highlight toggle
    const toggleDiffBtn = document.getElementById('toggleDiffBtn');
    if (toggleDiffBtn) toggleDiffBtn.addEventListener('click', toggleDiffHighlights);
}

// File Upload
async function uploadFile(file) {
    const allowedExtensions = ['xlsm', 'xls', 'xlsb', 'xlsx', 'xla', 'xlam'];
    const extension = file.name.split('.').pop().toLowerCase();

    if (!allowedExtensions.includes(extension)) {
        showStatus('error', `Invalid file type. Allowed: ${allowedExtensions.join(', ')}`);
        return;
    }

    // Store file for reuse in other sections
    lastUploadedFile = file;
    updateFileIndicator(file.name);

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

        if (!response.ok) {
            showStatus('error', getErrorMessage(data));
            return;
        }

        if (data.warning && (!data.modules || data.modules.length === 0)) {
            showStatus('warning', data.warning + ' Running full workbook analysis automatically...');
            // Automatically run workbook analysis instead of making user re-upload
            analyzeCompleteWorkbook(file);
            return;
        }

        if (data.warning) {
            showStatus('warning', data.warning);
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
    addLineNumbers(vbaCode);
    updateCodeStats('vbaStats', module.code, 'vba');
    updateCodeStats('pythonStats', '', 'python');
    
    // Scroll to comparison section
    comparisonSection.scrollIntoView({ behavior: 'smooth' });
}

// Convert Single Module
async function convertModule(index) {
    const module = extractedModules[index];
    await convertVBACode(module.code, module.name);
}

// Convert All Modules (with progress indicator)
async function convertAllModules() {
    if (extractedModules.length === 0) {
        showStatus('warning', 'No modules to convert');
        return;
    }

    const totalModules = extractedModules.length;
    showLoading('Converting all modules...');
    showBatchProgress(0, totalModules);
    const batchStartTime = performance.now();

    try {
        const providerSelect = document.getElementById('providerSelect');
        const provider = providerSelect ? providerSelect.value : null;

        // If only one module or using offline converter, do a single batch request
        // Otherwise, convert one at a time for progress feedback
        if (totalModules <= 1 || provider === 'offline') {
            const response = await fetch('/api/convert-all', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    modules: extractedModules,
                    target_library: targetLibrary.value,
                    provider: provider
                })
            });

            const data = await response.json();
            hideBatchProgress();
            hideLoading();

            if (!response.ok) {
                showStatus('error', getErrorMessage(data));
                return;
            }

            if (data.converted_modules && data.converted_modules.length > 0) {
                const batchElapsed = performance.now() - batchStartTime;
                finalizeBatchConversion(data.converted_modules, provider, batchElapsed);
            }
        } else {
            // Sequential conversion with per-module progress
            const converted = [];
            for (let i = 0; i < totalModules; i++) {
                const mod = extractedModules[i];
                updateBatchProgress(i, totalModules, mod.name);

                try {
                    const response = await fetch('/api/convert', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({
                            vba_code: mod.code,
                            module_name: mod.name,
                            target_library: targetLibrary.value,
                            provider: provider
                        })
                    });

                    const data = await response.json();
                    if (response.ok && data.success) {
                        converted.push({
                            name: mod.name,
                            type: mod.type,
                            original_code: mod.code,
                            python_code: data.python_code,
                            conversion_notes: data.conversion_notes || []
                        });
                    } else {
                        converted.push({
                            name: mod.name,
                            type: mod.type,
                            original_code: mod.code,
                            python_code: `# Conversion failed: ${getErrorMessage(data)}`,
                            conversion_notes: ['Conversion failed for this module']
                        });
                    }
                } catch (err) {
                    converted.push({
                        name: mod.name,
                        type: mod.type,
                        original_code: mod.code,
                        python_code: `# Error: ${err.message}`,
                        conversion_notes: ['Network error during conversion']
                    });
                }

                updateBatchProgress(i + 1, totalModules);
            }

            hideBatchProgress();
            hideLoading();
            const batchElapsed = performance.now() - batchStartTime;
            finalizeBatchConversion(converted, provider, batchElapsed);
        }

    } catch (error) {
        hideBatchProgress();
        hideLoading();
        showStatus('error', `Conversion failed: ${error.message}`);
    }
}

function finalizeBatchConversion(convertedModules, provider, totalDurationMs) {
    batchConvertedModules = convertedModules;
    currentBatchIndex = 0;
    const first = convertedModules[0];
    displayConversion(first.original_code, first.python_code, first.name, first.conversion_notes);

    const timeStr = totalDurationMs ? ` in ${formatTime(totalDurationMs)}` : '';
    showStatus('success', `Successfully converted ${convertedModules.length} module(s)${timeStr}`);
    if (totalDurationMs) showConversionTime(totalDurationMs);

    // Show module navigator tabs
    buildModuleTabs(convertedModules);

    // Show Download ZIP button
    const downloadAllZipBtn = document.getElementById('downloadAllZipBtn');
    if (downloadAllZipBtn) downloadAllZipBtn.classList.remove('hidden');

    // Add each to history
    const engine = provider || 'unknown';
    const perModuleMs = totalDurationMs ? totalDurationMs / convertedModules.length : null;
    convertedModules.forEach(m => {
        const success = !m.python_code.startsWith('# Conversion failed') && !m.python_code.startsWith('# Error');
        addToHistory(m.name, engine, success, perModuleMs);
    });
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
    const startTime = performance.now();
    
    comparisonSection.classList.remove('hidden');
    vbaModuleName.textContent = moduleName;
    vbaCode.textContent = vbaCodeText;
    conversionStatus.textContent = 'Converting...';
    conversionStatus.className = 'status-badge converting';
    pythonCode.textContent = '// Converting...';
    
    Prism.highlightElement(vbaCode);

    try {
        const providerSelect = document.getElementById('providerSelect');
        const response = await fetch('/api/convert', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                vba_code: vbaCodeText,
                module_name: moduleName,
                target_library: targetLibrary.value,
                provider: providerSelect ? providerSelect.value : null
            })
        });

        const data = await response.json();
        const elapsed = performance.now() - startTime;
        hideLoading();

        if (!response.ok) {
            conversionStatus.textContent = 'Error';
            conversionStatus.className = 'status-badge error';
            pythonCode.textContent = `# Error: ${getErrorMessage(data)}`;
            showStatus('error', getErrorMessage(data));
            return;
        }

        displayConversion(vbaCodeText, data.python_code, moduleName, data.conversion_notes);
        showConversionTime(elapsed);
        showStatus('success', `VBA code converted to Python in ${formatTime(elapsed)}`);
        announceToSR(`Conversion complete for ${moduleName} in ${formatTime(elapsed)}`);
        // Focus the Python code panel for keyboard users
        const pythonPanel = document.querySelector('.python-panel');
        if (pythonPanel) {
            pythonPanel.setAttribute('tabindex', '-1');
            pythonPanel.focus();
        }

        // Track in history
        const engine = data.engine || (document.getElementById('providerSelect')?.value) || 'unknown';
        addToHistory(moduleName, engine, true, elapsed);

    } catch (error) {
        const elapsed = performance.now() - startTime;
        hideLoading();
        conversionStatus.textContent = 'Error';
        conversionStatus.className = 'status-badge error';
        pythonCode.textContent = `# Error: ${error.message}`;
        showStatus('error', `Conversion failed: ${error.message}`);
        addToHistory(moduleName, 'unknown', false, elapsed);
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

    // Preserve time if already set, else default
    const existing = conversionStatus.dataset.time;
    if (existing) {
        conversionStatus.textContent = `Converted · ${existing}`;
    } else {
        conversionStatus.textContent = 'Converted';
    }
    conversionStatus.className = 'status-badge success';

    // Re-highlight code
    Prism.highlightElement(vbaCode);
    Prism.highlightElement(pythonCode);

    // Apply diff mapping highlights
    if (diffHighlightEnabled) {
        applyDiffHighlights(vbaCode, DIFF_MAPPINGS_VBA);
        applyDiffHighlights(pythonCode, DIFF_MAPPINGS_PYTHON);
    }

    // Add line numbers
    addLineNumbers(vbaCode);
    addLineNumbers(pythonCode);

    // Show code statistics
    updateCodeStats('vbaStats', originalVba, 'vba');
    updateCodeStats('pythonStats', convertedPython, 'python');

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
    announceToSR(message, type === 'error' ? 'assertive' : 'polite');
}

function showLoading(message) {
    loadingMessage.textContent = message;
    loadingOverlay.classList.remove('hidden');
    announceToSR(message, 'assertive');
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

        if (!response.ok) {
            showStatus('error', getErrorMessage(data));
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
                <strong>${escapeHtml(formula.sheet_name)} - ${escapeHtml(formula.cell_address)}</strong>
                <p class="formula-text">${escapeHtml(formula.formula)}</p>
                <div class="formula-functions">
                    ${formula.functions.map(f => `<span class="badge">${escapeHtml(f)}</span>`).join('')}
                </div>
            </div>
            <button class="btn btn-primary btn-small" data-formula-index="${index}">
                🔄 Convert to Python
            </button>
        `;
        formulaList.appendChild(card);
    });

    // Attach event listeners using data attributes (avoids escaping issues)
    formulaList.querySelectorAll('button[data-formula-index]').forEach(btn => {
        btn.addEventListener('click', () => {
            const i = parseInt(btn.dataset.formulaIndex, 10);
            const f = data.formulas[i];
            convertFormula(f.formula, f.cell_address, f.sheet_name);
        });
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

        if (!response.ok) {
            showStatus('error', getErrorMessage(data));
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

        if (!response.ok) {
            showStatus('error', getErrorMessage(data));
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

        if (!response.ok) {
            showStatus('error', getErrorMessage(data));
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

function updateFileIndicator(filename) {
    const indicator = document.getElementById('currentFileIndicator');
    const nameEl = document.getElementById('currentFileName');
    if (indicator && nameEl) {
        nameEl.textContent = filename;
        indicator.classList.remove('hidden');
    }
    // Update hints on other sections
    const analysisHint = document.getElementById('analysisHint');
    if (analysisHint) {
        analysisHint.textContent = `Using "${filename}" — click to analyze`;
    }
}

// ============================================================================
// DARK / LIGHT THEME TOGGLE
// ============================================================================

function initTheme() {
    const saved = localStorage.getItem('theme') || 'dark';
    document.documentElement.setAttribute('data-theme', saved);
    updateThemeIcon(saved);
}

function toggleTheme() {
    const current = document.documentElement.getAttribute('data-theme') || 'dark';
    const next = current === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('theme', next);
    updateThemeIcon(next);
}

function updateThemeIcon(theme) {
    const btn = document.getElementById('themeToggle');
    if (btn) btn.textContent = theme === 'dark' ? '🌙' : '☀️';
}

// ============================================================================
// CONVERSION HISTORY
// ============================================================================

function addToHistory(name, engine, success, durationMs) {
    conversionHistory.unshift({
        name,
        engine,
        success,
        timestamp: new Date().toISOString(),
        duration: durationMs ? Math.round(durationMs) : null
    });
    // Keep last 50 entries
    if (conversionHistory.length > 50) conversionHistory.length = 50;
    localStorage.setItem('conversionHistory', JSON.stringify(conversionHistory));
    displayHistory();
}

function displayHistory() {
    const historyList = document.getElementById('historyList');
    const historyCount = document.getElementById('historyCount');
    if (!historyList) return;

    // Apply search and filters
    const searchTerm = (document.getElementById('historySearch')?.value || '').toLowerCase().trim();
    const engineFilter = document.getElementById('historyEngineFilter')?.value || 'all';
    const statusFilter = document.getElementById('historyStatusFilter')?.value || 'all';

    let filtered = conversionHistory;

    if (searchTerm) {
        filtered = filtered.filter(item =>
            item.name.toLowerCase().includes(searchTerm)
        );
    }
    if (engineFilter !== 'all') {
        filtered = filtered.filter(item => item.engine === engineFilter);
    }
    if (statusFilter !== 'all') {
        const wantSuccess = statusFilter === 'success';
        filtered = filtered.filter(item => item.success === wantSuccess);
    }

    historyCount.textContent = `(${filtered.length}${filtered.length !== conversionHistory.length ? ' / ' + conversionHistory.length : ''})`;

    if (conversionHistory.length === 0) {
        historyList.innerHTML = '<p class="history-empty">No conversions yet. Convert a VBA module to see history here.</p>';
        return;
    }

    if (filtered.length === 0) {
        historyList.innerHTML = '<p class="history-no-results">No matching conversions found.</p>';
        return;
    }

    historyList.innerHTML = filtered.map((item, i) => {
        const date = new Date(item.timestamp);
        const timeStr = date.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
        const dateStr = date.toLocaleDateString([], { month: 'short', day: 'numeric' });
        const durationStr = item.duration ? formatTime(item.duration) : '';
        return `
            <div class="history-item">
                <div class="history-item-info">
                    <div class="history-item-name">${escapeHtml(item.name)}</div>
                    <div class="history-item-meta">
                        <span>${dateStr} ${timeStr}</span>
                        <span class="history-item-engine">${escapeHtml(item.engine)}</span>
                        ${durationStr ? `<span class="history-item-duration">⏱ ${durationStr}</span>` : ''}
                    </div>
                </div>
                <span class="history-item-status">${item.success ? '✅' : '❌'}</span>
            </div>
        `;
    }).join('');
}

function clearHistory() {
    conversionHistory = [];
    localStorage.removeItem('conversionHistory');
    displayHistory();
}

function toggleHistoryPanel() {
    const historyList = document.getElementById('historyList');
    const toggleBtn = document.getElementById('toggleHistoryBtn');
    if (!historyList || !toggleBtn) return;
    historyList.classList.toggle('collapsed');
    toggleBtn.textContent = historyList.classList.contains('collapsed') ? '▶' : '▼';
}

// ============================================================================
// DOWNLOAD ALL AS ZIP
// ============================================================================

async function downloadAllAsZip() {
    if (!batchConvertedModules || batchConvertedModules.length === 0) {
        showStatus('warning', 'No converted modules to download. Run "Convert All" first.');
        return;
    }

    showLoading('Packaging modules as ZIP...');

    const files = batchConvertedModules.map(m => ({
        filename: m.name,
        content: m.python_code
    }));

    try {
        const response = await fetch('/api/download-zip', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ files })
        });

        if (!response.ok) {
            const data = await response.json();
            hideLoading();
            showStatus('error', getErrorMessage(data));
            return;
        }

        const blob = await response.blob();
        hideLoading();

        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'converted_modules.zip';
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showStatus('success', `Downloaded ZIP with ${files.length} converted module(s)`);

    } catch (error) {
        hideLoading();
        showStatus('error', `ZIP download failed: ${error.message}`);
    }
}

// ============================================================================
// MODULE NAVIGATOR TABS
// ============================================================================

function buildModuleTabs(modules) {
    const tabsContainer = document.getElementById('moduleTabs');
    const tabsList = document.getElementById('moduleTabsList');
    if (!tabsContainer || !tabsList || modules.length <= 1) {
        if (tabsContainer) tabsContainer.classList.add('hidden');
        return;
    }

    tabsContainer.classList.remove('hidden');
    tabsList.innerHTML = modules.map((mod, i) => `
        <button class="module-tab${i === 0 ? ' active' : ''}" data-tab-index="${i}">
            ${escapeHtml(mod.name)}
        </button>
    `).join('');

    // Attach click handlers
    tabsList.querySelectorAll('.module-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            navigateBatchModule(parseInt(tab.dataset.tabIndex, 10));
        });
    });

    updateTabCounter();
}

function navigateBatchModule(index) {
    if (!batchConvertedModules || batchConvertedModules.length === 0) return;
    if (index < 0 || index >= batchConvertedModules.length) return;

    currentBatchIndex = index;
    const mod = batchConvertedModules[index];
    displayConversion(mod.original_code, mod.python_code, mod.name, mod.conversion_notes);

    // Update active tab
    const tabsList = document.getElementById('moduleTabsList');
    if (tabsList) {
        tabsList.querySelectorAll('.module-tab').forEach((tab, i) => {
            tab.classList.toggle('active', i === index);
        });
        // Scroll active tab into view
        const activeTab = tabsList.querySelector('.module-tab.active');
        if (activeTab) activeTab.scrollIntoView({ behavior: 'smooth', block: 'nearest', inline: 'center' });
    }

    updateTabCounter();
}

function updateTabCounter() {
    const counter = document.getElementById('tabCounter');
    const prevBtn = document.getElementById('prevModuleBtn');
    const nextBtn = document.getElementById('nextModuleBtn');
    if (!counter) return;

    const total = batchConvertedModules.length;
    counter.textContent = `${currentBatchIndex + 1} / ${total}`;
    if (prevBtn) prevBtn.disabled = currentBatchIndex <= 0;
    if (nextBtn) nextBtn.disabled = currentBatchIndex >= total - 1;
}

// ============================================================================
// LINE NUMBERS
// ============================================================================

function addLineNumbers(codeElement) {
    const pre = codeElement.closest('pre');
    if (!pre) return;

    // Remove existing line numbers
    const existing = pre.querySelector('.line-numbers');
    if (existing) existing.remove();

    const text = codeElement.textContent || '';
    const lineCount = text.split('\n').length;
    if (lineCount < 2) return;  // Don't add for single-line content

    const lineNumbersDiv = document.createElement('div');
    lineNumbersDiv.className = 'line-numbers';
    lineNumbersDiv.setAttribute('aria-hidden', 'true');
    lineNumbersDiv.textContent = Array.from({ length: lineCount }, (_, i) => i + 1).join('\n');

    pre.insertBefore(lineNumbersDiv, pre.firstChild);
}

// ============================================================================
// VBA SYNTAX VALIDATION
// ============================================================================

function validateVBASyntax() {
    const textarea = document.getElementById('vbaInput');
    const validation = document.getElementById('syntaxValidation');
    const icon = document.getElementById('validationIcon');
    const text = document.getElementById('validationText');
    if (!textarea || !validation || !icon || !text) return;

    const code = textarea.value.trim();

    if (!code) {
        validation.className = 'syntax-validation';
        icon.textContent = '⬜';
        text.textContent = 'Paste VBA code above';
        return;
    }

    const issues = [];
    const info = [];

    // Count VBA block keywords
    const lines = code.split('\n').map(l => l.trim());

    // Sub/Function matching
    const subOpens = (code.match(/\b(Sub|Function)\s+\w+/gi) || []).length;
    const subCloses = (code.match(/\bEnd\s+(Sub|Function)\b/gi) || []).length;
    if (subOpens > 0 && subCloses < subOpens) {
        issues.push(`${subOpens - subCloses} unclosed Sub/Function block(s)`);
    }

    // If/End If matching
    const ifOpens = (code.match(/\bIf\b.*\bThen\s*$/gim) || []).length;
    const ifCloses = (code.match(/\bEnd\s+If\b/gi) || []).length;
    if (ifOpens > 0 && ifCloses < ifOpens) {
        issues.push(`${ifOpens - ifCloses} unclosed If block(s)`);
    }

    // For/Next matching
    const forOpens = (code.match(/\bFor\s+(Each\s+)?\w+/gi) || []).length;
    const forCloses = (code.match(/\bNext\b/gi) || []).length;
    if (forOpens > 0 && forCloses < forOpens) {
        issues.push(`${forOpens - forCloses} unclosed For loop(s)`);
    }

    // Do/Loop matching
    const doOpens = (code.match(/\bDo\b/gi) || []).length;
    const doCloses = (code.match(/\bLoop\b/gi) || []).length;
    if (doOpens > 0 && doCloses < doOpens) {
        issues.push(`${doOpens - doCloses} unclosed Do loop(s)`);
    }

    // While/Wend matching
    const whileOpens = (code.match(/\bWhile\b(?!\s+\w+\s*\n)/gi) || []).length;
    const wendCloses = (code.match(/\bWend\b/gi) || []).length;
    if (whileOpens > wendCloses && wendCloses >= 0) {
        // Only flag if there are Wend keywords (indicating While/Wend style, not Do While)
        if (whileOpens > 0 && (code.match(/\bWend\b/gi) || code.match(/\bWhile\s+/gi))) {
            // Check for standalone While...Wend (not Do While)
            const standaloneWhile = (code.match(/^While\b/gim) || []).length;
            if (standaloneWhile > wendCloses) {
                issues.push(`${standaloneWhile - wendCloses} unclosed While/Wend block(s)`);
            }
        }
    }

    // Select Case matching
    const selectOpens = (code.match(/\bSelect\s+Case\b/gi) || []).length;
    const selectCloses = (code.match(/\bEnd\s+Select\b/gi) || []).length;
    if (selectOpens > 0 && selectCloses < selectOpens) {
        issues.push(`${selectOpens - selectCloses} unclosed Select Case block(s)`);
    }

    // With/End With matching
    const withOpens = (code.match(/\bWith\s+/gi) || []).length;
    const withCloses = (code.match(/\bEnd\s+With\b/gi) || []).length;
    if (withOpens > 0 && withCloses < withOpens) {
        issues.push(`${withOpens - withCloses} unclosed With block(s)`);
    }

    // Parentheses matching
    const openParens = (code.match(/\(/g) || []).length;
    const closeParens = (code.match(/\)/g) || []).length;
    if (openParens !== closeParens) {
        issues.push(`Mismatched parentheses (${openParens} open, ${closeParens} close)`);
    }

    // Quote matching (per line, skip comments)
    lines.forEach((line, idx) => {
        // Skip comment lines
        if (line.startsWith("'") || line.toLowerCase().startsWith('rem ')) return;
        // Remove string content to count outer quotes
        const quotes = (line.match(/"/g) || []).length;
        if (quotes % 2 !== 0) {
            issues.push(`Unmatched quote on line ${idx + 1}`);
        }
    });

    // Gather info
    const subs = (code.match(/\bSub\s+\w+/gi) || []).length;
    const funcs = (code.match(/\bFunction\s+\w+/gi) || []).length;
    const dims = (code.match(/\bDim\s+/gi) || []).length;
    if (subs > 0) info.push(`${subs} Sub(s)`);
    if (funcs > 0) info.push(`${funcs} Function(s)`);
    if (dims > 0) info.push(`${dims} variable(s)`);
    info.push(`${lines.length} line(s)`);

    // Set status
    if (issues.length > 0) {
        validation.className = 'syntax-validation warning';
        icon.textContent = '⚠️';
        text.textContent = issues[0] + (issues.length > 1 ? ` (+${issues.length - 1} more)` : '');
        text.title = issues.join('\n');
    } else if (subOpens > 0 || funcs > 0) {
        validation.className = 'syntax-validation valid';
        icon.textContent = '✅';
        text.textContent = info.join(' · ');
        text.title = 'VBA syntax looks valid';
    } else {
        validation.className = 'syntax-validation';
        icon.textContent = 'ℹ️';
        text.textContent = info.join(' · ') || 'Code detected';
        text.title = 'No Sub/Function blocks detected — may be a code fragment';
    }
}

// ============================================================================
// UTILITY: DEBOUNCE
// ============================================================================

function debounce(fn, delay) {
    let timer;
    return function (...args) {
        clearTimeout(timer);
        timer = setTimeout(() => fn.apply(this, args), delay);
    };
}

// ============================================================================
// COPY VBA CODE
// ============================================================================

async function copyVBAToClipboard() {
    const vbaEl = document.getElementById('vbaCode');
    const btn = document.getElementById('copyVbaBtn');
    if (!vbaEl || !vbaEl.textContent.trim()) {
        showStatus('warning', 'No VBA code to copy');
        return;
    }
    try {
        await navigator.clipboard.writeText(vbaEl.textContent);
        if (btn) {
            btn.textContent = '✅';
            setTimeout(() => { btn.textContent = '📋'; }, 2000);
        }
    } catch (err) {
        showStatus('error', 'Failed to copy VBA code');
    }
}

// ============================================================================
// CODE STATISTICS
// ============================================================================

function updateCodeStats(elementId, code, language) {
    const el = document.getElementById(elementId);
    if (!el || !code) { if (el) el.innerHTML = ''; return; }

    const lines = code.split('\n');
    const lineCount = lines.length;
    const nonEmpty = lines.filter(l => l.trim().length > 0).length;
    const charCount = code.length;

    let funcs = 0;
    let comments = 0;

    if (language === 'vba') {
        funcs = (code.match(/\b(Sub|Function)\s+\w+/gi) || []).length;
        comments = lines.filter(l => l.trim().startsWith("'") || l.trim().toLowerCase().startsWith('rem ')).length;
    } else if (language === 'python') {
        funcs = (code.match(/\bdef\s+\w+/g) || []).length;
        comments = lines.filter(l => l.trim().startsWith('#')).length;
    }

    const stats = [];
    stats.push(`<span class="stat"><span class="stat-value">${lineCount}</span><span class="stat-label">lines</span></span>`);
    if (funcs > 0) {
        stats.push(`<span class="stat"><span class="stat-value">${funcs}</span><span class="stat-label">${language === 'vba' ? 'procs' : 'funcs'}</span></span>`);
    }
    if (comments > 0) {
        stats.push(`<span class="stat"><span class="stat-value">${comments}</span><span class="stat-label">comments</span></span>`);
    }
    stats.push(`<span class="stat"><span class="stat-value">${formatNumber(charCount)}</span><span class="stat-label">chars</span></span>`);

    el.innerHTML = stats.join('');
}

function formatNumber(n) {
    if (n >= 1000) return (n / 1000).toFixed(1) + 'k';
    return n.toString();
}

// ================================================================
// DIFF HIGHLIGHTING — highlight VBA→Python keyword mappings
// ================================================================

const DIFF_MAPPINGS_VBA = [
    // VBA keywords that have Python equivalents
    { pattern: /\b(Sub|End Sub)\b/g, cls: 'diff-keyword' },
    { pattern: /\b(Function|End Function)\b/g, cls: 'diff-keyword' },
    { pattern: /\b(Dim|As\s+(?:String|Integer|Long|Double|Boolean|Variant|Object|Date|Currency|Single|Byte))\b/g, cls: 'diff-type' },
    { pattern: /\b(MsgBox)\b/g, cls: 'diff-call' },
    { pattern: /\b(InputBox)\b/g, cls: 'diff-call' },
    { pattern: /\b(Range|Cells|ActiveSheet|ActiveWorkbook|Worksheets|Sheets)\b/g, cls: 'diff-excel' },
    { pattern: /\b(On Error (?:GoTo|Resume Next))\b/g, cls: 'diff-error' },
    { pattern: /\b(For Each|For\b.*\bTo\b|Next)\b/g, cls: 'diff-loop' },
    { pattern: /\b(If|Then|ElseIf|Else|End If)\b/g, cls: 'diff-flow' },
    { pattern: /\b(Do While|Do Until|Loop|Wend|While)\b/g, cls: 'diff-loop' },
    { pattern: /\b(Select Case|Case|End Select)\b/g, cls: 'diff-flow' },
    { pattern: /\b(Set)\b/g, cls: 'diff-keyword' },
];

const DIFF_MAPPINGS_PYTHON = [
    // Python equivalents
    { pattern: /\b(def)\b/g, cls: 'diff-keyword' },
    { pattern: /\b(return)\b/g, cls: 'diff-keyword' },
    { pattern: /\b(str|int|float|bool|list|dict|Any|Optional)\b/g, cls: 'diff-type' },
    { pattern: /\b(print)\s*\(/g, cls: 'diff-call', matchGroup: 1 },
    { pattern: /\b(input)\s*\(/g, cls: 'diff-call', matchGroup: 1 },
    { pattern: /\b(pd\.read_excel|pd\.DataFrame|df\.\w+|openpyxl|xlwings)\b/g, cls: 'diff-excel' },
    { pattern: /\b(try|except|raise|finally)\b/g, cls: 'diff-error' },
    { pattern: /\b(for)\b/g, cls: 'diff-loop' },
    { pattern: /\b(while)\b/g, cls: 'diff-loop' },
    { pattern: /\b(if|elif|else)\b/g, cls: 'diff-flow' },
    { pattern: /\b(match|case)\b/g, cls: 'diff-flow' },
    { pattern: /\b(import|from)\b/g, cls: 'diff-keyword' },
];

function applyDiffHighlights(codeEl, mappings) {
    // Work on the already-highlighted HTML from Prism
    let html = codeEl.innerHTML;

    mappings.forEach(({ pattern, cls }) => {
        // Reset regex lastIndex
        pattern.lastIndex = 0;
        // Only highlight text NOT already inside an HTML tag
        html = html.replace(
            // Match text outside of HTML tags
            /(?<=>)([^<]+)(?=<)/g,
            (_, text) => {
                return text.replace(pattern, (match) => {
                    return `<span class="diff-hl ${cls}" title="${cls.replace('diff-', '').replace(/^\w/, c => c.toUpperCase())} mapping">${match}</span>`;
                });
            }
        );
    });

    codeEl.innerHTML = html;
}

let diffHighlightEnabled = true;

function toggleDiffHighlights() {
    diffHighlightEnabled = !diffHighlightEnabled;
    const btn = document.getElementById('toggleDiffBtn');
    if (btn) {
        const label = diffHighlightEnabled ? '🔍 Highlights On' : '🔍 Highlights Off';
        btn.textContent = label;
        btn.setAttribute('aria-pressed', diffHighlightEnabled);
        btn.setAttribute('aria-label', label.replace('🔍 ', ''));
    }
    // Re-render current view
    if (currentPythonCode) {
        Prism.highlightElement(vbaCode);
        Prism.highlightElement(pythonCode);
        if (diffHighlightEnabled) {
            applyDiffHighlights(vbaCode, DIFF_MAPPINGS_VBA);
            applyDiffHighlights(pythonCode, DIFF_MAPPINGS_PYTHON);
        }
        addLineNumbers(vbaCode);
        addLineNumbers(pythonCode);
    }
}

function formatTime(ms) {
    if (ms < 1000) return `${Math.round(ms)}ms`;
    if (ms < 60000) return `${(ms / 1000).toFixed(1)}s`;
    const mins = Math.floor(ms / 60000);
    const secs = ((ms % 60000) / 1000).toFixed(0);
    return `${mins}m ${secs}s`;
}

function showConversionTime(ms) {
    const timeStr = formatTime(ms);
    conversionStatus.textContent = `Converted · ${timeStr}`;
    conversionStatus.dataset.time = timeStr;
}

// ============================================================================
// BATCH PROGRESS INDICATOR
// ============================================================================

function showBatchProgress(current, total) {
    const container = document.getElementById('batchProgress');
    if (!container) return;
    container.classList.remove('hidden');
    updateBatchProgress(current, total);
}

function updateBatchProgress(current, total, moduleName) {
    const fill = document.getElementById('batchProgressFill');
    const text = document.getElementById('batchProgressText');
    if (!fill || !text) return;

    const pct = total > 0 ? Math.round((current / total) * 100) : 0;
    fill.style.width = `${pct}%`;

    if (moduleName) {
        text.textContent = `${current} / ${total} — ${moduleName}`;
    } else {
        text.textContent = `${current} / ${total} modules`;
    }
}

function hideBatchProgress() {
    const container = document.getElementById('batchProgress');
    const fill = document.getElementById('batchProgressFill');
    if (container) container.classList.add('hidden');
    if (fill) fill.style.width = '0%';
}

// ============================================================================
// EXPORT / IMPORT CONVERSION HISTORY
// ============================================================================

function exportHistory() {
    if (conversionHistory.length === 0) {
        showStatus('warning', 'No history to export');
        return;
    }

    const data = {
        exported_at: new Date().toISOString(),
        version: 1,
        entries: conversionHistory
    };

    const blob = new Blob([JSON.stringify(data, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    const dateStr = new Date().toISOString().slice(0, 10);
    a.download = `vba-converter-history-${dateStr}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showStatus('success', `Exported ${conversionHistory.length} history entries`);
}

function importHistory(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const content = e.target.result;
            let entries;

            if (file.name.endsWith('.csv')) {
                entries = parseCSVHistory(content);
            } else {
                const parsed = JSON.parse(content);
                entries = parsed.entries || parsed;
                if (!Array.isArray(entries)) {
                    throw new Error('Invalid format: expected an array of entries');
                }
            }

            // Validate entries
            const valid = entries.filter(entry =>
                entry && typeof entry.name === 'string' &&
                typeof entry.engine === 'string' &&
                typeof entry.success === 'boolean' &&
                entry.timestamp
            );

            if (valid.length === 0) {
                showStatus('error', 'No valid history entries found in file');
                return;
            }

            // Merge: add imported entries, deduplicate by timestamp, keep newest first
            const existing = new Set(conversionHistory.map(h => h.timestamp));
            let added = 0;
            valid.forEach(entry => {
                if (!existing.has(entry.timestamp)) {
                    conversionHistory.push(entry);
                    existing.add(entry.timestamp);
                    added++;
                }
            });

            // Sort newest first and cap at 50
            conversionHistory.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
            if (conversionHistory.length > 50) conversionHistory.length = 50;

            localStorage.setItem('conversionHistory', JSON.stringify(conversionHistory));
            displayHistory();
            showStatus('success', `Imported ${added} new entries (${valid.length - added} duplicates skipped)`);
        } catch (err) {
            showStatus('error', `Import failed: ${err.message}`);
        }
    };
    reader.readAsText(file);
    // Reset input so same file can be re-imported
    event.target.value = '';
}

function parseCSVHistory(csv) {
    const lines = csv.trim().split('\n');
    if (lines.length < 2) return [];
    // Expect header: name,engine,success,timestamp
    return lines.slice(1).map(line => {
        const parts = line.split(',');
        if (parts.length < 4) return null;
        return {
            name: parts[0].trim().replace(/^"|"$/g, ''),
            engine: parts[1].trim().replace(/^"|"$/g, ''),
            success: parts[2].trim().toLowerCase() === 'true',
            timestamp: parts[3].trim().replace(/^"|"$/g, '')
        };
    }).filter(Boolean);
}

// ============================================================================
// KEYBOARD SHORTCUTS
// ============================================================================

function setupKeyboardShortcuts() {
    // Shortcuts help button
    const shortcutsHelpBtn = document.getElementById('shortcutsHelpBtn');
    const closeShortcutsBtn = document.getElementById('closeShortcutsBtn');
    if (shortcutsHelpBtn) shortcutsHelpBtn.addEventListener('click', toggleShortcutsOverlay);
    if (closeShortcutsBtn) closeShortcutsBtn.addEventListener('click', hideShortcutsOverlay);

    // Close overlay on backdrop click
    const overlay = document.getElementById('shortcutsOverlay');
    if (overlay) {
        overlay.addEventListener('click', (e) => {
            if (e.target === overlay) hideShortcutsOverlay();
        });
    }

    // Global keyboard listener
    document.addEventListener('keydown', handleGlobalShortcut);
}

function handleGlobalShortcut(e) {
    const tag = e.target.tagName.toLowerCase();
    const isInput = tag === 'input' || tag === 'textarea' || tag === 'select' || e.target.isContentEditable;
    const overlay = document.getElementById('shortcutsOverlay');
    const shortcutsVisible = overlay && !overlay.classList.contains('hidden');

    // Esc — close any open overlay
    if (e.key === 'Escape') {
        if (shortcutsVisible) {
            hideShortcutsOverlay();
            e.preventDefault();
            return;
        }
        if (!loadingOverlay.classList.contains('hidden')) {
            // Don't close loading overlay (user can't cancel an API call)
            return;
        }
        return;
    }

    // ? key (no modifier, not in an input) — toggle shortcuts overlay
    if (e.key === '?' && !isInput && !e.ctrlKey && !e.metaKey && !e.altKey) {
        toggleShortcutsOverlay();
        e.preventDefault();
        return;
    }

    // Don't process other shortcuts if in an input (except Ctrl+Enter in textarea)
    const ctrlOrMeta = e.ctrlKey || e.metaKey;

    // Ctrl+Enter — convert pasted VBA (works even inside textarea)
    if (ctrlOrMeta && e.key === 'Enter' && !e.shiftKey) {
        e.preventDefault();
        if (vbaInput && vbaInput.value.trim()) {
            convertPastedVBA();
        }
        return;
    }

    // Ctrl+Shift+Enter — convert all modules
    if (ctrlOrMeta && e.shiftKey && e.key === 'Enter') {
        e.preventDefault();
        if (extractedModules.length > 0) {
            convertAllModules();
        }
        return;
    }

    // Below this: skip if user is typing in an input
    if (isInput) return;

    // Ctrl+S — download Python file
    if (ctrlOrMeta && e.key === 's' && !e.shiftKey) {
        e.preventDefault();
        downloadPythonFile();
        return;
    }

    // Ctrl+Shift+S — download all as ZIP
    if (ctrlOrMeta && e.shiftKey && e.key === 'S') {
        e.preventDefault();
        downloadAllAsZip();
        return;
    }

    // Ctrl+O — open file browser
    if (ctrlOrMeta && e.key === 'o') {
        e.preventDefault();
        fileInput.click();
        return;
    }

    // Ctrl+Shift+C — copy Python code
    if (ctrlOrMeta && e.shiftKey && e.key === 'C') {
        e.preventDefault();
        copyPythonToClipboard();
        return;
    }

    // Ctrl+D — toggle theme
    if (ctrlOrMeta && e.key === 'd') {
        e.preventDefault();
        toggleTheme();
        return;
    }

    // Ctrl+H — toggle diff highlights
    if (ctrlOrMeta && e.key === 'h') {
        e.preventDefault();
        toggleDiffHighlights();
        return;
    }

    // Alt+ArrowLeft — previous module tab
    if (e.altKey && e.key === 'ArrowLeft') {
        e.preventDefault();
        navigateBatchModule(currentBatchIndex - 1);
        return;
    }

    // Alt+ArrowRight — next module tab
    if (e.altKey && e.key === 'ArrowRight') {
        e.preventDefault();
        navigateBatchModule(currentBatchIndex + 1);
        return;
    }
}

function toggleShortcutsOverlay() {
    const overlay = document.getElementById('shortcutsOverlay');
    if (!overlay) return;
    if (overlay.classList.contains('hidden')) {
        showShortcutsOverlay();
    } else {
        hideShortcutsOverlay();
    }
}

function showShortcutsOverlay() {
    const overlay = document.getElementById('shortcutsOverlay');
    if (!overlay) return;
    overlay.classList.remove('hidden');
    // Focus the close button for keyboard accessibility
    const closeBtn = document.getElementById('closeShortcutsBtn');
    if (closeBtn) closeBtn.focus();
    announceToSR('Keyboard shortcuts panel opened');
}

function hideShortcutsOverlay() {
    const overlay = document.getElementById('shortcutsOverlay');
    if (!overlay) return;
    overlay.classList.add('hidden');
    // Return focus to the shortcuts help button
    const helpBtn = document.getElementById('shortcutsHelpBtn');
    if (helpBtn) helpBtn.focus();
}

// ============================================================================
// COLLAPSIBLE SECTIONS
// ============================================================================

function setupCollapsibleSections() {
    // Load persisted state from localStorage
    const savedState = JSON.parse(localStorage.getItem('collapsedSections') || '{}');

    document.querySelectorAll('.collapse-toggle').forEach(btn => {
        const targetId = btn.dataset.target;
        const body = document.getElementById(targetId);
        if (!body) return;

        // Restore saved collapsed state
        if (savedState[targetId] === true) {
            body.classList.add('collapsed');
            btn.setAttribute('aria-expanded', 'false');
            btn.textContent = '▶';
        }

        btn.addEventListener('click', () => {
            const isCollapsed = body.classList.toggle('collapsed');
            btn.setAttribute('aria-expanded', !isCollapsed);
            btn.textContent = isCollapsed ? '▶' : '▼';

            // Persist state
            const state = JSON.parse(localStorage.getItem('collapsedSections') || '{}');
            state[targetId] = isCollapsed;
            localStorage.setItem('collapsedSections', JSON.stringify(state));

            announceToSR(isCollapsed ? 'Section collapsed' : 'Section expanded');
        });
    });
}

// ============================================================================
// RESIZABLE CODE PANELS
// ============================================================================

function setupResizablePanels() {
    const handle = document.getElementById('panelResizeHandle');
    const panels = document.getElementById('codePanels');
    if (!handle || !panels) return;

    // Restore saved ratio from localStorage
    const savedRatio = parseFloat(localStorage.getItem('panelSplitRatio'));
    if (savedRatio && savedRatio > 0.15 && savedRatio < 0.85) {
        applyPanelRatio(panels, savedRatio);
    }

    let isDragging = false;

    handle.addEventListener('mousedown', (e) => {
        e.preventDefault();
        startResize(e.clientX);
    });

    handle.addEventListener('touchstart', (e) => {
        e.preventDefault();
        startResize(e.touches[0].clientX);
    }, { passive: false });

    // Keyboard support: left/right arrow keys to resize
    handle.addEventListener('keydown', (e) => {
        if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
        e.preventDefault();
        const rect = panels.getBoundingClientRect();
        const currentRatio = parseFloat(localStorage.getItem('panelSplitRatio')) || 0.5;
        const step = e.shiftKey ? 0.05 : 0.02; // Larger step with Shift
        const newRatio = e.key === 'ArrowLeft'
            ? Math.max(0.15, currentRatio - step)
            : Math.min(0.85, currentRatio + step);
        applyPanelRatio(panels, newRatio);
        localStorage.setItem('panelSplitRatio', newRatio.toString());
    });

    function startResize(startX) {
        isDragging = true;
        handle.classList.add('dragging');
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';

        // Switch to flex layout for precise control
        panels.classList.add('resizable');

        const rect = panels.getBoundingClientRect();
        const panelsWidth = rect.width;

        function onMove(clientX) {
            if (!isDragging) return;
            const offset = clientX - rect.left;
            const ratio = Math.max(0.15, Math.min(0.85, offset / panelsWidth));
            applyPanelRatio(panels, ratio);
        }

        function onMouseMove(e) { onMove(e.clientX); }
        function onTouchMove(e) { onMove(e.touches[0].clientX); }

        function onEnd() {
            isDragging = false;
            handle.classList.remove('dragging');
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
            document.removeEventListener('mousemove', onMouseMove);
            document.removeEventListener('mouseup', onEnd);
            document.removeEventListener('touchmove', onTouchMove);
            document.removeEventListener('touchend', onEnd);

            // Save ratio
            const vbaPanel = panels.querySelector('.vba-panel');
            if (vbaPanel) {
                const savedRatio = parseFloat(vbaPanel.style.flex) / 100;
                if (savedRatio > 0) {
                    localStorage.setItem('panelSplitRatio', savedRatio.toString());
                }
            }
        }

        document.addEventListener('mousemove', onMouseMove);
        document.addEventListener('mouseup', onEnd);
        document.addEventListener('touchmove', onTouchMove, { passive: false });
        document.addEventListener('touchend', onEnd);
    }
}

function applyPanelRatio(panels, ratio) {
    panels.classList.add('resizable');
    const vbaPanel = panels.querySelector('.vba-panel');
    const pythonPanel = panels.querySelector('.python-panel');
    if (!vbaPanel || !pythonPanel) return;

    const leftPct = ratio * 100;
    const rightPct = (1 - ratio) * 100;
    vbaPanel.style.flex = `${leftPct} 0 0%`;
    pythonPanel.style.flex = `${rightPct} 0 0%`;
}