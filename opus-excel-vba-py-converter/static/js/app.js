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
