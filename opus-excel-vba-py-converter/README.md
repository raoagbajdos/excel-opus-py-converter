# 🐍 Excel VBA to Python Converter

A web application that converts VBA/Macro code from Excel spreadsheets to idiomatic Python code — powered by AI (Claude/OpenAI) or a built-in offline rule-based engine that requires no API keys.

![Architecture](docs/architecture.svg)

## ✨ Features

### Core Features
- **📁 File Upload**: Drag-and-drop Excel files (.xlsm, .xls, .xlsb, .xlsx, .xla, .xlam)
- **🔍 VBA Extraction**: Automatically extracts all VBA modules, including VBA stored in worksheet cells
- **⚡ Offline Conversion**: Built-in rule-based converter — no API key needed
- **🤖 AI-Powered Conversion**: Optionally use Claude or OpenAI for higher fidelity output
- **📊 Modern Python Output**: Generates code using pandas, polars, or openpyxl
- **🎨 Side-by-Side View**: Compare original VBA with converted Python
- **💾 Download**: Export individual files or batch download as ZIP
- **✏️ Paste Mode**: Directly paste VBA code for quick conversion
- **🌗 Dark / Light Theme**: Toggle theme with one click, preference saved automatically
- **📜 Conversion History**: Track all conversions across sessions with timestamps and status

### UI & Productivity Features
- **🔍 Diff Highlighting**: Color-coded inline highlights showing VBA→Python keyword mappings (keywords, types, calls, Excel objects, error handling, loops, control flow) with toggleable on/off
- **⏱ Conversion Time Tracking**: Per-conversion and batch timing displayed in status bar and history
- **🗂 Module Navigator Tabs**: Browse batch-converted modules with prev/next navigation
- **📋 Copy VBA / Python**: Copy source or converted code to clipboard in one click
- **📊 Code Statistics**: Live line, character, function, and import counts in both panels
- **🔢 Line Numbers**: Automatic line numbering in VBA and Python code panels
- **📤 Export / Import History**: Save history as JSON; import to merge with existing
- **🔎 Search & Filter History**: Live search and filter by engine or success/failure
- **✅ Syntax Validation Preview**: Live VBA syntax checking as you type in paste mode
- **📦 Batch Progress Indicator**: Per-module progress shown during sequential conversions
- **⌨️ Keyboard Shortcuts**: Global shortcuts for conversion, files, clipboard, navigation, and theme — press `?` for overlay
- **📂 Collapsible Sections**: Collapse/expand Options, Paste, Formulas, Data Export, and Analysis panels with persisted state
- **↔️ Resizable Code Panels**: Drag the handle between VBA/Python panels to adjust widths (mouse, touch, keyboard); ratio persisted

### Accessibility
- **Skip-to-content link** visible on keyboard Tab focus
- **ARIA roles, labels, and live regions** for screen reader support
- **Keyboard navigation**: Enter/Space on drop zone, `:focus-visible` outlines
- **Focus management**: Python panel auto-focused after conversion
- **Screen reader announcements** for status changes and conversions

### Advanced Features
- **🔢 Formula Extraction & Conversion**: Extract all Excel formulas and convert them to Python/pandas equivalents
  - Supports VLOOKUP, SUMIF, IF, INDEX/MATCH, and 50+ Excel functions
  - Displays formula statistics and usage patterns
  - One-click conversion to pandas/numpy operations
  
- **📤 Data Export**: Export Excel data to pandas DataFrames
  - Automatically generates Python code to load data
  - Detects headers and data types
  - Creates ready-to-use DataFrame loading scripts
  
- **🔍 Complete Workbook Analysis**: Comprehensive analysis combining VBA + Formulas + Data
  - Analyzes dependencies between sheets, formulas, and VBA code
  - Generates complete Python recreation script
  - Provides detailed analysis report
  - Creates unified Python module with all workbook logic

- **📦 Download All as ZIP**: After batch conversion, download every converted module in a single ZIP archive

## 🚀 Quick Start

### 1. Clone the Repository

```bash
git clone https://github.com/yourusername/opus-excel-vba-py-converter.git
cd opus-excel-vba-py-converter
```

### 2. Install with uv (Recommended)

```bash
# Install uv if you don't have it
curl -LsSf https://astral.sh/uv/install.sh | sh

# Sync dependencies
uv sync

# Run with uv
uv run python app.py
```

### Alternative: Using pip

```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### 3. Configure API Keys (Optional)

The **Offline** conversion engine works without any API keys. If you want AI-powered conversion, copy the example environment file and add your API key:

```bash
cp .env.example .env
```

Edit `.env` and add your Anthropic or OpenAI API key:

```bash
# For Claude (recommended)
ANTHROPIC_API_KEY=sk-ant-your-api-key-here

# OR for OpenAI
OPENAI_API_KEY=sk-your-openai-key-here
LLM_PROVIDER=openai
```

### 4. Run the Application

```bash
# With uvicorn (recommended)
uvicorn app:app --host 127.0.0.1 --port 5000 --reload

# Or with uv
uv run uvicorn app:app --host 127.0.0.1 --port 5000 --reload
```

Open your browser to `http://localhost:5000`

## 📖 Usage

### 1. VBA Conversion

#### Upload Excel File
1. Drag and drop an Excel file with macros onto the upload area
2. The app extracts all VBA modules automatically
3. Click "Convert" on any module to generate Python code

#### Paste VBA Code
1. Scroll down to the "Paste VBA Code" section
2. Paste your VBA code into the text area
3. Click "Convert to Python"

### 2. Formula Extraction & Conversion (NEW!)

1. Click "Extract Formulas from Excel" button
2. Upload an Excel file
3. View all extracted formulas with statistics
4. Click "Convert to Python" on any formula to see the pandas equivalent
5. Supports VLOOKUP, SUMIF, IF, INDEX/MATCH, and 50+ Excel functions

### 3. Data Export (NEW!)

1. Click "Export Data to Python" button
2. Upload an Excel file
3. Automatically generates Python code to load all sheets as DataFrames
4. View data statistics and metadata
5. Copy the generated code to use in your projects

### 4. Complete Workbook Analysis (NEW!)

1. Click "Analyze Complete Workbook" button
2. Upload an Excel file with VBA, formulas, and data
3. Get comprehensive analysis including:
   - Number of VBA modules, formulas, and data sheets
   - Dependencies between sheets and formulas
   - Complete Python script recreating entire workbook logic
   - Detailed analysis report

### Conversion Options

- **Conversion Engine**: Choose `Offline` (no API key), `Anthropic Claude`, or `OpenAI GPT`
- **Target Library**: Choose between `pandas` (default) or `polars`
- **Type Hints**: Enable/disable Python type hints in output

### 5. Keyboard Shortcuts

Press `?` at any time to open the shortcuts overlay, or click the **⌨️ Shortcuts** button in the header.

| Shortcut | Action |
|----------|--------|
| `Ctrl+Enter` | Convert pasted VBA |
| `Ctrl+Shift+Enter` | Convert all modules |
| `Ctrl+S` | Download Python file |
| `Ctrl+Shift+S` | Download all as ZIP |
| `Ctrl+O` | Open file browser |
| `Ctrl+Shift+C` | Copy Python code |
| `Alt+←/→` | Previous/next module |
| `Ctrl+D` | Toggle dark/light theme |
| `Ctrl+H` | Toggle diff highlights |
| `?` | Toggle shortcuts overlay |
| `Esc` | Close overlays |

### 6. Customising the Layout

- **Collapse sections**: Click the ▼/▶ toggle on Options, Paste, Formulas, Data Export, or Analysis panels to hide them when not in use. State is remembered across sessions.
- **Resize code panels**: Drag the ⋮ handle between the VBA and Python panels to adjust widths. You can also focus the handle and use arrow keys. The ratio is remembered across sessions.

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                     Frontend (Web UI)                           │
│  HTML/CSS/JS • Drag & Drop • Prism.js • Dark/Light Theme       │
│  Conversion History • Download ZIP • Responsive Design         │
│  Keyboard Shortcuts • Collapsible Sections • Resizable Panels  │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                   FastAPI Backend (app.py)                      │
│  /api/upload • /api/convert • /api/convert-all                 │
│  /api/extract-formulas • /api/export-data                      │
│  /api/analyze-workbook • /api/download-zip                     │
└─────────────────────────────────────────────────────────────────┘
                              │
         ┌────────────────────┼────────────────────┐
         ▼                    ▼                    ▼
┌──────────────────┐ ┌──────────────────┐ ┌──────────────────────┐
│  VBA Extractor   │ │ Offline Converter │ │   LLM Converter      │
│ vba_extractor.py │ │offline_converter │ │  llm_converter.py    │
│ • oletools       │ │  .py             │ │  • Claude API        │
│ • Sheet-cell VBA │ │ • Rule-based     │ │  • OpenAI API        │
│ • OLE parsing    │ │ • No API key     │ │  • Prompt engineering│
└──────────────────┘ └──────────────────┘ └──────────────────────┘
```

## 📁 Project Structure

```
opus-excel-vba-py-converter/
├── app.py                    # FastAPI application with all API endpoints
├── config.py                 # Configuration, env vars, file-size limits
├── vba_extractor.py          # VBA extraction (OLE + sheet-cell scanning)
├── llm_converter.py          # LLM-powered VBA & formula conversion
├── offline_converter.py      # Rule-based VBA→Python converter (no API)
├── formula_extractor.py      # Excel formula extraction & analysis
├── data_exporter.py          # Data export to pandas DataFrames
├── workbook_analyzer.py      # Complete workbook analysis
├── requirements.txt          # Python dependencies
├── pyproject.toml            # UV / project configuration
├── .env.example              # Environment variables template
├── .gitignore
├── static/
│   ├── css/styles.css        # CSS with dark/light theme variables (~1860 lines)
│   └── js/app.js             # Frontend JS — shortcuts, panels, history, etc. (~1971 lines)
├── templates/
│   └── index.html            # Main page template (~354 lines)
├── docs/
│   ├── architecture.svg      # Architecture diagram
│   ├── agents.md             # AI agents documentation
│   ├── Claude.md             # Claude integration guide
│   └── skills.md             # Skills/capabilities reference
├── FEATURES_UPDATE.md        # Detailed feature changelog (Phases 9-13)
└── .github/
    └── copilot-instructions.md
```

## 🔧 API Endpoints

### VBA Conversion Endpoints

#### POST /api/upload
Upload an Excel file and extract VBA modules.

**Request**: `multipart/form-data` with file

**Response**:
```json
{
  "success": true,
  "filename": "workbook.xlsm",
  "modules": [
    {
      "name": "Module1",
      "type": "Standard Module",
      "code": "Sub Example()..."
    }
  ]
}
```

#### POST /api/convert
Convert a single VBA code snippet.

**Request**:
```json
{
  "vba_code": "Sub Example()...",
  "module_name": "Module1",
  "target_library": "pandas",
  "provider": "offline"
}
```

**Response**:
```json
{
  "success": true,
  "python_code": "def example():...",
  "conversion_notes": ["Converted Range to DataFrame"],
  "engine": "offline"
}
```

Set `provider` to `"offline"` for the rule-based engine (no API key), or `"anthropic"` / `"openai"` for AI conversion.

#### POST /api/convert-all
Batch convert all modules.

#### POST /api/download-zip
Package converted modules as a ZIP archive.

**Request**:
```json
{
  "files": [
    { "filename": "Module1", "content": "def example(): ..." },
    { "filename": "Module2", "content": "class MyClass: ..." }
  ]
}
```

**Response**: Binary ZIP file (`application/zip`)

### Formula Conversion Endpoints (NEW!)

#### POST /api/extract-formulas
Extract all formulas from an Excel file.

**Request**: `multipart/form-data` with file

**Response**:
```json
{
  "success": true,
  "formulas": [...],
  "statistics": {
    "total_formulas": 45,
    "sheets_with_formulas": 3,
    "unique_functions_used": 12
  }
}
```

#### POST /api/convert-formula
Convert a single Excel formula to Python.

**Request**:
```json
{
  "formula": "=VLOOKUP(A2,Sheet2!A:B,2,FALSE)",
  "cell_address": "B2",
  "sheet_name": "Sheet1"
}
```

**Response**:
```json
{
  "success": true,
  "python_code": "result = df.merge(lookup_df, ...)",
  "conversion_notes": ["Converted VLOOKUP to pandas merge"]
}
```

### Data Export Endpoints (NEW!)

#### POST /api/export-data
Export Excel data to Python/pandas code.

**Request**: `multipart/form-data` with file

**Response**:
```json
{
  "success": true,
  "python_code": "# Generated DataFrame loading code...",
  "metadata": {
    "total_sheets": 3,
    "total_rows": 150,
    "total_columns": 25
  }
}
```

### Complete Analysis Endpoint (NEW!)

#### POST /api/analyze-workbook
Perform comprehensive workbook analysis.

**Request**: `multipart/form-data` with file

**Response**:
```json
{
  "success": true,
  "has_vba": true,
  "vba_modules_count": 5,
  "has_formulas": true,
  "formulas_count": 45,
  "sheets_count": 3,
  "python_script": "# Complete Python recreation...",
  "report": "Detailed analysis report..."
}
```

## 🐛 Troubleshooting

### "No VBA code found"

- Ensure the file actually contains macros (check in Excel: Alt+F11)
- Try a `.xlsm` file (macro-enabled workbook)
- Files with VBA stored in worksheet cells (e.g. a "VBA_Code" sheet) are also supported
- If no macros are found, the app automatically runs a full workbook analysis

### "API key not found" / using without API keys

- Select **"Offline (no API key)"** in the Conversion Engine dropdown — no API key is needed
- For AI-powered conversion, check that `.env` file exists with your API key
- Verify the key format is correct

### File extension mismatch

- The app automatically detects OpenXML files saved with a `.xls` extension and normalizes them
- Supported formats: `.xlsm`, `.xlsx`, `.xlsb`, `.xls`, `.xla`, `.xlam`

### oletools not working

```bash
pip install oletools --upgrade
```

## 📝 License

MIT License - See LICENSE file for details.

## 🙏 Acknowledgments

- [Anthropic Claude](https://www.anthropic.com/) for AI-powered conversion
- [oletools](https://github.com/decalage2/oletools) for VBA extraction
- [FastAPI](https://fastapi.tiangolo.com/) for the backend framework
- [Prism.js](https://prismjs.com/) for syntax highlighting
