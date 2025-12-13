# 🐍 Excel VBA to Python Converter

An LLM-powered web application that converts VBA/Macro code from Excel spreadsheets to idiomatic Python code using Claude or OpenAI APIs.

![Architecture](docs/architecture.svg)

## ✨ Features

### Core Features
- **📁 File Upload**: Drag-and-drop Excel files (.xlsm, .xls, .xlsb, .xla, .xlam)
- **🔍 VBA Extraction**: Automatically extracts all VBA modules from uploaded files
- **🤖 AI-Powered Conversion**: Uses Claude or OpenAI to convert VBA to Python
- **📊 Modern Python Output**: Generates code using pandas, polars, or openpyxl
- **🎨 Side-by-Side View**: Compare original VBA with converted Python
- **💾 Download**: Export converted Python files
- **✏️ Paste Mode**: Directly paste VBA code for quick conversion

### Advanced Features (NEW!)
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

### 4. Configure API Keys

Copy the example environment file and add your API key:

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

### 3. Run the Application

```bash
# With uv
uv run python app.py

# Or with pip/venv
python app.py
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

- **Target Library**: Choose between `pandas` (default) or `polars`
- **Type Hints**: Enable/disable Python type hints in output

## 🏗️ Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        Frontend (Web UI)                        │
│  HTML/CSS/JS • Drag & Drop • Code Highlighting (Prism.js)      │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                     Flask Backend (app.py)                      │
│  POST /api/upload • POST /api/convert • POST /api/convert-all  │
└─────────────────────────────────────────────────────────────────┘
                              │
              ┌───────────────┴───────────────┐
              ▼                               ▼
┌─────────────────────────┐     ┌─────────────────────────────────┐
│    VBA Extractor        │     │     LLM Converter               │
│  vba_extractor.py       │     │  llm_converter.py               │
│  • oletools/olevba      │     │  • Claude API (Anthropic)       │
│  • olefile parsing      │     │  • OpenAI API (GPT-4)           │
│  • ZIP extraction       │     │  • Prompt engineering           │
└─────────────────────────┘     └─────────────────────────────────┘
```

## 📁 Project Structure

```
opus-excel-vba-py-converter/
├── app.py                    # Flask application with all API endpoints
├── vba_extractor.py          # VBA extraction from Excel files
├── llm_converter.py          # LLM-powered VBA & formula conversion
├── formula_extractor.py      # Excel formula extraction & analysis (NEW)
├── data_exporter.py          # Data export to pandas DataFrames (NEW)
├── workbook_analyzer.py      # Complete workbook analysis (NEW)
├── requirements.txt          # Python dependencies
├── pyproject.toml            # UV configuration
├── .env.example              # Environment variables template
├── .gitignore
├── static/
│   ├── css/styles.css        # Application styles
│   └── js/app.js             # Frontend JavaScript
├── templates/
│   └── index.html            # Main page template
├── docs/
│   ├── architecture.svg      # Architecture diagram
│   ├── agents.md             # AI agents documentation
│   └── Claude.md             # Claude integration guide
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
  "target_library": "pandas"
}
```

**Response**:
```json
{
  "success": true,
  "python_code": "def example():...",
  "conversion_notes": ["Converted Range to DataFrame"]
}
```

#### POST /api/convert-all
Batch convert all modules.

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

### "API key not found"

- Check that `.env` file exists with your API key
- Verify the key format is correct

### oletools not working

```bash
pip install oletools --upgrade
```

## 📝 License

MIT License - See LICENSE file for details.

## 🙏 Acknowledgments

- [Anthropic Claude](https://www.anthropic.com/) for AI-powered conversion
- [oletools](https://github.com/decalage2/oletools) for VBA extraction
- [Prism.js](https://prismjs.com/) for syntax highlighting
