# 🐍 Excel VBA to Python Converter

An LLM-powered web application that converts VBA/Macro code from Excel spreadsheets to idiomatic Python code using Claude or OpenAI APIs.

![Architecture](docs/architecture.svg)

## ✨ Features

- **📁 File Upload**: Drag-and-drop Excel files (.xlsm, .xls, .xlsb, .xla, .xlam)
- **🔍 VBA Extraction**: Automatically extracts all VBA modules from uploaded files
- **🤖 AI-Powered Conversion**: Uses Claude or OpenAI to convert VBA to Python
- **📊 Modern Python Output**: Generates code using pandas, polars, or openpyxl
- **🎨 Side-by-Side View**: Compare original VBA with converted Python
- **💾 Download**: Export converted Python files
- **✏️ Paste Mode**: Directly paste VBA code for quick conversion

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

### Upload Excel File

1. Drag and drop an Excel file with macros onto the upload area
2. The app extracts all VBA modules automatically
3. Click "Convert" on any module to generate Python code

### Paste VBA Code

1. Scroll down to the "Paste VBA Code" section
2. Paste your VBA code into the text area
3. Click "Convert to Python"

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
├── app.py                  # Flask application
├── vba_extractor.py        # VBA extraction from Excel files
├── llm_converter.py        # LLM-powered VBA→Python conversion
├── requirements.txt        # Python dependencies
├── .env.example            # Environment variables template
├── .gitignore
├── static/
│   ├── css/styles.css      # Application styles
│   └── js/app.js           # Frontend JavaScript
├── templates/
│   └── index.html          # Main page template
├── docs/
│   └── architecture.svg    # Architecture diagram
└── .github/
    └── copilot-instructions.md
```

## 🔧 API Endpoints

### POST /api/upload

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

### POST /api/convert

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

### POST /api/convert-all

Batch convert all modules.

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
