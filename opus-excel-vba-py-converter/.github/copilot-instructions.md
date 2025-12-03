# Copilot Instructions for Excel VBA to Python Converter

## Project Overview

This is an **LLM-Powered VBA to Python Conversion Application** that enables users to upload Excel spreadsheets containing VBA macros and convert them to idiomatic Python code. The application uses AI models (Claude/OpenAI) to intelligently translate VBA syntax while leveraging modern Python data libraries.

## Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                        Frontend (Web UI)                        │
│  - File upload interface for .xlsm, .xls, .xlsb files          │
│  - Code editor with VBA syntax highlighting                     │
│  - Side-by-side VBA → Python comparison view                   │
│  - Download converted Python files                              │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                     Flask Backend API                           │
│  - POST /api/upload - Extract VBA from Excel files             │
│  - POST /api/convert - Convert single VBA module               │
│  - POST /api/convert-all - Batch convert all modules           │
└─────────────────────────────────────────────────────────────────┘
                              │
              ┌───────────────┴───────────────┐
              ▼                               ▼
┌─────────────────────────┐     ┌─────────────────────────────────┐
│    VBA Extractor        │     │     LLM Conversion Engine       │
│  - oletools/olevba      │     │  - Claude API / OpenAI API      │
│  - Manual OLE parsing   │     │  - Prompt engineering           │
│  - Module classification│     │  - Code validation              │
└─────────────────────────┘     └─────────────────────────────────┘
```

## Tech Stack

- **Backend**: Python 3.10+, Flask
- **VBA Extraction**: oletools, olefile, zipfile
- **LLM Integration**: Anthropic Claude API or OpenAI API
- **Frontend**: HTML5, CSS3, JavaScript (vanilla or with Alpine.js)
- **Code Display**: Prism.js or highlight.js for syntax highlighting
- **Python Libraries for Conversion Targets**: pandas, polars, openpyxl

## Key Design Principles

### 1. LLM-First Conversion Strategy

The conversion engine should **always use an LLM API** for translation rather than rule-based parsing. This approach:
- Handles VBA syntax nuances and edge cases
- Produces idiomatic, Pythonic code
- Understands context and intent of the original code
- Can suggest modern library alternatives (pandas, polars)

### 2. Conversion Prompt Guidelines

When sending VBA code to the LLM for conversion, use structured prompts that:

```markdown
You are an expert VBA to Python converter. Convert the following VBA code to idiomatic Python.

**Conversion Rules:**
1. Use pandas for Excel/data operations (or polars if specified)
2. Replace VBA Range operations with pandas DataFrame operations
3. Convert VBA Subs to Python functions
4. Convert VBA Functions to Python functions with proper return types
5. Use type hints for all function parameters and returns
6. Replace MsgBox with print() or logging
7. Convert VBA error handling (On Error) to try/except blocks
8. Use pathlib for file operations
9. Replace VBA date functions with datetime module
10. Add docstrings explaining what each function does

**VBA Code:**
```vba
{vba_code}
```

**Output Requirements:**
- Provide complete, runnable Python code
- Include necessary imports at the top
- Add comments for any non-obvious conversions
- Note any functionality that cannot be directly converted
```

### 3. Module Type Handling

Different VBA module types require different conversion strategies:

| VBA Module Type | Python Conversion Approach |
|-----------------|---------------------------|
| Standard Module (.bas) | Convert to Python module with functions |
| Class Module (.cls) | Convert to Python class with methods |
| UserForm (.frm) | Convert to tkinter/PyQt or note as UI-only |
| ThisWorkbook | Convert to workbook event handlers or main script |
| Sheet Modules | Convert to sheet-specific operations with openpyxl |

### 4. Common VBA to Python Mappings

When reviewing or enhancing conversions, ensure these patterns are followed:

```python
# VBA: Range("A1:B10").Value
# Python:
df = pd.read_excel("file.xlsx", usecols="A:B", nrows=10)

# VBA: Cells(row, col).Value = x
# Python:
df.iloc[row-1, col-1] = x  # Note: VBA is 1-indexed, Python is 0-indexed

# VBA: For Each cell In Range("A1:A10")
# Python:
for value in df['A']:
    ...

# VBA: Application.WorksheetFunction.VLookup(...)
# Python:
result = df.merge(lookup_df, on='key', how='left')

# VBA: Dim arr() As Variant
# Python:
arr: list = []  # or numpy array if numerical

# VBA: Set ws = ThisWorkbook.Worksheets("Sheet1")
# Python:
df = pd.read_excel(workbook_path, sheet_name="Sheet1")
```

## File Structure

```
opus-excel-vba-py-converter/
├── app.py                      # Flask application entry point
├── vba_extractor.py            # VBA extraction from Excel files
├── llm_converter.py            # LLM-powered conversion engine
├── config.py                   # Configuration and API keys
├── requirements.txt            # Python dependencies
├── .env                        # Environment variables (API keys)
├── .env.example                # Example environment file
├── static/
│   ├── css/
│   │   └── styles.css          # Application styles
│   └── js/
│       └── app.js              # Frontend JavaScript
├── templates/
│   └── index.html              # Main application template
├── uploads/                    # Temporary file uploads (gitignored)
├── tests/
│   ├── test_extractor.py       # VBA extraction tests
│   ├── test_converter.py       # Conversion tests
│   └── sample_files/           # Sample Excel files for testing
└── .github/
    └── copilot-instructions.md # This file
```

## API Endpoints

### POST /api/upload
Upload an Excel file and extract VBA modules.

**Request:** `multipart/form-data` with file
**Response:**
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
Convert a single VBA code snippet to Python.

**Request:**
```json
{
  "vba_code": "Sub Example()...",
  "module_name": "Module1",
  "target_library": "pandas",  // or "polars"
  "include_type_hints": true
}
```

**Response:**
```json
{
  "success": true,
  "python_code": "def example():...",
  "conversion_notes": ["Converted Range to DataFrame", "..."]
}
```

### POST /api/convert-all
Batch convert all extracted modules.

## Environment Variables

```bash
# LLM API Configuration
ANTHROPIC_API_KEY=sk-ant-...      # For Claude API
OPENAI_API_KEY=sk-...              # For OpenAI API (alternative)
LLM_PROVIDER=anthropic             # 'anthropic' or 'openai'
LLM_MODEL=claude-sonnet-4-20250514        # or 'gpt-4-turbo'

# Application Settings
FLASK_ENV=development
FLASK_DEBUG=1
MAX_FILE_SIZE_MB=50
UPLOAD_FOLDER=uploads
```

## Code Style Guidelines

### Python Code
- Use Python 3.10+ features (match statements, union types with |)
- Always include type hints
- Use dataclasses or Pydantic for data models
- Follow PEP 8 style guide
- Use async/await for API calls where beneficial

### Error Handling
- Wrap LLM API calls in try/except blocks
- Provide meaningful error messages to users
- Log errors with appropriate severity levels
- Handle rate limiting gracefully with retries

### Security
- Never expose API keys in frontend code
- Validate and sanitize all uploaded files
- Limit file sizes and types
- Clean up uploaded files after processing
- Use CORS appropriately

## Testing Strategy

1. **Unit Tests**: Test VBA extraction with sample files
2. **Integration Tests**: Test full upload → convert flow
3. **LLM Response Mocking**: Mock API responses for CI/CD
4. **Sample Files**: Include various Excel file types with macros

## Common Conversion Challenges

When helping with this project, be aware of these VBA-specific challenges:

1. **1-indexed vs 0-indexed**: VBA uses 1-based indexing
2. **Variant types**: VBA's Variant maps to Python's Any or Union types
3. **ByRef vs ByVal**: Default pass-by-reference in VBA
4. **Optional parameters**: Handle with Python default arguments
5. **Error handling**: `On Error Resume Next` has no direct equivalent
6. **Excel object model**: Needs openpyxl or xlwings for similar functionality
7. **UserForms**: May need separate UI framework or removal

## Future Enhancements

- [ ] Support for xlwings integration for live Excel interaction
- [ ] Batch processing of multiple files
- [ ] Conversion history and versioning
- [ ] Custom prompt templates for specific use cases
- [ ] Export as Jupyter notebooks
- [ ] Integration with GitHub for converted code storage
- [ ] Polars-first conversion option for performance
