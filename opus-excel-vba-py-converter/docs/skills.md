# Skills Reference

This document describes the core capabilities (skills) of the Excel VBA to Python Converter application. Each skill maps to a module, API endpoint, and frontend feature.

---

## 1. VBA Extraction

**Module:** `vba_extractor.py` · **Class:** `VBAExtractor`

Extracts VBA macros from Excel files and classifies them by module type.

| Capability | Details |
|------------|---------|
| Supported formats | `.xlsm`, `.xlsb`, `.xlam`, `.xls`, `.xla`, `.xlsx` (for VBA-in-cells) |
| Extraction methods | ZIP-based parsing (modern formats), OLE/oletools (legacy `.xls`), xlrd (BIFF) |
| Sheet-cell VBA scanning | Detects `Sub`, `Function`, `Dim`, `MsgBox`, etc. embedded in cell values |
| Module classification | Standard Module, Class Module, UserForm, Document Module (ThisWorkbook/Sheet), Sheet-Cell VBA |
| File extension normalization | Auto-renames `.xls` → `.xlsx` when content is actually OpenXML |
| Fallback parsing | Manual OLE binary extraction when oletools is unavailable |

**API Endpoint:** `POST /api/upload`
- Input: `multipart/form-data` with an Excel file
- Output: JSON array of extracted modules (`name`, `type`, `code`)

**Usage:**
```python
from vba_extractor import VBAExtractor

extractor = VBAExtractor("workbook.xlsm")
modules = extractor.extract_all()
# [{'name': 'Module1', 'type': 'Standard Module', 'code': '...'}, ...]
```

---

## 2. VBA to Python Conversion

**Module:** `llm_converter.py` · **Classes:** `VBAToPythonConverter`, `AnthropicConverter`, `OpenAIConverter`

Converts VBA code to idiomatic Python using LLM APIs with a structured system prompt.

| Capability | Details |
|------------|---------|
| LLM providers | Anthropic Claude (default), OpenAI GPT-4 Turbo |
| Conversion targets | pandas (default), polars |
| Output features | Type hints, docstrings, conversion notes |
| Module type handling | Subs → functions, Classes → Python classes, UserForms → noted/tkinter |
| Index adjustment | Automatic 1-indexed (VBA) → 0-indexed (Python) translation |
| Error handling | `On Error` → `try/except`, rate-limit retries |

**API Endpoints:**
- `POST /api/convert` — Convert a single VBA module
- `POST /api/convert-all` — Batch-convert all extracted modules

**Conversion Rules (enforced via system prompt):**
1. pandas for Excel/data operations
2. `Range` → DataFrame operations
3. `MsgBox` → `print()` / `logging`
4. `On Error Resume Next` → `try/except`
5. `pathlib` for file operations
6. `datetime` for VBA date functions
7. VBA constants (`vbCrLf`, `vbTab`) → Python equivalents
8. Optional/ByRef parameters → default arguments

**Usage:**
```python
from llm_converter import VBAToPythonConverter

converter = VBAToPythonConverter(provider="anthropic")
result = converter.convert(vba_code, module_name="Module1", target_library="pandas")
print(result.python_code)
print(result.conversion_notes)
```

---

## 3. Excel Formula Extraction & Conversion

**Module:** `formula_extractor.py` · **Class:** `FormulaExtractor`

Extracts, classifies, and analyzes all formulas in a workbook, then converts them to pandas/numpy equivalents via LLM.

| Capability | Details |
|------------|---------|
| Formula types | Standard, array, shared |
| Functions recognized | 50+ Excel functions across Lookup, Math, Logical, Text, Date, Financial, Array categories |
| Dependency mapping | Identifies cell/range references each formula depends on |
| Statistics | Per-sheet and aggregate counts, function usage breakdown |
| LLM conversion | Dedicated `FORMULA_SYSTEM_PROMPT` with Excel → pandas/numpy mappings |

**API Endpoints:**
- `POST /api/extract-formulas` — Extract all formulas + statistics from a file
- `POST /api/convert-formula` — Convert a single formula to Python

**Recognized Function Categories:**

| Category | Examples |
|----------|---------|
| Lookup/Reference | `VLOOKUP`, `XLOOKUP`, `INDEX`, `MATCH`, `OFFSET`, `INDIRECT` |
| Math/Statistics | `SUM`, `SUMIF`, `SUMIFS`, `AVERAGE`, `COUNT`, `COUNTIF`, `STDEV` |
| Logical | `IF`, `IFS`, `AND`, `OR`, `IFERROR`, `IFNA` |
| Text | `CONCATENATE`, `LEFT`, `RIGHT`, `MID`, `TRIM`, `SUBSTITUTE` |
| Date/Time | `TODAY`, `NOW`, `DATEDIF`, `EOMONTH`, `WORKDAY`, `NETWORKDAYS` |
| Financial | `PMT`, `NPV`, `IRR`, `XIRR`, `FV`, `PV` |
| Array/Modern | `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `RANDARRAY` |

**Usage:**
```python
from formula_extractor import FormulaExtractor

extractor = FormulaExtractor("workbook.xlsx")
formulas = extractor.extract_all_formulas()
stats = extractor.get_formula_statistics(formulas)
```

---

## 4. Data Export to Pandas

**Module:** `data_exporter.py` · **Class:** `DataExporter`

Exports all worksheet data to pandas DataFrames and generates ready-to-run Python code.

| Capability | Details |
|------------|---------|
| Auto-header detection | Infers whether the first row is a header |
| Type inference | Maps each column to a pandas dtype |
| Code generation | Produces complete Python script with imports, `read_excel` calls, and summaries |
| Metadata | Sheet names, row/column counts, data ranges, dtype mappings |
| Export options | Exclude empty sheets, limit row count |

**API Endpoint:** `POST /api/export-data`
- Input: `multipart/form-data` with an Excel file
- Output: Generated Python code + sheet metadata

**Usage:**
```python
from data_exporter import DataExporter

exporter = DataExporter("workbook.xlsx")
result = exporter.export_all_sheets()
print(result.python_code)   # Ready-to-use Python script
print(result.metadata)      # Sheet statistics
```

---

## 5. Complete Workbook Analysis

**Module:** `workbook_analyzer.py` · **Class:** `WorkbookAnalyzer`

Orchestrates all other skills to produce a comprehensive Python recreation of an entire Excel workbook.

| Capability | Details |
|------------|---------|
| Combined analysis | VBA extraction + formula extraction + data export in one pass |
| Dependency graph | Maps relationships between sheets, formulas, and VBA modules |
| Script generation | Produces a single Python script with `WorkbookData`, `FormulaEngine`, converted VBA classes, and a `main()` entry point |
| Report generation | Text-based analysis report with counts, dependencies, and recommendations |

**API Endpoint:** `POST /api/analyze-workbook`
- Input: `multipart/form-data` with an Excel file
- Output: Analysis summary, complete Python script, dependency report, metadata

**Generated Script Structure:**
```python
class WorkbookData:       # Data loading from all sheets
class FormulaEngine:      # Formula logic as Python functions
class Module1:            # Converted VBA modules
def main():              # Orchestration entry point
```

**Usage:**
```python
from workbook_analyzer import WorkbookAnalyzer

analyzer = WorkbookAnalyzer("workbook.xlsm")
analysis = analyzer.analyze_complete()
report = analyzer.generate_analysis_report(analysis)
print(analysis.python_script)
```

---

## 6. Offline VBA → Python Conversion

**Module:** `offline_converter.py` · **Class:** `OfflineConverter`

Rule-based VBA to Python converter that works without any API key or internet connection.

| Capability | Details |
|------------|----------|
| No API key required | Fully deterministic, local-only execution |
| VBA patterns | `Sub`/`Function`, `If`/`Select Case`, `For`/`While`/`Do` loops, `With` blocks |
| Type mapping | VBA types (`Integer`, `String`, `Boolean`, etc.) → Python type hints |
| Constant mapping | `vbCrLf`, `vbTab`, `vbNullString`, `True`/`False` → Python equivalents |
| Built-in function mapping | `MsgBox` → `print()`, `InputBox` → `input()`, `UBound`/`LBound` → `len()` |
| Formula conversion | Converts common Excel formulas (`VLOOKUP`, `SUMIF`, `IF`, etc.) to pandas/numpy |
| Code size | ~881 lines of rule-based conversion logic |

**API Integration:** Used when `provider="offline"` is specified in:
- `POST /api/convert` — Single module conversion
- `POST /api/convert-all` — Batch conversion
- `POST /api/convert-formula` — Formula conversion

**Usage:**
```python
from offline_converter import OfflineConverter

converter = OfflineConverter()
result = converter.convert(vba_code, module_name="Module1")
print(result["python_code"])
print(result["conversion_notes"])
```

---

## 7. LLM Provider Management

**Module:** `llm_converter.py` · **Class:** `VBAToPythonConverter`

The orchestrator that routes conversion requests to the appropriate LLM backend.

| Capability | Details |
|------------|---------|
| Provider auto-detection | Checks available API keys and selects a provider |
| Provider fallback | Switches between Anthropic and OpenAI if one fails |
| Supported models | Claude Sonnet 3.5, Claude Opus, Claude Haiku, GPT-4 Turbo |
| Conversion notes | Extracts TODO/Warning/Note comments from LLM output |
| Configurable | Provider, model, and target library selectable per request |

**Environment Variables:**
```bash
ANTHROPIC_API_KEY=sk-ant-...
OPENAI_API_KEY=sk-...
LLM_PROVIDER=anthropic        # or 'openai'
LLM_MODEL=claude-sonnet-4-20250514    # or 'gpt-4-turbo'
```

---

## 8. Keyboard Shortcuts

**Module:** `static/js/app.js` · **Functions:** `setupKeyboardShortcuts()`, `handleGlobalShortcut()`, `toggleShortcutsOverlay()`

Global keyboard shortcuts for power-user workflows.

| Capability | Details |
|------------|---------|
| Conversion shortcuts | `Ctrl+Enter` (convert pasted), `Ctrl+Shift+Enter` (batch convert all) |
| File shortcuts | `Ctrl+S` (download .py), `Ctrl+O` (open file), `Ctrl+Shift+S` (download ZIP) |
| Clipboard | `Ctrl+Shift+C` (copy Python code) |
| Navigation | `Alt+←/→` (prev/next module tab), `Ctrl+H` (toggle highlights) |
| General | `Ctrl+D` (toggle theme), `?` (shortcuts overlay), `Esc` (close overlays) |
| Smart input detection | Single-key shortcuts suppressed when user is typing in text fields |
| Accessible overlay | Categorised help panel with `<kbd>` key badges, focus management, backdrop dismiss |

---

## 9. Collapsible Sections

**Module:** `static/js/app.js` · **Function:** `setupCollapsibleSections()`

Expand/collapse sidebar sections to manage screen space.

| Capability | Details |
|------------|---------|
| Collapsible panels | Options, Paste VBA, Formulas, Data Export, Analysis |
| Toggle control | ▼/▶ button on each section header |
| Persistence | Collapsed state saved to `localStorage` (`collapsedSections` key) |
| Animation | CSS `max-height`/`opacity`/`padding` transitions (0.3-0.4s) |
| Accessibility | `aria-expanded` attribute updated; screen reader announcement on toggle |

---

## 10. Resizable Code Panels

**Module:** `static/js/app.js` · **Functions:** `setupResizablePanels()`, `applyPanelRatio()`

Adjustable VBA/Python panel widths via a draggable handle.

| Capability | Details |
|------------|---------|
| Mouse drag | Click and drag the ⋮ handle between panels |
| Touch support | Works on mobile/tablet via `touchstart`/`touchmove` events |
| Keyboard support | Focus handle → Arrow keys (Shift for larger steps) |
| Persistence | Split ratio saved to `localStorage` (`panelSplitRatio` key) |
| Constraints | Min 15% / max 85% prevents panels from disappearing |
| Responsive | Handle hidden on viewports ≤1024px where panels stack vertically |

---

## Skills Matrix

| Skill | Module | API Endpoint | LLM Required |
|-------|--------|-------------|--------------|
| VBA Extraction | `vba_extractor.py` | `POST /api/upload` | No |
| VBA Conversion (LLM) | `llm_converter.py` | `POST /api/convert`, `/api/convert-all` | Yes |
| VBA Conversion (Offline) | `offline_converter.py` | `POST /api/convert`, `/api/convert-all` | No |
| Formula Extraction | `formula_extractor.py` | `POST /api/extract-formulas` | No |
| Formula Conversion | `llm_converter.py` or `offline_converter.py` | `POST /api/convert-formula` | Optional |
| Data Export | `data_exporter.py` | `POST /api/export-data` | No |
| Workbook Analysis | `workbook_analyzer.py` | `POST /api/analyze-workbook` | Yes (for VBA/formula conversion) |
| ZIP Download | `app.py` | `POST /api/download-zip` | No |
| Keyboard Shortcuts | `static/js/app.js` | — (client-side) | No |
| Collapsible Sections | `static/js/app.js` | — (client-side) | No |
| Resizable Code Panels | `static/js/app.js` | — (client-side) | No |

### Frontend Skills (JavaScript)

| Skill | Description |
|-------|-------------|
| Diff Highlighting | Color-coded inline keyword mapping highlights (VBA↔Python) with toggle |
| Conversion Time Tracking | `performance.now()` timing for single and batch conversions |
| Module Navigator Tabs | Prev/next navigation for batch-converted modules |
| Code Statistics | Live line/char/function/import counts in panel headers |
| Line Numbers | Auto-generated line numbering after Prism highlighting |
| Copy VBA / Python | Clipboard copy buttons for both code panels |
| Export / Import History | JSON export/import of conversion history |
| Search & Filter History | Live search + engine/status dropdown filters |
| Syntax Validation | Live VBA syntax checking with pattern-based validation |
| Batch Progress | Per-module progress updates during sequential conversion |
| Accessibility | ARIA, keyboard nav, SR announcements, skip link, focus management |
| Keyboard Shortcuts | Global shortcuts (Ctrl+Enter, Ctrl+S, ?, etc.) with help overlay |
| Collapsible Sections | Expand/collapse sidebar panels with localStorage persistence |
| Resizable Code Panels | Drag/keyboard resize handle between VBA and Python panels |

---

## Dependencies

| Package | Version | Used By |
|---------|---------|---------|
| FastAPI | ≥ 0.100.0 | `app.py` (web server & API) |
| Uvicorn | ≥ 0.29.0 | ASGI server for FastAPI |
| python-multipart | — | File upload handling in FastAPI |
| openpyxl | ≥ 3.1.0 | `formula_extractor.py`, `data_exporter.py`, `vba_extractor.py` |
| pandas | ≥ 2.0.0 | `data_exporter.py` |
| polars | ≥ 0.20.0 | Optional conversion target |
| oletools | — | `vba_extractor.py` (OLE VBA extraction) |
| olefile | — | `vba_extractor.py` (OLE binary parsing) |
| xlrd | ≥ 2.0.0 | `vba_extractor.py` (BIFF format `.xls` reading) |
| anthropic | — | `llm_converter.py` (Claude provider, optional) |
| openai | — | `llm_converter.py` (OpenAI provider, optional) |
| python-dotenv | — | Environment variable loading |
| xlsxwriter | — | Optional Excel file creation |
