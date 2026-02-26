# AI Agents Documentation

This document describes the AI agents and LLM integration used in the VBA to Python Converter application.

## Architecture Overview

```
┌─────────────────────────────────────────────────────────────────┐
│                      User Request                               │
│              (VBA Code + Conversion Options)                    │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                FastAPI Backend (app.py)                         │
│            Routes request based on provider                     │
└─────────────────────────────────────────────────────────────────┘
                              │
         ┌────────────────────┼────────────────────┐
         ▼                    ▼                    ▼
┌──────────────────┐ ┌──────────────────┐ ┌──────────────────┐
│ AnthropicConverter │ │  OpenAIConverter  │ │ OfflineConverter  │
│ (Claude Agent)     │ │  (GPT Agent)      │ │ (Rule-Based)     │
│                    │ │                   │ │                  │
│ claude-sonnet-4-20250514  │ │ gpt-4-turbo       │ │ No API key       │
│ Max Tokens: 4096  │ │ Max Tokens: 4096  │ │ Deterministic    │
└──────────────────┘ └──────────────────┘ └──────────────────┘
```

## Agent Classes

### 1. VBAToPythonConverter (Main Orchestrator)

**Location:** `llm_converter.py`

The main entry point that orchestrates the conversion process.

```python
from llm_converter import VBAToPythonConverter

converter = VBAToPythonConverter(provider="anthropic")
result = converter.convert(vba_code, module_name="MyModule", target_library="pandas")
```

**Responsibilities:**
- Auto-detect available LLM providers
- Route requests to appropriate converter
- Manage conversion state and notes
- Handle fallback between providers

**Configuration:**
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `provider` | str | `"anthropic"` | LLM provider: `"anthropic"` or `"openai"` |

---

### 2. AnthropicConverter (Claude Agent)

**Location:** `llm_converter.py`

Handles conversion using Anthropic's Claude API.

```python
from llm_converter import AnthropicConverter

converter = AnthropicConverter(
    api_key="sk-ant-...",  # Optional, uses env var
    model="claude-sonnet-4-20250514"
)
result = converter.convert(vba_code)
```

**Configuration:**
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `api_key` | str | `ANTHROPIC_API_KEY` env | API key for authentication |
| `model` | str | `claude-sonnet-4-20250514` | Claude model to use |

**Supported Models:**
- `claude-sonnet-4-20250514` (recommended)
- `claude-3-opus-20240229`
- `claude-3-haiku-20240307`

---

### 3. OpenAIConverter (GPT Agent)

**Location:** `llm_converter.py`

Handles conversion using OpenAI's API.

```python
from llm_converter import OpenAIConverter

converter = OpenAIConverter(
    api_key="sk-...",  # Optional, uses env var
    model="gpt-4-turbo"
)
result = converter.convert(vba_code)
```

**Configuration:**
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `api_key` | str | `OPENAI_API_KEY` env | API key for authentication |
| `model` | str | `gpt-4-turbo` | OpenAI model to use |

**Supported Models:**
- `gpt-4-turbo` (recommended)
- `gpt-4`
- `gpt-4o`
- `gpt-3.5-turbo`

---

### 4. OfflineConverter (Rule-Based Agent)

**Location:** `offline_converter.py`

A deterministic, rule-based converter that requires no API key or internet connection. Ideal for quick conversions and environments without LLM access.

```python
from offline_converter import OfflineConverter

converter = OfflineConverter()
result = converter.convert(vba_code, module_name="Module1")
print(result["python_code"])
```

**Capabilities:**
| Feature | Details |
|---------|---------|
| VBA Subs/Functions | Converted to Python `def` with type hints |
| Control flow | `If`/`Select Case`/`For`/`While`/`Do` loops |
| Type mapping | VBA types → Python type hints |
| Constants | `vbCrLf`, `vbTab`, `True`, `False`, etc. |
| Built-in functions | `MsgBox` → `print()`, `InputBox` → `input()` |
| Formula conversion | `VLOOKUP`, `SUMIF`, `IF`, etc. → pandas/numpy |
| `With` blocks | Flattened to explicit object references |
| Error handling | `On Error` → `try`/`except` |

**When to use Offline vs LLM:**
| Scenario | Recommended Agent |
|----------|-------------------|
| Simple macros, standard patterns | OfflineConverter |
| Complex business logic, context-dependent | AnthropicConverter / OpenAIConverter |
| No API key available | OfflineConverter |
| Highest fidelity needed | AnthropicConverter |
| Batch conversion of many modules | OfflineConverter (faster, no rate limits) |

---

## System Prompt

All agents use a shared system prompt that defines the conversion rules:

```markdown
You are an expert VBA to Python converter. Your task is to convert 
VBA/VBScript code to clean, idiomatic Python code.

**Conversion Rules:**
1. Use pandas for Excel/data operations by default
2. Replace VBA Range operations with pandas DataFrame operations
3. Convert VBA Subs to Python functions
4. Convert VBA Functions to Python functions with proper return types
5. Use type hints for all function parameters and returns
6. Replace MsgBox with print() or logging.info()
7. Convert VBA error handling (On Error) to try/except blocks
8. Use pathlib for file operations
9. Replace VBA date functions with datetime module
10. Add docstrings explaining what each function does
...
```

See full prompt in `llm_converter.py` → `BaseLLMConverter.SYSTEM_PROMPT`

---

## Data Flow

### Conversion Request Flow

```
1. User submits VBA code via UI or API
                    │
                    ▼
2. FastAPI endpoint receives request (/api/convert)
                    │
                    ▼
3. Provider routed based on request
   ├── provider="offline"    → OfflineConverter.convert()
   ├── provider="anthropic"  → AnthropicConverter.convert()
   └── provider="openai"     → OpenAIConverter.convert()
                    │
                    ▼
4. Conversion executed
   ├── Offline: Rule-based pattern matching (instant)
   └── LLM: System prompt + User prompt sent to API
                    │
                    ▼
5. Response parsed and validated
   ├── Extract code from markdown blocks (LLM)
   ├── Extract conversion notes
   └── Identify engine used
                    │
                    ▼
6. Result returned to user with engine label
```

### ConversionResult Schema

```python
@dataclass
class ConversionResult:
    success: bool              # Whether conversion succeeded
    python_code: str           # The converted Python code
    conversion_notes: list[str] # Notes about the conversion
    error: Optional[str]       # Error message if failed
    tokens_used: int           # API tokens consumed
```

---

## Prompt Engineering

### User Prompt Template

```python
def _build_user_prompt(self, vba_code: str, module_name: str, 
                       target_library: str) -> str:
    return f"""Convert the following VBA code to Python. 

**Module Name:** {module_name}
**Target Library:** {target_library} (use this for data operations)

**VBA Code:**
```vba
{vba_code}
```

Provide complete, runnable Python code with all necessary imports."""
```

### Key Prompt Features

1. **Structured Instructions**: Clear numbered rules for consistent output
2. **Type Mappings**: Explicit VBA → Python type conversions
3. **Library Targeting**: Support for pandas or polars
4. **Output Format**: Specifies imports, type hints, docstrings
5. **Edge Case Handling**: Notes on 1-indexed vs 0-indexed arrays

---

## Error Handling

### Agent-Level Errors

```python
try:
    result = converter.convert(vba_code)
except ValueError as e:
    # API key not configured
    print(f"Configuration error: {e}")
except ImportError as e:
    # SDK not installed
    print(f"Missing dependency: {e}")
```

### Conversion Errors

```python
result = converter.convert_with_result(vba_code)

if not result.success:
    print(f"Conversion failed: {result.error}")
else:
    print(result.python_code)
```

### Common Error Scenarios

| Error | Cause | Solution |
|-------|-------|----------|
| `API key not provided` | Missing env var | Set `ANTHROPIC_API_KEY` or `OPENAI_API_KEY` |
| `Rate limit exceeded` | Too many requests | Implement backoff/retry |
| `Context length exceeded` | VBA code too long | Split into smaller modules |
| `Invalid response` | LLM returned non-code | Retry with clearer prompt |

---

## Token Usage & Costs

### Estimated Token Usage

| VBA Code Size | Input Tokens | Output Tokens | Total |
|---------------|--------------|---------------|-------|
| Small (< 50 lines) | ~500 | ~800 | ~1,300 |
| Medium (50-200 lines) | ~1,500 | ~2,000 | ~3,500 |
| Large (200+ lines) | ~3,000 | ~4,000 | ~7,000 |

### Cost Optimization Tips

1. **Batch conversions**: Convert multiple small modules together
2. **Cache results**: Store conversions to avoid re-processing
3. **Use appropriate model**: Haiku/GPT-3.5 for simple code
4. **Trim unnecessary code**: Remove comments before conversion

---

## Extending the Agents

### Adding a New LLM Provider

1. Create a new class extending `BaseLLMConverter`:

```python
class GoogleConverter(BaseLLMConverter):
    DEFAULT_MODEL = "gemini-pro"
    
    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        super().__init__(model or self.DEFAULT_MODEL)
        self.api_key = api_key or os.getenv("GOOGLE_API_KEY")
        # Initialize client...
    
    def convert(self, vba_code: str, module_name: str = "converted_module",
                target_library: str = "pandas") -> ConversionResult:
        # Implement conversion...
        pass
```

2. Register in `VBAToPythonConverter._get_converter()`:

```python
elif self.provider == "google":
    self._converter = GoogleConverter()
```

### Customizing the System Prompt

Override the class attribute:

```python
class CustomConverter(AnthropicConverter):
    SYSTEM_PROMPT = """
    Your custom conversion instructions here...
    Focus on polars instead of pandas...
    """
```

---

## Environment Variables

| Variable | Required | Default | Description |
|----------|----------|---------|-------------|
| `ANTHROPIC_API_KEY` | No* | - | Anthropic API key |
| `OPENAI_API_KEY` | No* | - | OpenAI API key |
| `LLM_PROVIDER` | No | `anthropic` | Default LLM provider |
| `LLM_MODEL` | No | Provider default | Model override |

*API keys are only needed for LLM-powered conversion. The offline converter works without any keys.

---

## Testing Agents

### Unit Testing

```python
import pytest
from unittest.mock import Mock, patch
from llm_converter import VBAToPythonConverter, ConversionResult

@patch('llm_converter.AnthropicConverter')
def test_conversion(mock_converter):
    mock_converter.return_value.convert.return_value = ConversionResult(
        success=True,
        python_code="def example(): pass",
        conversion_notes=[],
        tokens_used=100
    )
    
    converter = VBAToPythonConverter()
    result = converter.convert("Sub Example()\nEnd Sub")
    
    assert "def example" in result
```

### Integration Testing

```python
@pytest.mark.integration
def test_real_conversion():
    """Test with actual API (requires valid API key)."""
    converter = VBAToPythonConverter()
    
    vba_code = """
    Sub HelloWorld()
        MsgBox "Hello, World!"
    End Sub
    """
    
    result = converter.convert_with_result(vba_code)
    
    assert result.success
    assert "def" in result.python_code
    assert "print" in result.python_code
```

---

## Monitoring & Logging

### Recommended Logging Setup

```python
import logging

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("vba_converter")

# In your converter
logger.info(f"Converting module: {module_name}")
logger.debug(f"Token usage: {result.tokens_used}")
logger.error(f"Conversion failed: {result.error}")
```

### Metrics to Track

- Conversion success rate
- Average tokens per conversion
- Response time per provider (now tracked automatically via `performance.now()`)
- Error distribution by type
- Per-module timing in batch conversions

---

## Frontend Agent Integration

The frontend (`app.js`) acts as a thin orchestration layer that:

1. **Times conversions** via `performance.now()` and displays elapsed time in the status bar and history
2. **Announces results** to screen readers via `announceToSR()` for accessibility
3. **Highlights mappings** using `applyDiffHighlights()` to mark VBA↔Python keyword correspondences inline
4. **Tracks history** with `addToHistory()`, persisting module name, engine, success status, and duration to `localStorage`
5. **Manages batches** with `finalizeBatchConversion()`, building module tabs and computing per-module timing
6. **Keyboard shortcuts** via `setupKeyboardShortcuts()` — global hotkeys for conversion, file operations, clipboard, navigation, and theme toggle; `?` opens a categorised shortcuts overlay
7. **Collapsible sections** via `setupCollapsibleSections()` — expand/collapse sidebar panels with persisted state in `localStorage`
8. **Resizable panels** via `setupResizablePanels()` — draggable/keyboard-driven handle between VBA and Python code panels with persisted split ratio
