# Claude Integration Guide

This document provides detailed information about integrating Claude (Anthropic) as the primary LLM for VBA to Python conversion.

## Overview

Claude is the recommended LLM provider for this application due to its:
- Superior code understanding and generation capabilities
- Large context window (200K tokens)
- Excellent instruction following for structured outputs
- Strong performance with legacy code formats like VBA

---

## Quick Start

### 1. Get API Key

1. Sign up at [console.anthropic.com](https://console.anthropic.com)
2. Navigate to **API Keys** section
3. Create a new API key
4. Copy the key (starts with `sk-ant-`)

### 2. Configure Environment

```bash
# .env file
ANTHROPIC_API_KEY=sk-ant-api03-your-key-here
LLM_PROVIDER=anthropic
LLM_MODEL=claude-sonnet-4-20250514
```

### 3. Test Connection

```python
from llm_converter import AnthropicConverter

converter = AnthropicConverter()
result = converter.convert("Sub Test()\nMsgBox \"Hello\"\nEnd Sub")
print(result.python_code)
```

---

## Available Models

| Model | Context | Best For | Cost |
|-------|---------|----------|------|
| `claude-sonnet-4-20250514` | 200K | **Recommended** - Best balance of quality/speed | $3/$15 per 1M tokens |
| `claude-3-opus-20240229` | 200K | Complex/large codebases | $15/$75 per 1M tokens |
| `claude-3-haiku-20240307` | 200K | Simple conversions, high volume | $0.25/$1.25 per 1M tokens |

### Model Selection Guide

```python
# For most conversions (default)
converter = AnthropicConverter(model="claude-sonnet-4-20250514")

# For complex enterprise VBA with many dependencies
converter = AnthropicConverter(model="claude-3-opus-20240229")

# For simple macros or high-volume batch processing
converter = AnthropicConverter(model="claude-3-haiku-20240307")
```

---

## API Configuration

### AnthropicConverter Class

```python
class AnthropicConverter(BaseLLMConverter):
    """VBA to Python converter using Anthropic's Claude API."""
    
    DEFAULT_MODEL = "claude-sonnet-4-20250514"
    
    def __init__(
        self, 
        api_key: Optional[str] = None,  # Uses ANTHROPIC_API_KEY env if not provided
        model: Optional[str] = None      # Uses DEFAULT_MODEL if not provided
    ):
        ...
```

### Configuration Options

| Option | Environment Variable | Default | Description |
|--------|---------------------|---------|-------------|
| API Key | `ANTHROPIC_API_KEY` | Required | Your Anthropic API key |
| Model | `LLM_MODEL` | `claude-sonnet-4-20250514` | Claude model to use |
| Max Tokens | - | `4096` | Maximum output tokens |

---

## Message Structure

### System Prompt

The system prompt establishes Claude's role and conversion rules:

```python
SYSTEM_PROMPT = """You are an expert VBA to Python converter. Your task is to convert 
VBA/VBScript code to clean, idiomatic Python code.

**Conversion Rules:**
1. Use pandas for Excel/data operations by default
2. Replace VBA Range operations with pandas DataFrame operations
3. Convert VBA Subs to Python functions (def function_name():)
4. Convert VBA Functions to Python functions with proper return types
5. Use type hints for all function parameters and returns
6. Replace MsgBox with print() or logging.info()
7. Convert VBA error handling (On Error) to try/except blocks
8. Use pathlib for file operations
9. Replace VBA date functions with datetime module
10. Add docstrings explaining what each function does
11. VBA is 1-indexed, Python is 0-indexed - adjust all array/range indices
12. Convert VBA constants (vbCrLf, vbTab) to Python equivalents
13. Handle Optional parameters with Python default arguments
14. Convert VBA collections to Python lists or dictionaries

**VBA to Python Type Mappings:**
- Integer, Long -> int
- Double, Single -> float
- String -> str
- Boolean -> bool
- Variant -> Any (from typing)
- Object -> appropriate class or Any
- Date -> datetime.datetime
- Collection -> list
- Dictionary -> dict

**Excel Object Model Mappings:**
- Range("A1:B10").Value -> df.loc[0:9, 'A':'B'] or df.iloc[0:10, 0:2]
- Cells(row, col) -> df.iloc[row-1, col-1]
- Worksheets("Sheet1") -> pd.read_excel(path, sheet_name="Sheet1")
- ActiveWorkbook -> workbook variable
- Application.WorksheetFunction.X -> pandas equivalent or numpy

**Output Format:**
1. Start with all necessary imports
2. Include type hints
3. Add docstrings to all functions
4. Include comments for non-obvious conversions
5. At the end, add a comment block listing any functionality that couldn't be directly converted"""
```

### User Prompt Template

```python
f"""Convert the following VBA code to Python. 

**Module Name:** {module_name}
**Target Library:** {target_library} (use this for data operations)

**VBA Code:**
```vba
{vba_code}
```

Provide complete, runnable Python code with all necessary imports."""
```

---

## API Call Implementation

### Making the Request

```python
def convert(self, vba_code: str, module_name: str = "converted_module",
            target_library: str = "pandas") -> ConversionResult:
    try:
        user_prompt = self._build_user_prompt(vba_code, module_name, target_library)
        
        message = self.client.messages.create(
            model=self.model,
            max_tokens=4096,
            system=self.SYSTEM_PROMPT,
            messages=[
                {"role": "user", "content": user_prompt}
            ]
        )
        
        response_text = message.content[0].text
        python_code = self._extract_python_code(response_text)
        notes = self._extract_notes_from_response(response_text)
        
        tokens_used = message.usage.input_tokens + message.usage.output_tokens
        
        return ConversionResult(
            success=True,
            python_code=python_code,
            conversion_notes=notes,
            tokens_used=tokens_used
        )
        
    except Exception as e:
        return ConversionResult(
            success=False,
            python_code="",
            error=str(e)
        )
```

### Response Parsing

```python
def _extract_python_code(self, response: str) -> str:
    """Extract Python code from Claude's response."""
    # Try to find code in markdown code blocks
    code_block = re.search(r'```python\n(.*?)```', response, re.DOTALL)
    if code_block:
        return code_block.group(1).strip()
    
    # Try without language specifier
    code_block = re.search(r'```\n(.*?)```', response, re.DOTALL)
    if code_block:
        return code_block.group(1).strip()
    
    # Return the whole response if no code blocks found
    return response.strip()
```

---

## Example Conversions

### Simple Sub Conversion

**Input VBA:**
```vba
Sub HelloWorld()
    MsgBox "Hello, World!"
End Sub
```

**Claude Output:**
```python
def hello_world() -> None:
    """Display a hello world message."""
    print("Hello, World!")
```

### Range Operations

**Input VBA:**
```vba
Sub CalculateSum()
    Dim ws As Worksheet
    Dim total As Double
    
    Set ws = ThisWorkbook.Worksheets("Data")
    total = Application.WorksheetFunction.Sum(ws.Range("A1:A100"))
    ws.Range("B1").Value = total
End Sub
```

**Claude Output:**
```python
import pandas as pd
from pathlib import Path

def calculate_sum(workbook_path: Path) -> None:
    """
    Calculate the sum of values in column A and write to cell B1.
    
    Args:
        workbook_path: Path to the Excel workbook
    """
    # Read the Data sheet
    df = pd.read_excel(workbook_path, sheet_name="Data")
    
    # Calculate sum of column A (first 100 rows)
    # Note: VBA Range("A1:A100") is 1-indexed, pandas is 0-indexed
    total: float = df.iloc[0:100, 0].sum()
    
    # Write result back to the workbook
    with pd.ExcelWriter(workbook_path, mode='a', if_sheet_exists='overlay') as writer:
        result_df = pd.DataFrame({'B': [total]})
        result_df.to_excel(writer, sheet_name="Data", startrow=0, startcol=1, 
                          index=False, header=False)
```

### Class Module Conversion

**Input VBA:**
```vba
' Class: Person
Private pName As String
Private pAge As Integer

Public Property Get Name() As String
    Name = pName
End Property

Public Property Let Name(value As String)
    pName = value
End Property

Public Property Get Age() As Integer
    Age = pAge
End Property

Public Property Let Age(value As Integer)
    If value >= 0 Then
        pAge = value
    End If
End Property

Public Function Greet() As String
    Greet = "Hello, my name is " & pName & " and I am " & pAge & " years old."
End Function
```

**Claude Output:**
```python
from dataclasses import dataclass
from typing import Optional

@dataclass
class Person:
    """
    Represents a person with name and age properties.
    
    Converted from VBA Class Module.
    """
    _name: str = ""
    _age: int = 0
    
    @property
    def name(self) -> str:
        """Get the person's name."""
        return self._name
    
    @name.setter
    def name(self, value: str) -> None:
        """Set the person's name."""
        self._name = value
    
    @property
    def age(self) -> int:
        """Get the person's age."""
        return self._age
    
    @age.setter
    def age(self, value: int) -> None:
        """Set the person's age (must be non-negative)."""
        if value >= 0:
            self._age = value
    
    def greet(self) -> str:
        """Return a greeting message."""
        return f"Hello, my name is {self._name} and I am {self._age} years old."
```

---

## Error Handling

### Common Errors

| Error | Cause | Solution |
|-------|-------|----------|
| `AuthenticationError` | Invalid API key | Check `ANTHROPIC_API_KEY` is correct |
| `RateLimitError` | Too many requests | Implement exponential backoff |
| `APIConnectionError` | Network issues | Check internet connection |
| `InvalidRequestError` | Malformed request | Validate input before sending |
| `OverloadedError` | API overloaded | Retry after delay |

### Implementing Retry Logic

```python
import time
from anthropic import RateLimitError, APIConnectionError

def convert_with_retry(self, vba_code: str, max_retries: int = 3) -> ConversionResult:
    """Convert with automatic retry on transient errors."""
    
    for attempt in range(max_retries):
        try:
            return self.convert(vba_code)
            
        except RateLimitError:
            if attempt < max_retries - 1:
                wait_time = 2 ** attempt  # Exponential backoff
                time.sleep(wait_time)
            else:
                return ConversionResult(
                    success=False,
                    python_code="",
                    error="Rate limit exceeded after retries"
                )
                
        except APIConnectionError:
            if attempt < max_retries - 1:
                time.sleep(1)
            else:
                return ConversionResult(
                    success=False,
                    python_code="",
                    error="Connection failed after retries"
                )
```

---

## Best Practices

### 1. Optimize Token Usage

```python
# Remove comments and unnecessary whitespace before conversion
import re

def preprocess_vba(code: str) -> str:
    """Clean VBA code before sending to Claude."""
    # Remove single-line comments
    code = re.sub(r"'.*$", "", code, flags=re.MULTILINE)
    # Remove empty lines
    code = re.sub(r"\n\s*\n", "\n", code)
    return code.strip()
```

### 2. Handle Large Files

```python
def convert_large_module(self, vba_code: str, chunk_size: int = 100) -> str:
    """Convert large VBA modules by splitting into chunks."""
    lines = vba_code.split('\n')
    
    if len(lines) <= chunk_size:
        return self.convert(vba_code).python_code
    
    # Split by Sub/Function boundaries
    chunks = self._split_by_procedures(vba_code)
    
    converted_chunks = []
    for chunk in chunks:
        result = self.convert(chunk)
        if result.success:
            converted_chunks.append(result.python_code)
    
    return self._merge_python_code(converted_chunks)
```

### 3. Validate Output

```python
import ast

def validate_python_code(code: str) -> bool:
    """Check if generated Python code is syntactically valid."""
    try:
        ast.parse(code)
        return True
    except SyntaxError:
        return False

# Usage
result = converter.convert(vba_code)
if result.success and validate_python_code(result.python_code):
    print("Valid Python code generated")
else:
    print("Conversion produced invalid Python - retry with different prompt")
```

### 4. Cache Conversions

```python
import hashlib
from functools import lru_cache

@lru_cache(maxsize=100)
def cached_convert(vba_code_hash: str, vba_code: str) -> str:
    """Cache conversion results to avoid duplicate API calls."""
    converter = AnthropicConverter()
    result = converter.convert(vba_code)
    return result.python_code if result.success else ""

def convert_with_cache(vba_code: str) -> str:
    """Convert VBA with caching."""
    code_hash = hashlib.md5(vba_code.encode()).hexdigest()
    return cached_convert(code_hash, vba_code)
```

---

## Rate Limits

### Anthropic Rate Limits (as of 2024)

| Tier | Requests/min | Tokens/min | Tokens/day |
|------|--------------|------------|------------|
| Free | 5 | 20,000 | 300,000 |
| Build | 50 | 40,000 | 1,000,000 |
| Scale | 1,000 | 80,000 | 5,000,000 |

### Handling Rate Limits

```python
from anthropic import RateLimitError
import time

class RateLimitedConverter(AnthropicConverter):
    """Converter with built-in rate limiting."""
    
    def __init__(self, requests_per_minute: int = 50, **kwargs):
        super().__init__(**kwargs)
        self.min_interval = 60 / requests_per_minute
        self.last_request_time = 0
    
    def convert(self, vba_code: str, **kwargs) -> ConversionResult:
        # Enforce minimum interval between requests
        elapsed = time.time() - self.last_request_time
        if elapsed < self.min_interval:
            time.sleep(self.min_interval - elapsed)
        
        self.last_request_time = time.time()
        return super().convert(vba_code, **kwargs)
```

---

## Monitoring & Debugging

### Enable Debug Logging

```python
import logging

# Enable Anthropic SDK logging
logging.getLogger("anthropic").setLevel(logging.DEBUG)

# Custom logging for conversions
logger = logging.getLogger("claude_converter")
logger.setLevel(logging.INFO)

handler = logging.StreamHandler()
handler.setFormatter(logging.Formatter(
    '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
))
logger.addHandler(handler)
```

### Track Token Usage

```python
class TokenTracker:
    """Track token usage across conversions."""
    
    def __init__(self):
        self.total_input_tokens = 0
        self.total_output_tokens = 0
        self.conversion_count = 0
    
    def record(self, result: ConversionResult):
        if result.success:
            self.total_input_tokens += result.tokens_used  # Simplified
            self.conversion_count += 1
    
    def report(self) -> dict:
        return {
            "total_tokens": self.total_input_tokens + self.total_output_tokens,
            "conversions": self.conversion_count,
            "avg_tokens_per_conversion": (
                (self.total_input_tokens + self.total_output_tokens) / 
                max(1, self.conversion_count)
            )
        }
```

---

## Security Considerations

1. **Never expose API keys in frontend code**
2. **Validate all VBA input before sending to API**
3. **Sanitize converted Python code before execution**
4. **Use environment variables for sensitive configuration**
5. **Implement request signing for production deployments**

```python
# Example: Input validation
def validate_vba_input(code: str) -> bool:
    """Validate VBA code before processing."""
    if len(code) > 100000:  # Limit code size
        return False
    if any(keyword in code.lower() for keyword in ['shell', 'wscript', 'createobject']):
        # Flag potentially dangerous VBA
        logging.warning("Potentially dangerous VBA code detected")
    return True
```

---

## Resources

- [Anthropic API Documentation](https://docs.anthropic.com)
- [Claude Prompt Engineering Guide](https://docs.anthropic.com/claude/docs/prompt-engineering)
- [Anthropic Python SDK](https://github.com/anthropics/anthropic-sdk-python)
- [Claude Model Comparison](https://docs.anthropic.com/claude/docs/models-overview)
