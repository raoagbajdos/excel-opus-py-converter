"""
LLM-Powered VBA to Python Converter

This module handles the conversion of VBA code to Python using
Claude (Anthropic) or OpenAI APIs.
"""
import logging
import os
import re
import time
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional
from dotenv import load_dotenv

logger = logging.getLogger(__name__)


class ConversionError(RuntimeError):
    """Raised when VBA-to-Python or formula conversion fails."""


# Retry defaults (can be overridden via config)
MAX_RETRIES = int(os.getenv('LLM_MAX_RETRIES', '3'))
RETRY_BASE_DELAY = float(os.getenv('LLM_RETRY_BASE_DELAY', '1.0'))

# Load environment variables
load_dotenv()


@dataclass
class ConversionResult:
    """Result of a VBA to Python conversion."""
    success: bool
    python_code: str
    conversion_notes: list[str] = field(default_factory=list)
    error: Optional[str] = None
    tokens_used: int = 0


class BaseLLMConverter(ABC):
    """Abstract base class for LLM converters."""
    
    VBA_SYSTEM_PROMPT = """You are an expert VBA to Python converter. Your task is to convert VBA/VBScript code to clean, idiomatic Python code.

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

    FORMULA_SYSTEM_PROMPT = """You are an expert Excel formula to Python converter. Convert Excel formulas to equivalent pandas/numpy operations.

**Excel Function to Python Mappings:**

**Lookup/Reference:**
- VLOOKUP(value, range, col, FALSE) -> df.merge() or df.set_index().loc[]
- XLOOKUP(lookup, array, return) -> df.set_index().loc[] or pd.merge()
- INDEX(array, row, col) -> df.iloc[row-1, col-1]
- MATCH(value, array, 0) -> (df == value).idxmax()

**Math/Statistics:**
- SUM(range) -> df['col'].sum() or df.sum()
- SUMIF(range, criteria, sum_range) -> df[df['col'] == criteria]['sum_col'].sum()
- SUMIFS(sum_range, criteria_range1, criteria1, ...) -> df[(df['col1'] == c1) & (df['col2'] == c2)]['sum_col'].sum()
- AVERAGE(range) -> df['col'].mean()
- COUNT(range) -> df['col'].count()
- COUNTIF(range, criteria) -> (df['col'] == criteria).sum()
- MAX/MIN(range) -> df['col'].max() / df['col'].min()

**Logical:**
- IF(condition, true_val, false_val) -> np.where(condition, true_val, false_val) or df['col'].apply(lambda x: true_val if condition else false_val)
- IFS(cond1, val1, cond2, val2, ...) -> np.select([cond1, cond2], [val1, val2])
- AND(cond1, cond2) -> (cond1) & (cond2)
- OR(cond1, cond2) -> (cond1) | (cond2)

**Text:**
- CONCATENATE/CONCAT(a, b, c) -> df['col1'] + df['col2'] or df['col1'].str.cat(df['col2'])
- LEFT(text, n) -> df['col'].str[:n]
- RIGHT(text, n) -> df['col'].str[-n:]
- MID(text, start, len) -> df['col'].str[start-1:start-1+len]
- LEN(text) -> df['col'].str.len()
- UPPER/LOWER(text) -> df['col'].str.upper() / df['col'].str.lower()
- TRIM(text) -> df['col'].str.strip()

**Date/Time:**
- TODAY() -> pd.Timestamp.today() or datetime.date.today()
- NOW() -> pd.Timestamp.now() or datetime.datetime.now()
- YEAR/MONTH/DAY(date) -> df['date_col'].dt.year / .dt.month / .dt.day
- DATEDIF(start, end, unit) -> (end - start).days or use relativedelta

**Array/Modern:**
- FILTER(array, condition) -> df[condition]
- SORT(array, col, order) -> df.sort_values(by='col', ascending=order)
- UNIQUE(array) -> df['col'].unique() or df.drop_duplicates()

**Conversion Guidelines:**
1. Use vectorized pandas operations instead of loops
2. Handle cell references by translating to DataFrame column names
3. Convert range references to DataFrame slicing
4. Use numpy for element-wise operations
5. Handle array formulas with apply() or vectorized operations
6. Add type hints and docstrings
7. Include error handling for edge cases

**Output Format:**
Provide Python code that:
1. Includes necessary imports (pandas, numpy, datetime)
2. Shows how to apply the formula to a DataFrame
3. Includes both the formula logic and usage example
4. Adds comments explaining the conversion"""

    def __init__(self, model: Optional[str] = None):
        """Initialize the converter with optional model override."""
        self.model = model
        self._conversion_notes: list[str] = []

    @staticmethod
    def _retry_with_backoff(fn, max_retries: int = MAX_RETRIES,
                            base_delay: float = RETRY_BASE_DELAY):
        """Execute *fn()* with exponential-backoff retry on transient errors."""
        last_exc: Exception | None = None
        for attempt in range(1, max_retries + 1):
            try:
                return fn()
            except Exception as exc:
                last_exc = exc
                err_str = str(exc).lower()
                # Retry on rate-limit (429), server errors (5xx), or overloaded
                retryable = any(tok in err_str for tok in (
                    '429', 'rate', 'overloaded', '529', '500', '502', '503',
                    'timeout', 'connection',
                ))
                if not retryable or attempt == max_retries:
                    raise
                delay = base_delay * (2 ** (attempt - 1))
                logger.warning(
                    "LLM call failed (attempt %d/%d): %s  — retrying in %.1fs",
                    attempt, max_retries, exc, delay,
                )
                time.sleep(delay)
        raise last_exc  # unreachable, but keeps type-checker happy

    @abstractmethod
    def convert(self, vba_code: str, module_name: str = "converted_module",
                target_library: str = "pandas") -> ConversionResult:
        """Convert VBA code to Python."""
        pass
    
    @abstractmethod
    def convert_formula(self, formula: str, cell_address: str = "A1",
                       sheet_name: str = "Sheet1") -> ConversionResult:
        """Convert Excel formula to Python code."""
        pass
    
    def _build_user_prompt(self, vba_code: str, module_name: str, 
                           target_library: str) -> str:
        """Build the user prompt for VBA conversion."""
        return f"""Convert the following VBA code to Python. 

**Module Name:** {module_name}
**Target Library:** {target_library} (use this for data operations)

**VBA Code:**
```vba
{vba_code}
```

Provide complete, runnable Python code with all necessary imports."""

    def _build_formula_prompt(self, formula: str, cell_address: str,
                             sheet_name: str) -> str:
        """Build the user prompt for formula conversion."""
        return f"""Convert the following Excel formula to Python code using pandas.

**Sheet:** {sheet_name}
**Cell:** {cell_address}
**Formula:** {formula}

Provide Python code that:
1. Shows how to implement this formula logic
2. Includes a function that can be applied to a DataFrame
3. Includes usage example
4. Handles edge cases"""

    def _extract_notes_from_response(self, response: str) -> list[str]:
        """Extract conversion notes from the LLM response."""
        notes = []
        
        # Look for notes in comments at the end
        note_patterns = [
            r'# Note: (.+)',
            r'# TODO: (.+)',
            r'# Warning: (.+)',
            r'# Conversion note: (.+)',
        ]
        
        for pattern in note_patterns:
            matches = re.findall(pattern, response, re.IGNORECASE)
            notes.extend(matches)
        
        # Look for a notes section
        notes_section = re.search(
            r'(?:# Notes?:|# Conversion Notes?:)\s*\n((?:#.+\n)+)',
            response,
            re.IGNORECASE
        )
        if notes_section:
            section_notes = re.findall(r'# - (.+)', notes_section.group(1))
            notes.extend(section_notes)
        
        return notes

    def _extract_python_code(self, response: str) -> str:
        """Extract Python code from the LLM response."""
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


class AnthropicConverter(BaseLLMConverter):
    """VBA to Python converter using Anthropic's Claude API."""
    
    DEFAULT_MODEL = "claude-sonnet-4-20250514"
    
    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        """
        Initialize the Anthropic converter.
        
        Args:
            api_key: Anthropic API key. If not provided, uses ANTHROPIC_API_KEY env var.
            model: Model to use. Defaults to claude-sonnet-4-20250514.
        """
        super().__init__(model or self.DEFAULT_MODEL)
        self.api_key = api_key or os.getenv("ANTHROPIC_API_KEY")
        
        if not self.api_key:
            raise ValueError(
                "Anthropic API key not provided. Set ANTHROPIC_API_KEY environment variable "
                "or pass api_key parameter."
            )
        
        try:
            import anthropic
            self.client = anthropic.Anthropic(api_key=self.api_key)
        except ImportError:
            raise ImportError("Please install anthropic: pip install anthropic")
    
    def convert(self, vba_code: str, module_name: str = "converted_module",
                target_library: str = "pandas") -> ConversionResult:
        """
        Convert VBA code to Python using Claude.
        
        Args:
            vba_code: The VBA code to convert
            module_name: Name for the output module
            target_library: Python library to use (pandas/polars)
            
        Returns:
            ConversionResult with the converted code
        """
        try:
            user_prompt = self._build_user_prompt(vba_code, module_name, target_library)

            def _call():
                return self.client.messages.create(
                    model=self.model,
                    max_tokens=4096,
                    system=self.VBA_SYSTEM_PROMPT,
                    messages=[{"role": "user", "content": user_prompt}],
                )

            message = self._retry_with_backoff(_call)

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
            logger.exception("Anthropic VBA conversion failed")
            return ConversionResult(
                success=False,
                python_code="",
                error=str(e)
            )
    
    def convert_formula(self, formula: str, cell_address: str = "A1",
                       sheet_name: str = "Sheet1") -> ConversionResult:
        """
        Convert Excel formula to Python using Claude.
        
        Args:
            formula: The Excel formula to convert
            cell_address: Cell address where formula is located
            sheet_name: Sheet name containing the formula
            
        Returns:
            ConversionResult with the converted code
        """
        try:
            user_prompt = self._build_formula_prompt(formula, cell_address, sheet_name)

            def _call():
                return self.client.messages.create(
                    model=self.model,
                    max_tokens=2048,
                    system=self.FORMULA_SYSTEM_PROMPT,
                    messages=[{"role": "user", "content": user_prompt}],
                )

            message = self._retry_with_backoff(_call)

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
            logger.exception("Anthropic formula conversion failed")
            return ConversionResult(
                success=False,
                python_code="",
                error=str(e)
            )


class OpenAIConverter(BaseLLMConverter):
    """VBA to Python converter using OpenAI's API."""
    
    DEFAULT_MODEL = "gpt-4-turbo"
    
    def __init__(self, api_key: Optional[str] = None, model: Optional[str] = None):
        """
        Initialize the OpenAI converter.
        
        Args:
            api_key: OpenAI API key. If not provided, uses OPENAI_API_KEY env var.
            model: Model to use. Defaults to gpt-4-turbo.
        """
        super().__init__(model or self.DEFAULT_MODEL)
        self.api_key = api_key or os.getenv("OPENAI_API_KEY")
        
        if not self.api_key:
            raise ValueError(
                "OpenAI API key not provided. Set OPENAI_API_KEY environment variable "
                "or pass api_key parameter."
            )
        
        try:
            from openai import OpenAI
            self.client = OpenAI(api_key=self.api_key)
        except ImportError:
            raise ImportError("Please install openai: pip install openai")
    
    def convert(self, vba_code: str, module_name: str = "converted_module",
                target_library: str = "pandas") -> ConversionResult:
        """
        Convert VBA code to Python using OpenAI.
        
        Args:
            vba_code: The VBA code to convert
            module_name: Name for the output module
            target_library: Python library to use (pandas/polars)
            
        Returns:
            ConversionResult with the converted code
        """
        try:
            user_prompt = self._build_user_prompt(vba_code, module_name, target_library)

            def _call():
                return self.client.chat.completions.create(
                    model=self.model,
                    max_tokens=4096,
                    messages=[
                        {"role": "system", "content": self.VBA_SYSTEM_PROMPT},
                        {"role": "user", "content": user_prompt},
                    ],
                )

            response = self._retry_with_backoff(_call)

            response_text = response.choices[0].message.content
            python_code = self._extract_python_code(response_text)
            notes = self._extract_notes_from_response(response_text)
            
            tokens_used = response.usage.total_tokens if response.usage else 0
            
            return ConversionResult(
                success=True,
                python_code=python_code,
                conversion_notes=notes,
                tokens_used=tokens_used
            )
            
        except Exception as e:
            logger.exception("OpenAI VBA conversion failed")
            return ConversionResult(
                success=False,
                python_code="",
                error=str(e)
            )
    
    def convert_formula(self, formula: str, cell_address: str = "A1",
                       sheet_name: str = "Sheet1") -> ConversionResult:
        """
        Convert Excel formula to Python using OpenAI.
        
        Args:
            formula: The Excel formula to convert
            cell_address: Cell address where formula is located
            sheet_name: Sheet name containing the formula
            
        Returns:
            ConversionResult with the converted code
        """
        try:
            user_prompt = self._build_formula_prompt(formula, cell_address, sheet_name)

            def _call():
                return self.client.chat.completions.create(
                    model=self.model,
                    max_tokens=2048,
                    messages=[
                        {"role": "system", "content": self.FORMULA_SYSTEM_PROMPT},
                        {"role": "user", "content": user_prompt},
                    ],
                )

            response = self._retry_with_backoff(_call)

            response_text = response.choices[0].message.content
            python_code = self._extract_python_code(response_text)
            notes = self._extract_notes_from_response(response_text)
            
            tokens_used = response.usage.total_tokens if response.usage else 0
            
            return ConversionResult(
                success=True,
                python_code=python_code,
                conversion_notes=notes,
                tokens_used=tokens_used
            )
            
        except Exception as e:
            logger.exception("OpenAI formula conversion failed")
            return ConversionResult(
                success=False,
                python_code="",
                error=str(e)
            )


class _OfflineConverterAdapter(BaseLLMConverter):
    """Adapter wrapping OfflineConverter to satisfy the BaseLLMConverter ABC."""

    def __init__(self) -> None:
        # Import here to avoid circular imports at module level
        from offline_converter import OfflineConverter as _OC
        self._engine = _OC()

    def convert(self, vba_code: str, module_name: str = "converted_module",
                target_library: str = "pandas") -> ConversionResult:
        r = self._engine.convert(vba_code, module_name, target_library)
        return ConversionResult(
            success=r.success, python_code=r.python_code,
            conversion_notes=r.conversion_notes,
            error=r.error, tokens_used=0,
        )

    def convert_formula(self, formula: str, cell_address: str = "A1",
                        sheet_name: str = "Sheet1") -> ConversionResult:
        r = self._engine.convert_formula(formula, cell_address, sheet_name)
        return ConversionResult(
            success=r.success, python_code=r.python_code,
            conversion_notes=r.conversion_notes,
            error=r.error, tokens_used=0,
        )


class VBAToPythonConverter:
    """
    Main converter class that automatically selects the appropriate LLM backend.
    """
    
    def __init__(self, provider: Optional[str] = None):
        """
        Initialize the converter.
        
        Args:
            provider: 'anthropic', 'openai', 'offline', or None for auto-detection
        """
        self.provider = provider or os.getenv("LLM_PROVIDER", "anthropic")
        self._converter: Optional[BaseLLMConverter] = None
        self._last_notes: list[str] = []
        
    def _get_converter(self) -> BaseLLMConverter:
        """Get or create the appropriate converter."""
        if self._converter is None:
            if self.provider == "offline":
                self._converter = _OfflineConverterAdapter()
            elif self.provider == "anthropic":
                self._converter = AnthropicConverter()
            elif self.provider == "openai":
                self._converter = OpenAIConverter()
            else:
                # Try Anthropic first, then OpenAI, then fall back to offline
                try:
                    self._converter = AnthropicConverter()
                except ValueError:
                    try:
                        self._converter = OpenAIConverter()
                    except ValueError:
                        logger.info("No LLM API key found — falling back to offline converter.")
                        self._converter = _OfflineConverterAdapter()
        return self._converter
    
    def convert(self, vba_code: str, module_name: str = "converted_module",
                target_library: str = "pandas") -> str:
        """
        Convert VBA code to Python.
        
        Args:
            vba_code: The VBA code to convert
            module_name: Name for the output module
            target_library: Python library to use (pandas/polars)
            
        Returns:
            Converted Python code as string
            
        Raises:
            ConversionError: If conversion fails
        """
        converter = self._get_converter()
        result = converter.convert(vba_code, module_name, target_library)
        
        self._last_notes = result.conversion_notes
        
        if not result.success:
            raise ConversionError(f"Conversion failed: {result.error}")
        
        return result.python_code
    
    def get_conversion_notes(self) -> list[str]:
        """Get notes from the last conversion."""
        return self._last_notes.copy()
    
    def convert_with_result(self, vba_code: str, module_name: str = "converted_module",
                            target_library: str = "pandas") -> ConversionResult:
        """
        Convert VBA code and return full result object.
        
        Args:
            vba_code: The VBA code to convert
            module_name: Name for the output module
            target_library: Python library to use (pandas/polars)
            
        Returns:
            ConversionResult object with all details
        """
        converter = self._get_converter()
        result = converter.convert(vba_code, module_name, target_library)
        self._last_notes = result.conversion_notes
        return result
    
    def convert_formula(self, formula: str, cell_address: str = "A1",
                       sheet_name: str = "Sheet1") -> str:
        """
        Convert Excel formula to Python code.
        
        Args:
            formula: The Excel formula to convert
            cell_address: Cell address where formula is located
            sheet_name: Sheet name containing the formula
            
        Returns:
            Converted Python code as string
            
        Raises:
            ConversionError: If conversion fails
        """
        converter = self._get_converter()
        result = converter.convert_formula(formula, cell_address, sheet_name)
        
        self._last_notes = result.conversion_notes
        
        if not result.success:
            raise ConversionError(f"Formula conversion failed: {result.error}")
        
        return result.python_code
    
    def convert_formula_with_result(self, formula: str, cell_address: str = "A1",
                                    sheet_name: str = "Sheet1") -> ConversionResult:
        """
        Convert Excel formula and return full result object.
        
        Args:
            formula: The Excel formula to convert
            cell_address: Cell address where formula is located
            sheet_name: Sheet name containing the formula
            
        Returns:
            ConversionResult object with all details
        """
        converter = self._get_converter()
        result = converter.convert_formula(formula, cell_address, sheet_name)
        self._last_notes = result.conversion_notes
        return result


# Convenience function for simple usage
def convert_vba_to_python(vba_code: str, 
                          module_name: str = "converted_module",
                          target_library: str = "pandas",
                          provider: Optional[str] = None) -> str:
    """
    Convert VBA code to Python.
    
    Args:
        vba_code: The VBA code to convert
        module_name: Name for the output module  
        target_library: Python library to use (pandas/polars)
        provider: LLM provider ('anthropic' or 'openai')
        
    Returns:
        Converted Python code as string
    """
    converter = VBAToPythonConverter(provider=provider)
    return converter.convert(vba_code, module_name, target_library)


if __name__ == "__main__":
    # Example usage
    sample_vba = '''
Sub CalculateTotal()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim total As Double
    Dim i As Long
    
    Set ws = ThisWorkbook.Worksheets("Sales")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    total = 0
    For i = 2 To lastRow
        total = total + ws.Cells(i, 3).Value
    Next i
    
    ws.Range("E1").Value = "Total:"
    ws.Range("F1").Value = total
    
    MsgBox "Total calculated: " & total
End Sub
'''
    
    try:
        converter = VBAToPythonConverter()
        result = converter.convert_with_result(sample_vba, "sales_calculator")
        
        print("Conversion successful!")
        print("=" * 50)
        print(result.python_code)
        print("=" * 50)
        print(f"Tokens used: {result.tokens_used}")
        if result.conversion_notes:
            print("\nConversion notes:")
            for note in result.conversion_notes:
                print(f"  - {note}")
    except Exception as e:
        print(f"Error: {e}")
