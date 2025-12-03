"""
LLM-Powered VBA to Python Converter

This module handles the conversion of VBA code to Python using
Claude (Anthropic) or OpenAI APIs.
"""
import os
import re
from abc import ABC, abstractmethod
from dataclasses import dataclass, field
from typing import Optional
from dotenv import load_dotenv

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
    
    SYSTEM_PROMPT = """You are an expert VBA to Python converter. Your task is to convert VBA/VBScript code to clean, idiomatic Python code.

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

    def __init__(self, model: Optional[str] = None):
        """Initialize the converter with optional model override."""
        self.model = model
        self._conversion_notes: list[str] = []
    
    @abstractmethod
    def convert(self, vba_code: str, module_name: str = "converted_module",
                target_library: str = "pandas") -> ConversionResult:
        """Convert VBA code to Python."""
        pass
    
    def _build_user_prompt(self, vba_code: str, module_name: str, 
                           target_library: str) -> str:
        """Build the user prompt for conversion."""
        return f"""Convert the following VBA code to Python. 

**Module Name:** {module_name}
**Target Library:** {target_library} (use this for data operations)

**VBA Code:**
```vba
{vba_code}
```

Provide complete, runnable Python code with all necessary imports."""

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
            
            response = self.client.chat.completions.create(
                model=self.model,
                max_tokens=4096,
                messages=[
                    {"role": "system", "content": self.SYSTEM_PROMPT},
                    {"role": "user", "content": user_prompt}
                ]
            )
            
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
            return ConversionResult(
                success=False,
                python_code="",
                error=str(e)
            )


class VBAToPythonConverter:
    """
    Main converter class that automatically selects the appropriate LLM backend.
    """
    
    def __init__(self, provider: Optional[str] = None):
        """
        Initialize the converter.
        
        Args:
            provider: 'anthropic', 'openai', or None for auto-detection
        """
        self.provider = provider or os.getenv("LLM_PROVIDER", "anthropic")
        self._converter: Optional[BaseLLMConverter] = None
        self._last_notes: list[str] = []
        
    def _get_converter(self) -> BaseLLMConverter:
        """Get or create the appropriate converter."""
        if self._converter is None:
            if self.provider == "anthropic":
                self._converter = AnthropicConverter()
            elif self.provider == "openai":
                self._converter = OpenAIConverter()
            else:
                # Try Anthropic first, then OpenAI
                try:
                    self._converter = AnthropicConverter()
                except ValueError:
                    try:
                        self._converter = OpenAIConverter()
                    except ValueError:
                        raise ValueError(
                            "No LLM API key found. Please set ANTHROPIC_API_KEY or "
                            "OPENAI_API_KEY environment variable."
                        )
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
            Exception: If conversion fails
        """
        converter = self._get_converter()
        result = converter.convert(vba_code, module_name, target_library)
        
        self._last_notes = result.conversion_notes
        
        if not result.success:
            raise Exception(f"Conversion failed: {result.error}")
        
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
