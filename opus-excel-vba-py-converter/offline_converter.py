"""
Offline (rule-based) VBA-to-Python converter.

No API key required.  Handles common VBA patterns via regex and AST-like
transforms.  Not as sophisticated as an LLM, but works immediately for
free and covers the majority of actuarial / business-logic VBA code.
"""
from __future__ import annotations

import re
import textwrap
from dataclasses import dataclass, field
from typing import Optional


# ---------------------------------------------------------------------------
# Public result type (mirrors llm_converter.ConversionResult)
# ---------------------------------------------------------------------------

@dataclass
class OfflineConversionResult:
    success: bool
    python_code: str
    conversion_notes: list[str] = field(default_factory=list)
    error: Optional[str] = None
    tokens_used: int = 0          # always 0 – no LLM involved


# ---------------------------------------------------------------------------
# Keyword / pattern tables
# ---------------------------------------------------------------------------

_VBA_TYPE_MAP: dict[str, str] = {
    "integer": "int",
    "long": "int",
    "longlong": "int",
    "byte": "int",
    "single": "float",
    "double": "float",
    "currency": "float",
    "string": "str",
    "boolean": "bool",
    "variant": "Any",
    "object": "Any",
    "date": "datetime.datetime",
    "collection": "list",
    "dictionary": "dict",
}

_VBA_CONST_MAP: dict[str, str] = {
    "vbcrlf": r'"\n"',
    "vblf": r'"\n"',
    "vbcr": r'"\r"',
    "vbtab": r'"\t"',
    "vbnullstring": '""',
    "vbnewline": r'"\n"',
    "true": "True",
    "false": "False",
    "nothing": "None",
    "empty": "None",
    "null": "None",
    "xlup": '"xlUp"',
    "xldown": '"xlDown"',
    "xltoleft": '"xlToLeft"',
    "xltoright": '"xlToRight"',
}

_BUILTIN_FUNC_MAP: dict[str, str] = {
    "msgbox": "print",
    "debug.print": "print",
    "ubound": "len",
    "lbound": "lambda a: 0  # LBound",
    "cstr": "str",
    "cint": "int",
    "clng": "int",
    "cdbl": "float",
    "csng": "float",
    "cbool": "bool",
    "cdate": "pd.to_datetime",
    "trim": "str.strip",
    "ltrim": "str.lstrip",
    "rtrim": "str.rstrip",
    "ucase": "str.upper",
    "lcase": "str.lower",
    "len": "len",
    "mid": "_mid",
    "left": "_left",
    "right": "_right",
    "instr": "_instr",
    "replace": "str.replace",
    "split": "str.split",
    "join": '", ".join',
    "isnumeric": "_is_numeric",
    "isempty": "_is_empty",
    "isnull": "_is_null",
    "isnothing": "_is_none",
    "now": "datetime.datetime.now",
    "date": "datetime.date.today",
    "time": "datetime.datetime.now().time",
    "year": "lambda d: d.year",
    "month": "lambda d: d.month",
    "day": "lambda d: d.day",
    "hour": "lambda d: d.hour",
    "minute": "lambda d: d.minute",
    "second": "lambda d: d.second",
    "datediff": "_datediff",
    "dateadd": "_dateadd",
    "abs": "abs",
    "sqr": "math.sqrt",
    "int": "int",
    "fix": "int",
    "round": "round",
    "rnd": "random.random",
    "sgn": "_sgn",
    "log": "math.log",
    "exp": "math.exp",
    "atn": "math.atan",
    "sin": "math.sin",
    "cos": "math.cos",
    "tan": "math.tan",
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

_HELPER_FUNCTIONS = '''\
# ── VBA helper stubs (adjust as needed) ──────────────────────────────────
def _mid(s: str, start: int, length: int | None = None) -> str:
    """VBA Mid(): 1-based start."""
    if length is None:
        return s[start - 1:]
    return s[start - 1: start - 1 + length]


def _left(s: str, n: int) -> str:
    return s[:n]


def _right(s: str, n: int) -> str:
    return s[-n:] if n else ""


def _instr(start_or_str, string_or_sub, sub=None):
    """VBA InStr (simplified)."""
    if sub is None:
        return (start_or_str.find(string_or_sub) + 1) or 0
    return (string_or_sub.find(sub, start_or_str - 1) + 1) or 0


def _is_numeric(v) -> bool:
    try:
        float(v)
        return True
    except (ValueError, TypeError):
        return False


def _is_empty(v) -> bool:
    return v is None or v == ""


def _is_null(v) -> bool:
    return v is None


def _is_none(v) -> bool:
    return v is None


def _sgn(x: float) -> int:
    return (x > 0) - (x < 0)


def _datediff(interval: str, d1, d2):
    """Simplified VBA DateDiff."""
    from dateutil.relativedelta import relativedelta
    delta = relativedelta(d2, d1)
    match interval.lower():
        case "d": return (d2 - d1).days
        case "m": return delta.years * 12 + delta.months
        case "yyyy": return delta.years
        case _: return (d2 - d1).days


def _dateadd(interval: str, number: int, date):
    """Simplified VBA DateAdd."""
    from dateutil.relativedelta import relativedelta
    match interval.lower():
        case "d": return date + datetime.timedelta(days=number)
        case "m": return date + relativedelta(months=number)
        case "yyyy": return date + relativedelta(years=number)
        case _: return date + datetime.timedelta(days=number)
'''


# ---------------------------------------------------------------------------
# Core converter
# ---------------------------------------------------------------------------

class OfflineConverter:
    """Rule-based VBA → Python converter (no API key needed)."""

    def __init__(self) -> None:
        self._notes: list[str] = []
        self._imports: set[str] = set()
        self._need_helpers: set[str] = set()

    # -- public API ---------------------------------------------------------

    def convert(self, vba_code: str, module_name: str = "converted_module",
                target_library: str = "pandas") -> OfflineConversionResult:
        """Convert a VBA module to Python."""
        self._notes = []
        self._imports = {"from __future__ import annotations"}
        self._need_helpers = set()

        try:
            lines = self._preprocess(vba_code)
            py_lines = self._convert_lines(lines)
            body = "\n".join(py_lines)

            # Determine required imports
            if "datetime" in body or "datetime" in str(self._imports):
                self._imports.add("import datetime")
            if "math." in body:
                self._imports.add("import math")
            if "random." in body:
                self._imports.add("import random")
            if "pd." in body or "pandas" in body:
                self._imports.add("import pandas as pd")
            if "np." in body or "numpy" in body:
                self._imports.add("import numpy as np")
            if "openpyxl" in body:
                self._imports.add("import openpyxl")

            # Assemble
            header = self._build_header(module_name)
            helpers = self._build_helpers()
            code = f"{header}\n\n{body}"
            if helpers:
                code = f"{header}\n\n{helpers}\n\n{body}"

            self._notes.insert(0,
                "Converted with offline (rule-based) engine — review carefully."
            )

            return OfflineConversionResult(
                success=True,
                python_code=code,
                conversion_notes=list(self._notes),
            )
        except Exception as exc:
            return OfflineConversionResult(
                success=False,
                python_code="",
                error=f"Offline conversion error: {exc}",
            )

    def convert_formula(self, formula: str, cell_address: str = "A1",
                        sheet_name: str = "Sheet1") -> OfflineConversionResult:
        """Convert an Excel formula to Python (best-effort)."""
        self._notes = []
        self._imports = {"from __future__ import annotations",
                         "import pandas as pd", "import numpy as np"}
        try:
            py = self._convert_formula_body(formula, cell_address, sheet_name)
            header = "\n".join(sorted(self._imports))
            code = f"{header}\n\n{py}"
            self._notes.insert(0,
                "Formula converted with offline engine — review carefully."
            )
            return OfflineConversionResult(
                success=True, python_code=code,
                conversion_notes=list(self._notes),
            )
        except Exception as exc:
            return OfflineConversionResult(
                success=False, python_code="",
                error=f"Offline formula conversion error: {exc}",
            )

    # -- preprocessing ------------------------------------------------------

    def _preprocess(self, vba: str) -> list[str]:
        """Normalise VBA source into a list of logical lines."""
        # Strip Attribute lines
        lines = [l for l in vba.splitlines()
                 if not l.strip().startswith("Attribute ")]
        # Join line continuations
        joined: list[str] = []
        buf = ""
        for line in lines:
            stripped = line.rstrip()
            if stripped.endswith(" _"):
                buf += stripped[:-2].rstrip() + " "
            else:
                buf += stripped
                joined.append(buf)
                buf = ""
        if buf:
            joined.append(buf)
        return joined

    # -- line-by-line converter ---------------------------------------------

    def _convert_lines(self, lines: list[str]) -> list[str]:
        out: list[str] = []
        indent = 0
        i = 0
        in_enum = False

        while i < len(lines):
            raw = lines[i]
            stripped = raw.strip()
            i += 1

            if not stripped or stripped.startswith("'"):
                # blank / comment
                cmt = stripped.lstrip("'").strip() if stripped.startswith("'") else ""
                out.append(f"{'    ' * indent}# {cmt}" if cmt else "")
                continue

            low = stripped.lower()

            # --- Option Explicit / Option Base ---
            if low.startswith("option "):
                out.append(f"{'    ' * indent}# {stripped}")
                continue

            # --- Enum ---
            m = re.match(r"(?:public\s+|private\s+)?enum\s+(\w+)", stripped, re.I)
            if m:
                out.append(f"{'    ' * indent}class {m.group(1)}:")
                indent += 1
                in_enum = True
                continue
            if low == "end enum":
                indent = max(indent - 1, 0)
                in_enum = False
                continue
            if in_enum:
                # Enum member assignment
                em = re.match(r"(\w+)\s*=\s*(.+)", stripped)
                if em:
                    out.append(f"{'    ' * indent}{em.group(1)} = {self._convert_expr(em.group(2))}")
                else:
                    out.append(f"{'    ' * indent}{stripped}")
                continue

            # --- Type / End Type  (VBA UDT → dataclass) ---
            m = re.match(r"(?:public\s+|private\s+)?type\s+(\w+)", stripped, re.I)
            if m:
                self._imports.add("from dataclasses import dataclass")
                out.append(f"\n{'    ' * indent}@dataclass")
                out.append(f"{'    ' * indent}class {m.group(1)}:")
                indent += 1
                continue
            if low == "end type":
                indent = max(indent - 1, 0)
                continue

            # --- Const ---
            m = re.match(
                r"(?:public\s+|private\s+|global\s+)?const\s+(\w+)"
                r"(?:\s+as\s+\w+)?\s*=\s*(.+)",
                stripped, re.I,
            )
            if m:
                name, val = m.group(1), self._convert_expr(m.group(2))
                out.append(f"{'    ' * indent}{name} = {val}")
                continue

            # --- Dim / Private / Public variable declarations ---
            m = re.match(
                r"(?:dim|private|public|global|static)\s+(.+)",
                stripped, re.I,
            )
            if m and not re.match(r"(?:dim|private|public|global|static)\s+(?:sub|function|property)", stripped, re.I):
                out.extend(self._convert_dim(m.group(1), indent))
                continue

            # --- Sub / Function / Property ---
            m = re.match(
                r"(?:public\s+|private\s+|friend\s+)?"
                r"(?:static\s+)?"
                r"(sub|function|property\s+(?:get|let|set))\s+"
                r"(\w+)\s*\(([^)]*)\)"
                r"(?:\s+as\s+(\w+))?",
                stripped, re.I,
            )
            if m:
                kind, name, params, ret_type = (
                    m.group(1).lower(), m.group(2),
                    m.group(3), m.group(4),
                )
                py_params = self._convert_params(params)
                ret = ""
                if ret_type:
                    ret = f" -> {self._map_type(ret_type)}"
                prefix = ""
                if "property" in kind:
                    if "get" in kind:
                        prefix = "@property\n" + "    " * indent
                        ret = ret or " -> Any"
                    elif "let" in kind or "set" in kind:
                        prefix = f"@{name}.setter\n" + "    " * indent
                out.append(f"{'    ' * indent}{prefix}def {self._to_snake(name)}({py_params}){ret}:")
                out.append(f"{'    ' * (indent + 1)}\"\"\"Converted from VBA {kind} {name}.\"\"\"")
                indent += 1
                continue

            # --- End Sub / Function / Property ---
            if re.match(r"end\s+(sub|function|property)", low):
                indent = max(indent - 1, 0)
                out.append("")
                continue

            # --- If / ElseIf / Else / End If ---
            # Single-line If: If cond Then statement
            m = re.match(
                r"if\s+(.+?)\s+then\s+(.+?)(?:\s+else\s+(.+))?$",
                stripped, re.I,
            )
            if m and not m.group(2).strip().lower().startswith("'"):
                cond = self._convert_expr(m.group(1))
                then_part = self._convert_statement(m.group(2))
                out.append(f"{'    ' * indent}if {cond}:")
                out.append(f"{'    ' * (indent + 1)}{then_part}")
                if m.group(3):
                    else_part = self._convert_statement(m.group(3))
                    out.append(f"{'    ' * indent}else:")
                    out.append(f"{'    ' * (indent + 1)}{else_part}")
                continue

            # Multi-line If
            m = re.match(r"if\s+(.+?)\s+then\s*$", stripped, re.I)
            if m:
                cond = self._convert_expr(m.group(1))
                out.append(f"{'    ' * indent}if {cond}:")
                indent += 1
                continue
            m = re.match(r"elseif\s+(.+?)\s+then", stripped, re.I)
            if m:
                indent = max(indent - 1, 0)
                cond = self._convert_expr(m.group(1))
                out.append(f"{'    ' * indent}elif {cond}:")
                indent += 1
                continue
            if low == "else":
                indent = max(indent - 1, 0)
                out.append(f"{'    ' * indent}else:")
                indent += 1
                continue
            if low == "end if":
                indent = max(indent - 1, 0)
                continue

            # --- Select Case ---
            m = re.match(r"select\s+case\s+(.+)", stripped, re.I)
            if m:
                expr = self._convert_expr(m.group(1))
                out.append(f"{'    ' * indent}match {expr}:")
                indent += 1
                continue
            m = re.match(r"case\s+else", stripped, re.I)
            if m:
                indent = max(indent - 1, 0)
                out.append(f"{'    ' * indent}case _:")
                indent += 1
                continue
            m = re.match(r"case\s+(.+)", stripped, re.I)
            if m:
                # handle first Case at same indent, subsequent need de-indent
                if out and "case " in out[-1]:
                    indent = max(indent - 1, 0)
                vals = m.group(1).strip()
                out.append(f"{'    ' * indent}case {self._convert_expr(vals)}:")
                indent += 1
                continue
            if low == "end select":
                indent = max(indent - 1, 0)
                continue

            # --- For / Next ---
            m = re.match(
                r"for\s+(\w+)\s*=\s*(.+?)\s+to\s+(.+?)(?:\s+step\s+(.+?))?\s*$",
                stripped, re.I,
            )
            if m:
                var = self._to_snake(m.group(1))
                start = self._convert_expr(m.group(2))
                end = self._convert_expr(m.group(3))
                step = self._convert_expr(m.group(4)) if m.group(4) else None
                if step:
                    out.append(f"{'    ' * indent}for {var} in range({start}, {end} + 1, {step}):")
                else:
                    out.append(f"{'    ' * indent}for {var} in range({start}, {end} + 1):")
                indent += 1
                continue

            # For Each
            m = re.match(r"for\s+each\s+(\w+)\s+in\s+(.+)", stripped, re.I)
            if m:
                var = self._to_snake(m.group(1))
                collection = self._convert_expr(m.group(2))
                out.append(f"{'    ' * indent}for {var} in {collection}:")
                indent += 1
                continue

            if re.match(r"next\b", low):
                indent = max(indent - 1, 0)
                continue

            # --- Do While / Loop ---
            m = re.match(r"do\s+while\s+(.+)", stripped, re.I)
            if m:
                cond = self._convert_expr(m.group(1))
                out.append(f"{'    ' * indent}while {cond}:")
                indent += 1
                continue
            m = re.match(r"do\s+until\s+(.+)", stripped, re.I)
            if m:
                cond = self._convert_expr(m.group(1))
                out.append(f"{'    ' * indent}while not ({cond}):")
                indent += 1
                continue
            if low in ("do", "do:"):
                out.append(f"{'    ' * indent}while True:  # Do...Loop")
                indent += 1
                self._notes.append("Converted Do...Loop to while True — add break condition.")
                continue
            m = re.match(r"loop\s+while\s+(.+)", stripped, re.I)
            if m:
                indent = max(indent - 1, 0)
                # Post-condition loop → requires restructuring
                self._notes.append("Do...Loop While converted; verify loop logic.")
                continue
            m = re.match(r"loop\s+until\s+(.+)", stripped, re.I)
            if m:
                indent = max(indent - 1, 0)
                self._notes.append("Do...Loop Until converted; verify loop logic.")
                continue
            if low == "loop":
                indent = max(indent - 1, 0)
                continue

            # --- While / Wend ---
            m = re.match(r"while\s+(.+)", stripped, re.I)
            if m and low != "wend":
                cond = self._convert_expr(m.group(1))
                out.append(f"{'    ' * indent}while {cond}:")
                indent += 1
                continue
            if low == "wend":
                indent = max(indent - 1, 0)
                continue

            # --- With ---
            m = re.match(r"with\s+(.+)", stripped, re.I)
            if m:
                out.append(f"{'    ' * indent}# With {m.group(1).strip()}")
                self._notes.append(f"With block for {m.group(1).strip()} — prefix member accesses manually.")
                continue
            if low == "end with":
                out.append(f"{'    ' * indent}# End With")
                continue

            # --- On Error ---
            if low.startswith("on error resume next"):
                out.append(f"{'    ' * indent}# On Error Resume Next — wrap individual calls in try/except")
                self._notes.append("'On Error Resume Next' has no Python equivalent; add try/except where needed.")
                continue
            if low.startswith("on error goto 0") or low.startswith("on error goto -1"):
                out.append(f"{'    ' * indent}# On Error GoTo 0 — error handling reset")
                continue
            m = re.match(r"on\s+error\s+goto\s+(\w+)", stripped, re.I)
            if m:
                label = m.group(1)
                out.append(f"{'    ' * indent}# On Error GoTo {label}")
                out.append(f"{'    ' * indent}try:")
                indent += 1
                self._notes.append(f"On Error GoTo {label} → try/except; move handler into except block.")
                continue

            # --- GoTo / labels ---
            if re.match(r"goto\s+\w+", low):
                out.append(f"{'    ' * indent}# {stripped}  (GoTo not supported in Python)")
                self._notes.append("GoTo statement requires manual refactoring.")
                continue
            if re.match(r"\w+:", stripped) and not re.match(r"(case|default)\s*:", low):
                out.append(f"{'    ' * indent}# Label: {stripped}")
                continue

            # --- Exit Sub / Function / For / Do ---
            if low.startswith("exit sub") or low.startswith("exit function") or low.startswith("exit property"):
                out.append(f"{'    ' * indent}return")
                continue
            if low.startswith("exit for") or low.startswith("exit do"):
                out.append(f"{'    ' * indent}break")
                continue

            # --- Set / Let assignments ---
            m = re.match(r"(?:set|let)\s+(\w[\w.]*)\s*=\s*(.+)", stripped, re.I)
            if m:
                lhs = self._convert_expr(m.group(1))
                rhs = self._convert_expr(m.group(2))
                out.append(f"{'    ' * indent}{lhs} = {rhs}")
                continue

            # --- ReDim ---
            m = re.match(r"redim\s+(?:preserve\s+)?(\w+)\((.+?)\)", stripped, re.I)
            if m:
                var = self._to_snake(m.group(1))
                size = self._convert_expr(m.group(2))
                if "preserve" in low:
                    out.append(f"{'    ' * indent}{var}.extend([None] * ({size} + 1 - len({var})))")
                else:
                    out.append(f"{'    ' * indent}{var} = [None] * ({size} + 1)")
                continue

            # --- Erase ---
            m = re.match(r"erase\s+(\w+)", stripped, re.I)
            if m:
                out.append(f"{'    ' * indent}{self._to_snake(m.group(1))} = []")
                continue

            # --- Call statement ---
            m = re.match(r"call\s+(\w+)\s*\(?(.*?)\)?\s*$", stripped, re.I)
            if m:
                func = self._to_snake(m.group(1))
                args = self._convert_expr(m.group(2)) if m.group(2) else ""
                out.append(f"{'    ' * indent}{func}({args})")
                continue

            # --- Generic assignment / statement ---
            out.append(f"{'    ' * indent}{self._convert_statement(stripped)}")

        # Ensure no empty function bodies
        final: list[str] = []
        for j, line in enumerate(out):
            final.append(line)
            if line.rstrip().endswith(":") and (j + 1 >= len(out) or not out[j + 1].strip()):
                final.append("    " * (line.index(line.lstrip()[0]) // 4 + 1) + "pass")

        return final

    # -- expression converter -----------------------------------------------

    def _convert_expr(self, expr: str) -> str:
        """Convert a VBA expression to Python."""
        if not expr:
            return expr
        s = expr.strip()

        # String concatenation: & → +
        s = re.sub(r'\s*&\s*', ' + ', s)

        # VBA constants
        for vba_c, py_c in _VBA_CONST_MAP.items():
            s = re.sub(rf'\b{re.escape(vba_c)}\b', py_c, s, flags=re.I)

        # Not / And / Or / Mod / Is
        s = re.sub(r'\bNot\b', 'not', s, flags=re.I)
        s = re.sub(r'\bAnd\b', 'and', s, flags=re.I)
        s = re.sub(r'\bOr\b', 'or', s, flags=re.I)
        s = re.sub(r'\bMod\b', '%', s, flags=re.I)
        s = re.sub(r'\bIs\s+Nothing\b', 'is None', s, flags=re.I)
        s = re.sub(r'\b<>\b', '!=', s)

        # VBA built-in function names
        for vba_fn, py_fn in _BUILTIN_FUNC_MAP.items():
            pattern = rf'\b{re.escape(vba_fn)}\s*\('
            if re.search(pattern, s, re.I):
                s = re.sub(pattern, f'{py_fn}(', s, flags=re.I)
                if py_fn.startswith("_"):
                    self._need_helpers.add(py_fn)

        return s

    # -- statement converter ------------------------------------------------

    def _convert_statement(self, stmt: str) -> str:
        """Convert a single VBA statement."""
        s = stmt.strip()

        # Assignment: var = expr
        m = re.match(r"(\w[\w.]*(?:\([^)]*\))?)\s*=\s*(.+)", s)
        if m:
            lhs = self._convert_expr(m.group(1))
            rhs = self._convert_expr(m.group(2))
            return f"{lhs} = {rhs}"

        # Bare function/sub call: FuncName arg1, arg2 → func_name(arg1, arg2)
        m = re.match(r"(\w+)\s+(.+)", s)
        if m and m.group(1).lower() not in (
            "if", "for", "do", "while", "select", "case", "dim",
            "public", "private", "sub", "function", "end", "exit",
            "set", "let", "with", "on", "goto", "redim", "erase",
            "const", "type", "enum", "option", "next", "loop", "wend",
            "elseif", "else", "call", "property", "class", "attribute",
            "static", "global", "friend",
        ):
            func = self._to_snake(m.group(1))
            args = self._convert_expr(m.group(2))
            return f"{func}({args})"

        return self._convert_expr(s)

    # -- type / declaration converters --------------------------------------

    def _map_type(self, vba_type: str) -> str:
        return _VBA_TYPE_MAP.get(vba_type.lower(), vba_type)

    def _convert_params(self, params_str: str) -> str:
        """Convert VBA parameter list to Python."""
        if not params_str.strip():
            return ""
        parts: list[str] = []
        for p in params_str.split(","):
            p = p.strip()
            if not p:
                continue
            # Remove ByVal / ByRef
            p = re.sub(r'\b(ByVal|ByRef)\b\s*', '', p, flags=re.I).strip()
            # Optional keyword
            is_optional = False
            m_opt = re.match(r'Optional\s+(.+)', p, re.I)
            if m_opt:
                is_optional = True
                p = m_opt.group(1).strip()
            # ParamArray
            m_pa = re.match(r'ParamArray\s+(\w+)\s*\(\s*\)', p, re.I)
            if m_pa:
                parts.append(f"*{self._to_snake(m_pa.group(1))}")
                continue
            # name As Type = default
            m_full = re.match(r'(\w+)(?:\s*\(\s*\))?\s+As\s+(\w+)(?:\s*=\s*(.+))?', p, re.I)
            if m_full:
                name = self._to_snake(m_full.group(1))
                typ = self._map_type(m_full.group(2))
                default = m_full.group(3)
                if default:
                    parts.append(f"{name}: {typ} = {self._convert_expr(default)}")
                elif is_optional:
                    parts.append(f"{name}: {typ} | None = None")
                else:
                    parts.append(f"{name}: {typ}")
            else:
                # Just a name
                name = self._to_snake(p.split()[0])
                if is_optional:
                    parts.append(f"{name}=None")
                else:
                    parts.append(name)
        return ", ".join(parts)

    def _convert_dim(self, decl: str, indent: int) -> list[str]:
        """Convert Dim/Private/Public variable declarations."""
        out: list[str] = []
        for part in decl.split(","):
            part = part.strip()
            if not part:
                continue
            # Array: name(size) As Type
            m = re.match(r'(\w+)\s*\(\s*(.*?)\s*\)\s*(?:As\s+(\w+))?', part, re.I)
            if m:
                name = self._to_snake(m.group(1))
                size = m.group(2)
                if size:
                    out.append(f"{'    ' * indent}{name}: list = [None] * ({self._convert_expr(size)} + 1)")
                else:
                    out.append(f"{'    ' * indent}{name}: list = []")
                continue
            # name As New ClassName
            m = re.match(r'(\w+)\s+As\s+New\s+(\w+)', part, re.I)
            if m:
                name = self._to_snake(m.group(1))
                cls = m.group(2)
                out.append(f"{'    ' * indent}{name} = {cls}()")
                continue
            # name As Type
            m = re.match(r'(\w+)\s+As\s+(\w+)', part, re.I)
            if m:
                name = self._to_snake(m.group(1))
                typ = self._map_type(m.group(2))
                default = {"int": "0", "float": "0.0", "str": '""', "bool": "False"}.get(typ, "None")
                out.append(f"{'    ' * indent}{name}: {typ} = {default}")
                continue
            # bare name
            name = self._to_snake(part.strip())
            if name:
                out.append(f"{'    ' * indent}{name} = None")
        return out

    # -- naming helpers -----------------------------------------------------

    @staticmethod
    def _to_snake(name: str) -> str:
        """Convert PascalCase/camelCase to snake_case (preserving existing underscores)."""
        if not name:
            return name
        # Don't convert if already has underscores or is all-caps
        if "_" in name or name.isupper():
            return name
        s = re.sub(r'([A-Z]+)([A-Z][a-z])', r'\1_\2', name)
        s = re.sub(r'([a-z0-9])([A-Z])', r'\1_\2', s)
        return s.lower()

    # -- header / helpers builders ------------------------------------------

    def _build_header(self, module_name: str) -> str:
        imports = sorted(self._imports)
        lines = [f'"""{module_name} — converted from VBA (offline engine)."""']
        lines.extend(imports)
        return "\n".join(lines)

    def _build_helpers(self) -> str:
        if not self._need_helpers:
            return ""
        # Include the full helpers block if any helper is referenced
        return _HELPER_FUNCTIONS.strip()

    # -- formula converter --------------------------------------------------

    def _convert_formula_body(self, formula: str, cell: str, sheet: str) -> str:
        """Best-effort Excel formula → Python."""
        f = formula.strip()
        if f.startswith("="):
            f = f[1:]

        lines = [
            f'# Original formula in {sheet}!{cell}: ={f}',
            '# Converted to Python/pandas:',
            '',
        ]

        fl = f.upper()
        # Common patterns
        if fl.startswith("SUM("):
            col = self._guess_col(f)
            lines.append(f"result = df['{col}'].sum()")
        elif fl.startswith("AVERAGE("):
            col = self._guess_col(f)
            lines.append(f"result = df['{col}'].mean()")
        elif fl.startswith("COUNT("):
            col = self._guess_col(f)
            lines.append(f"result = df['{col}'].count()")
        elif fl.startswith("MAX("):
            col = self._guess_col(f)
            lines.append(f"result = df['{col}'].max()")
        elif fl.startswith("MIN("):
            col = self._guess_col(f)
            lines.append(f"result = df['{col}'].min()")
        elif fl.startswith("VLOOKUP("):
            lines.append("# VLOOKUP → pandas merge/loc")
            lines.append("result = lookup_df.set_index('key_col').loc[search_value, 'return_col']")
            self._notes.append("VLOOKUP converted to stub — adjust column names.")
        elif fl.startswith("IF("):
            lines.append("result = np.where(condition, true_value, false_value)")
            self._notes.append("IF formula converted to np.where stub — fill in condition/values.")
        elif fl.startswith("SUMIF"):
            lines.append("result = df.loc[df['criteria_col'] == criteria, 'sum_col'].sum()")
            self._notes.append("SUMIF/SUMIFS converted to stub — adjust column names and criteria.")
        elif fl.startswith("COUNTIF"):
            lines.append("result = (df['criteria_col'] == criteria).sum()")
        else:
            # Generic passthrough
            lines.append("# TODO: Manually convert this formula")
            lines.append(f"# Formula: ={f}")
            lines.append("result = None  # Placeholder")
            self._notes.append(f"Formula ={f} not auto-convertible — manual conversion needed.")

        return "\n".join(lines)

    @staticmethod
    def _guess_col(formula: str) -> str:
        """Try to guess a column letter from a formula like SUM(A1:A100)."""
        m = re.search(r'\(([A-Z]+)\d+', formula, re.I)
        return m.group(1) if m else "A"
