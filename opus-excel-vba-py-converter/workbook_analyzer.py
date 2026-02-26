"""
Workbook Analyzer Module
Analyzes entire Excel workbooks and generates comprehensive Python recreations
"""
import logging
import os
import re
from dataclasses import dataclass
from typing import Dict, List, Optional, Any

logger = logging.getLogger(__name__)

from vba_extractor import VBAExtractor
from formula_extractor import FormulaExtractor, FormulaInfo
from data_exporter import DataExporter, ExportResult

# Reusable separator for generated Python scripts
_SECTION_SEP = '# ' + '=' * 76


@dataclass
class WorkbookAnalysis:
    """Complete analysis of an Excel workbook."""
    filename: str
    has_vba: bool
    vba_modules: List[Dict]
    has_formulas: bool
    formulas: List[FormulaInfo]
    data_export: ExportResult
    dependencies: Dict[str, List[str]]  # Sheet to list of dependencies
    python_script: str  # Complete Python script
    


class WorkbookAnalyzer:
    """Analyze Excel workbooks and generate comprehensive Python recreations."""
    
    def __init__(self, filepath: str):
        """
        Initialize the workbook analyzer.
        
        Args:
            filepath: Path to the Excel file
        """
        self.filepath = filepath
        self.filename = os.path.basename(filepath)
        self.extension = os.path.splitext(filepath)[1].lower()
        
    def analyze_complete(self) -> WorkbookAnalysis:
        """
        Perform complete analysis of the workbook.
        
        Returns:
            WorkbookAnalysis object with all information
        """
        # Extract VBA if present (try for all files — oletools auto-detects format)
        vba_modules = []
        has_vba = False
        
        try:
            vba_extractor = VBAExtractor(self.filepath)
            vba_modules = vba_extractor.extract_all()
            has_vba = len(vba_modules) > 0
        except Exception as e:
            logger.warning("Could not extract VBA: %s", e)
        
        # Extract formulas
        formulas = []
        has_formulas = False
        
        try:
            formula_extractor = FormulaExtractor(self.filepath)
            formulas = formula_extractor.extract_all_formulas()
            has_formulas = len(formulas) > 0
        except Exception as e:
            logger.warning("Could not extract formulas: %s", e)
        
        # Export data
        data_export = None
        try:
            data_exporter = DataExporter(self.filepath)
            data_export = data_exporter.export_all_sheets()
        except Exception as e:
            logger.warning("Could not export data: %s", e)
        
        # Analyze dependencies
        dependencies = self._analyze_dependencies(formulas, vba_modules)
        
        # Generate comprehensive Python script
        python_script = self._generate_complete_python_script(
            vba_modules=vba_modules,
            formulas=formulas,
            data_export=data_export,
            dependencies=dependencies
        )
        
        return WorkbookAnalysis(
            filename=self.filename,
            has_vba=has_vba,
            vba_modules=vba_modules,
            has_formulas=has_formulas,
            formulas=formulas,
            data_export=data_export,
            dependencies=dependencies,
            python_script=python_script
        )
    
    def _analyze_dependencies(self, 
                             formulas: List[FormulaInfo], 
                             vba_modules: List[Dict]) -> Dict[str, List[str]]:
        """
        Analyze dependencies between sheets, formulas, and VBA.
        
        Args:
            formulas: List of FormulaInfo objects
            vba_modules: List of VBA module dictionaries
            
        Returns:
            Dictionary mapping sheets to their dependencies
        """
        dependencies: Dict[str, List[str]] = {}
        self._collect_formula_deps(formulas, dependencies)
        self._collect_vba_deps(vba_modules, dependencies)
        return dependencies

    @staticmethod
    def _collect_formula_deps(formulas: List[FormulaInfo],
                              dependencies: Dict[str, List[str]]) -> None:
        """Populate *dependencies* with cross-sheet formula references."""
        for formula_info in formulas:
            sheet = formula_info.sheet_name
            if sheet not in dependencies:
                dependencies[sheet] = []
            for dep in formula_info.dependencies:
                if '!' in dep:
                    ref_sheet = dep.split('!')[0].strip("'\"")
                    if ref_sheet not in dependencies[sheet]:
                        dependencies[sheet].append(ref_sheet)

    @staticmethod
    def _collect_vba_deps(vba_modules: List[Dict],
                          dependencies: Dict[str, List[str]]) -> None:
        """Populate *dependencies* with sheet references found in VBA code."""
        for module in vba_modules:
            code = module.get('code', '')
            sheet_pattern = r'Worksheets?\s*\(\s*["\']([^"\']+)["\']\s*\)'
            sheet_refs = re.findall(sheet_pattern, code, re.IGNORECASE)
            if sheet_refs:
                if 'VBA' not in dependencies:
                    dependencies['VBA'] = []
                dependencies['VBA'].extend(sheet_refs)
                dependencies['VBA'] = list(set(dependencies['VBA']))
    
    def _generate_complete_python_script(self,
                                        vba_modules: List[Dict],
                                        formulas: List[FormulaInfo],
                                        data_export: Optional[ExportResult],
                                        dependencies: Dict[str, List[str]]) -> str:
        """Generate a complete Python script that recreates workbook logic."""
        lines: list[str] = []
        lines.extend(self._script_header(vba_modules, formulas, data_export, dependencies))
        lines.extend(self._script_data_section(data_export))

        formulas_by_sheet: Dict[str, List[FormulaInfo]] = {}
        for f in formulas:
            formulas_by_sheet.setdefault(f.sheet_name, []).append(f)

        lines.extend(self._script_formula_section(formulas, formulas_by_sheet))
        lines.extend(self._script_vba_section(vba_modules))
        lines.extend(self._script_main_section(formulas, formulas_by_sheet))
        return '\n'.join(lines)

    # -- script section builders ------------------------------------------------

    def _script_header(self, vba_modules: List[Dict], formulas: List[FormulaInfo],
                       data_export: Optional[ExportResult],
                       dependencies: Dict[str, List[str]]) -> List[str]:
        """Return the docstring + imports block for the generated script."""
        dep_lines = [f'  {s} -> {", ".join(d)}' for s, d in dependencies.items() if d]
        lines = [
            '"""',
            f'Complete Python Recreation of {self.filename}',
            '',
            'This script recreates the logic from the Excel workbook including:',
            f'- VBA Macros: {len(vba_modules)} modules' if vba_modules else '- No VBA macros',
            f'- Formulas: {len(formulas)} formulas across sheets' if formulas else '- No formulas',
            f'- Data: {len(data_export.sheet_data) if data_export else 0} sheets with data',
        ]
        if dep_lines:
            lines.extend(['', 'Sheet Dependencies:', *dep_lines])
        lines.extend([
            '', 'Generated by Excel VBA to Python Converter', '"""', '',
            '# Standard library imports',
            'from typing import Dict, List, Any, Optional',
            'from datetime import datetime, date',
            'from pathlib import Path', '',
            '# Data processing imports',
            'import pandas as pd',
            'import numpy as np', '',
            '# Excel interaction (optional - for reading/writing Excel files)',
            'import openpyxl', '', '',
        ])
        return lines

    def _script_data_section(self, data_export: Optional[ExportResult]) -> List[str]:
        """Return the DATA LOADING class block (or empty list)."""
        if not data_export or not data_export.sheet_data:
            return []
        lines = [
            _SECTION_SEP, '# DATA LOADING', _SECTION_SEP, '',
            'class WorkbookData:',
            '    """Container for all workbook data."""', '    ',
            '    def __init__(self, filepath: str):',
            '        """Load data from Excel file."""',
            '        self.filepath = filepath',
            '        self.sheets: Dict[str, pd.DataFrame] = {}',
            '        self._load_all_sheets()', '    ',
            '    def _load_all_sheets(self):',
            '        """Load all sheets from the workbook."""',
        ]
        for sd in data_export.sheet_data:
            lines.extend([
                f'        # Load {sd.sheet_name}',
                f'        self.sheets["{sd.sheet_name}"] = pd.read_excel(',
                '            self.filepath,',
                f'            sheet_name="{sd.sheet_name}",',
                f'            header=0 if {sd.has_header} else None',
                '        )',
            ])
        lines.extend([
            '    ', '    def get_sheet(self, name: str) -> pd.DataFrame:',
            '        """Get a sheet by name."""',
            '        return self.sheets.get(name)', '', '',
        ])
        return lines

    def _script_formula_section(self, formulas: List[FormulaInfo],
                                formulas_by_sheet: Dict[str, List[FormulaInfo]]) -> List[str]:
        """Return the FORMULA LOGIC class block (or empty list)."""
        if not formulas:
            return []
        lines = [
            _SECTION_SEP, '# FORMULA LOGIC', _SECTION_SEP, '',
            'class FormulaEngine:',
            '    """Recreates Excel formula logic in Python."""', '    ',
            '    def __init__(self, data: WorkbookData):',
            '        """Initialize with workbook data."""',
            '        self.data = data', '    ',
        ]
        for sheet_name, sheet_formulas in formulas_by_sheet.items():
            method_name = f"calculate_{self._clean_name(sheet_name)}"
            lines.extend([
                f'    def {method_name}(self):',
                f'        """Calculate formulas for {sheet_name}."""',
                f'        df = self.data.get_sheet("{sheet_name}")', '        ',
                f'        # TODO: Implement {len(sheet_formulas)} formulas',
            ])
            for formula in sheet_formulas[:5]:
                lines.append(f'        # {formula.cell_address}: {formula.formula}')
            if len(sheet_formulas) > 5:
                lines.append(f'        # ... and {len(sheet_formulas) - 5} more formulas')
            lines.extend(['        ', '        return df', '    '])
        lines.append('')
        return lines

    def _script_vba_section(self, vba_modules: List[Dict]) -> List[str]:
        """Return VBA placeholder classes (or empty list)."""
        if not vba_modules:
            return []
        lines = [
            _SECTION_SEP, '# VBA LOGIC (Converted to Python)', _SECTION_SEP, '',
            '# NOTE: VBA code requires LLM conversion for accurate translation',
            '# The following are placeholders for VBA modules:', '',
        ]
        for module in vba_modules:
            raw_name = self._clean_name(module.get('name', 'UnknownModule'))
            class_name = raw_name.replace('_', ' ').title().replace(' ', '_')
            module_type = module.get('type', 'Unknown')
            lines.extend([
                f'class {class_name}:', '    """',
                f'    Converted from VBA: {module.get("name")}',
                f'    Type: {module_type}', '    ',
                '    Original VBA code should be converted using LLM converter.',
                '    """', '    pass', '', '',
            ])
        return lines

    def _script_main_section(self, formulas: List[FormulaInfo],
                             formulas_by_sheet: Dict[str, List[FormulaInfo]]) -> List[str]:
        """Return the main() entry-point block."""
        lines = [
            _SECTION_SEP, '# MAIN EXECUTION', _SECTION_SEP, '',
            'def main():', '    """Main execution function."""', '    ',
            '    # Load workbook data',
            '    print("Loading workbook data...")',
            '    workbook = WorkbookData("path/to/your/file.xlsx")', '    ',
        ]
        if formulas:
            lines.extend([
                '    # Initialize formula engine',
                '    print("Calculating formulas...")',
                '    engine = FormulaEngine(workbook)', '    ',
            ])
            for sheet_name in formulas_by_sheet:
                method = f"calculate_{self._clean_name(sheet_name)}"
                lines.append(f'    result_{self._clean_name(sheet_name)} = engine.{method}()')
        lines.extend([
            '    ', '    print("Workbook processing complete!")', '    ',
            '    # Display summary', '    print("\\nData Summary:")',
            '    for sheet_name, df in workbook.sheets.items():',
            '        print(f"  {sheet_name}: {df.shape[0]} rows × {df.shape[1]} columns")',
            '', '', 'if __name__ == "__main__":', '    main()',
        ])
        return lines
    
    def _clean_name(self, name: str) -> str:
        """
        Clean a name to be a valid Python identifier.
        
        Args:
            name: Original name
            
        Returns:
            Cleaned name
        """
        if not name:
            return 'unnamed'
        
        # Replace spaces and special characters
        clean = str(name).strip()
        clean = clean.replace(' ', '_').replace('-', '_')
        clean = ''.join(c if c.isalnum() or c == '_' else '_' for c in clean)
        
        # Ensure it starts with a letter or underscore
        if clean and not (clean[0].isalpha() or clean[0] == '_'):
            clean = 'item_' + clean
        
        # Convert to lowercase for function/variable names
        return clean.lower() or 'unnamed'
    
    def generate_analysis_report(self, analysis: WorkbookAnalysis) -> str:
        """Generate a text report of the workbook analysis."""
        lines = [
            '=' * 80,
            f'WORKBOOK ANALYSIS REPORT: {analysis.filename}',
            '=' * 80,
            '',
            'OVERVIEW',
            '-' * 80,
            f'File: {analysis.filename}',
            f'Has VBA: {"Yes" if analysis.has_vba else "No"}',
            f'Has Formulas: {"Yes" if analysis.has_formulas else "No"}',
            f'Has Data: {"Yes" if analysis.data_export else "No"}',
            '',
        ]

        if analysis.has_vba:
            lines.extend(self._report_vba_section(analysis))
        if analysis.has_formulas:
            lines.extend(self._report_formula_section(analysis))
        if analysis.data_export:
            lines.extend(self._report_data_section(analysis))
        if analysis.dependencies:
            lines.extend(self._report_dependency_section(analysis))

        lines.extend(['=' * 80, 'END OF REPORT', '=' * 80])
        return '\n'.join(lines)

    # -- report section builders ------------------------------------------------

    @staticmethod
    def _report_vba_section(analysis: WorkbookAnalysis) -> List[str]:
        """Return the VBA modules section of the report."""
        lines = [
            'VBA MODULES', '-' * 80,
            f'Total modules: {len(analysis.vba_modules)}', '',
        ]
        for module in analysis.vba_modules:
            lines.append(f'  - {module.get("name")} ({module.get("type")})')
        lines.append('')
        return lines

    @staticmethod
    def _report_formula_section(analysis: WorkbookAnalysis) -> List[str]:
        """Return the formulas section of the report."""
        formulas_by_sheet: Dict[str, int] = {}
        for formula in analysis.formulas:
            formulas_by_sheet[formula.sheet_name] = formulas_by_sheet.get(formula.sheet_name, 0) + 1

        lines = [
            'FORMULAS', '-' * 80,
            f'Total formulas: {len(analysis.formulas)}',
            f'Sheets with formulas: {len(formulas_by_sheet)}', '',
        ]
        for sheet, count in formulas_by_sheet.items():
            lines.append(f'  - {sheet}: {count} formulas')
        lines.append('')
        return lines

    @staticmethod
    def _report_data_section(analysis: WorkbookAnalysis) -> List[str]:
        """Return the data section of the report."""
        lines = [
            'DATA', '-' * 80,
            f'Total sheets: {len(analysis.data_export.sheet_data)}', '',
        ]
        for sheet_data in analysis.data_export.sheet_data:
            lines.extend([
                f'  Sheet: {sheet_data.sheet_name}',
                f'    Range: {sheet_data.data_range}',
                f'    Rows: {len(sheet_data.dataframe)}',
                f'    Columns: {len(sheet_data.dataframe.columns)}',
                f'    Has header: {sheet_data.has_header}', '',
            ])
        return lines

    @staticmethod
    def _report_dependency_section(analysis: WorkbookAnalysis) -> List[str]:
        """Return the dependencies section of the report."""
        lines = ['DEPENDENCIES', '-' * 80]
        for source, targets in analysis.dependencies.items():
            if targets:
                lines.append(f'  {source} depends on: {", ".join(targets)}')
        lines.append('')
        return lines
