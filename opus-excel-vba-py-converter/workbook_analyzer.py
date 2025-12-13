"""
Workbook Analyzer Module
Analyzes entire Excel workbooks and generates comprehensive Python recreations
"""
from typing import Dict, List, Optional, Any
from dataclasses import dataclass
import os

from vba_extractor import VBAExtractor
from formula_extractor import FormulaExtractor, FormulaInfo
from data_exporter import DataExporter, ExportResult


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
        # Extract VBA if present
        vba_modules = []
        has_vba = False
        
        if self.extension in ['.xlsm', '.xlsb', '.xlam', '.xls', '.xla']:
            try:
                vba_extractor = VBAExtractor(self.filepath)
                vba_modules = vba_extractor.extract_all()
                has_vba = len(vba_modules) > 0
            except Exception as e:
                print(f"Warning: Could not extract VBA: {e}")
        
        # Extract formulas
        formulas = []
        has_formulas = False
        
        try:
            formula_extractor = FormulaExtractor(self.filepath)
            formulas = formula_extractor.extract_all_formulas()
            has_formulas = len(formulas) > 0
        except Exception as e:
            print(f"Warning: Could not extract formulas: {e}")
        
        # Export data
        data_export = None
        try:
            data_exporter = DataExporter(self.filepath)
            data_export = data_exporter.export_all_sheets()
        except Exception as e:
            print(f"Warning: Could not export data: {e}")
        
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
        dependencies = {}
        
        # Analyze formula dependencies
        for formula_info in formulas:
            sheet = formula_info.sheet_name
            if sheet not in dependencies:
                dependencies[sheet] = []
            
            # Extract sheet references from dependencies
            for dep in formula_info.dependencies:
                if '!' in dep:  # Cross-sheet reference
                    ref_sheet = dep.split('!')[0].strip("'\"")
                    if ref_sheet not in dependencies[sheet]:
                        dependencies[sheet].append(ref_sheet)
        
        # Analyze VBA dependencies (sheets referenced in VBA code)
        for module in vba_modules:
            module_name = module.get('name', 'Unknown')
            code = module.get('code', '')
            
            # Simple pattern matching for sheet references
            # This is a basic approach - could be enhanced
            import re
            sheet_pattern = r'Worksheets?\s*\(\s*["\']([^"\']+)["\']\s*\)'
            sheet_refs = re.findall(sheet_pattern, code, re.IGNORECASE)
            
            if sheet_refs:
                if 'VBA' not in dependencies:
                    dependencies['VBA'] = []
                dependencies['VBA'].extend(sheet_refs)
                dependencies['VBA'] = list(set(dependencies['VBA']))  # Remove duplicates
        
        return dependencies
    
    def _generate_complete_python_script(self,
                                        vba_modules: List[Dict],
                                        formulas: List[FormulaInfo],
                                        data_export: Optional[ExportResult],
                                        dependencies: Dict[str, List[str]]) -> str:
        """
        Generate a complete Python script that recreates workbook logic.
        
        Args:
            vba_modules: List of VBA modules
            formulas: List of formulas
            data_export: Exported data
            dependencies: Dependency mapping
            
        Returns:
            Complete Python script as string
        """
        lines = []
        
        # Header
        lines.extend([
            '"""',
            f'Complete Python Recreation of {self.filename}',
            '',
            'This script recreates the logic from the Excel workbook including:',
            f'- VBA Macros: {len(vba_modules)} modules' if vba_modules else '- No VBA macros',
            f'- Formulas: {len(formulas)} formulas across sheets' if formulas else '- No formulas',
            f'- Data: {len(data_export.sheet_data) if data_export else 0} sheets with data',
            '',
            'Generated by Excel VBA to Python Converter',
            '"""',
            '',
            '# Standard library imports',
            'from typing import Dict, List, Any, Optional',
            'from datetime import datetime, date',
            'from pathlib import Path',
            '',
            '# Data processing imports',
            'import pandas as pd',
            'import numpy as np',
            '',
            '# Excel interaction (optional - for reading/writing Excel files)',
            'import openpyxl',
            '',
            ''
        ])
        
        # Add data loading section
        if data_export and data_export.sheet_data:
            lines.extend([
                '# ============================================================================',
                '# DATA LOADING',
                '# ============================================================================',
                '',
                'class WorkbookData:',
                '    """Container for all workbook data."""',
                '    ',
                '    def __init__(self, filepath: str):',
                '        """Load data from Excel file."""',
                '        self.filepath = filepath',
                '        self.sheets: Dict[str, pd.DataFrame] = {}',
                '        self._load_all_sheets()',
                '    ',
                '    def _load_all_sheets(self):',
                '        """Load all sheets from the workbook."""',
            ])
            
            for sheet_data in data_export.sheet_data:
                var_name = self._clean_name(sheet_data.sheet_name)
                lines.extend([
                    f'        # Load {sheet_data.sheet_name}',
                    f'        self.sheets["{sheet_data.sheet_name}"] = pd.read_excel(',
                    f'            self.filepath,',
                    f'            sheet_name="{sheet_data.sheet_name}",',
                    f'            header=0 if {sheet_data.has_header} else None',
                    f'        )',
                ])
            
            lines.extend([
                '    ',
                '    def get_sheet(self, name: str) -> pd.DataFrame:',
                '        """Get a sheet by name."""',
                '        return self.sheets.get(name)',
                '',
                ''
            ])
        
        # Add formula logic section
        if formulas:
            lines.extend([
                '# ============================================================================',
                '# FORMULA LOGIC',
                '# ============================================================================',
                '',
                'class FormulaEngine:',
                '    """Recreates Excel formula logic in Python."""',
                '    ',
                '    def __init__(self, data: WorkbookData):',
                '        """Initialize with workbook data."""',
                '        self.data = data',
                '    ',
            ])
            
            # Group formulas by sheet
            formulas_by_sheet = {}
            for formula in formulas:
                sheet = formula.sheet_name
                if sheet not in formulas_by_sheet:
                    formulas_by_sheet[sheet] = []
                formulas_by_sheet[sheet].append(formula)
            
            # Generate methods for each sheet's formulas
            for sheet_name, sheet_formulas in formulas_by_sheet.items():
                method_name = f"calculate_{self._clean_name(sheet_name)}"
                lines.extend([
                    f'    def {method_name}(self):',
                    f'        """Calculate formulas for {sheet_name}."""',
                    f'        df = self.data.get_sheet("{sheet_name}")',
                    f'        ',
                    f'        # TODO: Implement {len(sheet_formulas)} formulas',
                ])
                
                # Add formula comments
                for i, formula in enumerate(sheet_formulas[:5], 1):  # Show first 5
                    lines.append(f'        # {formula.cell_address}: {formula.formula}')
                
                if len(sheet_formulas) > 5:
                    lines.append(f'        # ... and {len(sheet_formulas) - 5} more formulas')
                
                lines.extend([
                    f'        ',
                    f'        return df',
                    '    '
                ])
            
            lines.append('')
        
        # Add VBA logic section
        if vba_modules:
            lines.extend([
                '# ============================================================================',
                '# VBA LOGIC (Converted to Python)',
                '# ============================================================================',
                '',
                '# NOTE: VBA code requires LLM conversion for accurate translation',
                '# The following are placeholders for VBA modules:',
                '',
            ])
            
            for module in vba_modules:
                module_name = self._clean_name(module.get('name', 'UnknownModule'))
                module_type = module.get('type', 'Unknown')
                
                lines.extend([
                    f'class {module_name}:',
                    f'    """',
                    f'    Converted from VBA: {module.get("name")}',
                    f'    Type: {module_type}',
                    f'    ',
                    f'    Original VBA code should be converted using LLM converter.',
                    f'    """',
                    f'    pass',
                    '',
                    ''
                ])
        
        # Add main execution section
        lines.extend([
            '# ============================================================================',
            '# MAIN EXECUTION',
            '# ============================================================================',
            '',
            'def main():',
            '    """Main execution function."""',
            '    ',
            '    # Load workbook data',
            '    print("Loading workbook data...")',
            '    workbook = WorkbookData("path/to/your/file.xlsx")',
            '    ',
        ])
        
        if formulas:
            lines.extend([
                '    # Initialize formula engine',
                '    print("Calculating formulas...")',
                '    engine = FormulaEngine(workbook)',
                '    ',
            ])
            
            for sheet_name in formulas_by_sheet.keys():
                method_name = f"calculate_{self._clean_name(sheet_name)}"
                lines.append(f'    result_{self._clean_name(sheet_name)} = engine.{method_name}()')
        
        lines.extend([
            '    ',
            '    print("Workbook processing complete!")',
            '    ',
            '    # Display summary',
            '    print("\\nData Summary:")',
            '    for sheet_name, df in workbook.sheets.items():',
            '        print(f"  {sheet_name}: {df.shape[0]} rows × {df.shape[1]} columns")',
            '',
            '',
            'if __name__ == "__main__":',
            '    main()',
        ])
        
        return '\n'.join(lines)
    
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
        """
        Generate a text report of the workbook analysis.
        
        Args:
            analysis: WorkbookAnalysis object
            
        Returns:
            Report text
        """
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
        
        # VBA section
        if analysis.has_vba:
            lines.extend([
                'VBA MODULES',
                '-' * 80,
                f'Total modules: {len(analysis.vba_modules)}',
                '',
            ])
            
            for module in analysis.vba_modules:
                lines.append(f'  - {module.get("name")} ({module.get("type")})')
            lines.append('')
        
        # Formula section
        if analysis.has_formulas:
            formulas_by_sheet = {}
            for formula in analysis.formulas:
                sheet = formula.sheet_name
                formulas_by_sheet[sheet] = formulas_by_sheet.get(sheet, 0) + 1
            
            lines.extend([
                'FORMULAS',
                '-' * 80,
                f'Total formulas: {len(analysis.formulas)}',
                f'Sheets with formulas: {len(formulas_by_sheet)}',
                '',
            ])
            
            for sheet, count in formulas_by_sheet.items():
                lines.append(f'  - {sheet}: {count} formulas')
            lines.append('')
        
        # Data section
        if analysis.data_export:
            lines.extend([
                'DATA',
                '-' * 80,
                f'Total sheets: {len(analysis.data_export.sheet_data)}',
                '',
            ])
            
            for sheet_data in analysis.data_export.sheet_data:
                lines.extend([
                    f'  Sheet: {sheet_data.sheet_name}',
                    f'    Range: {sheet_data.data_range}',
                    f'    Rows: {len(sheet_data.dataframe)}',
                    f'    Columns: {len(sheet_data.dataframe.columns)}',
                    f'    Has header: {sheet_data.has_header}',
                    '',
                ])
        
        # Dependencies section
        if analysis.dependencies:
            lines.extend([
                'DEPENDENCIES',
                '-' * 80,
            ])
            
            for source, targets in analysis.dependencies.items():
                if targets:
                    lines.append(f'  {source} depends on: {", ".join(targets)}')
            lines.append('')
        
        lines.extend([
            '=' * 80,
            'END OF REPORT',
            '=' * 80,
        ])
        
        return '\n'.join(lines)
