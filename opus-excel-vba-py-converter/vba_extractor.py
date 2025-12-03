"""
VBA Extractor Module
Extracts VBA code from Excel files (.xlsm, .xls, .xlsb, .xla, .xlam)
"""
import os
import zipfile
import tempfile
import re
from typing import List, Dict, Optional


class VBAExtractor:
    """Extract VBA code from Excel files."""
    
    # Module type identifiers
    MODULE_TYPES = {
        1: 'Standard Module',
        2: 'Class Module', 
        3: 'UserForm',
        100: 'Document Module (ThisWorkbook/Sheet)'
    }
    
    def __init__(self, filepath: str):
        """
        Initialize the VBA extractor.
        
        Args:
            filepath: Path to the Excel file
        """
        self.filepath = filepath
        self.filename = os.path.basename(filepath)
        self.extension = os.path.splitext(filepath)[1].lower()
        
    def extract_all(self) -> List[Dict]:
        """
        Extract all VBA modules from the Excel file.
        
        Returns:
            List of dictionaries containing module information
        """
        if self.extension in ['.xlsm', '.xlsb', '.xlam']:
            return self._extract_from_xlsx_format()
        elif self.extension in ['.xls', '.xla']:
            return self._extract_from_xls_format()
        else:
            raise ValueError(f"Unsupported file format: {self.extension}")
    
    def _extract_from_xlsx_format(self) -> List[Dict]:
        """
        Extract VBA from modern Excel formats (.xlsm, .xlsb, .xlam).
        These are ZIP-based formats with vbaProject.bin inside.
        """
        modules = []
        
        try:
            with zipfile.ZipFile(self.filepath, 'r') as zf:
                # Look for vbaProject.bin
                vba_files = [f for f in zf.namelist() if 'vbaProject.bin' in f]
                
                if not vba_files:
                    return []
                
                for vba_file in vba_files:
                    vba_content = zf.read(vba_file)
                    extracted = self._parse_vba_project(vba_content)
                    modules.extend(extracted)
                    
        except zipfile.BadZipFile:
            raise ValueError("Invalid or corrupted Excel file")
        except Exception as e:
            raise ValueError(f"Error reading Excel file: {str(e)}")
        
        return modules
    
    def _extract_from_xls_format(self) -> List[Dict]:
        """
        Extract VBA from legacy Excel format (.xls, .xla).
        Uses oletools if available, otherwise attempts manual extraction.
        """
        try:
            # Try using oletools (olevba) if available
            from oletools.olevba import VBA_Parser
            
            vba_parser = VBA_Parser(self.filepath)
            modules = []
            
            if vba_parser.detect_vba_macros():
                for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                    if vba_code and vba_code.strip():
                        module_type = self._detect_module_type(vba_filename, vba_code)
                        modules.append({
                            'name': vba_filename or 'Unknown',
                            'type': module_type,
                            'code': vba_code,
                            'stream_path': stream_path
                        })
            
            vba_parser.close()
            return modules
            
        except ImportError:
            # Fallback to manual extraction for .xls files
            return self._manual_ole_extraction()
    
    def _parse_vba_project(self, vba_content: bytes) -> List[Dict]:
        """
        Parse the vbaProject.bin file to extract VBA code.
        
        Args:
            vba_content: Binary content of vbaProject.bin
            
        Returns:
            List of module dictionaries
        """
        modules = []
        
        try:
            # Try using oletools if available
            from oletools.olevba import VBA_Parser
            
            # Create a temporary file
            with tempfile.NamedTemporaryFile(suffix='.bin', delete=False) as tmp:
                tmp.write(vba_content)
                tmp_path = tmp.name
            
            try:
                vba_parser = VBA_Parser(tmp_path)
                
                if vba_parser.detect_vba_macros():
                    for (filename, stream_path, vba_filename, vba_code) in vba_parser.extract_macros():
                        if vba_code and vba_code.strip():
                            module_type = self._detect_module_type(vba_filename, vba_code)
                            modules.append({
                                'name': vba_filename or 'Unknown',
                                'type': module_type,
                                'code': vba_code,
                                'stream_path': stream_path
                            })
                
                vba_parser.close()
            finally:
                os.unlink(tmp_path)
                
        except ImportError:
            # Fallback to manual extraction
            modules = self._manual_vba_extraction(vba_content)
        
        return modules
    
    def _manual_vba_extraction(self, vba_content: bytes) -> List[Dict]:
        """
        Manual extraction of VBA code from binary content.
        This is a fallback when oletools is not available.
        
        Args:
            vba_content: Binary content containing VBA
            
        Returns:
            List of module dictionaries
        """
        modules = []
        
        # Try to find VBA code patterns in the binary
        # VBA code is typically stored as compressed or plain text
        
        # Look for "Attribute VB_Name" which indicates module start
        content_str = vba_content.decode('latin-1', errors='ignore')
        
        # Pattern to find module declarations
        module_pattern = r'Attribute\s+VB_Name\s*=\s*"([^"]+)"'
        matches = list(re.finditer(module_pattern, content_str))
        
        for i, match in enumerate(matches):
            module_name = match.group(1)
            start_pos = match.start()
            
            # Find the end of this module (start of next or end of content)
            if i + 1 < len(matches):
                end_pos = matches[i + 1].start()
            else:
                end_pos = len(content_str)
            
            # Extract the code
            code_section = content_str[start_pos:end_pos]
            
            # Clean up the code
            cleaned_code = self._clean_extracted_code(code_section)
            
            if cleaned_code.strip():
                modules.append({
                    'name': module_name,
                    'type': self._detect_module_type(module_name, cleaned_code),
                    'code': cleaned_code,
                    'stream_path': 'manual_extraction'
                })
        
        # If no modules found with pattern, try to extract any VBA-like code
        if not modules:
            vba_keywords = ['Sub ', 'Function ', 'Private Sub', 'Public Sub', 
                          'Private Function', 'Public Function', 'Dim ', 'End Sub', 'End Function']
            
            for keyword in vba_keywords:
                if keyword in content_str:
                    # Found some VBA code, extract it
                    code_start = content_str.find(keyword)
                    extracted = content_str[code_start:code_start + 5000]  # Extract chunk
                    cleaned = self._clean_extracted_code(extracted)
                    
                    if cleaned.strip():
                        modules.append({
                            'name': 'ExtractedCode',
                            'type': 'Unknown',
                            'code': cleaned,
                            'stream_path': 'manual_extraction'
                        })
                        break
        
        return modules
    
    def _manual_ole_extraction(self) -> List[Dict]:
        """
        Manual extraction for OLE compound files (.xls).
        """
        modules = []
        
        try:
            import olefile
            
            ole = olefile.OleFileIO(self.filepath)
            
            # Look for VBA storage
            if ole.exists('_VBA_PROJECT_CUR'):
                vba_root = '_VBA_PROJECT_CUR'
            elif ole.exists('Macros'):
                vba_root = 'Macros'
            else:
                # Try to find VBA in any storage
                for stream in ole.listdir():
                    stream_path = '/'.join(stream)
                    if 'VBA' in stream_path.upper():
                        content = ole.openstream(stream).read()
                        extracted = self._manual_vba_extraction(content)
                        modules.extend(extracted)
                
                ole.close()
                return modules
            
            # Extract from VBA storage
            for stream in ole.listdir():
                stream_path = '/'.join(stream)
                if vba_root in stream_path:
                    try:
                        content = ole.openstream(stream).read()
                        extracted = self._manual_vba_extraction(content)
                        modules.extend(extracted)
                    except:
                        pass
            
            ole.close()
            
        except ImportError:
            # If olefile is not available, try raw binary extraction
            with open(self.filepath, 'rb') as f:
                content = f.read()
            modules = self._manual_vba_extraction(content)
        
        return modules
    
    def _detect_module_type(self, module_name: str, code: str) -> str:
        """
        Detect the type of VBA module based on name and content.
        
        Args:
            module_name: Name of the module
            code: VBA code content
            
        Returns:
            Module type string
        """
        name_lower = module_name.lower() if module_name else ''
        
        if 'thisworkbook' in name_lower:
            return 'Document Module (ThisWorkbook)'
        elif 'sheet' in name_lower:
            return 'Document Module (Sheet)'
        elif 'userform' in name_lower or 'frm' in name_lower:
            return 'UserForm'
        elif 'class' in name_lower or 'cls' in name_lower:
            return 'Class Module'
        
        # Check code content
        if 'Attribute VB_Creatable' in code or 'Attribute VB_Exposed' in code:
            return 'Class Module'
        elif 'Begin {' in code or 'Begin VB.Form' in code:
            return 'UserForm'
        
        return 'Standard Module'
    
    def _clean_extracted_code(self, code: str) -> str:
        """
        Clean up extracted VBA code.
        
        Args:
            code: Raw extracted code
            
        Returns:
            Cleaned VBA code
        """
        # Remove null bytes and control characters
        cleaned = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', code)
        
        # Remove excessive whitespace
        lines = cleaned.split('\n')
        cleaned_lines = []
        
        for line in lines:
            # Skip lines that are just garbage
            if len(line) > 0 and len(line.strip()) > 0:
                # Check if line contains mostly printable characters
                printable_ratio = sum(c.isprintable() or c.isspace() for c in line) / len(line)
                if printable_ratio > 0.8:
                    cleaned_lines.append(line.rstrip())
        
        return '\n'.join(cleaned_lines)


def extract_vba_from_file(filepath: str) -> List[Dict]:
    """
    Convenience function to extract VBA from an Excel file.
    
    Args:
        filepath: Path to the Excel file
        
    Returns:
        List of module dictionaries
    """
    extractor = VBAExtractor(filepath)
    return extractor.extract_all()
