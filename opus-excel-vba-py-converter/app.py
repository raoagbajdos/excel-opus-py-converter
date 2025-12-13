"""
Excel VBA to Python Converter - Main Flask Application
"""
import os
from flask import Flask, request, jsonify, render_template, send_from_directory
from werkzeug.utils import secure_filename
from vba_extractor import VBAExtractor
from llm_converter import VBAToPythonConverter
from formula_extractor import FormulaExtractor
from data_exporter import DataExporter
from workbook_analyzer import WorkbookAnalyzer

app = Flask(__name__, static_folder='static', template_folder='templates')

# Configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'xlsm', 'xls', 'xlsb', 'xla', 'xlam'}
MAX_CONTENT_LENGTH = 50 * 1024 * 1024  # 50MB max file size

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# Ensure upload folder exists
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def allowed_file(filename):
    """Check if the file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Serve the main page."""
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """Handle file upload and extract VBA code."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({
            'error': f'Invalid file type. Allowed types: {", ".join(ALLOWED_EXTENSIONS)}'
        }), 400
    
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Extract VBA code
        extractor = VBAExtractor(filepath)
        vba_modules = extractor.extract_all()
        
        # Clean up uploaded file
        os.remove(filepath)
        
        if not vba_modules:
            return jsonify({
                'warning': 'No VBA code found in the uploaded file',
                'modules': []
            })
        
        return jsonify({
            'success': True,
            'filename': filename,
            'modules': vba_modules
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/convert', methods=['POST'])
def convert_vba():
    """Convert VBA code to Python."""
    data = request.get_json()
    
    if not data or 'vba_code' not in data:
        return jsonify({'error': 'No VBA code provided'}), 400
    
    vba_code = data['vba_code']
    module_name = data.get('module_name', 'converted_module')
    
    try:
        converter = VBAToPythonConverter()
        python_code = converter.convert(vba_code, module_name)
        
        return jsonify({
            'success': True,
            'python_code': python_code,
            'conversion_notes': converter.get_conversion_notes()
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/convert-all', methods=['POST'])
def convert_all_modules():
    """Convert all VBA modules to Python."""
    data = request.get_json()
    
    if not data or 'modules' not in data:
        return jsonify({'error': 'No modules provided'}), 400
    
    modules = data['modules']
    converter = VBAToPythonConverter()
    converted_modules = []
    
    try:
        for module in modules:
            python_code = converter.convert(
                module['code'], 
                module.get('name', 'module')
            )
            converted_modules.append({
                'name': module.get('name', 'module'),
                'type': module.get('type', 'unknown'),
                'original_code': module['code'],
                'python_code': python_code,
                'conversion_notes': converter.get_conversion_notes()
            })
        
        return jsonify({
            'success': True,
            'converted_modules': converted_modules
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/extract-formulas', methods=['POST'])
def extract_formulas():
    """Extract all formulas from an Excel file."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({
            'error': f'Invalid file type. Allowed types: {", ".join(ALLOWED_EXTENSIONS)}'
        }), 400
    
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Extract formulas
        extractor = FormulaExtractor(filepath)
        formulas = extractor.extract_all_formulas()
        
        # Get statistics
        statistics = extractor.get_formula_statistics(formulas)
        
        # Clean up uploaded file
        os.remove(filepath)
        
        # Convert FormulaInfo objects to dicts
        formulas_dict = [
            {
                'sheet_name': f.sheet_name,
                'cell_address': f.cell_address,
                'formula': f.formula,
                'formula_type': f.formula_type,
                'dependencies': f.dependencies,
                'functions': f.contains_functions
            }
            for f in formulas
        ]
        
        return jsonify({
            'success': True,
            'filename': filename,
            'formulas': formulas_dict,
            'statistics': statistics
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/convert-formula', methods=['POST'])
def convert_formula():
    """Convert a single Excel formula to Python."""
    data = request.get_json()
    
    if not data or 'formula' not in data:
        return jsonify({'error': 'No formula provided'}), 400
    
    formula = data['formula']
    cell_address = data.get('cell_address', 'A1')
    sheet_name = data.get('sheet_name', 'Sheet1')
    
    try:
        converter = VBAToPythonConverter()
        python_code = converter.convert_formula(formula, cell_address, sheet_name)
        
        return jsonify({
            'success': True,
            'python_code': python_code,
            'conversion_notes': converter.get_conversion_notes()
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/export-data', methods=['POST'])
def export_data():
    """Export Excel data to pandas DataFrames."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({
            'error': f'Invalid file type. Allowed types: {", ".join(ALLOWED_EXTENSIONS)}'
        }), 400
    
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Export data
        exporter = DataExporter(filepath)
        result = exporter.export_all_sheets()
        
        # Clean up uploaded file
        os.remove(filepath)
        
        return jsonify({
            'success': True,
            'filename': filename,
            'python_code': result.python_code,
            'metadata': result.metadata
        })
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/api/analyze-workbook', methods=['POST'])
def analyze_workbook():
    """Perform complete workbook analysis (VBA + formulas + data)."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if not allowed_file(file.filename):
        return jsonify({
            'error': f'Invalid file type. Allowed types: {", ".join(ALLOWED_EXTENSIONS)}'
        }), 400
    
    try:
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # Analyze workbook
        analyzer = WorkbookAnalyzer(filepath)
        analysis = analyzer.analyze_complete()
        
        # Generate report
        report = analyzer.generate_analysis_report(analysis)
        
        # Clean up uploaded file
        os.remove(filepath)
        
        # Convert analysis to dict for JSON
        response_data = {
            'success': True,
            'filename': analysis.filename,
            'has_vba': analysis.has_vba,
            'vba_modules_count': len(analysis.vba_modules) if analysis.vba_modules else 0,
            'has_formulas': analysis.has_formulas,
            'formulas_count': len(analysis.formulas) if analysis.formulas else 0,
            'sheets_count': len(analysis.data_export.sheet_data) if analysis.data_export else 0,
            'python_script': analysis.python_script,
            'report': report,
            'metadata': analysis.data_export.metadata if analysis.data_export else {}
        }
        
        return jsonify(response_data)
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    app.run(debug=True, port=5000)
