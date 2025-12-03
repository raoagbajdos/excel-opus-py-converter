"""
Excel VBA to Python Converter - Main Flask Application
"""
import os
from flask import Flask, request, jsonify, render_template, send_from_directory
from werkzeug.utils import secure_filename
from vba_extractor import VBAExtractor
from vba_to_python_converter import VBAToPythonConverter

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


if __name__ == '__main__':
    app.run(debug=True, port=5000)
