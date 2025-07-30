from flask import Flask, request, jsonify, render_template
from flask_cors import CORS
import os
import json
import logging
from werkzeug.utils import secure_filename
from io import StringIO
import sys
from final import (
    extract_castings_from_pdf,
    convert_extracted_data_to_castings,
    optimize_panels,
    print_results
)

# Configuration
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}
RESULTS_FILE = 'optimization_results.json'

# Setup
app = Flask(__name__)
CORS(app)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Logger setup
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/api/extract-castings', methods=['POST'])
def extract_castings():
    if 'pdf' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['pdf']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        try:
            file.save(file_path)
            castings_data = extract_castings_from_pdf(file_path)
            os.remove(file_path)
            return jsonify(castings_data), 200
        except Exception as e:
            logging.exception("Failed to process PDF")
            if os.path.exists(file_path):
                os.remove(file_path)
            return jsonify({'error': f"Failed to process PDF: {str(e)}"}), 500
    return jsonify({'error': 'Invalid file type'}), 400

@app.route('/api/optimize', methods=['POST']) 
def optimize():
    try:
        data = request.get_json()
        if not data or 'castings' not in data or 'primaryIdx' not in data:
            return jsonify({'error': 'Invalid request data'}), 400

        # Parse inputs
        castings = convert_extracted_data_to_castings(data['castings'])
        primary_idx = int(data['primaryIdx'])

        if primary_idx < 0 or primary_idx >= len(castings):
            return jsonify({'error': f"Invalid primaryIdx: {primary_idx}"}), 400

        # Perform optimization
        optimize_panels(castings, primary_idx)

        # ✅ Capture print_results output into a string
        original_stdout = sys.stdout
        sys.stdout = buffer = StringIO()
        print_results(castings, primary_idx)
        sys.stdout = original_stdout
        text_summary = buffer.getvalue()

        # Prepare summary results for JSON
        results = {
            'castings': [],
            'optimization_summary': {
                'total_castings': len(castings),
                'primary_casting': castings[primary_idx].name,
                'panel_sizes_used': list(set(
                    panel for casting in castings
                    for shape in casting.shapes
                    for panels in shape.panel_layout
                    for panel in panels
                )),
                'text_summary': text_summary  # ✅ Now it's defined
            }
        }

        for casting in castings:
            casting_data = {
                'name': casting.name,
                'shapes': []
            }
            for shape in casting.shapes:
                shape_data = {
                    'name': shape.name,
                    'sides': shape.sides,
                    'panel_layouts': shape.panel_layout
                }
                casting_data['shapes'].append(shape_data)
            results['castings'].append(casting_data)

        # Save summary to file
        with open(RESULTS_FILE, 'w') as f:
            json.dump(results, f, indent=2)

        return jsonify(results), 200

    except ValueError as ve:
        logging.error(f"ValueError: {str(ve)}")
        return jsonify({'error': f"Invalid data: {str(ve)}"}), 400
    except Exception as e:
        logging.exception("Optimization failed")
        return jsonify({'error': f"Optimization failed: {str(e)}"}), 500
    
if __name__ == '__main__':
    logging.info("Starting Flask server on http://localhost:5000")
    app.run(debug=True, port=5000)
