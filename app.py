"""
GMP Automation System - Main Web Application
Flask-based web interface for processing environmental measurement PDFs.
"""

import os
import json
import uuid
import traceback
from flask import Flask, render_template, request, jsonify, send_file, redirect, url_for
from werkzeug.utils import secure_filename
from config import UPLOAD_FOLDER, OUTPUT_FOLDER, get_semester_label, TEST_TYPES
from ocr_engine import EXTRACTORS as CLAUDE_EXTRACTORS
from deepseek_ocr.engine import EXTRACTORS as DEEPSEEK_EXTRACTORS
from excel_generator import GENERATORS

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024  # 100MB max

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

ALLOWED_EXTENSIONS = {'pdf'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """Open the online OCR workflow by default."""
    return redirect(url_for('online'))


@app.route('/online')
def online():
    """Online OCR workflow using the hosted API."""
    return render_template('index.html', ocr_mode='online')


@app.route('/offline')
def offline():
    """Offline OCR workflow using the configured OCR endpoint."""
    return render_template('index.html', ocr_mode='offline')


@app.route('/process', methods=['POST'])
def process():
    """Process uploaded PDFs and generate Excel files."""
    try:
        test_type = request.form.get('test_type')
        ocr_mode = request.form.get('ocr_mode', 'online').strip()
        api_key = request.form.get('api_key', '').strip()
        offline_endpoint = request.form.get('offline_endpoint', '').strip()

        if ocr_mode not in ('online', 'offline'):
            return jsonify({'error': 'Invalid OCR mode selected.'}), 400

        ocr_backend = 'claude' if ocr_mode == 'online' else 'deepseek'
        extractors = CLAUDE_EXTRACTORS if ocr_backend == 'claude' else DEEPSEEK_EXTRACTORS

        if not test_type or test_type not in extractors:
            return jsonify({'error': 'Invalid test type selected.'}), 400

        if ocr_backend == 'claude' and not api_key:
            return jsonify({'error': 'Anthropic API Key is required.'}), 400

        if ocr_backend == 'deepseek' and not offline_endpoint:
            return jsonify({'error': 'Offline OCR endpoint URL is required.'}), 400

        files = request.files.getlist('pdf_files')
        if not files or all(f.filename == '' for f in files):
            return jsonify({'error': 'No PDF files uploaded.'}), 400

        # Save uploaded files
        saved_paths = []
        for f in files:
            if f and allowed_file(f.filename):
                filename = secure_filename(f.filename)
                unique_name = f"{uuid.uuid4().hex}_{filename}"
                filepath = os.path.join(UPLOAD_FOLDER, unique_name)
                f.save(filepath)
                saved_paths.append(filepath)

        if not saved_paths:
            return jsonify({'error': 'No valid PDF files found.'}), 400

        # Extract data from each PDF
        extractor = extractors[test_type]
        all_ahu_data = {}
        errors = []

        for pdf_path in saved_paths:
            try:
                if ocr_backend == 'claude':
                    data = extractor(pdf_path, api_key=api_key)
                else:
                    data = extractor(pdf_path, endpoint_url=offline_endpoint)
                ahu_num = str(data.get('ahu', 'unknown'))
                date_str = data.get('date', '2025.08.01')
                semester_label = get_semester_label(date_str)

                # Organize data by AHU
                if ahu_num not in all_ahu_data:
                    all_ahu_data[ahu_num] = []

                # Build semester data based on test type
                sem_entry = {'semester': semester_label, 'date': date_str}

                if test_type == 'airborne_particle':
                    sem_entry['rooms'] = data.get('rooms', [])
                elif test_type == 'air_velocity':
                    sem_entry['machines'] = data.get('machines', [])
                elif test_type == 'air_change_rate':
                    sem_entry['rooms'] = data.get('rooms', [])
                elif test_type == 'hepa_filter':
                    sem_entry['items'] = data.get('items', [])
                elif test_type == 'airflow_pattern':
                    sem_entry['items'] = data.get('items', [])
                    # For airflow pattern, the AHU might not be in the PDF
                    # Try to get it from the filename
                    if ahu_num == 'unknown':
                        basename = os.path.basename(pdf_path).lower()
                        for part in basename.replace('-', '_').split('_'):
                            if part.startswith('ahu'):
                                ahu_num = part.replace('ahu', '')
                                break
                        if ahu_num == 'unknown':
                            ahu_num = 'unknown'

                        if ahu_num not in all_ahu_data:
                            all_ahu_data[ahu_num] = []

                all_ahu_data[ahu_num].append(sem_entry)

            except Exception as e:
                errors.append(f"Error processing {os.path.basename(pdf_path)}: {str(e)}")

        if not all_ahu_data:
            error_msg = "Failed to extract data from all PDFs."
            if errors:
                error_msg += "\n" + "\n".join(errors)
            return jsonify({'error': error_msg}), 400

        # Generate Excel
        generator = GENERATORS[test_type]
        test_config = TEST_TYPES[test_type]
        output_filename = test_config['excel_filename']
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)

        generator(all_ahu_data, output_path)

        # Clean up uploaded files
        for p in saved_paths:
            try:
                os.remove(p)
            except:
                pass

        result = {
            'success': True,
            'filename': output_filename,
            'download_url': f'/download/{output_filename}',
            'ahu_count': len(all_ahu_data),
            'ahu_list': sorted(all_ahu_data.keys(), key=lambda x: int(x) if x.isdigit() else x),
        }

        if errors:
            result['warnings'] = errors

        return jsonify(result)

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'Server error: {str(e)}'}), 500


@app.route('/download/<filename>')
def download(filename):
    """Download generated Excel file."""
    filepath = os.path.join(OUTPUT_FOLDER, filename)
    if os.path.exists(filepath):
        return send_file(filepath, as_attachment=True, download_name=filename)
    return jsonify({'error': 'File not found.'}), 404


if __name__ == '__main__':
    print("=" * 60)
    print("  GMP Automation System")
    print("  Open your browser and go to: http://localhost:5001")
    print("=" * 60)
    app.run(host='0.0.0.0', port=5001, debug=False)
