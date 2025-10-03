#!/usr/bin/env python
# -*- encoding: utf-8 -*-
from flask import Flask, render_template, request, jsonify, send_from_directory
from flask_cors import CORS
import docx
from simplify_docx import simplify
import mammoth
import json
import os
from werkzeug.utils import secure_filename
import traceback
from docx2pdf import convert
from pdf2image import convert_from_path
from PIL import Image
import base64
from io import BytesIO
import tempfile

app = Flask(__name__)
CORS(app)  # Enable CORS for all routes
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0  # Disable caching for development

# Create uploads directory if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        print("Upload request received")  # Debug log

        if 'file' not in request.files:
            print("No file in request")
            return jsonify({'error': 'No file part'}), 400

        file = request.files['file']
        print(f"File received: {file.filename}")

        if file.filename == '':
            return jsonify({'error': 'No selected file'}), 400

        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            print(f"Saving to: {filepath}")
            file.save(filepath)

            try:
                print("Converting to HTML...")
                # Convert to HTML using mammoth
                with open(filepath, "rb") as docx_file:
                    result = mammoth.convert_to_html(docx_file)
                    html_content = result.value
                print(f"HTML conversion complete: {len(html_content)} chars")

                print("Simplifying DOCX...")
                # Simplify using simplify-docx
                doc = docx.Document(filepath)
                simplified_json = simplify(doc)
                print("Simplification complete")

                print("Converting DOCX to image...")
                # Convert DOCX to PDF, then to image
                image_data = None
                pdf_path = None
                try:
                    # Set Poppler path
                    poppler_path = os.path.join(os.getcwd(), "poppler-24.08.0", "Library", "bin")
                    print(f"Using Poppler from: {poppler_path}")

                    # Create a temporary PDF file
                    pdf_path = filepath.replace('.docx', '.pdf')
                    convert(filepath, pdf_path)
                    print(f"PDF created: {pdf_path}")

                    # Convert PDF to images with Poppler path
                    images = convert_from_path(pdf_path, dpi=150, poppler_path=poppler_path)
                    print(f"Converted to {len(images)} image(s)")

                    # For simplicity, just use the first page
                    if images:
                        # Convert PIL Image to base64
                        buffered = BytesIO()
                        images[0].save(buffered, format="PNG")
                        image_data = base64.b64encode(buffered.getvalue()).decode('utf-8')
                        print("Image conversion complete")

                    # Clean up PDF
                    if pdf_path and os.path.exists(pdf_path):
                        os.remove(pdf_path)

                except Exception as img_error:
                    print(f"Warning: Could not convert to image: {str(img_error)}")
                    print(traceback.format_exc())
                    if pdf_path and os.path.exists(pdf_path):
                        os.remove(pdf_path)

                # Clean up uploaded file
                os.remove(filepath)

                response_data = {
                    'html': html_content,
                    'json': json.dumps(simplified_json, indent=2),
                    'image': image_data,
                    'filename': filename
                }
                print("Sending response")
                return jsonify(response_data)

            except Exception as e:
                print(f"Error processing file: {str(e)}")
                print(traceback.format_exc())
                if os.path.exists(filepath):
                    os.remove(filepath)
                return jsonify({'error': f'Processing error: {str(e)}'}), 500

        return jsonify({'error': 'Invalid file type. Please upload a .docx file'}), 400

    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        print(traceback.format_exc())
        return jsonify({'error': f'Server error: {str(e)}'}), 500

if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1', port=5000, use_reloader=False, threaded=True)
