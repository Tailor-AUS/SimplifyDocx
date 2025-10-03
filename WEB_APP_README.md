# DOCX Converter - Web Interface

A web-based fidelity comparison tool for DOCX to HTML/JSON conversion, built on top of the Simplify-Docx library.

## Features

- 📄 **3-Panel Split View**: Compare Original DOCX, HTML Preview, and Simplified JSON side-by-side
- 🎨 **Multiple View Options**: 7 different comparison modes
  - All 3 Panels
  - Original vs HTML
  - Original vs JSON
  - HTML vs JSON
  - Individual views for each format
- 🖼️ **Image Preview**: Visual representation of the original DOCX document
- 🌐 **HTML Rendering**: See how the document renders in HTML using Mammoth
- 📊 **JSON Structure**: View the simplified document structure
- 🚀 **Drag & Drop**: Easy file upload interface

## Installation

### Prerequisites

- Python 3.7+
- Microsoft Word (for DOCX to PDF conversion on Windows)
- Poppler (for PDF to image conversion)

### Install Dependencies

```bash
pip install -e .
pip install flask flask-cors mammoth docx2pdf pdf2image
```

### Install Poppler (Windows)

Download Poppler from the releases and extract to the project directory:

```bash
curl -L -o poppler.zip https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip
powershell -Command "Expand-Archive -Path poppler.zip -DestinationPath . -Force"
```

The app is configured to use `poppler-24.08.0/Library/bin` automatically.

### Run the Application

```bash
python app.py
```

The application will start on http://127.0.0.1:5000

## Usage

1. Open your browser to http://127.0.0.1:5000
2. Drag & drop a DOCX file or click to browse
3. Wait for processing (HTML, JSON, and image generation)
4. Compare the outputs using different view tabs

## Architecture

### Backend (Flask)
- **app.py**: Main Flask application
  - File upload handling
  - DOCX to HTML conversion (using Mammoth)
  - DOCX to JSON simplification (using Simplify-Docx)
  - DOCX to image conversion (using docx2pdf + pdf2image + Poppler)

### Frontend (HTML/CSS/JavaScript)
- **templates/index.html**: Single-page application
  - Responsive grid layout
  - Multiple view modes
  - Real-time file processing feedback

## File Structure

```
Simplify-Docx/
├── app.py                 # Flask web application
├── templates/
│   └── index.html        # Web interface
├── uploads/              # Temporary file storage (gitignored)
├── src/                  # Original Simplify-Docx library
└── poppler-24.08.0/      # Poppler binaries (gitignored)
```

## Contributing

This is a fork of [microsoft/Simplify-Docx](https://github.com/microsoft/Simplify-Docx) with added web interface capabilities.

## License

Same as the original Simplify-Docx project.
