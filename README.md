# DOCX Converter & Live Editor

A powerful web-based tool for converting DOCX files to HTML/JSON with real-time editing capabilities and fidelity comparison.

![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)
![Python](https://img.shields.io/badge/python-3.7+-green.svg)
![License](https://img.shields.io/badge/license-MIT-orange.svg)

## 🌟 Features

### 📊 Multi-Panel Comparison Views
- **All 3 Panels**: Side-by-side comparison of Original DOCX | HTML | JSON
- **Paired Views**: Compare any two formats (Original vs HTML, Original vs JSON, HTML vs JSON)
- **Single Views**: Focus on individual formats
- **8 Different View Modes** for comprehensive analysis

### ✨ Live JSON Editor
- **Real-time HTML Preview**: Edit JSON and see HTML changes instantly
- **Live Mode**: Auto-update preview as you type (500ms debounce)
- **Manual Mode**: Update preview on-demand
- **JSON Formatting**: Beautify/format JSON with one click
- **Validation**: Real-time JSON syntax validation
- **Status Feedback**: Visual indicators for success/errors

### 🎨 Visual Comparison
- **Original DOCX Preview**: View document as image (requires Poppler)
- **HTML Rendering**: See how document renders in browser
- **JSON Structure**: Explore simplified document structure

### 🚀 User-Friendly Interface
- **Drag & Drop Upload**: Easy file upload
- **Responsive Design**: Works on all screen sizes
- **Beautiful UI**: Modern gradient design with smooth animations

## 📦 Installation

### Prerequisites

- Python 3.7 or higher
- pip package manager
- Microsoft Word (optional, for DOCX to image conversion on Windows)
- Poppler (optional, for PDF to image conversion)

### Quick Start

1. **Clone the repository**
   ```bash
   git clone https://github.com/Tailor-AUS/SimplifyDocx.git
   cd SimplifyDocx
   ```

2. **Install Python dependencies**
   ```bash
   pip install -e .
   pip install flask flask-cors mammoth docx2pdf pdf2image
   ```

3. **Install Poppler (Optional - for image preview)**

   **Windows:**
   ```bash
   curl -L -o poppler.zip https://github.com/oschwartz10612/poppler-windows/releases/download/v24.08.0-0/Release-24.08.0-0.zip
   powershell -Command "Expand-Archive -Path poppler.zip -DestinationPath . -Force"
   ```

   **macOS:**
   ```bash
   brew install poppler
   ```

   **Linux:**
   ```bash
   sudo apt-get install poppler-utils
   ```

4. **Run the application**
   ```bash
   python app.py
   ```

5. **Open your browser**
   Navigate to: http://127.0.0.1:5000

## 🎯 Usage

### Basic Workflow

1. **Upload Document**
   - Drag & drop a DOCX file or click to browse
   - Supports files up to 16MB

2. **View Comparisons**
   - Switch between different view modes using tabs
   - Compare original formatting with conversions

3. **Live Editing**
   - Go to "✨ Live JSON Editor" tab
   - Edit JSON structure on the left
   - Click "Enable Live Preview" for auto-updates
   - Or click "Update Preview" for manual updates
   - Watch HTML preview update in real-time on the right

### View Modes

- **All 3 Panels**: Original | HTML | JSON side-by-side
- **Live JSON Editor**: Edit JSON with live HTML preview
- **Original vs HTML**: Compare original document with HTML rendering
- **Original vs JSON**: Compare original with simplified structure
- **HTML vs JSON**: Compare HTML output with JSON structure
- **Original Only**: Full view of original DOCX
- **HTML Only**: Full view of HTML conversion
- **JSON Only**: Full view of JSON structure

## 🏗️ Architecture

### Backend (Flask)

**app.py** - Main application server
- File upload handling with security
- DOCX to HTML conversion (Mammoth)
- DOCX to JSON simplification (Simplify-Docx library)
- DOCX to image conversion (docx2pdf + pdf2image + Poppler)
- JSON to HTML live conversion endpoint
- CORS enabled for development

### Frontend (HTML/CSS/JavaScript)

**templates/index.html** - Single-page application
- Responsive grid layout system
- 8 different view modes
- Real-time JSON editor with syntax validation
- Live preview with debounced updates
- Drag & drop file upload
- Status notifications and error handling

### Core Libraries

- **Flask**: Web framework
- **Mammoth**: DOCX to HTML conversion
- **python-docx**: DOCX parsing
- **Simplify-Docx**: Document structure simplification
- **docx2pdf**: DOCX to PDF conversion
- **pdf2image**: PDF to image conversion
- **Poppler**: PDF rendering engine

## 📁 Project Structure

```
SimplifyDocx/
├── app.py                    # Flask application
├── templates/
│   └── index.html           # Web interface
├── src/                     # Simplify-Docx library
│   └── simplify_docx/
├── uploads/                 # Temporary file storage (gitignored)
├── poppler-24.08.0/        # Poppler binaries (gitignored)
├── README.md               # This file
├── WEB_APP_README.md       # Additional documentation
├── setup.py                # Python package setup
└── .gitignore              # Git ignore rules
```

## 🔧 Configuration

### File Size Limit
Edit in `app.py`:
```python
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
```

### Live Preview Debounce
Edit in `templates/index.html`:
```javascript
updateTimeout = setTimeout(() => {
    updatePreview();
}, 500); // 500ms delay
```

### Server Settings
```python
app.run(
    debug=True,              # Set to False in production
    host='127.0.0.1',       # Change to '0.0.0.0' for network access
    port=5000,              # Change port if needed
    use_reloader=False,     # Prevent double loading
    threaded=True           # Enable threading
)
```

## 🐛 Troubleshooting

### Image Preview Not Available
- **Cause**: Poppler not installed or not in PATH
- **Solution**: Install Poppler following instructions above
- **Note**: Image preview is optional; HTML and JSON views work without it

### Microsoft Word Error (Windows)
- **Cause**: docx2pdf requires Microsoft Word
- **Solution**: Install Microsoft Word or accept that image preview won't work
- **Alternative**: All other features work without Word

### Port Already in Use
```bash
# Find process on port 5000
netstat -ano | findstr :5000

# Kill the process (Windows)
taskkill /F /PID <process_id>

# Kill the process (Mac/Linux)
kill -9 <process_id>
```

### JSON Validation Errors
- Ensure JSON is properly formatted
- Use "Format JSON" button to fix formatting
- Check browser console for detailed error messages

## 🚀 Deployment

### Development
```bash
python app.py
```

### Production (using Gunicorn)
```bash
pip install gunicorn
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

### Docker (coming soon)
```dockerfile
# Dockerfile example
FROM python:3.9
WORKDIR /app
COPY . .
RUN pip install -e .
RUN pip install flask flask-cors mammoth docx2pdf pdf2image gunicorn
EXPOSE 5000
CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:5000", "app:app"]
```

## 🤝 Contributing

This is a private repository. For internal contributions:

1. Create a feature branch
2. Make your changes
3. Test thoroughly
4. Submit a pull request

## 📄 License

This project is based on [microsoft/Simplify-Docx](https://github.com/microsoft/Simplify-Docx) and includes significant enhancements.

Original Simplify-Docx library: MIT License
Web application enhancements: MIT License

## 🙏 Acknowledgments

- Original Simplify-Docx library by Microsoft Research
- Mammoth.js for DOCX to HTML conversion
- Flask framework and community
- All open-source contributors

## 📞 Support

For questions or issues, please contact your development team lead.

## 🔄 Version History

### v1.0.0 (Current)
- ✨ Live JSON Editor with real-time HTML preview
- 📊 8 different comparison view modes
- 🎨 Modern responsive UI
- 🖼️ Image preview support
- 🚀 Drag & drop file upload
- ⚡ Real-time validation and feedback

---

**Built with ❤️ by the Tailor-AUS Team**
