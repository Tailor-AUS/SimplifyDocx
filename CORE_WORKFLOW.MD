# CORE_WORKFLOW.MD - Technical Reference for DOCX Converter & Live Editor

## Table of Contents
1. [Application Overview](#application-overview)
2. [Architecture](#architecture)
3. [Core Processing Workflows](#core-processing-workflows)
4. [API Endpoints](#api-endpoints)
5. [Data Flow Diagrams](#data-flow-diagrams)
6. [Key Components & Functions](#key-components--functions)
7. [Configuration & Setup](#configuration--setup)
8. [Common Development Tasks](#common-development-tasks)
9. [Debugging Guide](#debugging-guide)
10. [Performance Considerations](#performance-considerations)
11. [Security Considerations](#security-considerations)

---

## Application Overview

### Purpose
A web-based DOCX converter with real-time editing capabilities that provides:
- Multi-format document conversion (DOCX → HTML/JSON/Image)
- Side-by-side fidelity comparison
- Live JSON editing with instant HTML preview
- Page-aware document splitting

### Tech Stack
- **Backend**: Flask (Python 3.7+)
- **Frontend**: Vanilla JavaScript, HTML5, CSS3
- **Core Libraries**:
  - `simplify-docx`: Custom library for DOCX → JSON conversion
  - `mammoth`: DOCX → HTML conversion
  - `python-docx`: DOCX parsing
  - `docx2pdf`: DOCX → PDF (requires MS Word on Windows)
  - `pdf2image`: PDF → Image (requires Poppler)

### File Structure
```
Simplify-Docx/
├── app.py                    # Flask server (719 lines)
├── templates/
│   └── index.html           # Single-page frontend (1500+ lines)
├── src/simplify_docx/       # Core conversion library
│   ├── __init__.py          # Main entry point
│   ├── elements/            # DOCX element classes
│   ├── iterators/           # Document traversal logic
│   ├── types/               # Type definitions
│   └── utils/               # Helper functions
├── uploads/                 # Temp file storage (gitignored)
└── CLAUDE.md               # AI assistant instructions
```

---

## Architecture

### Component Interaction
```
┌─────────────────┐
│   Browser UI    │
│  (index.html)   │
└────────┬────────┘
         │ HTTP
         ↓
┌─────────────────┐
│   Flask Server  │
│    (app.py)     │
└────────┬────────┘
         │
    ┌────┴─────┬──────────┬─────────────┐
    ↓          ↓          ↓             ↓
┌─────────┐ ┌──────┐ ┌──────────┐ ┌──────────┐
│ Mammoth │ │Simpl.│ │ docx2pdf │ │pdf2image │
│  HTML   │ │ JSON │ │   PDF    │ │  Image   │
└─────────┘ └──────┘ └──────────┘ └──────────┘
```

### Request/Response Flow
1. User uploads DOCX via drag-drop or file picker
2. Flask receives file at `/upload` endpoint
3. Three parallel conversions execute:
   - Mammoth → HTML
   - simplify-docx → JSON
   - docx2pdf → PDF → Image (optional)
4. Results return as JSON with all formats
5. Frontend renders in selected view mode
6. Live editing sends JSON to `/json-to-html` for real-time preview

---

## Core Processing Workflows

### 1. File Upload & Processing (`app.py:415-575`)
```python
@app.route('/upload', methods=['POST'])
def upload_file():
```

**Workflow Steps:**
1. **Validation** (`app.py:420-430`)
   - Check file exists in request
   - Validate file extension (.docx only)
   - Apply `secure_filename()` sanitization

2. **Save to Temp** (`app.py:431-434`)
   ```python
   filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
   file.save(filepath)
   ```

3. **HTML Conversion** (`app.py:438-442`)
   ```python
   with open(filepath, "rb") as docx_file:
       result = mammoth.convert_to_html(docx_file)
       html_content = result.value
   ```

4. **JSON Simplification** (`app.py:450-453`)
   ```python
   doc = docx.Document(filepath)
   simplified_json = simplify(doc)
   ```

5. **Page Detection** (`app.py:461-474`)
   - Counts page breaks in document
   - Checks for `\f` or `\x0c` characters
   - Updates from PDF if available

6. **Image Generation** (`app.py:476-521`)
   ```python
   # Windows COM initialization
   pythoncom.CoInitialize()
   convert(filepath, pdf_path)  # DOCX → PDF
   images = convert_from_path(pdf_path, dpi=150, poppler_path=poppler_path)  # PDF → Images
   ```

7. **Page Splitting** (`app.py:527-532`)
   ```python
   html_pages = split_html_into_pages(html_content, num_pages, doc=doc, pdf_path=pdf_path)
   json_pages = split_json_into_pages(json_str, num_pages, doc=doc, pdf_path=pdf_path, html_content=html_content)
   ```

8. **Response Assembly** (`app.py:550-560`)
   ```python
   response_data = {
       'html': html_content,          # Full HTML
       'json': json_str,             # Full JSON
       'htmlPages': html_pages,      # Split by page
       'jsonPages': json_pages,      # Split by page
       'images': image_data,         # Base64 images
       'pageCount': num_pages,
       'filename': filename
   }
   ```

### 2. JSON to HTML Live Conversion (`app.py:576-606`)
```python
@app.route('/json-to-html', methods=['POST'])
def json_to_html():
    # Receives edited JSON, returns updated HTML
    html = convert_json_to_html(json_structure)
```

**Conversion Logic** (`app.py:634-715`):
```python
def convert_json_to_html(obj):
    # Recursive conversion with type mapping
    tag_map = {
        'document': 'div',
        'body': 'div',
        'paragraph': 'p',
        'table': 'table',
        'table-row': 'tr',
        'table-cell': 'td',
    }
```

### 3. Page Boundary Detection

**DOCX-Based** (`app.py:38-107`):
- Detects rendered page breaks
- Checks paragraph.rendered_page_breaks
- Searches for page breaks in runs

**PDF-Based** (`app.py:109-225`):
- Extracts text from each PDF page using pdfplumber
- Maps HTML paragraphs to PDF pages
- Handles unmatched paragraphs with proximity-based assignment

---

## API Endpoints

### POST /upload
**Purpose**: Upload and process DOCX file

**Request**:
- Method: `POST`
- Content-Type: `multipart/form-data`
- Body: `file` field with DOCX file
- Max Size: 16MB

**Response** (200 OK):
```json
{
  "html": "<p>Full HTML content...</p>",
  "json": "{ \"TYPE\": \"document\", ... }",
  "htmlPages": ["<p>Page 1...</p>", "<p>Page 2...</p>"],
  "jsonPages": ["{ page1 json }", "{ page2 json }"],
  "images": ["base64_page1", "base64_page2"],
  "pageCount": 2,
  "filename": "document.docx"
}
```

**Error Responses**:
- 400: Invalid file type or no file
- 500: Processing error

### POST /json-to-html
**Purpose**: Convert edited JSON back to HTML

**Request**:
```json
{
  "json": "{ \"TYPE\": \"document\", \"VALUE\": [...] }"
}
```

**Response** (200 OK):
```json
{
  "html": "<div><p>Converted HTML...</p></div>"
}
```

---

## Data Flow Diagrams

### Upload & Conversion Flow
```
DOCX File
    │
    ├──→ [Mammoth] ──→ HTML
    │
    ├──→ [simplify-docx] ──→ JSON Structure
    │                           │
    │                           ↓
    │                    [Type Mapping]
    │                           │
    │                           ↓
    │                      Simplified JSON
    │
    └──→ [docx2pdf] ──→ PDF ──→ [pdf2image] ──→ PNG Images
```

### Live Edit Flow
```
User Edits JSON
      │
      ↓ (Debounced 500ms)
POST /json-to-html
      │
      ↓
convert_json_to_html()
      │
      ↓
Recursive Mapping
      │
      ↓
Updated HTML
      │
      ↓
Preview Updates
```

### Page Splitting Algorithm
```
Document
    │
    ├──→ Check DOCX Page Breaks ─→ Found? ─→ Use DOCX Boundaries
    │                                 │
    │                                 ↓ Not Found
    ├──→ Check PDF Page Text ────→ Found? ─→ Use PDF Boundaries
    │                                 │
    │                                 ↓ Not Found
    └──→ Equal Division ─────────→ Split Evenly
```

---

## Key Components & Functions

### Core Functions

#### 1. `simplify()` (`src/simplify_docx/__init__.py:20`)
```python
def simplify(doc, options=None):
    """Main entry point for DOCX → JSON conversion"""
    # Uses visitor pattern to traverse document
    # Returns nested JSON structure
```

#### 2. `convert_json_to_html()` (`app.py:634-715`)
```python
def convert_json_to_html(obj):
    """Recursive JSON → HTML converter"""
    # Handles: strings, lists, dicts
    # Maps types to HTML tags
    # Detects headings via is_heading()
```

#### 3. `split_html_into_pages()` (`app.py:227-289`)
```python
def split_html_into_pages(html_content, num_pages, doc=None, pdf_path=None):
    """Intelligent page splitting with fallback strategies"""
    # Priority: DOCX breaks → PDF boundaries → Equal split
```

#### 4. `extract_page_boundaries_from_pdf()` (`app.py:109-225`)
```python
def extract_page_boundaries_from_pdf(pdf_path, html_content):
    """Maps HTML paragraphs to PDF pages using text matching"""
    # Uses flexible text matching strategies
    # Handles unmatched paragraphs with proximity assignment
```

### Frontend Components (`templates/index.html`)

#### File Upload Handler (Lines ~800-900)
```javascript
function handleFileUpload(file) {
    const formData = new FormData();
    formData.append('file', file);
    fetch('/upload', { method: 'POST', body: formData })
        .then(response => response.json())
        .then(data => displayResults(data));
}
```

#### Live Preview System (Lines ~1200-1350)
```javascript
let updateTimeout;
jsonEditor.addEventListener('input', () => {
    clearTimeout(updateTimeout);
    updateTimeout = setTimeout(() => {
        updatePreview();
    }, 500); // 500ms debounce
});

function updatePreview() {
    fetch('/json-to-html', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ json: jsonEditor.value })
    })
    .then(response => response.json())
    .then(data => htmlPreview.innerHTML = data.html);
}
```

---

## Configuration & Setup

### Quick Start for New Developers

1. **Clone & Navigate**
   ```bash
   git clone [repository]
   cd Simplify-Docx
   ```

2. **Install Core Dependencies**
   ```bash
   pip install -e .  # Install simplify-docx library
   pip install flask flask-cors mammoth python-docx
   ```

3. **Optional: Image Preview Setup**
   ```bash
   # Install for DOCX → PDF → Image conversion
   pip install docx2pdf pdf2image

   # Windows: Install Poppler
   curl -L -o poppler.zip [poppler-url]
   powershell -Command "Expand-Archive -Path poppler.zip -DestinationPath ."
   ```

4. **Run Application**
   ```bash
   python app.py
   # Server starts at http://127.0.0.1:5000
   ```

### Configuration Options

#### File Size Limit (`app.py:27`)
```python
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
```

#### Upload Directory (`app.py:26`)
```python
app.config['UPLOAD_FOLDER'] = 'uploads'
```

#### Simplification Options (`src/simplify_docx/__init__.py:44-84`)
```python
__default_options__ = {
    'remove-leading-white-space': True,
    'remove-trailing-white-space': True,
    'flatten-hyperlinks': True,
    'flatten-smart-tags': True,
    # ... more options
}
```

---

## Common Development Tasks

### 1. Add New Document Type Mapping
**Location**: `app.py:681-691`
```python
tag_map = {
    'document': 'div',
    'paragraph': 'p',
    'your-new-type': 'your-html-tag',  # Add here
}
```

### 2. Modify Heading Detection
**Location**: `app.py:608-632`
```python
def is_heading(text):
    # Add your heading patterns here
    if 'YourPattern' in text:
        return 'h1'
```

### 3. Add New Simplification Element
1. Create element class in `src/simplify_docx/elements/`
2. Implement `to_json()` method
3. Add iterator in `src/simplify_docx/iterators/`
4. Import in `__init__.py`

### 4. Change Live Preview Debounce
**Location**: `templates/index.html` (~Line 1250)
```javascript
updateTimeout = setTimeout(() => {
    updatePreview();
}, 500); // Change 500 to desired milliseconds
```

### 5. Add New View Mode
1. Add tab button in `templates/index.html` (~Line 400)
2. Create tab content div (~Line 500)
3. Add JavaScript handler (~Line 1100)

### 6. Customize Image DPI
**Location**: `app.py:500`
```python
images = convert_from_path(pdf_path, dpi=150)  # Change DPI here
```

---

## Debugging Guide

### Common Issues & Solutions

#### 1. "No module named simplify_docx"
```bash
# Solution: Install library in editable mode
pip install -e .
```

#### 2. Port 5000 Already in Use
```bash
# Find process
netstat -ano | findstr :5000

# Kill process (Windows)
taskkill /F /PID [process_id]

# Kill process (Mac/Linux)
kill -9 [process_id]
```

#### 3. Image Preview Not Working
**Check Poppler Installation**:
```python
# app.py:484
poppler_path = os.path.join(os.getcwd(), "poppler-24.08.0", "Library", "bin")
print(f"Poppler path exists: {os.path.exists(poppler_path)}")
```

**Check MS Word (Windows)**:
```python
# Requires MS Word for docx2pdf
# Test: Convert manually in Word
```

#### 4. Page Splitting Issues
**Debug Output** (`app.py`):
```python
# Enable debug prints
print(f"  Total HTML paragraphs: {len(html_paragraphs)}")
print(f"  Page {page_num + 1}: paragraphs {page_para_indices}")
```

#### 5. JSON Conversion Errors
**Add Logging** (`app.py:589-591`):
```python
print(f"Input JSON structure: {str(json_structure)[:500]}")
print(f"Output HTML: {html[:500]}")
```

### Debug Mode Settings
```python
# app.py:718
app.run(debug=True,  # Enable debug mode
        use_reloader=False,  # Prevent double loading
        threaded=True)  # Enable threading
```

### Browser Console Commands
```javascript
// Check uploaded data
console.log(uploadedData);

// Test JSON conversion
fetch('/json-to-html', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({json: testJson})
}).then(r => r.json()).then(console.log);
```

---

## Performance Considerations

### Bottlenecks

1. **DOCX to PDF Conversion** (Windows)
   - Uses MS Word COM automation
   - Can take 2-5 seconds per document
   - Solution: Make optional, cache results

2. **Large Document Processing**
   - Memory usage scales with document size
   - Current limit: 16MB
   - Solution: Stream processing for large files

3. **Page Detection**
   - PDF text extraction is expensive
   - Multiple regex operations
   - Solution: Cache boundaries, optimize regex

### Optimization Opportunities

#### 1. Parallel Processing
```python
# Current: Sequential
html = convert_html()
json = convert_json()
image = convert_image()

# Optimized: Parallel with threading
from concurrent.futures import ThreadPoolExecutor
with ThreadPoolExecutor() as executor:
    html_future = executor.submit(convert_html)
    json_future = executor.submit(convert_json)
    image_future = executor.submit(convert_image)
```

#### 2. Caching
```python
# Add Redis or in-memory cache
from functools import lru_cache

@lru_cache(maxsize=100)
def convert_cached(filepath_hash):
    return process_document(filepath)
```

#### 3. Lazy Loading
```javascript
// Load images only when needed
function loadImageForPage(pageNum) {
    if (!pageImages[pageNum]) {
        pageImages[pageNum] = new Image();
        pageImages[pageNum].src = imageData[pageNum];
    }
}
```

### Memory Management

1. **Clean Up Temp Files** (`app.py:535-546`)
   ```python
   os.remove(pdf_path)  # Remove PDF
   os.remove(filepath)  # Remove uploaded DOCX
   ```

2. **Limit Concurrent Uploads**
   ```python
   from threading import Semaphore
   upload_semaphore = Semaphore(5)  # Max 5 concurrent
   ```

---

## Security Considerations

### Current Security Measures

1. **File Type Validation** (`app.py:33-36`)
   ```python
   ALLOWED_EXTENSIONS = {'docx'}
   def allowed_file(filename):
       return '.' in filename and
              filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
   ```

2. **Filename Sanitization** (`app.py:431`)
   ```python
   filename = secure_filename(file.filename)
   ```

3. **Size Limit** (`app.py:27`)
   ```python
   app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB
   ```

4. **CORS Configuration** (`app.py:25`)
   ```python
   CORS(app)  # Currently allows all origins
   ```

### Recommended Enhancements

1. **Restrict CORS**
   ```python
   CORS(app, origins=['http://localhost:3000', 'https://yourdomain.com'])
   ```

2. **Add Rate Limiting**
   ```python
   from flask_limiter import Limiter
   limiter = Limiter(app, key_func=lambda: request.remote_addr)

   @app.route('/upload', methods=['POST'])
   @limiter.limit("10 per minute")
   def upload_file():
   ```

3. **Virus Scanning**
   ```python
   import pyclamd
   cd = pyclamd.ClamdAgnostic()
   if cd.scan_stream(file.read()):
       return jsonify({'error': 'File infected'}), 400
   ```

4. **Input Validation**
   ```python
   def validate_json_structure(json_obj):
       # Validate structure before processing
       required_keys = ['TYPE', 'VALUE']
       if not all(k in json_obj for k in required_keys):
           raise ValueError("Invalid JSON structure")
   ```

5. **Session Management**
   ```python
   from flask_session import Session
   app.config['SESSION_TYPE'] = 'filesystem'
   Session(app)
   ```

---

## Testing Recommendations

### Unit Tests
```python
# test_app.py
def test_allowed_file():
    assert allowed_file('test.docx') == True
    assert allowed_file('test.pdf') == False

def test_is_heading():
    assert is_heading('1.0 Introduction') == 'h2'
    assert is_heading('Regular text') == None
```

### Integration Tests
```python
def test_upload_endpoint():
    with open('test.docx', 'rb') as f:
        response = client.post('/upload',
            data={'file': (f, 'test.docx')},
            content_type='multipart/form-data')
    assert response.status_code == 200
    assert 'html' in response.json
```

### Load Testing
```bash
# Using Apache Bench
ab -n 100 -c 10 -p test.docx http://127.0.0.1:5000/upload

# Using locust
locust -f locustfile.py --host=http://127.0.0.1:5000
```

---

## Deployment Checklist

### Development → Production

- [ ] Set `debug=False` in `app.py:718`
- [ ] Configure proper CORS origins
- [ ] Set up HTTPS/SSL certificates
- [ ] Implement rate limiting
- [ ] Add logging/monitoring
- [ ] Set up error tracking (Sentry)
- [ ] Configure production WSGI server (Gunicorn)
- [ ] Set up reverse proxy (Nginx)
- [ ] Implement backup strategy
- [ ] Add health check endpoint
- [ ] Document API with OpenAPI/Swagger
- [ ] Set up CI/CD pipeline

### Production Configuration
```python
# config.py
class ProductionConfig:
    DEBUG = False
    TESTING = False
    SECRET_KEY = os.environ.get('SECRET_KEY')
    MAX_CONTENT_LENGTH = 16 * 1024 * 1024
    SESSION_COOKIE_SECURE = True
    SESSION_COOKIE_HTTPONLY = True
    SESSION_COOKIE_SAMESITE = 'Lax'
```

---

## Monitoring & Logging

### Add Logging
```python
import logging
logging.basicConfig(
    filename='app.log',
    level=logging.INFO,
    format='%(asctime)s %(levelname)s: %(message)s'
)

# Log uploads
logger.info(f'File uploaded: {filename} ({file_size} bytes)')
```

### Performance Monitoring
```python
import time

def timed_operation(func):
    def wrapper(*args, **kwargs):
        start = time.time()
        result = func(*args, **kwargs)
        duration = time.time() - start
        logger.info(f'{func.__name__} took {duration:.2f}s')
        return result
    return wrapper

@timed_operation
def process_document(filepath):
    # Processing code
```

---

## Contact & Resources

### Internal Resources
- **Source Code**: `C:\Users\knoxh\Documents\tailor\Simplify-Docx\`
- **Original Library**: [microsoft/Simplify-Docx](https://github.com/microsoft/Simplify-Docx)
- **Team Lead**: [Contact Information]

### External Documentation
- [Flask Documentation](https://flask.palletsprojects.com/)
- [Mammoth.js](https://github.com/mwilliamson/mammoth.js)
- [python-docx](https://python-docx.readthedocs.io/)
- [Poppler](https://poppler.freedesktop.org/)

### Version History
- **v1.0.0**: Initial release with live editing
- **Future**: WebSocket support, collaborative editing

---

*Last Updated: October 2025*
*Document Version: 1.0.0*
*Maintainer: Development Team*