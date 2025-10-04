# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a web-based DOCX converter with live editing capabilities, built on top of Microsoft's Simplify-Docx library. The project combines a Flask backend with a single-page HTML/JavaScript frontend to provide real-time document conversion and editing.

**Two main components:**
1. **Simplify-Docx Library** (`src/simplify_docx/`) - Core library that converts python-docx Document objects to simplified JSON
2. **Web Application** (`app.py` + `templates/index.html`) - Flask app providing DOCX → HTML/JSON conversion with live editing

## Development Commands

### Running the Application
```bash
python app.py
```
Server runs at http://127.0.0.1:5000

### Installing Dependencies
```bash
# Install the simplify-docx library in editable mode
pip install -e .

# Install web app dependencies
pip install flask flask-cors mammoth docx2pdf pdf2image
```

### Optional: Image Preview Support
Requires Microsoft Word (Windows) and Poppler for PDF conversion. See lines 83-115 in app.py:223 - this functionality is currently commented out.

To enable:
- Uncomment the image conversion code block in app.py:83-115
- Install Poppler and ensure it's in the `poppler-24.08.0/Library/bin` directory

### Finding and Killing Processes on Port 5000
```bash
# Windows
netstat -ano | findstr :5000
taskkill /F /PID <process_id>
```

## Architecture

### Request Flow
1. User uploads DOCX via drag-and-drop or file picker
2. `app.py:/upload` endpoint receives file (max 16MB)
3. File is processed through three conversions in parallel:
   - **Mammoth** → HTML rendering (app.py:60-64)
   - **simplify-docx** → Simplified JSON structure (app.py:66-70)
   - **docx2pdf + pdf2image** → Preview image (optional, commented out)
4. Frontend receives all three formats and displays in selected view mode
5. User can edit JSON in live editor, which POSTs to `app.py:/json-to-html`
6. Backend converts JSON back to HTML via `convert_json_to_html()` (app.py:168-219)

### Simplify-Docx Library Architecture

The library follows a visitor pattern to traverse and convert DOCX elements:

- **Entry point**: `simplify()` function in `src/simplify_docx/__init__.py:20`
- **Element classes** (`src/simplify_docx/elements/`): Each DOCX element type (document, body, paragraph, table, run) has a corresponding class with a `to_json()` method
- **Iterators** (`src/simplify_docx/iterators/`): Custom iterators for traversing nested DOCX structures
- **Type mapping**: The `convert_json_to_html()` function in app.py:168 maps simplified types back to HTML tags

Key element types:
- `document` → root container
- `body` → document body
- `paragraph` → text paragraphs (may contain indentation, numbering)
- `table`, `table-row`, `table-cell` → table structures
- `text`/`run` → text content with formatting

### JSON-to-HTML Conversion

The `convert_json_to_html()` function (app.py:168-219) recursively processes the simplified JSON:
- Maps simplified types to HTML tags (paragraph → p, table-row → tr, etc.)
- Handles nested children recursively
- Supports direct text values and text properties

### Configuration Options

Simplify-docx behavior is controlled via options dict (see `__default_options__` in `src/simplify_docx/__init__.py:44-84`):
- Flattening: hyperlinks, smartTags, customXml
- Whitespace handling: merge consecutive text, trim leading/trailing
- Forms: checkboxes, dropdowns, text inputs
- Special characters: smart quotes → dumb quotes, special symbols

## File Structure

- `app.py` - Flask server with upload, conversion, and live editing endpoints
- `templates/index.html` - Single-page app with 8 view modes and live JSON editor
- `src/simplify_docx/` - Core library for DOCX → JSON conversion
  - `elements/` - Element type classes with `to_json()` methods
  - `iterators/` - Custom iterators for traversing DOCX structures
  - `utils/` - Helper functions (friendly names, option handling, walking)
  - `types/` - Type definitions (documentPart fragment)
- `uploads/` - Temporary file storage (auto-created, gitignored)

## Key Implementation Details

### Security
- `secure_filename()` used for all uploads (app.py:53)
- File type validation restricts to `.docx` only (app.py:28)
- Files cleaned up after processing (app.py:118)
- Max upload size: 16MB (app.py:22)

### Live Editing
- JavaScript debounces updates (500ms) to avoid excessive requests
- JSON validation happens client-side before sending to backend
- Backend parses JSON string and converts via recursive HTML generation
- Preview updates in real-time in right panel

### View Modes
8 different comparison modes accessible via tabs:
- All 3 panels, Live JSON Editor, Original vs HTML, Original vs JSON, HTML vs JSON, and individual views

## Common Tasks

### Modifying JSON-to-HTML Conversion
Edit the `convert_json_to_html()` function in app.py:168 and update the `tag_map` dictionary at line 190 to change how simplified types map to HTML tags.

### Adding New Element Types to Simplify-Docx
1. Create element class in `src/simplify_docx/elements/` with `to_json()` method
2. Add corresponding iterator in `src/simplify_docx/iterators/`
3. Update element registry/imports as needed

### Changing Upload Limits
Modify `app.config['MAX_CONTENT_LENGTH']` in app.py:22

### Adjusting Live Preview Debounce
Edit timeout value in `templates/index.html` JavaScript (search for "500ms delay")
