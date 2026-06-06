# LinguaDoc Specification (Spec-Driven Development)

## 1. Project Overview
LinguaDoc is a web application that translates documents (.docx, .xlsx, .txt) into a preferred language and produces professional reports in Word (.docx) and/or PowerPoint (.pptx) formats.

The project features a single unified frontend and **two distinct backend implementations**:
- **Python/FastAPI Backend (pp/)**: Full-featured, local or Render-deployed. Supports all file types, custom templates, and has no strict upload limits.
- **Cloudflare Worker Backend (src/)**: Lightweight, edge-deployed. Supports .txt and .docx, limited to 5MB. Includes Workers AI integration for generating structured reports and filling existing .docx form templates.

## 2. Architecture & Data Flow

### Core Pipeline
Regardless of the backend, the core translation pipeline follows a strict, decoupled flow:
1. **Parse**: Extract content from the uploaded file (.docx, .xlsx, .txt) into a unified, format-agnostic **Content Model**.
2. **Translate**: Iterate through the content model and translate every text segment into the target language.
3. **Generate**: Rebuild the translated content model into the requested output formats (.docx and/or .pptx).

### Content Model
The intermediate content model is a dictionary containing a 	itle and a list of locks. Each block has a 	ype (heading, paragraph, 	able) and contains the text or rows of text.
This decoupling ensures parsers, translators, and generators remain independent and extensible.

## 3. Directory Structure
- pp/: Python/FastAPI backend source code.
  - main.py: API routes and application entry point.
  - parsers.py: Readers for .txt, .docx, .xlsx to generate the content model.
  - 	ranslator.py: Translation logic (via deep-translator) and language definitions.
  - generators.py: Builders for .docx and .pptx outputs.
- src/: Cloudflare Worker backend source code (JavaScript).
  - index.js: Worker entry point and routing.
  - parse.js, docx.js, pptx.js, 	ranslate.js, ill.js, eport.js: JS equivalents of the Python pipeline, plus AI features.
- static/: Frontend assets (index.html, logo.svg). Served by both backends.
- scripts/: Utility scripts.
  - monitoring_report.py: Maps Excel data into a .docx template based on heuristics.
  - compare_report.py: Additional reporting script.
  - 	est-*.mjs: Offline tests for Worker logic.

## 4. API Endpoints

### 4.1. GET /api/languages
Returns a JSON object mapping language codes to their localized names.

### 4.2. GET /api/capabilities
Returns a JSON object detailing backend capabilities. This allows the unified frontend to adapt to the running backend.
- Python example: {"xlsx": true, "templates": true, "maxBytes": null, "build": "python"}
- Cloudflare example: {"xlsx": false, "templates": true, "report": true, "fill": true, "maxBytes": 5242880, "build": "cloudflare"}

### 4.3. POST /api/translate
Translates an uploaded document.
- **Payload (multipart/form-data)**:
  - ile: The document to translate.
  - 	arget_lang: Target language code.
  - ormats: Comma-separated list of formats (docx, pptx).
  - docx_template (optional): Custom Word template.
  - pptx_template (optional): Custom PowerPoint template.
  - mode (optional): Set to aw to bypass AI reporting on Cloudflare.
- **Response**: A file download (.docx or .pptx), or a .zip if both formats are requested.

### 4.4. POST /api/fill (Cloudflare Only)
Fills a .docx template using Workers AI.
- **Payload (multipart/form-data)**:
  - ile: Source document.
  - 	emplate: The .docx form to fill.
- **Response**: The filled .docx file.

## 5. Development Guidelines & Constraints
- **Backend Parity**: When adding general translation features, ensure both the Python and JavaScript backends are updated to maintain feature parity where possible.
- **Worker Limits**: Cloudflare Worker code (src/) must respect free-tier constraints (10ms CPU, 50 subrequests). Avoid heavy dependencies; use lightweight libraries like flate for zip manipulation.
- **Translations**: Translation occurs on a per-segment basis. Always preserve the content model's structure.
- **Styling**: Generated documents should adhere to the built-in professional template (covers, styled headings, page numbers) unless a custom template is provided.
- **Dependencies**:
  - Python: astapi, python-docx, python-pptx, deep-translator, openpyxl.
  - JS: Native fetch, flate for zip.
