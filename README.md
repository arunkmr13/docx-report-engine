# Docx Report Engine

A Python-based Word document generation engine exposed as a REST API using FastAPI. Send a JSON request — get back a fully formatted `.docx` report with dynamic content, custom branding, watermarks, page borders, and more.

---

## Overview

The Docx Report Engine takes a JSON input via a REST API and generates a professionally formatted Word document (`.docx`). It is built on `python-docx` for document generation and `FastAPI` for the API layer.

Every field in the request is optional — sensible defaults are loaded from `config.py` — making it easy to generate a standard report with just `{}` or a fully custom one with a complete payload.

---

## Features

- Dynamic sections — pass any number of sections (up to 10) with custom headings and content
- Logo upload & placement — upload PNG/JPG via `/upload-logo`, control position (left/center/right) and size
- Text watermark — add DRAFT, CONFIDENTIAL, SAMPLE or any custom text diagonally across all pages
- 3-column header — logo, company name, and report title/quarter
- Dynamic footer — custom page label with auto page numbers
- Page borders — professional borders via correctly ordered OOXML XML
- Two-column layout — Section 3 automatically renders in two columns
- Compatibility mode fix — documents open correctly in all Word versions
- Auto cleanup — files older than 1 hour deleted automatically from `output/` and `assets/`
- Security — logo path validation prevents directory traversal, max section limit prevents abuse
- Error handling — all exceptions return clean HTTP error responses via `HTTPException`
- Direct file download — API streams the `.docx` directly, no manual file retrieval needed

---

## Project Structure

```
docx_report_engine/
├── app.py                  # FastAPI application — all endpoints
├── config.py               # All constants and default values
├── content_generator.py    # Builds document sections dynamically
├── requirements.txt        # Python dependencies
├── README.md               # This file
├── .gitignore
├── assets/
│   └── logo.jpeg           # Default logo
├── modules/
│   ├── styles.py           # Global font and heading styles
│   ├── header_footer.py    # Header/footer with logo and page numbers
│   ├── formatting.py       # Line spacing and paragraph spacing
│   ├── sections.py         # Two-column layout via XML
│   ├── borders.py          # Page borders via OOXML
│   └── watermark.py        # Text watermark via VML XML
└── output/
    └── .gitkeep            # Keeps folder tracked in git
```

---

## Tech Stack

| Technology | Version | Purpose |
|---|---|---|
| Python | 3.13 | Core language |
| FastAPI | Latest | REST API framework |
| python-docx | Latest | Word document generation |
| Pydantic | v2 | Request validation and schema |
| uvicorn | Latest | ASGI server |
| lxml | Latest | Direct XML manipulation |
| python-multipart | Latest | File upload support |

---

## Setup & Installation

1. Clone the repository:
```bash
git clone https://github.com/arunkmr13/docx-report-engine.git
cd docx-report-engine
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

`requirements.txt` contents:
```
fastapi
uvicorn
python-docx
pydantic
python-multipart
```

---

## Running the API

```bash
python3 -m uvicorn app:app --reload --port 8001
```

The API will be available at:
- Swagger UI: `http://127.0.0.1:8001/docs`
- OpenAPI JSON: `http://127.0.0.1:8001/openapi.json`

---

## API Endpoints

### `POST /upload-logo`
Upload a logo image to the `assets/` folder.

- Content-Type: `multipart/form-data`
- Accepted formats: `.png`, `.jpg`, `.jpeg`
- Response:
```json
{
  "status": "success",
  "logo_path": "assets/your_logo.png"
}
```

---

### `POST /generate`
Generate a fully formatted `.docx` report.

- Content-Type: `application/json`
- Response: Direct `.docx` file download
- All fields are optional — defaults loaded from `config.py`

---

### `GET /cleanup`
Manually trigger cleanup of files older than 1 hour from `output/` and `assets/`.

- Response:
```json
{
  "status": "success",
  "message": "Old files cleaned up"
}
```

> Note: Cleanup also runs automatically on every `/generate` call.

---

## Security

- Logo path validation — `logo_path` must resolve to inside the `assets/` directory. Any attempt to traverse outside (e.g. `../../etc/passwd`) returns a `400` error.
- Section limit — maximum 10 sections per request to prevent abuse and oversized documents.
- Error handling — all unhandled exceptions are caught and returned as `500` with a detail message, never a raw Python traceback.
- File cleanup — generated files auto-delete after 1 hour, preventing disk exhaustion.

---

## Known Limitations

- No authentication on endpoints — suitable for local/internal use; add API key middleware for production
- Uploaded logos are cleaned up after 1 hour — re-upload needed for recurring use
- Only text-based watermarks supported — image watermarks not yet implemented
- Report output is always a new unique file — no versioning or overwrite support
- Two-column layout is fixed to Section 3 — not configurable per section yet

---

## GitHub

[https://github.com/arunkmr13/docx-report-engine](https://github.com/arunkmr13/docx-report-engine)
