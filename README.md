# docx-report-engine

> Enhance any `.docx` file with a branded header, footer, watermark, logo, and styles — via a single API call.

![Python](https://img.shields.io/badge/Python-3.11-blue)
![FastAPI](https://img.shields.io/badge/FastAPI-latest-green)
![Docker](https://img.shields.io/badge/Docker-ready-blue)
![License](https://img.shields.io/badge/License-MIT-purple)
![Cost](https://img.shields.io/badge/Cost-Free-brightgreen)

---

## What It Does

Upload any `.docx` file and get back a professionally enhanced document with:

1. **Header** — logo, company name, and title/quarter laid out left-to-right
2. **Footer** — page number on the left, document reference on the right
3. **Watermark** — diagonal or horizontal overlay text (e.g. `DRAFT`, `REVIEW`, `CONFIDENTIAL`)
4. **Styles** — global font and paragraph styles applied cleanly without altering original spacing
5. **Word compatibility** — output targets Office 2013+ (`compatibilityMode=15`)

---

## Why This Exists

Research teams and publishers handle dozens of manuscript revisions. Manually adding consistent headers, footers, and watermarks to every `.docx` is tedious and error-prone. This engine automates it — upload the file, pass a JSON config, download the branded document.

---

## Quick Start

### 1. Clone and install

```bash
git clone https://github.com/arunkmr13/docx-report-engine.git
cd docx-report-engine
pip install -r requirements.txt
```

### 2. Run the server

```bash
uvicorn app:app --reload --port 8000
```

### 3. Open Swagger UI

```
http://127.0.0.1:8000/docs
```

---

## Docker

```bash
docker-compose up --build
```

---

## API Reference

### `POST /enhance`

Enhance an existing `.docx` file with header, footer, watermark, and styles.

**Request** — `multipart/form-data`

| Field  | Type   | Required | Description                        |
|--------|--------|----------|------------------------------------|
| `file` | binary | ✅        | The `.docx` file to enhance        |
| `logo` | binary | ❌        | Logo image (`.png` / `.jpg`)       |
| `req`  | object | ✅        | JSON config (see fields below)     |

**`req` JSON config fields**

| Field                  | Type    | Default      | Description                                      |
|------------------------|---------|--------------|--------------------------------------------------|
| `title`                | string  | —            | Document title shown in header                   |
| `quarter`              | string  | —            | Period label shown in header (e.g. `May 2026`)   |
| `author`               | string  | —            | Document author (saved to `.docx` metadata)      |
| `company_name`         | string  | —            | Organisation name shown in header                |
| `prepared_by`          | string  | —            | Reference shown in footer (right side)           |
| `header_size`          | integer | `12`         | Header font size in pt                           |
| `footer_size`          | integer | `10`         | Footer font size in pt                           |
| `logo_size`            | float   | `0.5`        | Logo width in inches                             |
| `logo_position`        | string  | `"left"`     | Logo position: `"left"`, `"center"`, `"right"`  |
| `page_label`           | string  | `"Page no: "`| Prefix before the page number in footer         |
| `watermark`            | string  | `null`       | Watermark text (e.g. `"DRAFT"`, `"REVIEW"`)     |
| `watermark_orientation`| string  | `"diagonal"` | `"diagonal"` or `"horizontal"`                  |
| `logo_path`            | string  | `null`       | Server-side logo path (must be inside `assets/`)|
| `extra_fields`         | object  | `null`       | Additional key-value pairs in footer row 2      |
| `sections`             | object  | `null`       | Content sections to generate inside the doc     |

**Example `req`**

```json
{
  "title": "Toward an Integrated Model for Dengue Management: A Scoping Review",
  "quarter": "May 2026",
  "author": "Mehrdad Salehi, Ehsan Mousa Farkhani, Javad Moghri et al.",
  "company_name": "Mashhad University of Medical Sciences",
  "prepared_by": "BMJ Open — bmjopen-2025-111920",
  "header_size": 11,
  "footer_size": 9,
  "logo_size": 0.5,
  "logo_position": "left",
  "page_label": "Page no: ",
  "watermark": "REVIEW",
  "watermark_orientation": "diagonal",
  "logo_path": null,
  "extra_fields": null,
  "sections": null
}
```

**Response** — `.docx` file download

---

### `GET /cleanup`

Deletes output files and uploaded assets older than 1 hour.

```bash
curl http://127.0.0.1:8000/cleanup
```

---

## Project Structure

```
docx-report-engine/
├── app.py                  # FastAPI app — endpoints and request parsing
├── config.py               # Default constants (title, author, company etc.)
├── content_generator.py    # Renders section blocks into the document
├── main.py                 # Uvicorn entry point
├── assets/                 # Logo files (logo.jpeg used as default)
├── output/                 # Generated .docx files (auto-cleaned after 1hr)
└── modules/
    ├── styles.py           # Global document styles
    ├── header_footer.py    # Header/footer table layout and alignment
    ├── formatting.py       # Paragraph formatting (preserves original spacing)
    ├── sections.py         # Multi-column layout
    ├── borders.py          # Page borders
    └── watermark.py        # VML watermark overlay with orientation support
```

---

## Notes

- Original document spacing (`line_spacing`, `space_after`) is fully preserved
- `logo_path` must resolve inside the `assets/` folder (path traversal protection)
- Maximum 10 sections allowed per request
- Logo falls back to `assets/logo.jpeg` if no logo is uploaded and `logo_path` is `null`
- Swagger UI (`/docs`) renders the `req` field as a structured JSON editor

---

## License

MIT