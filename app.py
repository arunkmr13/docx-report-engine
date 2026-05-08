import sys
import os
import uuid
import shutil
import time
import json

sys.path.insert(0, os.path.dirname(__file__))

from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Optional
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from content_generator import generate_content
from modules.styles import apply_global_styles
from modules.header_footer import apply_header_footer
from modules.formatting import apply_formatting
from modules.sections import apply_columns
from modules.borders import apply_page_borders
from modules.watermark import apply_watermark
from config import *

app = FastAPI()

# ── Schema ────────────────────────────────────────────────────────────────────

class ReportConfig(BaseModel):
    title: str = TITLE
    quarter: str = QUARTER
    author: str = AUTHOR
    company_name: str = COMPANY_NAME
    prepared_by: str = PREPARED_BY
    header_size: int = 12
    footer_size: int = 10
    logo_path: Optional[str] = None
    logo_size: float = 0.5
    logo_position: str = "left"
    page_label: str = "Page no: "
    watermark: Optional[str] = None
    extra_fields: Optional[dict] = None
    sections: Optional[dict[str, list[dict]]] = None

    model_config = {
        "json_schema_extra": {
            "example": {
                "title": "BMJ Open Research Summary",
                "quarter": "May 2026",
                "author": "Arun Kumar",
                "company_name": "Molecular Connections",
                "prepared_by": "INT 1331",
                "logo_size": 0.5,
                "logo_position": "left",
                "page_label": "Page no: ",
                "watermark": "SAMPLE",
                "header_size": 12,
                "footer_size": 10
            }
        }
    }

# ── Helpers ───────────────────────────────────────────────────────────────────

def validate_asset_path(path: str) -> str:
    assets_dir = os.path.abspath("assets")
    full_path = os.path.abspath(path)
    if not full_path.startswith(assets_dir):
        raise HTTPException(status_code=400, detail="logo_path must be inside the assets/ folder")
    if not os.path.exists(full_path):
        raise HTTPException(status_code=400, detail=f"Logo file not found: {path}")
    return path


def cleanup_files():
    now = time.time()
    for folder in ["output", "assets"]:
        for f in os.listdir(folder):
            if folder == "assets" and f in ("logo.jpeg", "logo.jpg", "logo.png"):
                continue
            path = os.path.join(folder, f)
            if os.path.isfile(path) and os.path.getmtime(path) < now - 3600:
                os.remove(path)


def save_logo(logo_file: UploadFile) -> str:
    allowed = [".png", ".jpg", ".jpeg"]
    ext = os.path.splitext(logo_file.filename)[1].lower()
    if ext not in allowed:
        raise HTTPException(status_code=400, detail="Logo must be .png, .jpg, or .jpeg")
    save_path = f"assets/{logo_file.filename}"
    with open(save_path, "wb") as buffer:
        shutil.copyfileobj(logo_file.file, buffer)
    return save_path


def build_document(req: ReportConfig, doc: Document) -> str:
    if req.sections and len(req.sections) > 10:
        raise HTTPException(status_code=400, detail="Maximum 10 sections allowed")

    if req.logo_path:
        validate_asset_path(req.logo_path)

    cleanup_files()

    # Fix compatibility mode
    settings = doc.settings.element
    existing = settings.findall(
        './/{http://schemas.openxmlformats.org/wordprocessingml/2006/main}CompatSetting'
    )
    if not existing:
        compat = OxmlElement('w:CompatSetting')
        compat.set(qn('w:name'), 'compatibilityMode')
        compat.set(qn('w:uri'), 'http://schemas.microsoft.com/office/word')
        compat.set(qn('w:val'), '15')
        settings.append(compat)

    if req.sections:
        generate_content(doc, req.sections)

    apply_global_styles(doc)
    apply_header_footer(
        doc,
        req.title,
        req.quarter,
        req.company_name,
        req.prepared_by,
        header_font_size=req.header_size,
        footer_font_size=req.footer_size,
        logo_path=req.logo_path,
        logo_size=req.logo_size,
        logo_position=req.logo_position,
        page_label=req.page_label
    )
    apply_formatting(doc)
    apply_columns(doc)
    apply_page_borders(doc)

    if req.watermark:
        apply_watermark(doc, req.watermark)

    filename = f"output/report_{uuid.uuid4().hex[:8]}.docx"
    doc.core_properties.author = req.author
    doc.core_properties.title = req.title
    doc.save(filename)
    return filename


# ── Endpoints ─────────────────────────────────────────────────────────────────

@app.post(
    "/enhance",
    summary="Enhance Document",
    description="Upload an existing `.docx` and optionally a logo, then apply header, footer, watermark, and styles."
)
async def enhance_document(
    file: UploadFile = File(..., description="Existing .docx file to enhance"),
    logo: Optional[UploadFile] = File(None, description="Logo image for header (.png / .jpg) — optional"),
    config: str = Form(
        ...,
        description="JSON config string",
        example='{"title":"BMJ Open Research Summary","quarter":"May 2026","author":"Arun Kumar","company_name":"Molecular Connections","prepared_by":"INT 1331","logo_size":0.5,"logo_position":"left","page_label":"Page no: ","watermark":"SAMPLE"}'
    )
):
    try:
        # Parse config JSON
        try:
            data = json.loads(config)
            req = ReportConfig(**data)
        except Exception:
            raise HTTPException(
                status_code=422,
                detail="'config' must be a valid JSON string. See endpoint description for fields."
            )

        # Validate docx
        if not file.filename.endswith(".docx"):
            raise HTTPException(status_code=400, detail="Only .docx files are accepted")

        # Handle logo — uploaded file takes priority, then fall back to default
        if logo and logo.filename:
            req.logo_path = save_logo(logo)
        elif req.logo_path is None:
            for default in ["assets/logo.jpeg", "assets/logo.jpg", "assets/logo.png"]:
                if os.path.exists(default):
                    req.logo_path = default
                    break

        # Save uploaded docx temporarily
        tmp_path = f"output/tmp_{uuid.uuid4().hex[:8]}.docx"
        with open(tmp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        doc = Document(tmp_path)
        os.remove(tmp_path)

        filename = build_document(req, doc)

        return FileResponse(
            path=filename,
            filename=os.path.basename(filename),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get(
    "/cleanup",
    summary="Cleanup",
    description="Delete output files and temporary uploads older than 1 hour."
)
def cleanup():
    try:
        cleanup_files()
        return {"status": "success", "message": "Old files cleaned up"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))