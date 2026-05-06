import sys
import os
import uuid
import shutil
import time
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


def validate_asset_path(path: str) -> str:
    assets_dir = os.path.abspath("assets")
    full_path = os.path.abspath(path)
    if not full_path.startswith(assets_dir):
        raise HTTPException(
            status_code=400,
            detail="logo_path must be inside the assets/ folder"
        )
    if not os.path.exists(full_path):
        raise HTTPException(
            status_code=400,
            detail=f"Logo file not found: {path}"
        )
    return path


def cleanup_files():
    now = time.time()
    for f in os.listdir("output"):
        if f.endswith(".docx"):
            path = os.path.join("output", f)
            if os.path.getmtime(path) < now - 3600:
                os.remove(path)
    for f in os.listdir("assets"):
        if f == "logo.jpeg":
            continue
        path = os.path.join("assets", f)
        if os.path.getmtime(path) < now - 3600:
            os.remove(path)


class ReportRequest(BaseModel):
    title: str = TITLE
    quarter: str = QUARTER
    author: str = AUTHOR
    company_name: str = COMPANY_NAME
    prepared_by: str = PREPARED_BY
    header_size: int = 12
    footer_size: int = 10
    logo_path: Optional[str] = LOGO_PATH
    logo_size: float = 0.5
    logo_position: str = "left"
    page_label: str = "Page no: "
    watermark: Optional[str] = None
    sections: Optional[dict[str, list[dict]]] = None


@app.post("/upload-logo")
async def upload_logo(file: UploadFile = File(...)):
    try:
        allowed = [".png", ".jpg", ".jpeg"]
        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in allowed:
            raise HTTPException(
                status_code=400,
                detail="Only .png, .jpg, .jpeg allowed"
            )
        save_path = f"assets/{file.filename}"
        with open(save_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        return {"status": "success", "logo_path": save_path}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/cleanup")
def cleanup():
    try:
        cleanup_files()
        return {"status": "success", "message": "Old files cleaned up"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/enhance")
async def enhance_document(
    file: UploadFile = File(...),
    title: str = Form(TITLE),
    quarter: str = Form(QUARTER),
    company_name: str = Form(COMPANY_NAME),
    prepared_by: str = Form(PREPARED_BY),
    logo_path: str = Form(LOGO_PATH),
    logo_size: float = Form(0.5),
    logo_position: str = Form("left"),
    page_label: str = Form("Page no: "),
    watermark: Optional[str] = Form(None)
):
    try:
        # Validate file type
        if not file.filename.endswith(".docx"):
            raise HTTPException(
                status_code=400,
                detail="Only .docx files are supported"
            )

        # Save uploaded file temporarily
        temp_path = f"output/temp_{uuid.uuid4().hex[:8]}.docx"
        with open(temp_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        # Open existing document
        doc = Document(temp_path)

        # Validate logo path
        if logo_path:
            validate_asset_path(logo_path)

        # Apply branding
        apply_header_footer(
            doc,
            title,
            quarter,
            company_name,
            prepared_by,
            logo_path=logo_path,
            logo_size=logo_size,
            logo_position=logo_position,
            page_label=page_label
        )

        apply_page_borders(doc)

        if watermark:
            apply_watermark(doc, watermark)

        # Save enhanced document
        filename = f"output/enhanced_{uuid.uuid4().hex[:8]}.docx"
        doc.save(filename)

        # Clean up temp file
        os.remove(temp_path)

        return FileResponse(
            path=filename,
            filename=os.path.basename(filename),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/generate")
def generate_report(req: ReportRequest):
    try:
        if req.sections and len(req.sections) > 10:
            raise HTTPException(
                status_code=400,
                detail="Maximum 10 sections allowed"
            )

        if req.logo_path:
            validate_asset_path(req.logo_path)

        cleanup_files()

        doc = Document()

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

        return FileResponse(
            path=filename,
            filename=os.path.basename(filename),
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))