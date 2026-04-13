import sys
import os
import uuid
import shutil
sys.path.insert(0, os.path.dirname(__file__))

from fastapi import FastAPI, HTTPException, UploadFile, File
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
from config import *

app = FastAPI()

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
    page_label: str = "Page no: "
    sections: Optional[dict[str, str]] = None


@app.post("/upload-logo")
async def upload_logo(file: UploadFile = File(...)):
    try:
        allowed = [".png", ".jpg", ".jpeg"]
        ext = os.path.splitext(file.filename)[1].lower()
        if ext not in allowed:
            raise HTTPException(status_code=400, detail="Only .png, .jpg, .jpeg allowed")

        save_path = f"assets/{file.filename}"
        with open(save_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)

        return {"status": "success", "logo_path": save_path}

    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/generate")
def generate_report(req: ReportRequest):
    try:
        doc = Document()

        # Fix compatibility mode
        compat = OxmlElement('w:CompatSetting')
        compat.set(qn('w:name'), 'compatibilityMode')
        compat.set(qn('w:uri'), 'http://schemas.microsoft.com/office/word')
        compat.set(qn('w:val'), '15')
        doc.settings.element.append(compat)

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
            page_label=req.page_label
            )

        apply_formatting(doc)
        apply_columns(doc)
        apply_page_borders(doc)

        filename = f"output/report_{uuid.uuid4().hex[:8]}.docx"
        doc.core_properties.author = req.author
        doc.core_properties.title = req.title
        doc.save(filename)

        return {"status": "success", "file": filename}

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))