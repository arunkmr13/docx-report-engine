from fastapi import FastAPI
from pydantic import BaseModel
from docx import Document

from modules.header_footer import apply_header_footer
from modules.styles import apply_global_styles
from modules.formatting import apply_formatting
from modules.sections import apply_columns
from modules.borders import apply_page_borders
from content_generator import generate_content
from config import *

app = FastAPI()

class HeaderRequest(BaseModel):
    font_size: int


@app.post("/generate-report")
def generate_report(req: HeaderRequest):

    doc = Document()

    generate_content(doc)
    apply_global_styles(doc)

    apply_header_footer(doc, TITLE, QUARTER, req.font_size)

    apply_formatting(doc)
    apply_columns(doc)
    apply_page_borders(doc)

    output_path = "output/api_report.docx"
    doc.save(output_path)

    return {"message": "Report generated", "file": output_path}