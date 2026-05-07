from docx.shared import Pt


def apply_formatting(doc):
    for para in doc.paragraphs:
        try:
            if para.text.startswith("Section"):
                para.paragraph_format.line_spacing = 1
            else:
                para.paragraph_format.line_spacing = 1.5
            para.paragraph_format.space_after = Pt(8)
        except Exception:
            pass  # Skip paragraphs that don't support formatting (e.g. tables, text boxes)