from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

DEFAULT_SECTIONS = {
    "Executive Summary": "Overview of key business outcomes, major achievements, and strategic direction.",
    "Market Analysis": "Insights into current market trends, competitor positioning, and growth opportunities.",
    "Department Highlights": "Performance review of each department with key metrics and accomplishments.",
    "Financial Overview": "Breakdown of revenue, expenses, and profitability across business units.",
    "Future Strategy": "Planned initiatives, expansion goals, and strategic priorities for upcoming quarters."
}

def generate_content(doc, sections: dict = None):
    if sections is None:
        sections = DEFAULT_SECTIONS

    for idx, sec in enumerate(sections):

        if idx != 0:
            doc.add_page_break()

        heading = doc.add_heading(f"Section {idx+1}: {sec}", level=1)

        

        doc.add_paragraph(sections[sec], style='List Bullet')