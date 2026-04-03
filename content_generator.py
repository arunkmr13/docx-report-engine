from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def generate_content(doc):
    sections = {
        "Executive Summary": "Overview of key business outcomes, major achievements, and strategic direction.",
        "Market Analysis": "Insights into current market trends, competitor positioning, and growth opportunities.",
        "Department Highlights": "Performance review of each department with key metrics and accomplishments.",
        "Financial Overview": "Breakdown of revenue, expenses, and profitability across business units.",
        "Future Strategy": "Planned initiatives, expansion goals, and strategic priorities for upcoming quarters."
    }

    for idx, sec in enumerate(sections):

        if idx != 0:
            doc.add_page_break()

        # Section Heading
        heading = doc.add_heading(f"Section {idx+1}: {sec}", level=1)

        

        #  Styled heading (BOLD + UNDERLINE)
        para = doc.add_paragraph()
        run = para.add_run(f"{sec.upper()} PARAGRAPH")
        run.bold = True
        run.underline = True

        para.paragraph_format.space_before = Pt(10)
        para.paragraph_format.space_after = Pt(6)

        #  Bullet points (use description)
        doc.add_paragraph(sections[sec], style='List Bullet')