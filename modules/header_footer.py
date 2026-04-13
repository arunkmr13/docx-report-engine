from docx.shared import Inches, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def apply_header_footer(doc, title, quarter, company_name, prepared_by, header_font_size=10, footer_font_size=10, logo_path=None, logo_size=0.5, page_label="Page no: "):

    for section in doc.sections:

        usable_width = section.page_width - section.left_margin - section.right_margin

        # ---------------- HEADER ----------------
        header = section.header
        p = header.paragraphs[0]._element
        p.getparent().remove(p)

        if logo_path:
            table = header.add_table(rows=1, cols=3, width=usable_width)
            logo_cell  = table.rows[0].cells[0]
            mid_cell   = table.rows[0].cells[1]
            right_cell = table.rows[0].cells[2]

            logo_cell.width  = int(usable_width * 0.12)
            mid_cell.width   = int(usable_width * 0.35)
            right_cell.width = int(usable_width * 0.53)

            logo_para = logo_cell.paragraphs[0]
            run = logo_para.add_run()
            run.add_picture(logo_path, width=Inches(logo_size))

            mid_cell.text = company_name
            right_cell.text = f"{title} | {quarter}"
            right_cell.paragraphs[0].alignment = 2  # RIGHT

            for cell in [mid_cell, right_cell]:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(header_font_size)

        else:
            table = header.add_table(rows=1, cols=2, width=usable_width)
            left_cell  = table.rows[0].cells[0]
            right_cell = table.rows[0].cells[1]

            left_cell.width  = int(usable_width * 0.35)
            right_cell.width = int(usable_width * 0.65)

            left_cell.text = company_name
            right_cell.text = f"{title} | {quarter}"
            right_cell.paragraphs[0].alignment = 2  # RIGHT

            for cell in [left_cell, right_cell]:
                for para in cell.paragraphs:
                    for run in para.runs:
                        run.font.size = Pt(header_font_size)

        # ---------------- FOOTER ----------------
        footer = section.footer
        p = footer.paragraphs[0]._element
        p.getparent().remove(p)

        table = footer.add_table(rows=1, cols=2, width=usable_width)

        left_cell  = table.rows[0].cells[0]
        right_cell = table.rows[0].cells[1]

        left_cell.width  = int(usable_width * 0.35)
        right_cell.width = int(usable_width * 0.65)

        para = left_cell.paragraphs[0]
        run = para.add_run(page_label)

        fldChar1 = OxmlElement('w:fldChar')
        fldChar1.set(qn('w:fldCharType'), 'begin')

        instrText = OxmlElement('w:instrText')
        instrText.text = "PAGE"

        fldChar2 = OxmlElement('w:fldChar')
        fldChar2.set(qn('w:fldCharType'), 'end')

        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)

        right_cell.text = prepared_by
        right_cell.paragraphs[0].alignment = 2  # RIGHT

        for cell in [left_cell, right_cell]:
            for para in cell.paragraphs:
                for run in para.runs:
                    run.font.size = Pt(footer_font_size)