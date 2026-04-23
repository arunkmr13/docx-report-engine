from docx.shared import Inches, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def apply_header_footer(doc, title, quarter, company_name, prepared_by, header_font_size=10, footer_font_size=10, logo_path=None, logo_size=0.5, logo_position="left", page_label="Page no: "):

    for section in doc.sections:

        usable_width = section.page_width - section.left_margin - section.right_margin

        # ---------------- HEADER ----------------
        header = section.header
        p = header.paragraphs[0]._element
        p.getparent().remove(p)

        if logo_path:
            table = header.add_table(rows=1, cols=3, width=usable_width)
            cell_a = table.rows[0].cells[0]
            cell_b = table.rows[0].cells[1]
            cell_c = table.rows[0].cells[2]

            if logo_position == "left":
                cell_a.width = int(usable_width * 0.12)  # logo — small
                cell_b.width = int(usable_width * 0.33)  # company name
                cell_c.width = int(usable_width * 0.55)  # title | quarter

                logo_cell  = cell_a
                mid_cell   = cell_b
                text_cell  = cell_c

                mid_cell.text = company_name
                mid_cell.paragraphs[0].alignment = 0   # LEFT
                text_cell.text = f"{title} | {quarter}"
                text_cell.paragraphs[0].alignment = 2  # RIGHT

            elif logo_position == "center":
                cell_a.width = int(usable_width * 0.35)  # company name
                cell_b.width = int(usable_width * 0.12)  # logo — small
                cell_c.width = int(usable_width * 0.53)  # title | quarter

                logo_cell  = cell_b
                mid_cell   = cell_a
                text_cell  = cell_c

                mid_cell.text = company_name
                mid_cell.paragraphs[0].alignment = 0   # LEFT
                text_cell.text = f"{title} | {quarter}"
                text_cell.paragraphs[0].alignment = 2  # RIGHT

            elif logo_position == "right":
                cell_a.width = int(usable_width * 0.35)  # company name
                cell_b.width = int(usable_width * 0.53)  # title | quarter
                cell_c.width = int(usable_width * 0.12)  # logo — small

                logo_cell  = cell_c
                mid_cell   = cell_a
                text_cell  = cell_b

                mid_cell.text = company_name
                mid_cell.paragraphs[0].alignment = 0   # LEFT
                text_cell.text = f"{title} | {quarter}"
                text_cell.paragraphs[0].alignment = 2  # RIGHT

            else:
                raise ValueError(f"Invalid logo_position '{logo_position}'. Use 'left', 'center', or 'right'.")

            logo_para = logo_cell.paragraphs[0]
            logo_para.alignment = 1  # CENTER within its cell
            run = logo_para.add_run()
            run.add_picture(logo_path, width=Inches(logo_size))

            for cell in [mid_cell, text_cell]:
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