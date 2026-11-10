"""
renderer.py - Render shuffled exam data back to a .docx file.

Format:
- Times New Roman, 12pt body
- Page margins: top=2cm, bottom=2cm, left=2.5cm, right=2cm
- Header: school name (bold center), exam title (bold center 14pt),
          "Mã đề: 001" (right-aligned bold), duration (center), horizontal rule
- Each question: "Câu N:" bold, then question text normal
- Options A-D: indented 4 spaces, "    A. text"
- Group headers: bold, centered, with separators between groups
"""

from io import BytesIO
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml
import copy


def _add_footer_field(paragraph, fld_name: str, default_val: str = "1"):
    """Thêm dynamic Word field (PAGE hoặc NUMPAGES) chuẩn ISO/IEC 29500 OpenXML."""
    r_begin = OxmlElement("w:r")
    rPr = parse_xml(r'<w:rPr %s><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr>' % nsdecls("w"))
    r_begin.append(rPr)
    fldChar_begin = OxmlElement("w:fldChar")
    fldChar_begin.set(qn("w:fldCharType"), "begin")
    r_begin.append(fldChar_begin)
    paragraph._p.append(r_begin)

    r_instr = OxmlElement("w:r")
    rPr = parse_xml(r'<w:rPr %s><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr>' % nsdecls("w"))
    r_instr.append(rPr)
    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = f" {fld_name} "
    r_instr.append(instrText)
    paragraph._p.append(r_instr)

    r_sep = OxmlElement("w:r")
    rPr = parse_xml(r'<w:rPr %s><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr>' % nsdecls("w"))
    r_sep.append(rPr)
    fldChar_sep = OxmlElement("w:fldChar")
    fldChar_sep.set(qn("w:fldCharType"), "separate")
    r_sep.append(fldChar_sep)
    paragraph._p.append(r_sep)

    r_val = OxmlElement("w:r")
    rPr = parse_xml(r'<w:rPr %s><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr>' % nsdecls("w"))
    r_val.append(rPr)
    t = OxmlElement("w:t")
    t.text = str(default_val)
    r_val.append(t)
    paragraph._p.append(r_val)

    r_end = OxmlElement("w:r")
    rPr = parse_xml(r'<w:rPr %s><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/><w:sz w:val="20"/></w:rPr>' % nsdecls("w"))
    r_end.append(rPr)
    fldChar_end = OxmlElement("w:fldChar")
    fldChar_end.set(qn("w:fldCharType"), "end")
    r_end.append(fldChar_end)
    paragraph._p.append(r_end)


def _set_margins(doc: Document, top: float, bottom: float, left: float, right: float) -> None:
    """Set page margins (in centimetres)."""
    for section in doc.sections:
        section.top_margin = Cm(top)
        section.bottom_margin = Cm(bottom)
        section.left_margin = Cm(left)
        section.right_margin = Cm(right)


def _add_horizontal_rule(doc: Document) -> None:
    """Add a paragraph with a bottom border that acts as a horizontal rule."""
    p = doc.add_paragraph()
    pPr = p._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "6")
    bottom.set(qn("w:space"), "1")
    bottom.set(qn("w:color"), "000000")
    pBdr.append(bottom)
    pPr.append(pBdr)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.space_before = Pt(4)


def _run_font(run, size_pt: float = 12, bold: bool = False, color: RGBColor = None) -> None:
    """Apply font settings to a run."""
    run.font.name = "Times New Roman"
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    if color:
        run.font.color.rgb = color
    # Ensure East-Asian font also uses Times New Roman
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.get_or_add_rFonts()
    rFonts.set(qn("w:eastAsia"), "Times New Roman")


def _add_paragraph(
    doc: Document,
    text: str = "",
    bold: bool = False,
    size: float = 12,
    alignment=WD_ALIGN_PARAGRAPH.LEFT,
    space_before: float = 0,
    space_after: float = 2,
) -> None:
    """Add a simple single-run paragraph."""
    p = doc.add_paragraph()
    p.alignment = alignment
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    if text:
        run = p.add_run(text)
        _run_font(run, size_pt=size, bold=bold)
    return p


def _add_question_paragraph(doc: Document, num: int, text: str) -> None:
    """Add a question paragraph with bold number and normal text."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(1)

    # Bold "Câu N: "
    run_num = p.add_run(f"Câu {num}: ")
    _run_font(run_num, size_pt=12, bold=True)

    # Normal question text
    if text:
        run_text = p.add_run(text)
        _run_font(run_text, size_pt=12, bold=False)


def _add_option_paragraph(doc: Document, letter: str, text: str) -> None:
    """Add an answer option paragraph with 4-space indent."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(1)
    p.paragraph_format.left_indent = Cm(1)

    run = p.add_run(f"{letter}. {text}")
    _run_font(run, size_pt=12, bold=False)


def render_docx(shuffled_data: dict, ma_de: int) -> bytes:
    """
    Render shuffled exam data into a .docx file and return as bytes.

    Args:
        shuffled_data: Output of shuffler.shuffle_exam
        ma_de: Exam code (mã đề) integer

    Returns:
        bytes of the .docx document
    """
    doc = Document()

    # Page margins
    _set_margins(doc, top=1.5, bottom=1.5, left=2.0, right=1.5)

    # Dynamic Footer
    for sec in doc.sections:
        footer = sec.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.text = ""
        p.paragraph_format.tab_stops.add_tab_stop(Cm(17.5), WD_TAB_ALIGNMENT.RIGHT)
        r_made = p.add_run(f"Mã đề {ma_de:03d}")
        r_made.font.name = "Times New Roman"
        r_made.font.size = Pt(10)
        r_tab = p.add_run("\tTrang ")
        r_tab.font.name = "Times New Roman"
        r_tab.font.size = Pt(10)
        _add_footer_field(p, "PAGE", default_val="1")
        r_sep = p.add_run(" / ")
        r_sep.font.name = "Times New Roman"
        r_sep.font.size = Pt(10)
        _add_footer_field(p, "NUMPAGES", default_val="4")

    try:
        update_fields = parse_xml(r'<w:updateFields %s w:val="true"/>' % nsdecls("w"))
        doc.settings.element.append(update_fields)
    except Exception:
        pass

    metadata = shuffled_data.get("metadata", {})
    school = metadata.get("school", "")
    title = metadata.get("title", "ĐỀ KIỂM TRA")
    duration = metadata.get("duration", "")
    subject = metadata.get("subject", "")

    # ---- Header ----
    if school:
        _add_paragraph(
            doc, school,
            bold=True,
            size=12,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            space_before=0,
            space_after=2,
        )

    if title:
        _add_paragraph(
            doc, title,
            bold=True,
            size=14,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            space_before=2,
            space_after=2,
        )

    if subject:
        _add_paragraph(
            doc, f"Môn: {subject}",
            bold=False,
            size=12,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            space_before=0,
            space_after=2,
        )

    # Mã đề (right-aligned, bold)
    p_made = doc.add_paragraph()
    p_made.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_made.paragraph_format.space_before = Pt(0)
    p_made.paragraph_format.space_after = Pt(2)
    run_made = p_made.add_run(f"Mã đề: {ma_de:03d}")
    _run_font(run_made, size_pt=12, bold=True)

    if duration:
        _add_paragraph(
            doc, f"Thời gian làm bài: {duration} phút (không tính thời gian phát đề)",
            bold=False,
            size=12,
            alignment=WD_ALIGN_PARAGRAPH.CENTER,
            space_before=0,
            space_after=4,
        )

    # Horizontal rule separator
    _add_horizontal_rule(doc)

    # ---- Questions ----
    groups = shuffled_data.get("groups", [])
    total_groups = len(groups)

    for g_idx, group in enumerate(groups):
        group_name = group.get("name", "")

        # Group header (only if more than one group or group name is not default)
        if total_groups > 1 or (group_name and group_name.upper() not in ("PHẦN I", "PHAN I")):
            if g_idx > 0:
                # Separator between groups
                _add_paragraph(doc, "", space_before=4, space_after=4)

            p_grp = doc.add_paragraph()
            p_grp.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_grp.paragraph_format.space_before = Pt(6)
            p_grp.paragraph_format.space_after = Pt(4)
            run_grp = p_grp.add_run(group_name.upper())
            _run_font(run_grp, size_pt=12, bold=True)

        for question in group.get("questions", []):
            q_num = question.get("num", 0)
            q_text = question.get("text", "")
            options = question.get("options", {})

            _add_question_paragraph(doc, q_num, q_text)

            for letter in ["A", "B", "C", "D"]:
                opt_text = options.get(letter, "")
                _add_option_paragraph(doc, letter, opt_text)

    # ---- Dòng kết thúc đề thi ----
    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_end.paragraph_format.space_before = Pt(14)
    p_end.paragraph_format.space_after = Pt(12)
    run_end = p_end.add_run("-------------------- HẾT --------------------")
    _run_font(run_end, size_pt=11, bold=True)

    # ---- Serialize to bytes ----
    buf = BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()
