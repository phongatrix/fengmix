"""
ym_core.py - Core engine for parsing, shuffling, rendering, and exporting
exams according to the YoungMix 2025 / ChamThi QM App specification.
"""

import copy
import io
import json
import os
import random
import re
from pathlib import Path

import docx
import openpyxl
import qrcode
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Inches, Pt, RGBColor
from docx.table import Table
from docx.text.paragraph import Paragraph
from PIL import Image

_ASSETS_DIR = Path(__file__).parent


def is_header_table(table_el) -> bool:
    """Kiểm tra bảng có phải là bảng tiêu đề / thông tin thí sinh hay không."""
    try:
        xml_str = table_el.xml if hasattr(table_el, "xml") else str(table_el)
        keywords = ["Họ và tên", "Số báo danh", "Mã đề", "SỞ GD", "TRƯỜNG THPT", "Thời gian làm bài"]
        return any(kw in xml_str for kw in keywords)
    except Exception:
        return False


def _append_element(doc, element):
    """Chèn element XML vào đúng thứ tự văn bản Word trước sectPr."""
    sectPr = doc._body._body.get_or_add_sectPr()
    sectPr.addprevious(element)


def get_full_para_text(p) -> str:
    """Extract full paragraph text including Office Math ML (<m:t>) and tabs."""
    parts = []
    for node in p._p.iter():
        if node.tag.endswith("}t"):
            parts.append(node.text or "")
        elif node.tag.endswith("}tab"):
            parts.append("\t")
    return "".join(parts).strip()


def get_underlined_letters(p) -> set:
    """Find capital letters A-D that have font.underline in paragraph runs."""
    underlined = set()
    for r in p.runs:
        if r.font.underline and r.text.strip():
            for ch in r.text.strip():
                if ch.upper() in "ABCD":
                    underlined.add(ch.upper())
    return underlined


def find_mcq_options(text: str, expected_start: str = "A") -> list:
    """
    Sequentially search for MCQ option markers (A., B., C., D.) to avoid false
    positives from units or math symbols (like '1MB=1024 B.' or 'a=b+5').
    """
    expected = ["A", "B", "C", "D"]
    start_idx = expected.index(expected_start) if expected_start in expected else 0
    matches = []
    curr_target_idx = start_idx
    search_pos = 0

    while curr_target_idx < 4 and search_pos < len(text):
        target_letter = expected[curr_target_idx]
        pattern = rf"(?:^|[\s\t])({target_letter})\.\s*"
        m = re.search(pattern, text[search_pos:])
        if m:
            abs_start = search_pos + m.start()
            abs_end = search_pos + m.end()
            matches.append((target_letter, abs_start, abs_end))
            search_pos = abs_end
            curr_target_idx += 1
        else:
            break

    return matches


def parse_exam_bytes(file_bytes: bytes) -> dict:
    """Parse a .docx exam file into structured sections, questions, and tables."""
    doc = docx.Document(io.BytesIO(file_bytes))
    body = doc._body._element

    part = 0  # 0: header, 1: Part I, 2: Part II, 3: Part III
    header_tables = []
    p1_items, p2_items, p3_items = [], [], []
    p1_header = "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn."
    p2_header = "PHẦN II. Câu trắc nghiệm đúng sai."
    p3_header = "PHẦN III. Câu trắc nghiệm trả lời ngắn."

    for child in body:
        if child.tag.endswith("p"):
            p = Paragraph(child, doc)
            t = get_full_para_text(p)
            if not t or t.startswith("<g"):
                continue
            if re.search(r"PHẦN\s+I\b", t, re.IGNORECASE):
                part = 1
                p1_header = t
                continue
            elif re.search(r"PHẦN\s+II\b", t, re.IGNORECASE):
                part = 2
                p2_header = t
                continue
            elif re.search(r"PHẦN\s+III\b", t, re.IGNORECASE):
                part = 3
                p3_header = t
                continue

            if part == 1:
                p1_items.append(("p", p))
            elif part == 2:
                p2_items.append(("p", p))
            elif part == 3:
                p3_items.append(("p", p))
        elif child.tag.endswith("tbl"):
            t = Table(child, doc)
            if is_header_table(child):
                header_tables.append(copy.deepcopy(child))
                continue
            if part == 1:
                p1_items.append(("tbl", t))
            elif part == 2:
                p2_items.append(("tbl", t))
            elif part == 3:
                p3_items.append(("tbl", t))

    # Parse P1
    q1 = []
    curr = None
    for el_type, item in p1_items:
        if el_type == "p":
            p = item
            text = get_full_para_text(p)
            m = re.match(r"^Câu\s+(\d+)[:.]\s*(.*)", text)
            if m:
                if curr:
                    q1.append(curr)
                curr = {
                    "num": int(m.group(1)),
                    "text": m.group(2).strip(),
                    "options": {},
                    "correct": "A",
                    "tables": [],
                }
                continue
            if curr:
                next_needed = "A"
                if curr["options"]:
                    keys = ["A", "B", "C", "D"]
                    last_key = list(curr["options"].keys())[-1]
                    if last_key in keys and keys.index(last_key) < 3:
                        next_needed = keys[keys.index(last_key) + 1]
                    else:
                        next_needed = None
                if next_needed:
                    opt_matches = find_mcq_options(text, expected_start=next_needed)
                    if opt_matches:
                        underlined = get_underlined_letters(p)
                        for i, om in enumerate(opt_matches):
                            letter = om[0]
                            s_idx = om[2]
                            e_idx = opt_matches[i + 1][1] if i + 1 < len(opt_matches) else len(text)
                            opt_txt = text[s_idx:e_idx].strip()
                            hm = re.search(r"[\[\(][Đđ]áp\s+án\s*:\s*([A-D])[\]\)]", opt_txt)
                            if hm:
                                curr["correct"] = hm.group(1).upper()
                                opt_txt = re.sub(r"[\[\(][Đđ]áp\s+án\s*:\s*[A-D][\]\)]", "", opt_txt).strip()
                            curr["options"][letter] = opt_txt
                            if letter in underlined:
                                curr["correct"] = letter
                        continue
                hm = re.search(r"[\[\(][Đđ]áp\s+án\s*:\s*([A-D])[\]\)]", text)
                if hm:
                    curr["correct"] = hm.group(1).upper()
                elif not curr["options"]:
                    curr["text"] += " " + text
        elif el_type == "tbl" and curr:
            curr["tables"].append(copy.deepcopy(item._element))
    if curr:
        q1.append(curr)

    # Parse P2
    q2 = []
    curr = None
    for el_type, item in p2_items:
        if el_type == "p":
            p = item
            text = get_full_para_text(p)
            m = re.match(r"^Câu\s+(\d+)[:.]\s*(.*)", text)
            if m:
                if curr:
                    q2.append(curr)
                curr = {
                    "num": int(m.group(1)),
                    "intro": [m.group(2).strip()] if m.group(2).strip() else [],
                    "tables": [],
                    "statements": [],
                }
                continue
            sm = re.match(r"^([a-d])\)\s*(.*)", text)
            if sm and curr:
                lbl = sm.group(1).lower()
                st_txt = sm.group(2).strip()
                is_true = any(r.font.underline and lbl in r.text.lower() for r in p.runs)
                curr["statements"].append(
                    {
                        "label": lbl,
                        "text": st_txt,
                        "is_true": is_true,
                        "val": "D" if is_true else "S",
                    }
                )
            elif curr:
                curr["intro"].append(text)
        elif el_type == "tbl" and curr:
            curr["tables"].append(copy.deepcopy(item._element))
    if curr:
        q2.append(curr)

    # Parse P3
    q3 = []
    curr = None
    for el_type, item in p3_items:
        if el_type == "p":
            p = item
            text = get_full_para_text(p)
            m = re.match(r"^Câu\s+(\d+)[:.]\s*(.*)", text)
            if m:
                if curr:
                    q3.append(curr)
                curr = {
                    "num": int(m.group(1)),
                    "text_lines": [m.group(2).strip()] if m.group(2).strip() else [],
                    "tables": [],
                    "answer": None,
                }
                continue
            am = re.match(r"^(?:[A-D]\.\s*|\[?Đ[áa]p\s*án\s*:\s*)(.+?)(?:\]|\))?$", text, re.IGNORECASE)
            if am and curr:
                curr["answer"] = am.group(1).strip()
            elif curr and not curr["answer"] and re.match(r"^\d+(?:[.,]\d+)?$", text):
                curr["answer"] = text
            elif curr:
                curr["text_lines"].append(text)
        elif el_type == "tbl" and curr:
            curr["tables"].append(copy.deepcopy(item._element))
    if curr:
        q3.append(curr)

    # Fallback options for P1
    for q in q1:
        for letter in "ABCD":
            q["options"].setdefault(letter, "")

    return {
        "header_tables": header_tables,
        "p1_header": p1_header,
        "p1": q1,
        "p2_header": p2_header,
        "p2": q2,
        "p3_header": p3_header,
        "p3": q3,
    }


def apply_modified_data(
    ym_parsed: dict,
    modified_answers: dict = None,
    modified_questions: dict = None,
) -> dict:
    """
    Apply user-modified answers and question contents from the preview UI
    into the parsed exam structure before shuffling and exporting.

    modified_answers schema:
    {
        "p1": {"1": "B", ...},
        "p2": {"1": {"a": "D", "b": "S", ...}, ...},
        "p3": {"1": "15", ...}
    }

    modified_questions schema:
    {
        "p1": {"1": {"text": "...", "options": {"A": "...", ...}}, ...},
        "p2": {"1": {"text": "...", "statements": {"a": "...", ...}}, ...},
        "p3": {"1": {"text": "..."}, ...}
    }
    """
    if not ym_parsed or not isinstance(ym_parsed, dict):
        return ym_parsed

    # Merge questions if questions was embedded inside modified_answers
    if modified_answers and isinstance(modified_answers, dict):
        if "questions" in modified_answers and not modified_questions:
            modified_questions = modified_answers.get("questions")

    p1_ans = (modified_answers or {}).get("p1", {}) if isinstance(modified_answers, dict) else {}
    p2_ans = (modified_answers or {}).get("p2", {}) if isinstance(modified_answers, dict) else {}
    p3_ans = (modified_answers or {}).get("p3", {}) if isinstance(modified_answers, dict) else {}

    p1_q = (modified_questions or {}).get("p1", {}) if isinstance(modified_questions, dict) else {}
    p2_q = (modified_questions or {}).get("p2", {}) if isinstance(modified_questions, dict) else {}
    p3_q = (modified_questions or {}).get("p3", {}) if isinstance(modified_questions, dict) else {}

    # --- Part 1 (MCQ) ---
    for q in ym_parsed.get("p1", []):
        q_str = str(q.get("num", ""))

        # 1. Answer key update
        if q_str in p1_ans:
            val = p1_ans[q_str]
            if isinstance(val, dict):
                if "correct" in val:
                    new_correct = str(val["correct"]).upper().strip()
                    if new_correct in ["A", "B", "C", "D"]:
                        q["correct"] = new_correct
                if "text" in val and val["text"] is not None:
                    q["text"] = str(val["text"]).strip()
                if "options" in val and isinstance(val["options"], dict):
                    for k, txt in val["options"].items():
                        k_upper = str(k).upper().strip()
                        if k_upper in q.get("options", {}):
                            q["options"][k_upper] = str(txt)
            else:
                new_correct = str(val).upper().strip()
                if new_correct in ["A", "B", "C", "D"]:
                    q["correct"] = new_correct

        # 2. Question text / options update
        if q_str in p1_q:
            q_mod = p1_q[q_str]
            if isinstance(q_mod, dict):
                if "text" in q_mod and q_mod["text"] is not None:
                    q["text"] = str(q_mod["text"]).strip()
                if "options" in q_mod and isinstance(q_mod["options"], dict):
                    for k, txt in q_mod["options"].items():
                        k_upper = str(k).upper().strip()
                        if k_upper in q.get("options", {}):
                            q["options"][k_upper] = str(txt)
            elif isinstance(q_mod, str) and q_mod.strip():
                q["text"] = q_mod.strip()

    # --- Part 2 (True/False) ---
    for q in ym_parsed.get("p2", []):
        q_str = str(q.get("num", ""))

        # 1. Answer key update
        if q_str in p2_ans:
            mod_stmts = p2_ans[q_str]
            if isinstance(mod_stmts, dict):
                for stmt in q.get("statements", []):
                    lbl = stmt["label"].lower()
                    if lbl in mod_stmts:
                        val_item = mod_stmts[lbl]
                        val_str = (
                            val_item.get("val") if isinstance(val_item, dict) else str(val_item)
                        ).strip().upper()
                        is_true = val_str in ["D", "Đ", "TRUE", "T", "1"]
                        stmt["is_true"] = is_true
                        stmt["val"] = "D" if is_true else "S"
                        if isinstance(val_item, dict) and "text" in val_item:
                            stmt["text"] = str(val_item["text"])

        # 2. Question intro / statement text update
        if q_str in p2_q:
            q_mod = p2_q[q_str]
            if isinstance(q_mod, dict):
                if "text" in q_mod and q_mod["text"] is not None:
                    new_text = str(q_mod["text"]).strip()
                    if q.get("intro"):
                        q["intro"][0] = new_text
                    else:
                        q["intro"] = [new_text]
                # Allow 'statements' or 'options' dict
                st_mods = q_mod.get("statements") or q_mod.get("options")
                if isinstance(st_mods, dict):
                    for stmt in q.get("statements", []):
                        lbl = stmt["label"].lower()
                        if lbl in st_mods:
                            item = st_mods[lbl]
                            stmt["text"] = (item.get("text") if isinstance(item, dict) else str(item)).strip()
            elif isinstance(q_mod, str) and q_mod.strip():
                new_text = q_mod.strip()
                if q.get("intro"):
                    q["intro"][0] = new_text
                else:
                    q["intro"] = [new_text]

    # --- Part 3 (Short answer) ---
    for q in ym_parsed.get("p3", []):
        q_str = str(q.get("num", ""))

        # 1. Answer update
        if q_str in p3_ans:
            val = p3_ans[q_str]
            if isinstance(val, dict):
                if "answer" in val:
                    q["answer"] = str(val["answer"]).strip()
                if "text" in val and val["text"] is not None:
                    new_text = str(val["text"]).strip()
                    if q.get("text_lines"):
                        q["text_lines"][0] = new_text
                    else:
                        q["text_lines"] = [new_text]
            else:
                q["answer"] = str(val).strip()

        # 2. Question text update
        if q_str in p3_q:
            q_mod = p3_q[q_str]
            if isinstance(q_mod, dict):
                if "text" in q_mod and q_mod["text"] is not None:
                    new_text = str(q_mod["text"]).strip()
                    if q.get("text_lines"):
                        q["text_lines"][0] = new_text
                    else:
                        q["text_lines"] = [new_text]
            elif isinstance(q_mod, str) and q_mod.strip():
                new_text = q_mod.strip()
                if q.get("text_lines"):
                    q["text_lines"][0] = new_text
                else:
                    q["text_lines"] = [new_text]

    return ym_parsed


def apply_modified_answers(ym_parsed: dict, modified_answers: dict) -> dict:
    """
    Backward-compatible wrapper for apply_modified_data.
    """
    return apply_modified_data(ym_parsed, modified_answers=modified_answers)


def shuffle_exam(parsed: dict, ma_de: int, shuffle_mode: str = "ca_cau_va_dap_an") -> tuple:
    """Shuffle questions and choices deterministically based on ma_de."""
    rnd = random.Random(ma_de * 31337 + 42)

    p1_orig = parsed["p1"]
    p2_orig = parsed["p2"]
    p3_orig = parsed["p3"]

    # --- Part 1 ---
    p1_s = [copy.deepcopy(q) for q in p1_orig]
    if shuffle_mode in ["ca_cau_va_dap_an", "chi_cau_hoi", "theo_nhom"]:
        rnd.shuffle(p1_s)

    if shuffle_mode in ["ca_cau_va_dap_an", "chi_dap_an", "theo_nhom"]:
        for q in p1_s:
            opts = q["options"]
            keys = ["A", "B", "C", "D"]
            vals = [opts[k] for k in keys]
            perm = list(range(4))
            rnd.shuffle(perm)
            new_opts = {keys[i]: vals[perm[i]] for i in range(4)}
            correct_idx = keys.index(q["correct"]) if q["correct"] in keys else 0
            new_correct_idx = perm.index(correct_idx)
            q["options"] = new_opts
            q["correct"] = keys[new_correct_idx]

    # --- Part 2 ---
    p2_s = [copy.deepcopy(q) for q in p2_orig]
    if shuffle_mode in ["ca_cau_va_dap_an", "chi_cau_hoi", "theo_nhom"]:
        rnd.shuffle(p2_s)

    if shuffle_mode in ["ca_cau_va_dap_an", "chi_dap_an", "theo_nhom"]:
        for q in p2_s:
            stmts = list(q["statements"])
            rnd.shuffle(stmts)
            labels = ["a", "b", "c", "d"]
            new_stmts = []
            for i, s in enumerate(stmts):
                new_stmts.append({
                    "label": labels[i] if i < 4 else f"st{i}",
                    "text": s["text"],
                    "is_true": s["is_true"],
                    "val": s["val"],
                })
            q["statements"] = new_stmts

    # --- Part 3 ---
    p3_s = [copy.deepcopy(q) for q in p3_orig]
    if shuffle_mode in ["ca_cau_va_dap_an", "chi_cau_hoi", "theo_nhom"]:
        rnd.shuffle(p3_s)

    return p1_s, p2_s, p3_s


def add_footer_field(paragraph, fld_name: str, default_val: str = "1"):
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


def render_docx(p1_list: list, p2_list: list, p3_list: list, ma_de_str: str, headers_info: dict = None) -> bytes:
    """Render exam data to docx bytes matching official YoungMix layout."""
    doc = docx.Document()

    # Calculate total questions to estimate default pages
    total_q = len(p1_list) + len(p2_list) + len(p3_list)
    est_pages = 4 if total_q >= 20 else (3 if total_q >= 10 else 2)

    # Margins and Dynamic Footer (Mã đề [left] ... Trang [PAGE] / [NUMPAGES] [right])
    for sec in doc.sections:
        sec.top_margin = Cm(1.5)
        sec.bottom_margin = Cm(1.5)
        sec.left_margin = Cm(2.0)
        sec.right_margin = Cm(1.5)
        footer = sec.footer
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.text = ""
        
        # Tab stop right aligned at 17.5cm (21cm - 2.0cm left - 1.5cm right)
        p.paragraph_format.tab_stops.add_tab_stop(Cm(17.5), WD_TAB_ALIGNMENT.RIGHT)
        
        r_made = p.add_run(f"Mã đề {ma_de_str}")
        r_made.font.name = "Times New Roman"
        r_made.font.size = Pt(10)
        
        r_tab = p.add_run("\tTrang ")
        r_tab.font.name = "Times New Roman"
        r_tab.font.size = Pt(10)
        
        # Dynamic PAGE field
        add_footer_field(p, "PAGE", default_val="1")
        
        r_sep = p.add_run(" / ")
        r_sep.font.name = "Times New Roman"
        r_sep.font.size = Pt(10)
        
        # Dynamic NUMPAGES field
        add_footer_field(p, "NUMPAGES", default_val=str(est_pages))

    # Enable updateFields in settings.xml
    try:
        update_fields = parse_xml(r'<w:updateFields %s w:val="true"/>' % nsdecls("w"))
        doc.settings.element.append(update_fields)
    except Exception:
        pass

    # Header tables (đặt ở đầu trang 1 của tài liệu)
    header_tables = (headers_info.get("header_tables") if headers_info else None) or []
    if not header_tables:
        template_path = _ASSETS_DIR.parent / "samples" / "De_goc.docx"
        if template_path.exists():
            ref_doc = docx.Document(template_path)
            header_tables = [copy.deepcopy(t._element) for t in ref_doc.tables if is_header_table(t._element)]

    if header_tables:
        for tbl_el in header_tables:
            tbl_copy = copy.deepcopy(tbl_el)
            tbl_obj = Table(tbl_copy, doc)
            for r in tbl_obj.rows:
                for c in r.cells:
                    if "Mã đề" in c.text:
                        for p in c.paragraphs:
                            p.text = f"Mã đề {ma_de_str}"
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            for run in p.runs:
                                run.bold = True
                                run.font.name = "Times New Roman"
            _append_element(doc, tbl_copy)
    else:
        tbl = doc.add_table(rows=1, cols=2)
        tbl.rows[0].cells[0].paragraphs[0].text = "TRƯỜNG THPT\n--------------------"
        tbl.rows[0].cells[1].paragraphs[0].text = f"ĐỀ THI KIỂM TRA\nMã đề: {ma_de_str}"

    p_sep = doc.add_paragraph()
    p_sep.paragraph_format.space_after = Pt(4)

    # --- Part I ---
    if p1_list:
        p_h1 = doc.add_paragraph()
        r = p_h1.add_run(
            headers_info.get("p1_header", "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.")
            if headers_info
            else "PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn."
        )
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

        for i, q in enumerate(p1_list, start=1):
            p_q = doc.add_paragraph()
            p_q.paragraph_format.space_before = Pt(3)
            p_q.paragraph_format.space_after = Pt(2)
            r_num = p_q.add_run(f"Câu {i}. ")
            r_num.bold = True
            r_num.font.name = "Times New Roman"
            r_num.font.size = Pt(11)
            r_txt = p_q.add_run(q["text"])
            r_txt.font.name = "Times New Roman"
            r_txt.font.size = Pt(11)

            # Options
            opts = q.get("options", {})
            keys = ["A", "B", "C", "D"]
            texts = [f"{k}. {opts.get(k, '')}" for k in keys]
            total_len = sum(len(t) for t in texts)

            if total_len <= 75 and all(len(t) < 22 for t in texts):
                p_opt = doc.add_paragraph()
                p_opt.paragraph_format.space_after = Pt(2)
                p_opt.paragraph_format.left_indent = Inches(0.2)
                for j, t in enumerate(texts):
                    r = p_opt.add_run(t + ("\t" if j < 3 else ""))
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(11)
            elif total_len <= 140:
                p_opt1 = doc.add_paragraph()
                p_opt1.paragraph_format.space_after = Pt(1)
                p_opt1.paragraph_format.left_indent = Inches(0.2)
                r1 = p_opt1.add_run(texts[0] + "\t" + texts[1])
                r1.font.name = "Times New Roman"
                r1.font.size = Pt(11)

                p_opt2 = doc.add_paragraph()
                p_opt2.paragraph_format.space_after = Pt(2)
                p_opt2.paragraph_format.left_indent = Inches(0.2)
                r2 = p_opt2.add_run(texts[2] + "\t" + texts[3])
                r2.font.name = "Times New Roman"
                r2.font.size = Pt(11)
            else:
                for t in texts:
                    p_opt = doc.add_paragraph()
                    p_opt.paragraph_format.space_after = Pt(1)
                    p_opt.paragraph_format.left_indent = Inches(0.2)
                    r = p_opt.add_run(t)
                    r.font.name = "Times New Roman"
                    r.font.size = Pt(11)

            for tbl_el in q.get("tables", []):
                _append_element(doc, copy.deepcopy(tbl_el))
                p_space = doc.add_paragraph()
                p_space.paragraph_format.space_after = Pt(2)

    # --- Part II ---
    if p2_list:
        p_h2 = doc.add_paragraph()
        p_h2.paragraph_format.space_before = Pt(8)
        r = p_h2.add_run(
            headers_info.get("p2_header", "PHẦN II. Câu trắc nghiệm đúng sai.")
            if headers_info
            else "PHẦN II. Câu trắc nghiệm đúng sai."
        )
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

        for i, q in enumerate(p2_list, start=1):
            p_q = doc.add_paragraph()
            p_q.paragraph_format.space_before = Pt(3)
            p_q.paragraph_format.space_after = Pt(2)
            r_num = p_q.add_run(f"Câu {i}. ")
            r_num.bold = True
            r_num.font.name = "Times New Roman"
            r_num.font.size = Pt(11)
            if q.get("intro"):
                r_txt = p_q.add_run(q["intro"][0])
                r_txt.font.name = "Times New Roman"
                r_txt.font.size = Pt(11)

            for line in q.get("intro", [])[1:]:
                p_line = doc.add_paragraph()
                p_line.paragraph_format.space_after = Pt(2)
                r = p_line.add_run(line)
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)

            for tbl_el in q.get("tables", []):
                _append_element(doc, copy.deepcopy(tbl_el))
                p_space = doc.add_paragraph()
                p_space.paragraph_format.space_after = Pt(2)

            for st in q.get("statements", []):
                p_st = doc.add_paragraph()
                p_st.paragraph_format.space_after = Pt(1)
                p_st.paragraph_format.left_indent = Inches(0.2)
                r_lbl = p_st.add_run(f"{st['label']}) ")
                r_lbl.bold = True
                r_lbl.font.name = "Times New Roman"
                r_lbl.font.size = Pt(11)
                r_txt = p_st.add_run(st["text"])
                r_txt.font.name = "Times New Roman"
                r_txt.font.size = Pt(11)

    # --- Part III ---
    if p3_list:
        p_h3 = doc.add_paragraph()
        p_h3.paragraph_format.space_before = Pt(8)
        r = p_h3.add_run(
            headers_info.get("p3_header", "PHẦN III. Câu trắc nghiệm trả lời ngắn.")
            if headers_info
            else "PHẦN III. Câu trắc nghiệm trả lời ngắn."
        )
        r.bold = True
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

        for i, q in enumerate(p3_list, start=1):
            p_q = doc.add_paragraph()
            p_q.paragraph_format.space_before = Pt(3)
            p_q.paragraph_format.space_after = Pt(2)
            r_num = p_q.add_run(f"Câu {i}. ")
            r_num.bold = True
            r_num.font.name = "Times New Roman"
            r_num.font.size = Pt(11)
            if q.get("text_lines"):
                r_txt = p_q.add_run(q["text_lines"][0])
                r_txt.font.name = "Times New Roman"
                r_txt.font.size = Pt(11)

            for line in q.get("text_lines", [])[1:]:
                p_line = doc.add_paragraph()
                p_line.paragraph_format.space_after = Pt(2)
                r = p_line.add_run(line)
                r.font.name = "Times New Roman"
                r.font.size = Pt(11)

            for tbl_el in q.get("tables", []):
                _append_element(doc, copy.deepcopy(tbl_el))
                p_space = doc.add_paragraph()
                p_space.paragraph_format.space_after = Pt(2)

    # Dòng kết thúc đề thi ở cuối trang, in đậm và căn giữa
    p_end = doc.add_paragraph()
    p_end.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_end.paragraph_format.space_before = Pt(14)
    p_end.paragraph_format.space_after = Pt(12)
    r_end = p_end.add_run("-------------------- HẾT --------------------")
    r_end.bold = True
    r_end.font.name = "Times New Roman"
    r_end.font.size = Pt(11)

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf.read()


def export_excel_ym(all_exam_data: dict, exam_codes: list) -> bytes:
    """Generate Excel workbook matching official YoungMix format."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # Header row: 'Đề\câu', 1..40, 1a..8d, 1..8
    headers = ["Đề\\câu"]
    headers += [str(i) for i in range(1, 41)]
    for i in range(1, 9):
        for lbl in ["a", "b", "c", "d"]:
            headers.append(f"{i}{lbl}")
    headers += [str(i) for i in range(1, 9)]

    ws.append(headers)

    # All codes, with 0 (đề gốc) first if present
    codes_order = ([0] if 0 in all_exam_data else []) + [c for c in exam_codes if c != 0]

    for code in codes_order:
        code_str = "000" if code == 0 else str(code)
        p1_s, p2_s, p3_s = all_exam_data[code]
        row = [code_str]

        # Part 1 (cols 1..40)
        for i in range(1, 41):
            if i <= len(p1_s):
                row.append(p1_s[i - 1].get("correct", ""))
            else:
                row.append(None)

        # Part 2 (cols 1a..8d)
        for q_idx in range(1, 9):
            if q_idx <= len(p2_s):
                q = p2_s[q_idx - 1]
                stmt_dict = {s["label"]: s["val"] for s in q.get("statements", [])}
                for lbl in ["a", "b", "c", "d"]:
                    row.append(stmt_dict.get(lbl, None))
            else:
                for _ in range(4):
                    row.append(None)

        # Part 3 (cols 1..8)
        for i in range(1, 9):
            if i <= len(p3_s):
                row.append(p3_s[i - 1].get("answer", None))
            else:
                row.append(None)

        ws.append(row)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def export_qr_code(ma_de: int, p1_list: list, p2_list: list, p3_list: list) -> bytes:
    """Generate QR PNG bytes for an exam code conforming to QM App schema."""
    p1_dict = {str(i): q["correct"] for i, q in enumerate(p1_list, 1)}
    p2_dict = {}
    for i, q in enumerate(p2_list, 1):
        for s in q.get("statements", []):
            p2_dict[f"({i},{s['label']})"] = s["val"]
    p3_dict = {str(i): q["answer"] for i, q in enumerate(p3_list, 1) if q.get("answer") is not None}

    payload = {
        "đề": "QM",
        "các đề": {
            str(ma_de): {
                "p1": p1_dict,
                "p2": p2_dict,
                "p3": p3_dict,
            }
        },
    }
    data_str = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    qr = qrcode.QRCode(box_size=8, border=4)
    qr.add_data(data_str)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.read()


def parse_full_preview(file_bytes: bytes, filename: str = "de_thi.docx") -> dict:
    """
    Parse a docx file into groups and questions specifically for the
    interactive YoungMix preview workspace (pre-shuffle review).
    Faithfully detects missing answers based on exam part (MCQ underline, True/False, Short answer).
    """
    doc = docx.Document(io.BytesIO(file_bytes))
    groups = []
    curr_group = None
    curr_q = None
    group_idx = 0
    total_q = 0
    warnings = []
    curr_part = 1

    for p in doc.paragraphs:
        t = get_full_para_text(p)
        if not t:
            continue

        # Tag <g...> triggers a new group
        if re.match(r"^<g\d*(?:#\d+)?>", t):
            if curr_q and curr_group:
                curr_group["questions"].append(curr_q)
                curr_q = None
            if curr_group:
                groups.append(curr_group)
            group_idx += 1
            curr_group = {
                "group_num": group_idx,
                "title": "",
                "questions": [],
            }
            continue

        # Section headers
        if re.search(r"^PHẦN\s+I\b", t, re.IGNORECASE):
            curr_part = 1
            if curr_group is None:
                group_idx += 1
                curr_group = {"group_num": group_idx, "title": t, "questions": []}
                continue
            elif not curr_group["title"]:
                curr_group["title"] = t
                continue
        elif re.search(r"^PHẦN\s+II\b", t, re.IGNORECASE):
            curr_part = 2
            if curr_group is None:
                group_idx += 1
                curr_group = {"group_num": group_idx, "title": t, "questions": []}
                continue
            elif not curr_group["title"]:
                curr_group["title"] = t
                continue
        elif re.search(r"^PHẦN\s+III\b", t, re.IGNORECASE):
            curr_part = 3
            if curr_group is None:
                group_idx += 1
                curr_group = {"group_num": group_idx, "title": t, "questions": []}
                continue
            elif not curr_group["title"]:
                curr_group["title"] = t
                continue

        # If group has no title yet
        if curr_group and not curr_group["title"] and not t.startswith("Câu "):
            curr_group["title"] = t
            continue

        # Ensure a group exists
        if curr_group is None:
            group_idx += 1
            curr_group = {"group_num": group_idx, "title": f"Nhóm {group_idx}", "questions": []}

        # Check question: Câu X. or Câu X:
        m = re.match(r"^Câu\s+(\d+)[:.]\s*(.*)", t)
        if m:
            if curr_q:
                # Check previous question warning
                if curr_q["type"] == "mcq" and not any(o["is_correct"] for o in curr_q["options"]):
                    w = f"Câu {curr_q['num']}: Chưa đánh dấu đáp án đúng"
                    curr_q["warning"] = w
                    warnings.append(w)
                elif curr_q["type"] == "sa" and not curr_q["options"]:
                    w = f"Câu {curr_q['num']}: Chưa có đáp án"
                    curr_q["warning"] = w
                    warnings.append(w)
                curr_group["questions"].append(curr_q)

            total_q += 1
            curr_q = {
                "num": int(m.group(1)),
                "text": m.group(2).strip(),
                "options": [],
                "type": "mcq" if curr_part == 1 else ("tf" if curr_part == 2 else "sa"),
                "warning": None,
            }
            continue

        if not curr_q:
            continue

        # Handle MCQ (Part I)
        if curr_q["type"] == "mcq":
            next_needed = "A"
            if curr_q["options"]:
                last_key = curr_q["options"][-1]["key"]
                keys = ["A", "B", "C", "D"]
                if last_key in keys and keys.index(last_key) < 3:
                    next_needed = keys[keys.index(last_key) + 1]
                else:
                    next_needed = None
            if next_needed:
                opt_matches = find_mcq_options(t, expected_start=next_needed)
                if opt_matches:
                    underlined_letters = get_underlined_letters(p)
                    for i, om in enumerate(opt_matches):
                        letter = om[0]
                        s_idx = om[2]
                        e_idx = opt_matches[i + 1][1] if i + 1 < len(opt_matches) else len(t)
                        opt_txt = t[s_idx:e_idx].strip()
                        hm = re.search(r"[\[\(][Đđ]áp\s+án\s*:\s*([A-D])[\]\)]", opt_txt)
                        is_correct = letter in underlined_letters
                        if hm:
                            is_correct = hm.group(1).upper() == letter
                            opt_txt = re.sub(r"[\[\(][Đđ]áp\s+án\s*:\s*[A-D][\]\)]", "", opt_txt).strip()
                        curr_q["options"].append({
                            "key": letter,
                            "text": opt_txt,
                            "is_correct": is_correct,
                        })
                    continue
            if not curr_q["options"]:
                curr_q["text"] += " " + t
            continue

        # Handle True/False (Part II)
        elif curr_q["type"] == "tf":
            sm = re.match(r"^([a-d])\)\s*(.*)", t)
            if sm:
                lbl = sm.group(1).lower()
                stmt_txt = sm.group(2).strip()
                is_u = any(r.font.underline and lbl in r.text.lower() for r in p.runs)
                curr_q["options"].append({
                    "key": lbl,
                    "text": stmt_txt,
                    "is_correct": is_u,
                    "val": "D" if is_u else "S",
                })
                continue
            else:
                if not curr_q["options"]:
                    curr_q["text"] += " " + t
                continue

        # Handle Short Answer (Part III)
        elif curr_q["type"] == "sa":
            am = re.match(r"^(?:[A-D]\.\s*|\[?Đ[áa]p\s*án\s*:\s*)(.+?)(?:\]|\))?$", t, re.IGNORECASE)
            if am:
                curr_q["options"].append({
                    "key": "A",
                    "text": am.group(1).strip(),
                    "is_correct": True,
                })
                continue
            elif not curr_q["options"] and re.match(r"^\d+(?:[.,]\d+)?$", t):
                curr_q["options"].append({
                    "key": "A",
                    "text": t,
                    "is_correct": True,
                })
                continue
            else:
                if not curr_q["options"]:
                    curr_q["text"] += " " + t
                continue

    # Flush last question and group
    if curr_q and curr_group:
        if curr_q["type"] == "mcq" and not any(o["is_correct"] for o in curr_q["options"]):
            w = f"Câu {curr_q['num']}: Chưa đánh dấu đáp án đúng"
            curr_q["warning"] = w
            warnings.append(w)
        elif curr_q["type"] == "sa" and not curr_q["options"]:
            w = f"Câu {curr_q['num']}: Chưa có đáp án"
            curr_q["warning"] = w
            warnings.append(w)
        curr_group["questions"].append(curr_q)
    if curr_group:
        groups.append(curr_group)

    return {
        "status": "ok",
        "filename": filename,
        "total_groups": len(groups),
        "total_questions": total_q,
        "warnings": warnings,
        "groups": groups,
    }
