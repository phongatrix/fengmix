"""
exporter.py - Export answer keys to Excel (QM format) and QR code PNG images.
"""

import json
import io
import qrcode
from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side


def export_excel(all_answers: dict, exam_codes: list) -> bytes:
    """
    Export all exam answer keys to an Excel workbook in QM (Quản Môn) format.

    Args:
        all_answers: {ma_de (int): {"answer_key": {q_num (int): letter (str)}}}
        exam_codes:  Ordered list of exam code integers

    Returns:
        bytes of the .xlsx file
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Dap an"

    # ---- Header row ----
    headers = ["Mã đề", "p1", "p2", "p3", "Ghi chú"]
    header_font = Font(name="Times New Roman", bold=True, size=12)
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font_white = Font(name="Times New Roman", bold=True, size=12, color="FFFFFF")

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for col_idx, h in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = header_font_white
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border

    ws.row_dimensions[1].height = 20

    # ---- Data rows ----
    for row_idx, ma_de in enumerate(exam_codes, start=2):
        answer_data = all_answers.get(ma_de, {})
        answer_key: dict = answer_data.get("answer_key", {})

        # Sort by question number and join letters into a single string (p1)
        sorted_nums = sorted(answer_key.keys(), key=lambda x: int(x))
        p1_str = "".join(answer_key[n] for n in sorted_nums)

        row_data = [ma_de, p1_str, "", "", ""]
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = Font(name="Times New Roman", size=11)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        # Alternate row color for readability
        if row_idx % 2 == 0:
            fill = PatternFill("solid", fgColor="DCE6F1")
            for col_idx in range(1, 6):
                ws.cell(row=row_idx, column=col_idx).fill = fill

    # ---- Column widths ----
    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = max(
        30,
        max(
            (len(all_answers.get(c, {}).get("answer_key", {})) for c in exam_codes),
            default=10,
        ) + 2,
    )
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 15
    ws.column_dimensions["E"].width = 20

    # ---- Second sheet: per-question breakdown ----
    ws2 = wb.create_sheet(title="Chi tiet")
    if exam_codes:
        # Determine max question count
        all_q_nums = set()
        for ma_de in exam_codes:
            ak = all_answers.get(ma_de, {}).get("answer_key", {})
            all_q_nums.update(ak.keys())

        sorted_q_nums = sorted(all_q_nums, key=lambda x: int(x))

        # Header row
        header2 = ["Mã đề"] + [f"Câu {n}" for n in sorted_q_nums]
        for col_idx, h in enumerate(header2, start=1):
            cell = ws2.cell(row=1, column=col_idx, value=h)
            cell.font = header_font_white
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        for row_idx, ma_de in enumerate(exam_codes, start=2):
            ak = all_answers.get(ma_de, {}).get("answer_key", {})
            ws2.cell(row=row_idx, column=1, value=ma_de).font = Font(
                name="Times New Roman", bold=True, size=11
            )
            ws2.cell(row=row_idx, column=1).alignment = Alignment(horizontal="center")
            ws2.cell(row=row_idx, column=1).border = border

            for col_idx, q_num in enumerate(sorted_q_nums, start=2):
                letter = ak.get(q_num, "")
                cell = ws2.cell(row=row_idx, column=col_idx, value=letter)
                cell.font = Font(name="Times New Roman", size=11)
                cell.alignment = Alignment(horizontal="center")
                cell.border = border

                if row_idx % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="DCE6F1")

        # Freeze first row
        ws2.freeze_panes = "B2"

    # Freeze first row on sheet 1
    ws.freeze_panes = "A2"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.read()


def export_qr(ma_de: int, answer_key: dict, p2: dict = None, p3: dict = None) -> bytes:
    """
    Generate a QR code PNG for an exam code and its answer key.
    Follows standard ChamThi QM App format:
    {"đề": "QM", "các đề": {de: {"p1": {câu: đáp án}, "p2": {(câu,ý): Đ/S}, "p3": {câu: số}}}}

    Args:
        ma_de: Exam code integer
        answer_key: {q_num (int): letter (str)}
        p2: Optional part 2 True/False answers dict
        p3: Optional part 3 short answer dict

    Returns:
        bytes of the PNG image
    """
    # Build QM-standard QR payload
    payload = {
        "đề": "QM",
        "các đề": {
            str(ma_de): {
                "p1": {str(k): v for k, v in sorted(answer_key.items(), key=lambda x: int(x[0]))},
                "p2": p2 or {},
                "p3": p3 or {},
            }
        },
    }
    data_str = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))

    qr = qrcode.QRCode(
        version=None,  # auto-size
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=8,
        border=4,
    )
    qr.add_data(data_str)
    qr.make(fit=True)

    img: Image.Image = qr.make_image(fill_color="black", back_color="white").convert("RGB")

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.read()
