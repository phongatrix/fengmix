"""
main.py - FastAPI application for YoungMix Clone (Trộn đề trắc nghiệm).

Endpoints:
  POST /api/process     - Upload .docx exam, get job_id + download URL
  GET  /api/download/{job_id} - Download the result zip
  GET  /api/health      - Health check

The result zip contains:
  de_001.docx, de_002.docx, ...
  dap_an.xlsx
  qr/qr_001.png, qr/qr_002.png, ...
"""

import io
import json
import os
import random
import re
import uuid
import zipfile
from datetime import datetime
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, JSONResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from parser import parse_docx
from shuffler import shuffle_exam
from renderer import render_docx
from exporter import export_excel, export_qr
import ym_core

# ---------------------------------------------------------------------------
# App setup
# ---------------------------------------------------------------------------

app = FastAPI(
    title="Feng Mix API",
    description="Hệ thống trộn đề thi trắc nghiệm trực tuyến",
    version="1.0.0",
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Serve frontend static files if the directory exists
_FRONTEND_DIR = Path(__file__).parent.parent / "frontend"

# Samples directory
_SAMPLES_DIR = Path(__file__).parent.parent / "samples"

# In-memory job registry: job_id -> {"zip_path": str, "total": int, "status": str}
JOBS: dict[str, dict] = {}

TMP_DIR = Path("/tmp")


# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------

def _zip_path(job_id: str) -> Path:
    return TMP_DIR / f"youngmix_{job_id}.zip"


def _cleanup_old_jobs(keep_latest: int = 50) -> None:
    """Remove old zip files if we have too many stored jobs."""
    if len(JOBS) > keep_latest:
        # Remove oldest entries
        sorted_ids = list(JOBS.keys())
        for old_id in sorted_ids[: len(sorted_ids) - keep_latest]:
            old_path = _zip_path(old_id)
            if old_path.exists():
                old_path.unlink(missing_ok=True)
            JOBS.pop(old_id, None)


# ---------------------------------------------------------------------------
# Routes
# ---------------------------------------------------------------------------

@app.get("/")
async def index():
    """Serve the frontend index.html if it exists, else show API info."""
    index_file = _FRONTEND_DIR / "index.html"
    if index_file.exists():
        return FileResponse(str(index_file))
    return JSONResponse(
        {
            "app": "YoungMix Clone API",
            "version": "1.0.0",
            "docs": "/docs",
            "health": "/api/health",
        }
    )


@app.get("/style.css")
async def serve_css():
    """Serve frontend CSS."""
    f = _FRONTEND_DIR / "style.css"
    return FileResponse(str(f), media_type="text/css") if f.exists() else JSONResponse({}, status_code=404)


@app.get("/app.js")
async def serve_js():
    """Serve frontend JS."""
    f = _FRONTEND_DIR / "app.js"
    return FileResponse(str(f), media_type="application/javascript") if f.exists() else JSONResponse({}, status_code=404)


# ---------------------------------------------------------------------------
# Sample file routes
# ---------------------------------------------------------------------------

@app.get("/api/samples")
async def list_samples():
    """Danh sách các file mẫu để tải về (chỉ gồm 2 file đề tham khảo chuẩn)."""
    if not _SAMPLES_DIR.exists():
        return {"samples": []}
    allowed_samples = {
        "2025_De Tin Hoc Tham Khao YM.docx",
        "2025_De Dia li Tham Khao YM.docx",
    }
    files = []
    for f in sorted(_SAMPLES_DIR.iterdir()):
        if f.is_file() and f.name in allowed_samples:
            files.append({
                "name": f.name,
                "path": f.name,
                "size": f.stat().st_size,
                "download_url": f"/api/samples/download/{f.name}",
            })
    return {"samples": files}


@app.get("/api/samples/download/{file_path:path}")
async def download_sample(file_path: str):
    """Tải về file mẫu theo đường dẫn tương đối."""
    target = (_SAMPLES_DIR / file_path).resolve()
    # Security: prevent path traversal
    try:
        target.relative_to(_SAMPLES_DIR.resolve())
    except ValueError:
        raise HTTPException(status_code=403, detail="Truy cập bị từ chối.")
    if not target.exists() or not target.is_file():
        raise HTTPException(status_code=404, detail=f"Không tìm thấy file mẫu: {file_path}")
    media = "application/octet-stream"
    if target.suffix == ".docx":
        media = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    elif target.suffix == ".xlsx":
        media = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    elif target.suffix == ".pdf":
        media = "application/pdf"
    elif target.suffix == ".zip":
        media = "application/zip"
    elif target.suffix == ".png":
        media = "image/png"
    return FileResponse(
        str(target),
        media_type=media,
        filename=target.name,
        headers={"Content-Disposition": f'attachment; filename="{target.name}"'},
    )


@app.get("/api/health")
async def health():
    """Health check endpoint."""
    return {"status": "ok"}


@app.post("/api/parse")
async def parse_exam_endpoint(
    file: UploadFile = File(..., description="File đề thi .docx"),
):
    """
    Parse file .docx và trả về dữ liệu xem trước (groups, questions, options, checkmarks, warnings).
    Dành cho giao diện cấu hình & xem trước đề trước khi bấm Trộn đề.
    """
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(
            status_code=400,
            detail="Chỉ chấp nhận file .docx. Vui lòng kiểm tra lại file tải lên.",
        )
    try:
        file_bytes = await file.read()
        if not file_bytes:
            raise ValueError("File rỗng.")
        preview_data = ym_core.parse_full_preview(file_bytes, filename=file.filename)
        return JSONResponse(preview_data)
    except Exception as exc:
        raise HTTPException(
            status_code=422,
            detail=f"Không thể phân tích file đề thi: {exc}",
        )


@app.post("/api/process")
async def process_exam(
    file: UploadFile = File(..., description="File đề thi .docx"),
    so_ma_de: int = Form(default=None, ge=1, le=100, description="Số lượng mã đề cần tạo"),
    num_codes: int = Form(default=None, ge=1, le=100, description="Số lượng mã đề (alias)"),
    shuffle_mode: str = Form(
        default="ca_cau_va_dap_an",
        description=(
            "Chế độ trộn: ca_cau_va_dap_an | chi_cau_hoi | chi_dap_an | theo_nhom"
        ),
    ),
    bat_dau_tu: int = Form(default=101, ge=1, le=999, description="Mã đề bắt đầu từ số"),
    random_codes: str = Form(default="false", description="Tạo mã đề ngẫu nhiên (true | false)"),
    modified_answers: str = Form(default=None, description="JSON of user-modified answers"),
    modified_questions: str = Form(default=None, description="JSON of user-modified question texts/options"),
):
    """
    Xử lý file đề thi .docx, trộn thành nhiều mã đề, trả về link tải.

    Returns JSON:
        {
            "job_id": "...",
            "status": "done",
            "total": N,
            "download_url": "/api/download/{job_id}"
        }
    """
    # Determine requested number of exam codes (default 4 if not given)
    effective_count = 4
    if so_ma_de is not None:
        effective_count = so_ma_de
    elif num_codes is not None:
        effective_count = num_codes

    # ---- Validate file ----
    if not file.filename or not file.filename.lower().endswith(".docx"):
        raise HTTPException(
            status_code=400,
            detail="Chỉ chấp nhận file .docx. Vui lòng kiểm tra lại file tải lên.",
        )

    valid_modes = {"ca_cau_va_dap_an", "chi_cau_hoi", "chi_dap_an", "theo_nhom"}
    if shuffle_mode not in valid_modes:
        raise HTTPException(
            status_code=400,
            detail=f"Chế độ trộn không hợp lệ: '{shuffle_mode}'. "
                   f"Chọn một trong: {', '.join(sorted(valid_modes))}.",
        )

    # ---- Read & parse ----
    try:
        file_bytes = await file.read()
        if not file_bytes:
            raise ValueError("File rỗng.")
        ym_parsed = ym_core.parse_exam_bytes(file_bytes)

        # Apply user modifications from interactive preview UI if provided
        if modified_answers or modified_questions:
            try:
                mods_ans_dict = (
                    json.loads(modified_answers)
                    if isinstance(modified_answers, str)
                    else (modified_answers or {})
                )
            except Exception as e:
                print("Lỗi parse modified_answers:", e)
                mods_ans_dict = {}

            try:
                mods_q_dict = (
                    json.loads(modified_questions)
                    if isinstance(modified_questions, str)
                    else (modified_questions or {})
                )
            except Exception as e:
                print("Lỗi parse modified_questions:", e)
                mods_q_dict = {}

            try:
                ym_parsed = ym_core.apply_modified_data(
                    ym_parsed,
                    modified_answers=mods_ans_dict,
                    modified_questions=mods_q_dict,
                )
            except Exception as e:
                print("Lỗi khi áp dụng nội dung/đáp án đã chỉnh sửa:", e)

        total_questions = len(ym_parsed["p1"]) + len(ym_parsed["p2"]) + len(ym_parsed["p3"])
    except Exception as exc:
        raise HTTPException(
            status_code=422,
            detail=f"Không thể phân tích file đề thi: {exc}",
        )

    # Sanity check: must have at least one question
    if total_questions == 0:
        raise HTTPException(
            status_code=422,
            detail=(
                "Không tìm thấy câu hỏi nào trong file. "
                "Hãy kiểm tra định dạng: câu hỏi phải bắt đầu bằng 'Câu X:' hoặc 'Câu X.'"
            ),
        )

    # ---- Generate each exam code ----
    is_random = str(random_codes).strip().lower() in ("true", "1", "yes", "on")
    if is_random:
        all_possible = list(range(101, 1000))
        exam_codes = sorted(random.sample(all_possible, effective_count))
    else:
        exam_codes = list(range(bat_dau_tu, bat_dau_tu + effective_count))
    all_exam_data: dict[int, tuple] = {}
    docx_files: dict[str, bytes] = {}
    qr_files: dict[str, bytes] = {}

    try:
        # Render De_goc (000)
        degoc_bytes = ym_core.render_docx(
            ym_parsed["p1"], ym_parsed["p2"], ym_parsed["p3"], "000", ym_parsed
        )
        docx_files["De_goc.docx"] = degoc_bytes
        all_exam_data[0] = (ym_parsed["p1"], ym_parsed["p2"], ym_parsed["p3"])

        # Render each shuffled code
        for ma_de in exam_codes:
            code_str = f"{ma_de:03d}"
            p1_s, p2_s, p3_s = ym_core.shuffle_exam(ym_parsed, ma_de, shuffle_mode)
            all_exam_data[ma_de] = (p1_s, p2_s, p3_s)

            docx_bytes = ym_core.render_docx(p1_s, p2_s, p3_s, str(ma_de), ym_parsed)
            docx_files[f"Ma_de_{code_str}.docx"] = docx_bytes

            qr_bytes = ym_core.export_qr_code(ma_de, p1_s, p2_s, p3_s)
            qr_files[f"qr_{code_str}.png"] = qr_bytes

        # Export Excel
        excel_bytes = ym_core.export_excel_ym(all_exam_data, exam_codes)
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"Lỗi khi xử lý trộn đề: {exc}",
        )

    # ---- Build zip ----
    job_id = str(uuid.uuid4())
    zip_path = _zip_path(job_id)

    try:
        with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            # Word documents (De_goc.docx và các Ma_de_xxx.docx)
            for fname, bdata in docx_files.items():
                zf.writestr(fname, bdata)

            # Excel answer key chuẩn app chấm thi QM 2025
            zf.writestr("Dap_an_excel_app_QM_2025.xlsx", excel_bytes)

            # QR codes
            for qrfname, qrdata in qr_files.items():
                zf.writestr(f"qr/{qrfname}", qrdata)

            # Phiếu trả lời trắc nghiệm chuẩn QM 2025 A4
            pdf_path = _SAMPLES_DIR / "QM_2025_A4_Ky-kiem-tra.pdf"
            if pdf_path.exists():
                zf.writestr("QM_2025_A4_Ky-kiem-tra.pdf", pdf_path.read_bytes())
    except Exception as exc:
        raise HTTPException(
            status_code=500,
            detail=f"Lỗi khi tạo file zip kết quả: {exc}",
        )

    # ---- Tên file zip kết quả: fengmix_ngày_giờ_Tên file (bỏ .docx) ----
    now = datetime.now()
    date_str = now.strftime("%Y%m%d")
    time_str = now.strftime("%H%M%S")
    orig_filename = file.filename or "de_thi.docx"
    clean_name = re.sub(r"\.docx$", "", orig_filename, flags=re.IGNORECASE).strip()
    if not clean_name:
        clean_name = "de_thi"
    zip_filename = f"fengmix_{date_str}_{time_str}_{clean_name}.zip"

    # ---- Register job ----
    JOBS[job_id] = {
        "zip_path": str(zip_path),
        "total": effective_count,
        "status": "done",
        "exam_codes": exam_codes,
        "total_questions": total_questions,
        "zip_filename": zip_filename,
    }
    _cleanup_old_jobs()

    return JSONResponse(
        {
            "job_id": job_id,
            "status": "done",
            "total": effective_count,
            "total_questions": total_questions,
            "exam_codes": exam_codes,
            "download_url": f"/api/download/{job_id}",
            "zip_filename": zip_filename,
        }
    )


@app.get("/api/download/{job_id}")
async def download_result(job_id: str):
    """
    Tải file zip kết quả chứa tất cả đề thi, đáp án Excel và mã QR.
    Định dạng: fengmix_ngày_giờ_Tên file (bỏ .docx)
    """
    if job_id not in JOBS:
        raise HTTPException(
            status_code=404,
            detail="Không tìm thấy job. Có thể đã hết hạn hoặc job_id không đúng.",
        )

    job = JOBS[job_id]
    zip_path = Path(job["zip_path"])

    if not zip_path.exists():
        raise HTTPException(
            status_code=410,
            detail="File kết quả đã bị xóa. Vui lòng xử lý lại.",
        )

    zip_filename = job.get("zip_filename")
    if not zip_filename:
        now = datetime.now()
        date_str = now.strftime("%Y%m%d")
        time_str = now.strftime("%H%M%S")
        zip_filename = f"fengmix_{date_str}_{time_str}_{job_id[:8]}.zip"

    def _stream():
        with open(zip_path, "rb") as f:
            while chunk := f.read(65536):
                yield chunk

    import unicodedata
    nfkd = unicodedata.normalize("NFKD", zip_filename)
    ascii_name = nfkd.encode("ascii", "ignore").decode("ascii").strip()
    if not ascii_name.endswith(".zip"):
        ascii_name = "fengmix_result.zip"

    quoted_name = quote(zip_filename)
    content_disp = f'attachment; filename="{ascii_name}"; filename*=UTF-8\'\'{quoted_name}'

    return StreamingResponse(
        _stream(),
        media_type="application/zip",
        headers={
            "Content-Disposition": content_disp,
            "Content-Length": str(zip_path.stat().st_size),
        },
    )


@app.get("/api/jobs/{job_id}")
async def get_job_status(job_id: str):
    """Trả về thông tin chi tiết của một job."""
    if job_id not in JOBS:
        raise HTTPException(status_code=404, detail="Không tìm thấy job.")
    return JOBS[job_id]


@app.delete("/api/jobs/{job_id}")
async def delete_job(job_id: str):
    """Xóa một job và file zip liên quan."""
    if job_id not in JOBS:
        raise HTTPException(status_code=404, detail="Không tìm thấy job.")
    job = JOBS.pop(job_id)
    zip_path = Path(job["zip_path"])
    if zip_path.exists():
        zip_path.unlink(missing_ok=True)
    return {"status": "deleted", "job_id": job_id}
