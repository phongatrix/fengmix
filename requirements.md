# 📦 Yêu Cầu Hệ Thống & Thư Viện Cần Thiết (Requirements)

Tài liệu này liệt kê chi tiết các yêu cầu hệ thống và thư viện cần thiết để cài đặt và vận hành ứng dụng **Feng Mix**.

---

## 1. Yêu Cầu Phần Cứng & Hệ Điều Hành

- **Hệ điều hành:**
  - Windows 10 / 11 (64-bit)
  - Linux (Ubuntu, Debian, Fedora, Arch...)
  - macOS (Apple Silicon M1/M2/M3 hoặc Intel)
- **Môi trường:**
  - **Python 3.9 trở lên** (Khuyên dùng Python 3.10, 3.11 hoặc 3.12).
  - Trình duyệt Web hiện đại: Google Chrome, Microsoft Edge, Mozilla Firefox, Brave, Safari...

> **⚠️ Lưu ý quan trọng cho người dùng Windows:** Khi cài đặt Python từ [python.org](https://www.python.org/downloads/), ở màn hình đầu tiên bắt buộc phải tích chọn ô **`Add python.exe to PATH`** trước khi bấm Install.

---

## 2. Danh Sách Thư Viện Python Phụ Thuộc

Các thư viện được quản lý trong file `requirements.txt`:

| Tên thư viện | Phiên bản | Mục đích sử dụng |
|--------------|-----------|------------------|
| `fastapi` | `==0.111.0` | Framework API backend tốc độ cao, hỗ trợ router và validation |
| `uvicorn[standard]` | `==0.29.0` | ASGI Web Server chuẩn để chạy ứng dụng FastAPI |
| `python-docx` | `==1.1.2` | Đọc, bóc tách cấu trúc câu hỏi và render các file đề thi `.docx` |
| `openpyxl` | `==3.1.2` | Tạo và định dạng bảng đáp án thi file Excel `.xlsx` |
| `qrcode[pil]` | `==7.4.2` | Sinh mã QR code cho từng mã đề dùng để chấm tự động |
| `Pillow` | `>=10.3.0` | Thư viện xử lý ảnh phục vụ tạo mã QR và xuất tài liệu |
| `python-multipart` | `==0.0.9` | Tiếp nhận và xử lý file Word tải lên qua form HTTP multipart |

---

## 3. Cài Đặt Tự Động

### Trên Windows
Chỉ cần nhấp đúp (double-click) chuột vào file:
👉 **`start.bat`** (hoặc `run_windows.bat`)

File script sẽ tự động:
1. Kiểm tra Python trong máy (nếu chưa có sẽ tự mở link tải Python).
2. Tự tạo môi trường ảo `.venv` để tránh xung đột thư viện.
3. Tự động chạy lệnh cài đặt toàn bộ danh sách thư viện từ `requirements.txt`.
4. Tự động bật máy chủ và mở trình duyệt web tại `http://localhost:8000`.

### Trên Linux / macOS
Mở terminal và chạy lệnh:
```bash
bash start.sh
```

---

## 4. Cài Đặt Thủ Công Bằng Dòng Lệnh

Nếu bạn muốn cài đặt thủ công:

```bash
# 1. Tạo và kích hoạt môi trường ảo (tùy chọn)
python -m venv .venv

# Kích hoạt trên Windows:
.venv\Scripts\activate

# Kích hoạt trên Linux/macOS:
source .venv/bin/activate

# 2. Cài đặt các thư viện cần thiết
pip install -r requirements.txt

# 3. Khởi chạy ứng dụng
cd backend
python -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```
