# 🎓 Feng Mix – Hệ Thống Trộn Đề Thi Trắc Nghiệm Thông Minh

**Feng Mix** là ứng dụng trộn đề thi trắc nghiệm trực tuyến hiện đại, hỗ trợ chuẩn hóa đề thi theo cấu trúc đề thi tốt nghiệp THPT từ năm 2025 của Bộ Giáo dục & Đào tạo. 

Hệ thống cho phép nạp file Microsoft Word (`.docx`), xem trước cấu trúc đề, **chỉnh sửa trực tiếp nội dung câu hỏi, phương án và đáp án đúng**, cảnh báo câu hỏi thiếu đáp án, xáo đề ngẫu nhiên theo nhiều chế độ và tự động xuất trọn bộ gói kết quả (File Word từng mã đề, Bảng đáp án Excel, Mã QR chấm bài, Phiếu thi PDF).

---

## 🌟 Tính Năng Nổi Bật

- **Hỗ trợ đầy đủ 3 phần thi theo chuẩn Bộ GD&ĐT 2025:**
  - **Phần I:** Trắc nghiệm nhiều phương án lựa chọn (A, B, C, D).
  - **Phần II:** Trắc nghiệm Đúng / Sai (mỗi câu gồm các ý a, b, c, d).
  - **Phần III:** Trắc nghiệm trả lời ngắn (điền kết quả số hoặc chuỗi).
- **Xem trước & Chỉnh sửa trực quan ngay trên Web:**
  - Bấm trực tiếp vào nội dung câu hỏi sau nhãn `Câu X.` để sửa văn bản, sửa lỗi chính tả.
  - Bấm vào bất kỳ phương án nào để thay đổi đáp án đúng (`✓`) hoặc sửa nội dung chữ của phương án.
  - Bấm nhãn `a)`, `b)`, `c)`, `d)` ở Phần II để đổi trạng thái Đúng / Sai nhanh chóng.
  - Nút **💾 Lưu** và **🔄 Reset** giúp quản lý các thay đổi an toàn trước khi bấm trộn.
- **Cảnh báo thông minh (Smart Warnings):**
  - Tự động phát hiện và cảnh báo các câu hỏi chưa được gạch chân / chưa chọn đáp án đúng trước khi trộn.
- **Nhiều chế độ trộn đề linh hoạt:**
  - Trộn cả câu hỏi và đảo thứ tự phương án.
  - Chỉ trộn câu hỏi (giữ nguyên thứ tự A, B, C, D).
  - Chỉ đảo phương án (giữ nguyên thứ tự câu hỏi).
  - Trộn theo từng nhóm câu hỏi (`<g...>` tags).
- **Tùy chỉnh mã đề thi:**
  - Tự chọn số lượng mã đề (từ 1 đến 100 mã đề).
  - Tùy chọn mã đề bắt đầu (ví dụ: 101, 201...) hoặc tạo mã đề ngẫu nhiên 3 chữ số không trùng nhau.
- **Đóng gói kết quả toàn diện trong 1 file ZIP:**
  - File đề thi gốc `De_goc.docx` (mã 000).
  - Các file đề thi hoán vị đẹp chuẩn Word `Ma_de_xxx.docx` (có tiêu đề trường, mã đề, số trang động `PAGE / NUMPAGES`).
  - File bảng đáp án Excel chuẩn format `Dap_an_excel_app_QM_2025.xlsx`.
  - Thư mục ảnh mã QR code cho từng mã đề dùng để chấm tự động bằng app điện thoại.
  - File biểu mẫu phiếu làm bài thi PDF.

---

## 📁 Cấu Trúc Thư Mục Dự Án

```
fengmix/
├── backend/                  # Mã nguồn Backend (FastAPI, Python)
│   ├── main.py               # FastAPI entrypoint & router API
│   ├── ym_core.py            # Parser Word, bộ hoán vị và xuất docx/xlsx/qr
│   ├── parser.py             # Parser xử lý văn bản và bảng biểu
│   ├── shuffler.py           # Thuật toán xáo câu hỏi và đáp án
│   ├── renderer.py           # Render tài liệu Word hoàn chỉnh
│   ├── exporter.py           # Xuất báo cáo Excel và mã QR
│   └── requirements.txt      # Danh sách thư viện Python cần thiết
├── frontend/                 # Giao diện người dùng Web (Vanilla HTML/CSS/JS)
│   ├── index.html            # Giao diện chính responsive
│   ├── style.css             # Thiết kế giao diện hiện đại, tinh tế
│   └── app.js                # Xử lý tương tác, xem trước, sửa đề và gọi API
├── samples/                  # Các file đề mẫu chuẩn (Địa lý, Tin học...)
├── start.sh                  # Script khởi động nhanh server trên Linux / macOS
├── .gitignore                # Cấu hình bỏ qua các file tạm, cache
└── README.md                 # Hướng dẫn cài đặt và sử dụng
```

---

## 🚀 Hướng Dẫn Cài Đặt & Chạy

### 1. Yêu cầu hệ thống
- **Python:** Phiên bản `3.9` trở lên (Khuyên dùng Python 3.10, 3.11 hoặc 3.12).
- **Hệ điều hành:** Linux, macOS, hoặc Windows.

### 2. Cài đặt các thư viện phụ thuộc

Mở Terminal (hoặc Command Prompt / PowerShell trên Windows) và chuyển vào thư mục dự án:

```bash
cd fengmix/backend
pip install -r requirements.txt
```

> **Lưu ý trên Ubuntu / Debian đời mới:** Nếu hệ thống yêu cầu môi trường ảo hoặc gắn cờ hệ thống, bạn có thể tạo venv:
> ```bash
> python3 -m venv venv
> source venv/bin/activate
> pip install -r backend/requirements.txt
> ```

---

### 3. Khởi động ứng dụng

#### Cách 1: Sử dụng script `start.sh` (Linux / macOS)
Từ thư mục gốc của dự án, chạy lệnh:

```bash
bash start.sh
```

Mặc định server sẽ khởi chạy tại cổng **8000**. Nếu bạn muốn chỉ định cổng khác (ví dụ: cổng 8080):
```bash
bash start.sh 8080
```

#### Cách 2: Khởi chạy trực tiếp bằng Uvicorn (Mọi hệ điều hành)
Chuyển vào thư mục `backend` và chạy:

```bash
cd backend
python3 -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

---

### 4. Truy cập giao diện Web
Sau khi khởi động server, mở trình duyệt web và truy cập:

👉 **[http://localhost:8000](http://localhost:8000)** (hoặc cổng bạn đã thiết lập)

---

## 📝 Hướng Dẫn Định Dạng Đề Thi Word Đầu Vào

Để hệ thống nhận diện chính xác các phần và đáp án, file Word `.docx` cần tuân thủ một số quy ước đơn giản sau:

### 1. Phân chia các phần:
- **Phần I:** Ghi tiêu đề bắt đầu bằng `PHẦN I.` hoặc `Phần 1:` (ví dụ: `PHẦN I. Câu trắc nghiệm nhiều phương án lựa chọn.`)
- **Phần II:** Ghi tiêu đề bắt đầu bằng `PHẦN II.` (ví dụ: `PHẦN II. Câu trắc nghiệm đúng sai.`)
- **Phần III:** Ghi tiêu đề bắt đầu bằng `PHẦN III.` (ví dụ: `PHẦN III. Câu trắc nghiệm trả lời ngắn.`)

### 2. Quy cách câu hỏi:
- Bắt đầu mỗi câu bằng: `Câu 1.`, `Câu 2:`, `Câu 3.`...
- **Phần I (MCQ):** 
  - Các phương án ghi rõ `A.`, `B.`, `C.`, `D.`
  - **Đáp án đúng:** Được đánh dấu bằng cách **gạch chân (Underline)** chữ cái phương án (ví dụ: <u>A</u>. Phương án đúng) hoặc ghi chú cuối câu `[Đáp án: A]`.
- **Phần II (Đúng/Sai):**
  - Các ý ghi nhãn `a)`, `b)`, `c)`, `d)`
  - **Ý đúng:** Gạch chân chữ cái nhãn ý đó (ví dụ: <u>a</u>) Nội dung đúng...). Nếu không gạch chân, mặc định là Sai.
- **Phần III (Trả lời ngắn):**
  - Ghi đáp án ở dòng kế tiếp dạng `Đáp án: 15` hoặc `[Đáp án: -3.5]` hoặc ghi trực tiếp giá trị số.

---

## 🔌 API Endpoints chính

| Phương thức | Đường dẫn | Mô tả |
|-------------|-----------|-------|
| `GET` | `/` | Trả về giao diện chính của ứng dụng |
| `GET` | `/api/health` | Kiểm tra tình trạng hoạt động của API |
| `POST` | `/api/parse` | Tải file `.docx`, trả về dữ liệu JSON cấu trúc câu hỏi, phương án và cảnh báo |
| `POST` | `/api/process` | Tiếp nhận file đề, cấu hình số mã đề, nội dung sửa đổi và thực hiện trộn đề |
| `GET` | `/api/download/{job_id}` | Tải về file kết quả nén `.zip` |
| `GET` | `/api/samples` | Lấy danh sách các đề thi mẫu có sẵn |
| `GET` | `/api/samples/download/{filename}` | Tải file đề mẫu để chạy thử |

---

## 👥 Đóng Góp & Phát Triển

Mọi đóng góp, báo cáo lỗi (issues) và đề xuất tính năng (pull requests) đều được hoan nghênh. Vui lòng tạo issue trên GitHub repository:
👉 [https://github.com/phongatrix/fengmix](https://github.com/phongatrix/fengmix)

---

## 📄 Bản Quyền

Dự án được phát triển và phát hành miễn phí cho cộng đồng giáo viên và học sinh Việt Nam.
