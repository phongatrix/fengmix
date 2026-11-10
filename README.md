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
├── api/                      # Serverless entrypoint cho Vercel Cloud
│   ├── index.py              # Vercel Serverless Function (FastAPI)
│   └── requirements.txt      # Dependency cho Vercel runtime
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
├── start.bat                 # Script tự động cài thư viện & chạy trên Windows
├── run_windows.bat           # Phím tắt khởi chạy nhanh trên Windows
├── start.sh                  # Script khởi động nhanh server trên Linux / macOS
├── vercel.json               # Cấu hình triển khai Vercel (routing, functions)
├── requirements.txt          # Danh sách thư viện Python dùng cho cài đặt nhanh
├── requirements.md           # Hướng dẫn chi tiết về các thư viện phụ thuộc
├── .gitignore                # Cấu hình bỏ qua các file tạm, cache
└── README.md                 # Hướng dẫn cài đặt và sử dụng
```

---

## 🚀 Hướng Dẫn Cài Đặt & Chạy Tự Động

### 1. Dành cho người dùng Windows (Tự động 100%)

Chỉ cần nhấp đúp (double-click) chuột vào file:
👉 **`start.bat`** (hoặc **`run_windows.bat`**)

Hệ thống sẽ **tự động hoàn toàn**:
1. Kiểm tra Python trên máy (nếu chưa cài, script sẽ tự mở link tải Python chính thức).
2. Tự khởi tạo môi trường ảo `.venv`.
3. Tự động kiểm tra và chạy cài đặt toàn bộ thư viện cần thiết từ `requirements.txt` (theo mô tả tại `requirements.md`).
4. Khởi động máy chủ và **tự động mở trình duyệt web** tại `http://localhost:8000`.

---

### 2. Dành cho người dùng Linux / macOS

Mở Terminal và chạy lệnh:

```bash
bash start.sh
# Hoặc chỉ định cổng tùy ý: bash start.sh 8080
```

---

### 3. Cài đặt và khởi chạy thủ công (Mọi hệ điều hành)

Nếu bạn muốn tự quản lý bằng dòng lệnh:

```bash
# 1. Cài đặt các thư viện phụ thuộc
pip install -r requirements.txt

# 2. Khởi chạy server
cd backend
python3 -m uvicorn main:app --host 0.0.0.0 --port 8000 --reload
```

---

### 4. Triển khai trực tuyến lên Vercel (`fengmix.vercel.app`)

Dự án đã được cấu hình sẵn sàng cho **Vercel Serverless Architecture** (`vercel.json` + `api/index.py`):

1. **Cách 1: Triển khai 1-Click qua GitHub (Khuyên dùng - Tự động cập nhật khi push code):**
   - Đăng nhập vào [Vercel](https://vercel.com).
   - Chọn **Add New...** ➔ **Project**.
   - Chọn import repository `phongatrix/fengmix`.
   - Đặt **Project Name** là `fengmix` (Vercel sẽ tự động cấp domain: `https://fengmix.vercel.app`).
   - Bấm **Deploy**. Vercel sẽ tự động build và online trong vòng 30 giây!
   - Từ nay về sau, mỗi khi bạn `git push` code mới lên GitHub, Vercel sẽ tự động build lại phiên bản mới nhất.

2. **Cách 2: Triển khai nhanh bằng Vercel CLI:**
   ```bash
   npx vercel login
   npx vercel --prod
   ```

3. **Lưu ý về tên miền:**
   - Vercel cấp tên miền miễn phí dạng `*.vercel.app` (ví dụ: `fengmix.vercel.app`).
   - Nếu bạn muốn dùng tên miền riêng (ví dụ: `fengmix.com` hoặc `fengmix.edu.vn`), chỉ cần vào mục **Settings ➔ Domains** trên Vercel Dashboard và thêm tên miền của bạn.

---

### 5. Truy cập giao diện Web nội bộ

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
