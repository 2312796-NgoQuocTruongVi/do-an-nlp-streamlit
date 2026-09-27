# VietNLP Studio - Ứng dụng NLP tiếng Việt dùng Streamlit

Hệ thống phân tích xử lý ngôn ngữ tự nhiên (NLP) tiếng Việt hiện đại kết hợp Streamlit, mô hình học sâu PhoBERT và cơ sở dữ liệu SQLite.

## 1. Chức năng chính
- **Nhập văn bản linh hoạt:**
  - Nhập trực tiếp qua ô văn bản hoặc chọn các mẫu văn bản gợi ý nhanh.
  - **Tải lên từ tập tin:** Hỗ trợ định dạng `.docx` (Microsoft Word), `.pdf` (PDF), `.txt`, `.csv`, `.md`.
  - **Phân đoạn & Chọn đoạn tùy ý:** Tự động trích xuất các đoạn văn bản (paragraphs) trong tệp tin, cho phép chọn bất kỳ đoạn nào để phân tích và **tiếp tục chọn đoạn khác để phân tích liên tục** mà không cần tải lại tệp tin.
- **Phân tích NLP chuyên sâu:**
  - Thống kê định lượng: Số ký tự, số từ, số câu, số từ khác nhau, từ có nội dung.
  - Phân tích cảm xúc tiếng Việt bằng mô hình Deep Learning **PhoBERT** (`wonrax/phobert-base-vietnamese-sentiment`).
  - Trích xuất 10 từ khóa quan trọng nhất theo tần suất có trọng số.
  - Trực quan hóa tần suất từ với biểu đồ cột.
  - Tóm tắt trích xuất câu quan trọng từ văn bản gốc.
- **Cơ sở dữ liệu SQLite & Quản lý lịch sử:**
  - Tự động lưu trữ mọi kết quả phân tích vào database `history.db` (thời gian, nguồn gốc, số liệu, cảm xúc, từ khóa, tóm tắt).
  - **Chức năng xóa riêng lẻ từng bản ghi:** Cho phép người dùng xóa đúng 1 bản ghi bất kỳ, không nhất thiết phải xóa tất cả.
  - Tra cứu, tìm kiếm theo từ khóa và lọc theo nhãn cảm xúc.
  - Nạp lại văn bản từ lịch sử vào trang phân tích.
  - Xuất báo cáo kết quả ra tệp CSV.
- **Giao diện Menu & Sidebar nâng cấp:**
  - Thiết kế bảng điều khiển chuyên nghiệp, màu sắc hài hòa.
  - Giám sát trạng thái AI Engine, trạng thái Database và số lượng bản ghi thời gian thực.

## 2. Cài đặt

Mở PowerShell tại thư mục dự án:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Nếu PowerShell chặn `Activate.ps1`, có thể dùng CMD:

```cmd
venv\Scripts\activate.bat
python -m pip install -r requirements.txt
```

## 3. Chạy ứng dụng

```powershell
streamlit run app.py
```

Ứng dụng sẽ tự động mở trên trình duyệt tại địa chỉ `http://localhost:8501`.

## 4. Cấu trúc dự án

```text
DO AN STREAMLIT/
├── .streamlit/
│   └── config.toml           # Cấu hình theme và hệ thống Streamlit
├── app.py                    # Giao diện chính Streamlit và điều hướng
├── database.py               # Module quản lý cơ sở dữ liệu SQLite (CRUD lịch sử)
├── file_handler.py           # Module đọc tệp tin (Word, PDF, TXT, CSV) & phân đoạn
├── history.db                # Cơ sở dữ liệu SQLite lưu trữ lịch sử
├── requirements.txt          # Danh sách thư viện phụ thuộc
├── README.md                 # Hướng dẫn dự án
└── venv/                     # Môi trường ảo Python
```
