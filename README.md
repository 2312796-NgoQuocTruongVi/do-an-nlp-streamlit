# Ứng dụng NLP tiếng Việt dùng Streamlit

## 1. Chức năng
- Nhập văn bản tiếng Việt.
- Thống kê ký tự, từ, câu, từ khác nhau.
- Phân tích cảm xúc bằng mô hình NLP tiếng Việt.
- Trích xuất từ khóa.
- Biểu đồ tần suất từ.
- Tóm tắt trích xuất.
- Lịch sử phân tích trong phiên.
- Xuất kết quả CSV.

## 2. Cài đặt

Mở PowerShell tại thư mục dự án:

```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Nếu PowerShell chặn Activate.ps1, có thể dùng CMD:

```cmd
venv\Scripts\activate.bat
python -m pip install -r requirements.txt
```

## 3. Chạy ứng dụng

```powershell
streamlit run app.py
```

Lần đầu dùng chức năng phân tích cảm xúc, Transformers sẽ tải mô hình về máy.
Cần Internet ở lần tải đầu tiên.

## 4. Cấu trúc

```text
DO AN STREAMLIT/
├── app.py
├── requirements.txt
├── README.md
└── venv/
```

## 5. Gợi ý nội dung báo cáo
- Chương 1: Tổng quan NLP và bài toán phân tích văn bản.
- Chương 2: Python, Streamlit, Transformers, mô hình sentiment tiếng Việt.
- Chương 3: Phân tích yêu cầu và thiết kế hệ thống.
- Chương 4: Xây dựng ứng dụng.
- Chương 5: Kiểm thử, đánh giá và hướng phát triển.

## 6. Hướng phát triển
- Thêm phân loại chủ đề.
- Thêm nhận diện thực thể (NER).
- Thêm tóm tắt bằng mô hình Transformer.
- Lưu dữ liệu vào SQLite/MySQL.
- Đăng nhập và phân quyền.
- Triển khai ứng dụng lên máy chủ/cloud.
