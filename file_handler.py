import io
import re
import csv
import docx

try:
    import pypdf
except ImportError:
    pypdf = None


def decode_bytes(raw_bytes: bytes) -> str:
    """Giải mã chuỗi bytes văn bản với nhiều encoding fallback."""
    for enc in ["utf-8-sig", "utf-8", "utf-16", "cp1252", "latin-1"]:
        try:
            return raw_bytes.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw_bytes.decode("utf-8", errors="ignore")


def extract_paragraphs_from_text(text: str) -> list[str]:
    """Tách văn bản thuần thành danh sách các đoạn văn bản (paragraphs)."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    # Tách theo dòng trống kép hoặc ngắt đoạn
    raw_paras = re.split(r"\n\s*\n+", text)
    cleaned = []
    for p in raw_paras:
        p_clean = " ".join(p.split())
        if p_clean and len(p_clean) >= 3:
            cleaned.append(p_clean)
    
    # Nếu chỉ có 1 đoạn lớn nhưng có nhiều dòng đơn, xem xét tách theo dòng
    if len(cleaned) <= 1 and "\n" in text:
        single_lines = [l.strip() for l in text.split("\n") if l.strip()]
        if len(single_lines) > 1:
            cleaned = single_lines
            
    return cleaned


def parse_uploaded_file(uploaded_file) -> dict:
    """
    Xử lý tệp tải lên từ Streamlit, trích xuất toàn bộ văn bản và danh sách các đoạn.
    Trả về dict: {
        'filename': str,
        'extension': str,
        'paragraphs': list[str],
        'full_text': str,
        'error': str or None
    }
    """
    filename = uploaded_file.name
    file_bytes = uploaded_file.getvalue()
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""

    paragraphs = []
    full_text = ""
    error = None

    try:
        if ext in ["txt", "md"]:
            full_text = decode_bytes(file_bytes)
            paragraphs = extract_paragraphs_from_text(full_text)

        elif ext == "docx":
            doc = docx.Document(io.BytesIO(file_bytes))
            for p in doc.paragraphs:
                p_text = p.text.strip()
                if p_text:
                    paragraphs.append(p_text)
            full_text = "\n\n".join(paragraphs)

        elif ext == "pdf":
            if pypdf is None:
                raise ImportError("Thư viện pypdf chưa được cài đặt.")
            reader = pypdf.PdfReader(io.BytesIO(file_bytes))
            pages_text = []
            for page in reader.pages:
                t = page.extract_text()
                if t:
                    pages_text.append(t)
            full_text = "\n\n".join(pages_text)
            paragraphs = extract_paragraphs_from_text(full_text)

        elif ext == "csv":
            text_data = decode_bytes(file_bytes)
            reader = csv.reader(io.StringIO(text_data))
            rows = list(reader)
            if rows:
                # Lấy tất cả nội dung ô có độ dài > 5 ký tự
                for r in rows:
                    row_content = " | ".join([cell.strip() for cell in r if cell.strip()])
                    if row_content:
                        paragraphs.append(row_content)
                full_text = "\n".join(paragraphs)
        else:
            # Thử đọc dạng text thuần nếu không rõ định dạng
            full_text = decode_bytes(file_bytes)
            paragraphs = extract_paragraphs_from_text(full_text)

        # Đảm bảo full_text chứa toàn bộ các đoạn
        if paragraphs:
            full_text = "\n\n".join(paragraphs)
        elif full_text.strip():
            paragraphs = [full_text.strip()]

    except Exception as e:
        error = f"Lỗi đọc tệp: {str(e)}"

    return {
        "filename": filename,
        "extension": ext,
        "paragraphs": paragraphs,
        "full_text": full_text,
        "error": error,
    }
