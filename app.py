import re
import io
import csv
import json
from collections import Counter
from datetime import datetime

import streamlit as st

import database
import file_handler

# ============================================================
# CẤU HÌNH TRANG
# ============================================================
st.set_page_config(
    page_title="VietNLP Studio - Phân tích văn bản tiếng Việt",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Khởi tạo database nếu chưa có
database.init_db()

# ============================================================
# CSS GIAO DIỆN HIỆN ĐẠI & CHUYÊN NGHIỆP
# ============================================================
st.markdown(
    """
    <style>
    /* Tổng thể và typography */
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
    }

    /* Tiêu đề chính */
    .main-hero {
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 50%, #3b82f6 100%);
        color: white;
        padding: 1.6rem 2rem;
        border-radius: 16px;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 20px -2px rgba(37, 99, 235, 0.25);
    }
    .main-hero h1 {
        color: white !important;
        font-size: 2.1rem;
        font-weight: 800;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .main-hero p {
        color: #dbeafe !important;
        font-size: 1.02rem;
        margin: 0.35rem 0 0 0;
        font-weight: 400;
    }

    /* Menu điều hướng Sidebar dạng hàng bo góc, không có ô tick */
    div[data-testid="stSidebar"] div.stButton > button {
        display: flex !important;
        justify-content: flex-start !important;
        align-items: center !important;
        text-align: left !important;
        padding: 0.8rem 1.15rem !important;
        font-size: 0.96rem !important;
        border-radius: 12px !important;
        border: 1px solid #e2e8f0 !important;
        transition: all 0.2s cubic-bezier(0.4, 0, 0.2, 1) !important;
        margin-bottom: 0.2rem !important;
        width: 100% !important;
    }

    /* Nút chưa chọn: bo góc nguyên hàng chữ, nền trắng/nhẹ, viền mềm */
    div[data-testid="stSidebar"] div.stButton > button[kind="secondary"] {
        background-color: #ffffff !important;
        color: #334155 !important;
        font-weight: 500 !important;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.03) !important;
    }
    div[data-testid="stSidebar"] div.stButton > button[kind="secondary"]:hover {
        background-color: #eff6ff !important;
        border-color: #93c5fd !important;
        color: #1d4ed8 !important;
        transform: translateX(4px);
    }
    div[data-testid="stSidebar"] div.stButton > button[kind="secondary"] * {
        color: #334155 !important;
        font-weight: 500 !important;
    }
    div[data-testid="stSidebar"] div.stButton > button[kind="secondary"]:hover * {
        color: #1d4ed8 !important;
    }

    /* Nút đang được chọn: bo góc nguyên hàng chữ, nền xanh nổi bật */
    div[data-testid="stSidebar"] div.stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #2563eb, #1d4ed8) !important;
        border-color: #1d4ed8 !important;
        color: #ffffff !important;
        font-weight: 700 !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.28) !important;
    }
    div[data-testid="stSidebar"] div.stButton > button[kind="primary"] * {
        color: #ffffff !important;
        font-weight: 700 !important;
    }

    /* Card kết quả và phân đoạn */
    .file-segment-box {
        background: #ffffff;
        padding: 1rem 1.2rem;
        border-radius: 12px;
        border: 1px solid #e2e8f0;
        margin-bottom: 1rem;
    }
    .file-badge {
        background: #e0f2fe;
        color: #0369a1;
        padding: 3px 10px;
        border-radius: 6px;
        font-size: 0.85rem;
        font-weight: 600;
        display: inline-block;
    }

    /* Status indicator */
    .status-dot {
        height: 9px;
        width: 9px;
        background-color: #10b981;
        border-radius: 50%;
        display: inline-block;
        margin-right: 5px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# TỪ DỪNG TIẾNG VIỆT CƠ BẢN
# ============================================================
STOPWORDS = {
    "và", "là", "của", "cho", "với", "các", "một", "những", "trong",
    "được", "có", "không", "rất", "này", "đó", "tôi", "bạn", "mình",
    "thì", "mà", "khi", "đã", "sẽ", "đang", "như", "về", "từ", "đến",
    "theo", "nên", "hay", "hoặc", "vì", "nữa", "ra", "vào", "trên",
    "dưới", "sau", "trước", "lại", "cũng", "chỉ", "để", "bị", "do",
    "nó", "họ", "em", "anh", "chị", "ông", "bà", "we", "you", "the",
}

# ============================================================
# HÀM XỬ LÝ VĂN BẢN
# ============================================================
def normalize_text(text: str) -> str:
    """Chuẩn hóa khoảng trắng và xuống dòng."""
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def tokenize_words(text: str):
    """Tách từ đơn giản dựa trên Unicode chữ/số."""
    return re.findall(r"[^\W_]+(?:[-'][^\W_]+)*", text.lower(), flags=re.UNICODE)


def split_sentences(text: str):
    """Tách câu theo dấu câu tiếng Việt cơ bản."""
    parts = re.split(r"(?<=[.!?…])\s+|\n+", text.strip())
    return [p.strip() for p in parts if p.strip()]


def get_word_frequencies(text: str):
    words = tokenize_words(text)
    filtered = [
        w for w in words
        if len(w) > 1 and w not in STOPWORDS and not w.isdigit()
    ]
    return Counter(filtered)


def extract_keywords(text: str, top_n: int = 10):
    """Trích xuất từ khóa theo tần suất có trọng số."""
    freq = get_word_frequencies(text)
    if not freq:
        return []

    max_count = max(freq.values())
    scored = []
    for word, count in freq.items():
        score = count / max_count
        if len(word) >= 6:
            score *= 1.15
        scored.append((word, count, score))

    scored.sort(key=lambda x: (x[2], x[1]), reverse=True)
    return scored[:top_n]


def extractive_summary(text: str, max_sentences: int = 3):
    """Tóm tắt trích xuất đơn giản dựa trên chấm điểm câu."""
    sentences = split_sentences(text)
    if len(sentences) <= max_sentences:
        return sentences

    freq = get_word_frequencies(text)
    if not freq:
        return sentences[:max_sentences]

    max_freq = max(freq.values())
    sentence_scores = []

    for idx, sentence in enumerate(sentences):
        words = tokenize_words(sentence)
        score = sum(freq.get(w, 0) / max_freq for w in words if w not in STOPWORDS)
        if idx == 0:
            score *= 1.10
        sentence_scores.append((idx, score, sentence))

    best = sorted(sentence_scores, key=lambda x: x[1], reverse=True)[:max_sentences]
    best = sorted(best, key=lambda x: x[0])
    return [item[2] for item in best]


# ============================================================
# MODEL PHÂN TÍCH CẢM XÚC
# ============================================================
@st.cache_resource(show_spinner="Đang khởi động mô hình PhoBERT tiếng Việt...")
def load_sentiment_model():
    from transformers import pipeline

    return pipeline(
        "sentiment-analysis",
        model="wonrax/phobert-base-vietnamese-sentiment",
    )


def split_text_into_sentiment_chunks(text: str, max_words: int = 150) -> list[str]:
    """Chia văn bản thành các đoạn nhỏ hơn max_words để PhoBERT phân tích toàn diện."""
    paras = [p.strip() for p in text.split("\n") if p.strip()]
    chunks = []
    for p in paras:
        words = p.split()
        if len(words) <= max_words:
            chunks.append(p)
        else:
            sentences = re.split(r"(?<=[.!?…])\s+", p)
            curr = []
            curr_len = 0
            for s in sentences:
                s_words = len(s.split())
                if curr_len + s_words > max_words and curr:
                    chunks.append(" ".join(curr))
                    curr = [s]
                    curr_len = s_words
                else:
                    curr.append(s)
                    curr_len += s_words
            if curr:
                chunks.append(" ".join(curr))
    return chunks or [text]


def analyze_sentiment(text: str):
    """
    Phân tích cảm xúc văn bản bằng PhoBERT.
    Tự động phân đoạn để phân tích toàn bộ nội dung của tệp hoặc văn bản dài,
    không bị cắt cụt ở đoạn đầu tiên.
    """
    classifier = load_sentiment_model()

    mapping = {
        "NEG": ("😡", "Tiêu cực", "error"),
        "NEU": ("😐", "Trung tính", "info"),
        "POS": ("😊", "Tích cực", "success"),
        "LABEL_0": ("😡", "Tiêu cực", "error"),
        "LABEL_1": ("😐", "Trung tính", "info"),
        "LABEL_2": ("😊", "Tích cực", "success"),
    }

    chunks = split_text_into_sentiment_chunks(text, max_words=150)

    # Nếu văn bản ngắn (chỉ 1 đoạn), phân tích trực tiếp
    if len(chunks) <= 1:
        result = classifier(
            text,
            truncation=True,
            max_length=256,
        )[0]
        label = str(result["label"]).upper()
        score = float(result["score"])
        emoji, vietnamese_label, box_type = mapping.get(
            label,
            ("🔍", label, "info"),
        )
        return {
            "raw_label": label,
            "label": vietnamese_label,
            "emoji": emoji,
            "score": score,
            "box_type": box_type,
            "is_aggregated": False,
            "total_chunks": 1,
            "chunk_results": [],
        }

    # Nếu văn bản dài hoặc gồm nhiều đoạn (toàn bộ file), phân tích từng đoạn rồi tổng hợp
    chunk_results = []
    label_counts = {"Tích cực": 0, "Tiêu cực": 0, "Trung tính": 0}
    label_scores = {"Tích cực": [], "Tiêu cực": [], "Trung tính": []}

    for c in chunks:
        if not c.strip():
            continue
        res = classifier(c, truncation=True, max_length=256)[0]
        raw_label = str(res["label"]).upper()
        score = float(res["score"])
        c_emoji, c_vlabel, c_box = mapping.get(raw_label, ("🔍", raw_label, "info"))
        chunk_results.append({
            "text": c[:120] + ("..." if len(c) > 120 else ""),
            "label": c_vlabel,
            "score": score,
            "emoji": c_emoji,
        })
        if c_vlabel in label_counts:
            label_counts[c_vlabel] += 1
            label_scores[c_vlabel].append(score)

    total_valid_chunks = len(chunk_results)
    if total_valid_chunks == 0:
        return {
            "raw_label": "NEU",
            "label": "Trung tính",
            "emoji": "😐",
            "score": 0.5,
            "box_type": "info",
            "is_aggregated": True,
            "total_chunks": 0,
            "chunk_results": [],
            "breakdown": label_counts,
        }

    # Chọn nhãn chiếm ưu thế cao nhất
    sorted_labels = sorted(
        label_counts.items(),
        key=lambda item: (item[1], sum(label_scores[item[0]]) if label_scores[item[0]] else 0),
        reverse=True,
    )
    winner_label, winner_count = sorted_labels[0]

    if label_scores[winner_label]:
        avg_score = sum(label_scores[winner_label]) / len(label_scores[winner_label])
        ratio = winner_count / total_valid_chunks
        final_score = (avg_score * 0.7) + (ratio * 0.3)
    else:
        final_score = winner_count / total_valid_chunks

    box_type_map = {
        "Tích cực": ("😊", "success"),
        "Tiêu cực": ("😡", "error"),
        "Trung tính": ("😐", "info"),
    }
    emoji, box_type = box_type_map.get(winner_label, ("🔍", "info"))

    return {
        "raw_label": winner_label,
        "label": winner_label,
        "emoji": emoji,
        "score": min(0.9999, max(0.5, final_score)),
        "box_type": box_type,
        "is_aggregated": True,
        "total_chunks": total_valid_chunks,
        "chunk_results": chunk_results,
        "breakdown": label_counts,
    }


# ============================================================
# EXPORT CSV
# ============================================================
def make_csv(stats, sentiment, keywords):
    output = io.StringIO()
    writer = csv.writer(output)

    writer.writerow(["Thông tin", "Giá trị"])
    for key, value in stats.items():
        writer.writerow([key, value])

    if sentiment:
        writer.writerow(["Cảm xúc", sentiment["label"]])
        writer.writerow(["Độ tin cậy", f"{sentiment['score']:.2%}"])

    writer.writerow([])
    writer.writerow(["Từ khóa", "Tần suất", "Điểm"])
    for word, count, score in keywords:
        writer.writerow([word, count, f"{score:.4f}"])

    return output.getvalue().encode("utf-8-sig")


# ============================================================
# QUẢN LÝ SESSION STATE
# ============================================================
if "selected_page" not in st.session_state:
    st.session_state.selected_page = "Phân tích văn bản"

if "manual_text_input" not in st.session_state:
    st.session_state.manual_text_input = ""

if "file_data" not in st.session_state:
    st.session_state.file_data = None

if "current_para_index" not in st.session_state:
    st.session_state.current_para_index = 0

# ============================================================
# SIDEBAR GIAO DIỆN MỚI: ĐẸP MẮT & CHUYÊN NGHIỆP
# ============================================================
with st.sidebar:
    # 1. Danh mục chức năng dạng hàng bo góc (không có ô tick)
    st.markdown("### 🧭 Danh mục chức năng")
    st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)

    nav_items = [
        ("📝  Phân tích văn bản", "Phân tích văn bản"),
        ("📊  Lịch sử phân tích", "Lịch sử phân tích"),
        ("ℹ️  Giới thiệu đề tài", "Giới thiệu đề tài"),
    ]

    for label, page_key in nav_items:
        is_active = (st.session_state.selected_page == page_key)
        if st.button(
            label,
            key=f"nav_btn_{page_key}",
            type="primary" if is_active else "secondary",
            width="stretch",
        ):
            st.session_state.selected_page = page_key
            st.rerun()

    page = st.session_state.selected_page

    st.markdown("<div style='height: 12px;'></div>", unsafe_allow_html=True)

    # 3. Widget Thông tin Hệ thống & Database
    db_stats = database.get_history_stats()
    with st.container(border=True):
        st.markdown("**⚡ Trạng thái hệ thống**")
        st.caption(f"💾 Cơ sở dữ liệu: **SQLite**")
        st.caption(f"📊 Lịch sử đã lưu: **{db_stats['total']} bản ghi**")
        st.caption(f"🤖 Mô hình: **PhoBERT Sentiment**")
        st.caption(f"📁 Định dạng: **.docx, .pdf, .txt, .csv**")

    # 4. Footer sidebar
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    with st.container():
        st.caption("🎓 **Đồ án Chuyên ngành KTPM**")
        st.caption("Công nghệ: Streamlit • Transformers • SQLite")


# ============================================================
# TRANG 1: PHÂN TÍCH VĂN BẢN (NHẬP TAY HOẶC TẢI TỆP TIN & CHỌN ĐOẠN)
# ============================================================
if page == "Phân tích văn bản":
    st.markdown(
        """
        <div class="main-hero">
            <h1>🧠 Phân tích Văn bản NLP Tiếng Việt</h1>
            <p>Hỗ trợ nhập trực tiếp hoặc tải tệp tin (Word, PDF, TXT, CSV), trích xuất và chọn từng đoạn văn bản để phân tích chuyên sâu.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # Lựa chọn phương thức nhập văn bản
    input_method = st.segmented_control(
        "Phương thức nhập văn bản",
        options=["✍️ Nhập văn bản trực tiếp", "📁 Tải tệp tin (File DOCX, PDF, TXT, CSV)"],
        default="✍️ Nhập văn bản trực tiếp",
    )

    text_to_analyze = ""
    analysis_source = "Nhập trực tiếp"

    # ------------------------------------------------------------
    # CHẾ ĐỘ 1: NHẬP VĂN BẢN TRỰC TIẾP
    # ------------------------------------------------------------
    if input_method == "✍️ Nhập văn bản trực tiếp":
        with st.container(border=True):
            st.markdown("##### ✍️ Nhập hoặc dán văn bản tiếng Việt")
            
            # Các nút văn bản mẫu gợi ý nhanh
            col_ex1, col_ex2, col_ex3 = st.columns(3)
            with col_ex1:
                if st.button("💡 Mẫu tích cực (Khen ngợi)", width="stretch"):
                    st.session_state.manual_text_input = (
                        "Tôi rất hài lòng với chất lượng dịch vụ của công ty. "
                        "Giao diện ứng dụng đẹp mắt, tốc độ xử lý nhanh chóng và đội ngũ nhân viên rất nhiệt tình."
                    )
                    st.rerun()
            with col_ex2:
                if st.button("💡 Mẫu tiêu cực (Phàn nàn)", width="stretch"):
                    st.session_state.manual_text_input = (
                        "Sản phẩm quá tệ, đóng gói sơ sài và bị hư hỏng khi nhận hàng. "
                        "Hệ thống thường xuyên gặp lỗi kết nối khiến trải nghiệm của tôi rất thất vọng."
                    )
                    st.rerun()
            with col_ex3:
                if st.button("💡 Mẫu trung tính (Thông tin)", width="stretch"):
                    st.session_state.manual_text_input = (
                        "Trường Đại học Công nghệ Thông tin tổ chức hội thảo khoa học vào sáng thứ Năm tuần này. "
                        "Các sinh viên và giảng viên có thể đăng ký tham gia trực tuyến qua cổng thông tin."
                    )
                    st.rerun()

            text_input = st.text_area(
                "Nội dung văn bản",
                value=st.session_state.manual_text_input,
                height=180,
                placeholder="Nhập hoặc dán nội dung văn bản tiếng Việt cần phân tích tại đây...",
                label_visibility="collapsed",
            )
            # Cập nhật state nếu người dùng gõ
            st.session_state.manual_text_input = text_input
            text_to_analyze = text_input
            analysis_source = "Nhập trực tiếp"

    # ------------------------------------------------------------
    # CHẾ ĐỘ 2: NHẬP TỪ TỆP TIN & CHỌN ĐOẠN TÙY Ý ĐỂ PHÂN TÍCH
    # ------------------------------------------------------------
    else:
        with st.container(border=True):
            st.markdown("##### 📁 Tải lên tệp tin văn bản")
            st.caption("Hỗ trợ các định dạng: **.docx (Microsoft Word), .pdf (PDF), .txt, .md, .csv**")

            uploaded_file = st.file_uploader(
                "Chọn tệp tin từ máy tính",
                type=["docx", "pdf", "txt", "md", "csv"],
                label_visibility="collapsed",
            )

            if uploaded_file is not None:
                # Đọc và phân đoạn file
                with st.spinner("Đang đọc và phân tích cấu trúc tệp tin..."):
                    parsed = file_handler.parse_uploaded_file(uploaded_file)

                if parsed["error"]:
                    st.error(f"❌ {parsed['error']}")
                else:
                    st.session_state.file_data = parsed
                    paragraphs = parsed["paragraphs"]
                    total_paras = len(paragraphs)

                    if total_paras == 0:
                        st.warning("⚠️ Không tìm thấy nội dung văn bản hợp lệ trong tệp tin.")
                    else:
                        # Thông tin tệp tin đã tải
                        st.success(
                            f"✅ Đã tải thành công tệp **{parsed['filename']}** | "
                            f"Trích xuất được **{total_paras} đoạn văn**"
                        )

                        # Bộ chọn phạm vi phân tích: Toàn bộ tệp hoặc Từng đoạn
                        st.markdown("---")
                        st.markdown("##### 🎯 Chọn phạm vi nội dung để phân tích")

                        file_scope = st.radio(
                            "Phạm vi phân tích:",
                            options=[
                                f"📄 Phân tích toàn bộ nội dung tệp ({total_paras} đoạn)",
                                "🔍 Chọn một đoạn cụ thể trong tệp để phân tích",
                            ],
                            key=f"file_scope_{parsed['filename']}",
                        )

                        if file_scope.startswith("📄 Phân tích toàn bộ"):
                            analysis_source = f"Tệp: {parsed['filename']} (Toàn bộ {total_paras} đoạn)"
                            st.caption(
                                f"Đang chọn: **Toàn bộ nội dung tệp ({total_paras} đoạn, {len(parsed['full_text']):,} ký tự)**. "
                                "Hệ thống sẽ phân tích toàn diện tất cả các đoạn văn trong tệp."
                            )
                            editable_full = st.text_area(
                                "Nội dung toàn bộ tệp (bạn có thể đọc hoặc chỉnh sửa trước khi phân tích):",
                                value=parsed["full_text"],
                                height=240,
                                key=f"file_full_text_{parsed['filename']}",
                            )
                            text_to_analyze = editable_full

                        else:
                            st.caption(
                                "Bạn có thể duyệt qua từng đoạn trong danh sách dưới đây để phân tích. "
                                "Sau khi phân tích xong, bạn có thể **tiếp tục chọn một đoạn khác** để phân tích tiếp."
                            )

                            para_options = []
                            for idx, p in enumerate(paragraphs):
                                preview = p[:70].replace("\n", " ") + ("..." if len(p) > 70 else "")
                                para_options.append(f"Đoạn #{idx + 1} ({len(p)} ký tự): {preview}")

                            curr_idx_key = f"para_idx_{parsed['filename']}"
                            if curr_idx_key not in st.session_state or st.session_state[curr_idx_key] >= total_paras:
                                st.session_state[curr_idx_key] = 0

                            # Thanh điều hướng nhanh
                            nav_col1, nav_col2, nav_col3 = st.columns([1.5, 1.5, 4])
                            with nav_col1:
                                if st.button("⏮️ Đoạn trước", disabled=(st.session_state[curr_idx_key] <= 0), width="stretch"):
                                    st.session_state[curr_idx_key] = max(0, st.session_state[curr_idx_key] - 1)
                                    st.rerun()
                            with nav_col2:
                                if st.button("Đoạn sau ⏭️", disabled=(st.session_state[curr_idx_key] >= total_paras - 1), width="stretch"):
                                    st.session_state[curr_idx_key] = min(total_paras - 1, st.session_state[curr_idx_key] + 1)
                                    st.rerun()
                            with nav_col3:
                                st.caption(f"Đang duyệt đoạn **{st.session_state[curr_idx_key] + 1}** / **{total_paras}**")

                            selected_para_idx = st.selectbox(
                                "Chọn đoạn văn trong tệp:",
                                range(total_paras),
                                index=st.session_state[curr_idx_key],
                                format_func=lambda i: para_options[i],
                                key=f"select_box_{parsed['filename']}_{st.session_state[curr_idx_key]}",
                            )
                            st.session_state[curr_idx_key] = selected_para_idx

                            chosen_para = paragraphs[selected_para_idx]
                            analysis_source = f"Tệp: {parsed['filename']} (Đoạn {selected_para_idx + 1}/{total_paras})"

                            st.markdown(f"**Nội dung đang chọn:** `({analysis_source})`")
                            editable_para = st.text_area(
                                "Nội dung đoạn văn được chọn (bạn có thể chỉnh sửa trước khi phân tích):",
                                value=chosen_para,
                                height=160,
                                key=f"para_edit_box_{parsed['filename']}_{selected_para_idx}",
                            )
                            text_to_analyze = editable_para
            else:
                st.info("💡 Hãy tải lên một tệp tin (.docx, .pdf, .txt, .csv) để hệ thống tự động tách đoạn và hỗ trợ chọn đoạn phân tích.")

    # ------------------------------------------------------------
    # NÚT BẤM THỰC HIỆN PHÂN TÍCH
    # ------------------------------------------------------------
    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
    btn_col1, btn_col2 = st.columns([2, 5])
    with btn_col1:
        analyze_btn = st.button("🔍 Tiến hành phân tích", type="primary", width="stretch")
    with btn_col2:
        if text_to_analyze:
            st.caption(f"Nguồn: **{analysis_source}** | Chiều dài: **{len(text_to_analyze)} ký tự**")

    # Khi người dùng nhấn nút Phân tích
    if analyze_btn:
        clean_text = normalize_text(text_to_analyze)

        if not clean_text:
            st.warning("⚠️ Vui lòng nhập hoặc chọn đoạn văn bản trước khi phân tích.")
            st.stop()

        with st.spinner("Đang chạy phân tích NLP tiếng Việt..."):
            words = tokenize_words(clean_text)
            sentences = split_sentences(clean_text)
            freq = get_word_frequencies(clean_text)
            keywords = extract_keywords(clean_text, top_n=10)
            summary_list = extractive_summary(clean_text, max_sentences=3)
            summary_str = " ".join(summary_list) if summary_list else ""

            stats = {
                "Số ký tự": len(clean_text),
                "Số từ": len(words),
                "Số câu": len(sentences),
                "Số từ khác nhau": len(set(words)),
                "Từ có nội dung": len(freq),
            }

            # Phân tích cảm xúc
            sentiment = None
            try:
                sentiment = analyze_sentiment(clean_text)
            except Exception as exc:
                st.error("Không thể chạy mô hình phân tích cảm xúc.")
                st.exception(exc)

            # LƯU VÀO DATABASE SQLITE
            inserted_id = database.insert_history(
                source_type=analysis_source,
                text_content=clean_text,
                char_count=stats["Số ký tự"],
                word_count=stats["Số từ"],
                sentence_count=stats["Số câu"],
                unique_words=stats["Số từ khác nhau"],
                sentiment_label=sentiment["label"] if sentiment else "Chưa xác định",
                sentiment_score=sentiment["score"] if sentiment else 0.0,
                sentiment_emoji=sentiment["emoji"] if sentiment else "🔍",
                keywords=keywords,
                summary_text=summary_str,
            )

        st.toast(f"✅ Đã phân tích và lưu vào lịch sử (Mã #{inserted_id})!", icon=":material/check_circle:")

        # HIỂN THỊ KẾT QUẢ PHÂN TÍCH
        st.markdown("---")
        st.markdown("### 📊 Kết quả phân tích chi tiết")
        st.caption(f"Bản ghi **#{inserted_id}** • Nguồn: **{analysis_source}**")

        # 1. Thống kê số liệu KPI Cards
        c1, c2, c3, c4, c5 = st.columns(5)
        kpi_cols = [c1, c2, c3, c4, c5]
        for col, (label, value) in zip(kpi_cols, stats.items()):
            with col:
                with st.container(border=True):
                    st.metric(label, f"{value:,}")

        st.markdown("<div style='height: 8px;'></div>", unsafe_allow_html=True)

        # 2. Tabs chi tiết: Cảm xúc, Từ khóa, Tóm tắt (Đã bỏ tab Biểu đồ Tần suất)
        tab_sent, tab_kw, tab_sum = st.tabs(
            [
                "😊 Phân tích Cảm xúc",
                "🔑 Từ khóa Nổi bật",
                "📝 Tóm tắt Trích xuất",
            ]
        )

        with tab_sent:
            if sentiment:
                with st.container(border=True):
                    sc1, sc2 = st.columns([1.5, 4])
                    with sc1:
                        if sentiment["box_type"] == "success":
                            st.success(f"### {sentiment['emoji']} {sentiment['label']}")
                        elif sentiment["box_type"] == "error":
                            st.error(f"### {sentiment['emoji']} {sentiment['label']}")
                        else:
                            st.info(f"### {sentiment['emoji']} {sentiment['label']}")
                    with sc2:
                        st.markdown(f"**Độ tin cậy của mô hình:** `{sentiment['score']:.2%}`")
                        st.progress(sentiment["score"])
                        st.caption("Dự đoán bởi mô hình PhoBERT-base Vietnamese Sentiment")

                    # Nếu phân tích toàn bộ tệp hoặc văn bản nhiều đoạn, hiển thị thống kê phân bố
                    if sentiment.get("is_aggregated") and sentiment.get("breakdown"):
                        st.markdown("---")
                        bd = sentiment["breakdown"]
                        total_c = sentiment.get("total_chunks", 1)
                        st.markdown(f"**Phân tích toàn diện trên tất cả {total_c} đoạn nội dung:**")
                        st.markdown(
                            f":green-badge[{bd['Tích cực']} đoạn Tích cực]  "
                            f":red-badge[{bd['Tiêu cực']} đoạn Tiêu cực]  "
                            f":gray-badge[{bd['Trung tính']} đoạn Trung tính]"
                        )
                        if sentiment.get("chunk_results"):
                            with st.expander("📖 Xem chi tiết cảm xúc từng đoạn trong tệp"):
                                for c_idx, cr in enumerate(sentiment["chunk_results"], 1):
                                    st.markdown(
                                        f"**Đoạn #{c_idx}:** {cr['emoji']} **{cr['label']}** ({cr['score']:.1%}) — _{cr['text']}_"
                                    )
            else:
                st.info("Chưa có kết quả cảm xúc.")

        with tab_kw:
            with st.container(border=True):
                if keywords:
                    st.markdown("**Danh sách 10 từ khóa mang trọng số cao nhất:**")
                    kw_table = [
                        {
                            "Từ khóa": w,
                            "Tần suất": c,
                            "Trọng số": round(s, 4),
                        }
                        for w, c, s in keywords
                    ]
                    st.dataframe(
                        kw_table,
                        width="stretch",
                        hide_index=True,
                    )
                else:
                    st.info("Không tìm thấy từ khóa phù hợp.")

        with tab_sum:
            with st.container(border=True):
                if summary_str:
                    st.markdown("**Bản tóm tắt trích xuất thông minh:**")
                    st.info(summary_str)
                else:
                    st.info("Chưa đủ độ dài để thực hiện tóm tắt câu.")

        # Nút xuất CSV
        st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
        csv_bytes = make_csv(stats, sentiment, keywords)
        st.download_button(
            "⬇️ Tải xuống kết quả phân tích (CSV)",
            data=csv_bytes,
            file_name=f"ket_qua_phan_tich_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
        )

        if input_method != "✍️ Nhập văn bản trực tiếp":
            st.info(
                "💡 **Mẹo:** Bạn có thể kéo lên trên, chọn tiếp **Đoạn tiếp theo** hoặc đoạn bất kỳ khác trong tệp "
                "và bấm **Tiến hành phân tích** để phân tích tiếp tục mà không cần tải lại tệp!"
            )


# ============================================================
# TRANG 2: LỊCH SỬ PHÂN TÍCH (LƯU DB & XÓA TỪNG BẢN GHI HOẶC TẤT CẢ)
# ============================================================
elif page == "Lịch sử phân tích":
    st.markdown(
        """
        <div class="main-hero" style="background: linear-gradient(135deg, #0f766e 0%, #0d9488 50%, #14b8a6 100%); box-shadow: 0 4px 20px -2px rgba(13, 148, 136, 0.25);">
            <h1>📊 Cơ sở Dữ liệu Lịch sử Phân tích</h1>
            <p>Toàn bộ kết quả được lưu trữ bền vững trong SQLite. Bạn có thể tra cứu, xem chi tiết, nạp lại hoặc xóa riêng lẻ từng bản ghi.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    # 1. Thống kê tổng quan từ Database
    stats = database.get_history_stats()
    m1, m2, m3, m4 = st.columns(4)
    with m1:
        with st.container(border=True):
            st.metric("Tổng lượt phân tích", stats["total"], delta=None)
    with m2:
        with st.container(border=True):
            st.metric("😊 Tích cực", stats["pos"])
    with m3:
        with st.container(border=True):
            st.metric("😡 Tiêu cực", stats["neg"])
    with m4:
        with st.container(border=True):
            st.metric("😐 Trung tính", stats["neu"])

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    # 2. Thanh tìm kiếm và bộ lọc cảm xúc
    with st.container(border=True):
        f_col1, f_col2 = st.columns([3, 2])
        with f_col1:
            search_query = st.text_input(
                "🔍 Tìm kiếm lịch sử",
                placeholder="Nhập từ khóa tìm kiếm trong nội dung hoặc tên tệp...",
                label_visibility="collapsed",
            )
        with f_col2:
            sentiment_filter = st.segmented_control(
                "Lọc cảm xúc",
                options=["Tất cả", "Tích cực", "Tiêu cực", "Trung tính"],
                default="Tất cả",
                label_visibility="collapsed",
            )

    # Lấy dữ liệu từ SQLite
    history_items = database.get_all_history(
        search_query=search_query,
        sentiment_filter=sentiment_filter,
    )

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    if not history_items:
        with st.container(border=True):
            st.info("ℹ️ Không tìm thấy bản ghi lịch sử phân tích nào phù hợp với bộ lọc.")
    else:
        st.markdown(f"Đang hiển thị **{len(history_items)}** bản ghi phân tích gần nhất:")

        # 3. Danh sách từng bản ghi lịch sử
        for item in history_items:
            item_id = item["id"]

            with st.container(border=True):
                # Header hàng trên: Thời gian, Nguồn, Cảm xúc
                hdr_left, hdr_right = st.columns([3.5, 2.5])
                with hdr_left:
                    st.markdown(
                        f"**#{item_id}** • 🕒 `{item['created_at']}` • 📌 `{item['source_type']}`"
                    )
                with hdr_right:
                    # Hiển thị badge cảm xúc
                    label = item["sentiment_label"]
                    score = item["sentiment_score"]
                    if label == "Tích cực":
                        st.markdown(f":green-badge[😊 {label} ({score:.1%})]")
                    elif label == "Tiêu cực":
                        st.markdown(f":red-badge[😡 {label} ({score:.1%})]")
                    else:
                        st.markdown(f":gray-badge[😐 {label} ({score:.1%})]")

                # Nội dung văn bản (rút gọn nếu dài)
                content = item["text_content"]
                if len(content) > 280:
                    st.write(content[:280] + "...")
                    with st.expander("📖 Xem toàn bộ nội dung bản ghi"):
                        st.write(content)
                else:
                    st.write(content)

                # Thống kê nhanh
                s1, s2, s3, s4 = st.columns(4)
                s1.caption(f"Ký tự: **{item['char_count']}**")
                s2.caption(f"Từ: **{item['word_count']}**")
                s3.caption(f"Câu: **{item['sentence_count']}**")
                s4.caption(f"Từ khác nhau: **{item['unique_words']}**")

                # Từ khóa & Tóm tắt nếu có
                if item["keywords"]:
                    kw_names = [f"`{w[0]}` ({w[1]})" for w in item["keywords"][:6]]
                    st.caption("Từ khóa: " + " • ".join(kw_names))

                if item["summary_text"]:
                    st.caption(f"💡 Tóm tắt: _{item['summary_text']}_")

                st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)

                # CÁC NÚT THAO TÁC CHO TỪNG BẢN GHI
                act_col1, act_col2, act_col3 = st.columns([2, 2, 4])
                
                # Nút Xóa 1 bản ghi cụ thể (Chức năng người dùng yêu cầu!)
                with act_col1:
                    if st.button("🗑️ Xóa bản ghi này", key=f"btn_del_{item_id}", width="stretch"):
                        success = database.delete_history_item(item_id)
                        if success:
                            st.toast(f"Đã xóa thành công bản ghi #{item_id}!", icon=":material/delete:")
                            st.rerun()
                        else:
                            st.error("Không thể xóa bản ghi này.")

                # Nút Nạp lại để phân tích tiếp
                with act_col2:
                    if st.button("🔄 Nạp lại văn bản này", key=f"btn_reload_{item_id}", width="stretch"):
                        st.session_state.manual_text_input = item["text_content"]
                        st.session_state.selected_page = "Phân tích văn bản"
                        st.rerun()

        # Khu vực quản lý toàn cục & Xóa tất cả
        st.markdown("---")
        with st.expander("⚙️ Quản lý dữ liệu lịch sử nâng cao"):
            st.markdown("Tại đây bạn có thể xuất toàn bộ dữ liệu hoặc xóa toàn bộ lịch sử nếu cần thiết.")
            
            danger_col1, danger_col2 = st.columns([2, 3])
            with danger_col1:
                if st.button("⚠️ Xóa toàn bộ lịch sử", type="secondary", width="stretch"):
                    database.delete_all_history()
                    st.toast("Đã dọn dẹp sạch toàn bộ lịch sử phân tích!", icon=":material/delete_forever:")
                    st.rerun()
            with danger_col2:
                st.caption("Lưu ý: Hành động xóa toàn bộ sẽ không thể hoàn tác.")


# ============================================================
# TRANG 3: GIỚI THIỆU ĐỀ TÀI & KIẾN TRÚC HỆ THỐNG
# ============================================================
else:
    st.markdown(
        """
        <div class="main-hero" style="background: linear-gradient(135deg, #3730a3 0%, #4f46e5 50%, #6366f1 100%); box-shadow: 0 4px 20px -2px rgba(79, 70, 229, 0.25);">
            <h1>ℹ️ Giới thiệu Đề tài Đồ án Chuyên ngành</h1>
            <p>Hệ thống Xử lý Ngôn ngữ Tự nhiên Tiếng Việt kết hợp Streamlit, PhoBERT và SQLite Database.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        st.markdown(
            """
            ### 📌 Thông tin Đồ án
            - **Tên đề tài:** Xây dựng ứng dụng NLP tiếng Việt dùng Streamlit
            - **Ngành đào tạo:** Kỹ thuật Phần mềm / Công nghệ Thông tin
            - **Mục tiêu:** Xây dựng giải pháp phần mềm hiện đại, trực quan, phục vụ việc tiếp nhận dữ liệu văn bản từ nhiều nguồn (nhập tay, tệp Word, PDF, TXT, CSV), trích xuất và chọn lọc đoạn văn bản tùy ý, áp dụng các kỹ thuật NLP và mô hình Deep Learning để phân tích cảm xúc, từ khóa, tần suất, tóm tắt và quản lý lịch sử bền vững trên SQLite.

            ---

            ### 🎯 Mục tiêu đề tài và chức năng của ứng dụng

            #### 1. Mục tiêu đề tài
            - **Nghiên cứu & Ứng dụng NLP tiếng Việt:** Khảo sát và tích hợp các kỹ thuật xử lý ngôn ngữ tự nhiên hiện đại cùng mô hình học sâu PhoBERT tối ưu cho tiếng Việt để giải quyết bài toán phân tích sắc thái cảm xúc, trích xuất từ khóa và tóm tắt văn bản.
            - **Xây dựng giải pháp phần mềm hoàn chỉnh:** Phát triển ứng dụng Web trực quan, thân thiện bằng Python và Streamlit, cho phép tiếp nhận dữ liệu linh hoạt từ nhập trực tiếp và đọc đa dạng tệp tin (Word, PDF, TXT, CSV).
            - **Quản lý dữ liệu thông minh & bền vững:** Tích hợp cơ sở dữ liệu SQLite giúp lưu trữ an toàn, tra cứu nhanh, tìm kiếm và cung cấp khả năng xóa riêng lẻ từng bản ghi lịch sử phân tích.

            #### 2. Các chức năng chính của ứng dụng
            - **Tiếp nhận dữ liệu đa kênh:**
              - *Nhập văn bản trực tiếp:* Hỗ trợ soạn thảo tự do kèm 3 mẫu văn bản gợi ý nhanh (khen ngợi, phàn nàn, thông tin).
              - *Tải lên từ tập tin:* Hỗ trợ đọc tệp tin Word (`.docx`), tài liệu PDF (`.pdf`), tệp văn bản thuần (`.txt`, `.md`) và dữ liệu bảng (`.csv`).
            - **Phân đoạn thông minh & Phân tích linh hoạt:**
              - *Phân tích toàn bộ tệp:* Tự động bóc tách và phân tích toàn diện tất cả các đoạn văn có trong tệp tin, không bị cắt cụt.
              - *Chọn đoạn tùy ý:* Duyệt và chọn bất kỳ đoạn văn bản nào trong tệp để phân tích, sau đó có thể tiếp tục chọn đoạn khác để phân tích liên tục mà không cần tải lại tệp.
            - **Phân tích NLP chuyên sâu:**
              - *Thống kê định lượng:* Đo lường chính xác số ký tự, số từ, số câu, số từ khác nhau và số từ có nội dung.
              - *Phân tích cảm xúc (Sentiment Analysis):* Dự đoán nhãn cảm xúc (Tích cực, Tiêu cực, Trung tính) kèm độ tin cậy bằng PhoBERT. Khi phân tích toàn bộ tệp, hệ thống tự động phân tích từng đoạn và tổng hợp phân bố cảm xúc tổng thể.
              - *Trích xuất từ khóa (Keyword Extraction):* Chấm điểm và xếp hạng Top 10 từ khóa cốt lõi theo tần suất có trọng số.
              - *Tóm tắt trích xuất (Extractive Summarization):* Trích xuất các câu tiêu biểu mang thông tin quan trọng nhất từ văn bản gốc.
            - **Cơ sở dữ liệu & Quản lý lịch sử phân tích:**
              - Tự động lưu trữ chi tiết từng lượt phân tích vào SQLite (`history.db`).
              - Thanh tìm kiếm nội dung và bộ lọc phân loại cảm xúc.
              - **Chức năng xóa 1 bản ghi riêng lẻ:** Xóa đúng bản ghi được chỉ định mà không ảnh hưởng đến các dữ liệu khác.
              - Nạp lại văn bản từ lịch sử trở lại trang phân tích.
              - Xuất toàn bộ kết quả phân tích ra tệp CSV.

            ---

            ### 🛠️ Công nghệ & Thư viện sử dụng
            - **Ngôn ngữ:** Python 3.12
            - **Giao diện Web:** Streamlit 1.61+ (Clean theme, container cards, segmented control)
            - **Mô hình NLP:** Transformers + PhoBERT (`wonrax/phobert-base-vietnamese-sentiment`)
            - **Cơ sở dữ liệu:** SQLite3 (Native, zero-configuration)
            - **Xử lý tệp tin:** `python-docx` (Word), `pypdf` (PDF), `csv`, `re` (Unicode Regular Expressions)
            """
        )
