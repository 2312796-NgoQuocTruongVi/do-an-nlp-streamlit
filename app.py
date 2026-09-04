import re
import io
import csv
from collections import Counter

import streamlit as st

# ============================================================
# CẤU HÌNH TRANG
# ============================================================
st.set_page_config(
    page_title="Ứng dụng NLP tiếng Việt",
    page_icon="🧠",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ============================================================
# CSS GIAO DIỆN
# ============================================================
st.markdown(
    """
    <style>
    .main-title {
        font-size: 2.5rem;
        font-weight: 800;
        margin-bottom: 0.2rem;
    }
    .subtitle {
        color: #6b7280;
        font-size: 1.05rem;
        margin-bottom: 1.5rem;
    }
    .result-card {
        padding: 1rem 1.2rem;
        border-radius: 12px;
        border: 1px solid rgba(128,128,128,.25);
        margin: .5rem 0;
    }
    .small-note {
        color: #6b7280;
        font-size: .9rem;
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
    """
    Trích xuất từ khóa theo tần suất có trọng số.
    Đây là phương pháp nhẹ, không cần huấn luyện model.
    """
    freq = get_word_frequencies(text)
    if not freq:
        return []

    max_count = max(freq.values())
    scored = []
    for word, count in freq.items():
        score = count / max_count
        # Ưu tiên nhẹ các từ dài hơn vì thường mang nhiều thông tin hơn.
        if len(word) >= 6:
            score *= 1.15
        scored.append((word, count, score))

    scored.sort(key=lambda x: (x[2], x[1]), reverse=True)
    return scored[:top_n]


def extractive_summary(text: str, max_sentences: int = 3):
    """
    Tóm tắt trích xuất đơn giản:
    - chấm điểm câu dựa trên từ khóa quan trọng
    - giữ nguyên các câu nổi bật từ văn bản gốc
    """
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
        # Ưu tiên nhẹ câu đầu vì thường chứa chủ đề.
        if idx == 0:
            score *= 1.10
        sentence_scores.append((idx, score, sentence))

    best = sorted(sentence_scores, key=lambda x: x[1], reverse=True)[:max_sentences]
    best = sorted(best, key=lambda x: x[0])
    return [item[2] for item in best]


# ============================================================
# MODEL PHÂN TÍCH CẢM XÚC
# ============================================================
@st.cache_resource(show_spinner="Đang tải mô hình NLP tiếng Việt... Lần đầu có thể mất vài phút.")
def load_sentiment_model():
    from transformers import pipeline

    return pipeline(
        "sentiment-analysis",
        model="wonrax/phobert-base-vietnamese-sentiment",
    )


def analyze_sentiment(text: str):
    classifier = load_sentiment_model()
    result = classifier(
        text,
        truncation=True,
        max_length=256,
    )[0]

    label = str(result["label"]).upper()
    score = float(result["score"])

    mapping = {
        "NEG": ("😡", "Tiêu cực", "error"),
        "NEU": ("😐", "Trung tính", "info"),
        "POS": ("😊", "Tích cực", "success"),
        "LABEL_0": ("😡", "Tiêu cực", "error"),
        "LABEL_1": ("😐", "Trung tính", "info"),
        "LABEL_2": ("😊", "Tích cực", "success"),
    }

    emoji, vietnamese_label, box_type = mapping.get(
        label,
        ("🔎", label, "info"),
    )

    return {
        "raw_label": label,
        "label": vietnamese_label,
        "emoji": emoji,
        "score": score,
        "box_type": box_type,
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
# SESSION STATE
# ============================================================
if "history" not in st.session_state:
    st.session_state.history = []

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.header("🧭 MENU")
    page = st.radio(
        "Chọn chức năng",
        [
            "🏠 Phân tích văn bản",
            "📊 Lịch sử phân tích",
            "ℹ️ Giới thiệu đề tài",
        ],
    )

    st.divider()
    st.caption("Đồ án chuyên ngành CNTT")
    st.caption("Python + Streamlit + Transformers")

# ============================================================
# TRANG CHÍNH
# ============================================================
if page == "🏠 Phân tích văn bản":
    st.markdown('<div class="main-title">🧠 Ứng dụng NLP tiếng Việt</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="subtitle">Phân tích văn bản, từ khóa, thống kê và cảm xúc bằng Python + Streamlit.</div>',
        unsafe_allow_html=True,
    )

    text = st.text_area(
        "✍️ Nhập văn bản cần phân tích",
        height=220,
        placeholder=(
            "Ví dụ: Tôi rất thích ứng dụng này. "
            "Giao diện đẹp, dễ sử dụng và tốc độ xử lý khá nhanh."
        ),
    )

    col_a, col_b = st.columns([1, 5])
    with col_a:
        analyze = st.button("🔍 Phân tích", type="primary", use_container_width=True)
    with col_b:
        st.caption("Bạn có thể nhập một câu hoặc một đoạn văn tiếng Việt.")

    if analyze:
        clean_text = normalize_text(text)

        if not clean_text:
            st.warning("⚠️ Vui lòng nhập văn bản trước khi phân tích.")
            st.stop()

        words = tokenize_words(clean_text)
        sentences = split_sentences(clean_text)
        freq = get_word_frequencies(clean_text)
        keywords = extract_keywords(clean_text, top_n=10)

        stats = {
            "Số ký tự": len(clean_text),
            "Số từ": len(words),
            "Số câu": len(sentences),
            "Số từ khác nhau": len(set(words)),
            "Từ có nội dung": len(freq),
        }

        # Lưu kết quả cơ bản vào lịch sử.
        history_item = {
            "text": clean_text[:500],
            "characters": stats["Số ký tự"],
            "words": stats["Số từ"],
            "sentences": stats["Số câu"],
        }

        st.subheader("📊 Thống kê văn bản")
        c1, c2, c3, c4, c5 = st.columns(5)
        for col, (label, value) in zip(
            [c1, c2, c3, c4, c5],
            stats.items(),
        ):
            with col:
                st.metric(label, value)

        st.divider()

        tab1, tab2, tab3, tab4 = st.tabs(
            [
                "😊 Cảm xúc",
                "🔑 Từ khóa",
                "📈 Tần suất",
                "📝 Tóm tắt",
            ]
        )

        sentiment = None

        with tab1:
            try:
                sentiment = analyze_sentiment(clean_text)

                if sentiment["box_type"] == "success":
                    st.success(
                        f'{sentiment["emoji"]} {sentiment["label"]} — '
                        f'Độ tin cậy: {sentiment["score"]:.2%}'
                    )
                elif sentiment["box_type"] == "error":
                    st.error(
                        f'{sentiment["emoji"]} {sentiment["label"]} — '
                        f'Độ tin cậy: {sentiment["score"]:.2%}'
                    )
                else:
                    st.info(
                        f'{sentiment["emoji"]} {sentiment["label"]} — '
                        f'Độ tin cậy: {sentiment["score"]:.2%}'
                    )

                st.progress(sentiment["score"])
                st.caption(
                    "Kết quả được dự đoán bởi mô hình sentiment tiếng Việt "
                    "đã được huấn luyện trước."
                )

            except Exception as exc:
                st.error(
                    "Không thể tải/chạy mô hình cảm xúc. "
                    "Hãy kiểm tra Internet và các thư viện trong requirements.txt."
                )
                st.exception(exc)

        with tab2:
            if keywords:
                st.write("**Các từ khóa nổi bật:**")
                keyword_table = [
                    {
                        "Từ khóa": word,
                        "Tần suất": count,
                        "Điểm": round(score, 4),
                    }
                    for word, count, score in keywords
                ]
                st.dataframe(
                    keyword_table,
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("Không tìm thấy từ khóa phù hợp.")

        with tab3:
            if freq:
                chart_data = dict(freq.most_common(10))
                st.bar_chart(chart_data)
                st.write("**Top 10 từ xuất hiện nhiều nhất:**")
                st.write(chart_data)
            else:
                st.info("Không có dữ liệu tần suất.")

        with tab4:
            summary = extractive_summary(clean_text, max_sentences=3)
            if summary:
                st.write("**Bản tóm tắt trích xuất:**")
                st.info(" ".join(summary))
            else:
                st.info("Chưa đủ dữ liệu để tạo tóm tắt.")

        # Export kết quả.
        csv_data = make_csv(stats, sentiment, keywords)
        st.download_button(
            "⬇️ Xuất kết quả CSV",
            data=csv_data,
            file_name="ket_qua_phan_tich_nlp.csv",
            mime="text/csv",
        )

        # Cập nhật lịch sử sau khi phân tích.
        history_item["sentiment"] = sentiment["label"] if sentiment else "Chưa xác định"
        history_item["confidence"] = sentiment["score"] if sentiment else 0
        st.session_state.history.insert(0, history_item)
        st.session_state.history = st.session_state.history[:20]


elif page == "📊 Lịch sử phân tích":
    st.title("📊 Lịch sử phân tích")

    if not st.session_state.history:
        st.info("Chưa có lượt phân tích nào trong phiên làm việc này.")
    else:
        st.caption("Lịch sử được lưu trong phiên Streamlit hiện tại.")
        for idx, item in enumerate(st.session_state.history, start=1):
            with st.expander(
                f"#{idx} — {item['sentiment']} — "
                f"{item['confidence']:.2%}"
            ):
                st.write(item["text"])
                a, b, c = st.columns(3)
                a.metric("Ký tự", item["characters"])
                b.metric("Từ", item["words"])
                c.metric("Câu", item["sentences"])

        if st.button("🗑️ Xóa lịch sử"):
            st.session_state.history = []
            st.rerun()


else:
    st.title("ℹ️ Giới thiệu đề tài")

    st.markdown(
        """
        ### Tên đề tài
        **Xây dựng ứng dụng NLP dùng Streamlit**

        ### Mục đích
        Ứng dụng được xây dựng nhằm minh họa khả năng kết hợp **Python,
        Streamlit và các mô hình xử lý ngôn ngữ tự nhiên** để tiếp nhận,
        phân tích và trực quan hóa dữ liệu văn bản tiếng Việt.

        ### Các chức năng chính
        - Nhập văn bản tiếng Việt từ người dùng.
        - Thống kê số ký tự, số từ, số câu và số từ khác nhau.
        - Phân tích cảm xúc bằng mô hình NLP tiếng Việt.
        - Trích xuất các từ khóa nổi bật.
        - Thống kê tần suất từ và biểu đồ trực quan.
        - Tóm tắt trích xuất dựa trên nội dung văn bản.
        - Lưu lịch sử phân tích trong phiên làm việc.
        - Xuất kết quả phân tích ra file CSV.

        ### Công nghệ sử dụng
        - **Python**: ngôn ngữ lập trình chính.
        - **Streamlit**: xây dựng giao diện web.
        - **Transformers**: tích hợp mô hình NLP.
        - **PhoBERT-based sentiment model**: phân tích cảm xúc tiếng Việt.

        ### Luồng xử lý
        `Người dùng nhập văn bản → Tiền xử lý → Phân tích NLP → Trực quan hóa → Xuất kết quả`

        ### Lưu ý
        Kết quả phân tích cảm xúc là dự đoán của mô hình học máy,
        vì vậy không nên xem là kết luận tuyệt đối trong mọi trường hợp.
        """
    )
