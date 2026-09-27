import sqlite3
import json
from datetime import datetime
import os

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "history.db")


def get_connection():
    """Tạo kết nối đến cơ sở dữ liệu SQLite."""
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """Khởi tạo bảng lưu lịch sử phân tích nếu chưa tồn tại."""
    conn = get_connection()
    try:
        with conn:
            conn.execute(
                """
                CREATE TABLE IF NOT EXISTS analysis_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    created_at TEXT NOT NULL,
                    source_type TEXT NOT NULL,
                    text_content TEXT NOT NULL,
                    char_count INTEGER DEFAULT 0,
                    word_count INTEGER DEFAULT 0,
                    sentence_count INTEGER DEFAULT 0,
                    unique_words INTEGER DEFAULT 0,
                    sentiment_label TEXT DEFAULT 'Chưa xác định',
                    sentiment_score REAL DEFAULT 0.0,
                    sentiment_emoji TEXT DEFAULT '🔍',
                    keywords_json TEXT DEFAULT '[]',
                    summary_text TEXT DEFAULT ''
                )
                """
            )
            # Tạo index để tìm kiếm và sắp xếp nhanh hơn
            conn.execute(
                "CREATE INDEX IF NOT EXISTS idx_history_created_at ON analysis_history(id DESC)"
            )
    finally:
        conn.close()


def insert_history(
    source_type: str,
    text_content: str,
    char_count: int,
    word_count: int,
    sentence_count: int,
    unique_words: int,
    sentiment_label: str = "Chưa xác định",
    sentiment_score: float = 0.0,
    sentiment_emoji: str = "🔍",
    keywords: list = None,
    summary_text: str = "",
) -> int:
    """Lưu một kết quả phân tích vào cơ sở dữ liệu và trả về ID."""
    init_db()
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    keywords_json = json.dumps(keywords or [], ensure_ascii=False)

    conn = get_connection()
    try:
        with conn:
            cursor = conn.execute(
                """
                INSERT INTO analysis_history (
                    created_at, source_type, text_content,
                    char_count, word_count, sentence_count, unique_words,
                    sentiment_label, sentiment_score, sentiment_emoji,
                    keywords_json, summary_text
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """,
                (
                    now_str,
                    source_type,
                    text_content,
                    char_count,
                    word_count,
                    sentence_count,
                    unique_words,
                    sentiment_label,
                    sentiment_score,
                    sentiment_emoji,
                    keywords_json,
                    summary_text,
                ),
            )
            return cursor.lastrowid
    finally:
        conn.close()


def get_all_history(search_query: str = None, sentiment_filter: str = None, limit: int = 100):
    """
    Lấy danh sách lịch sử phân tích, hỗ trợ tìm kiếm và lọc theo cảm xúc.
    """
    init_db()
    conn = get_connection()
    try:
        query = "SELECT * FROM analysis_history WHERE 1=1"
        params = []

        if search_query:
            query += " AND (text_content LIKE ? OR source_type LIKE ?)"
            wildcard = f"%{search_query.strip()}%"
            params.extend([wildcard, wildcard])

        if sentiment_filter and sentiment_filter != "Tất cả":
            query += " AND sentiment_label = ?"
            params.append(sentiment_filter)

        query += " ORDER BY id DESC LIMIT ?"
        params.append(limit)

        cursor = conn.execute(query, params)
        rows = cursor.fetchall()

        results = []
        for r in rows:
            results.append(
                {
                    "id": r["id"],
                    "created_at": r["created_at"],
                    "source_type": r["source_type"],
                    "text_content": r["text_content"],
                    "char_count": r["char_count"],
                    "word_count": r["word_count"],
                    "sentence_count": r["sentence_count"],
                    "unique_words": r["unique_words"],
                    "sentiment_label": r["sentiment_label"],
                    "sentiment_score": r["sentiment_score"],
                    "sentiment_emoji": r["sentiment_emoji"],
                    "keywords": json.loads(r["keywords_json"] or "[]"),
                    "summary_text": r["summary_text"],
                }
            )
        return results
    finally:
        conn.close()


def get_history_by_id(item_id: int):
    """Lấy chi tiết một mục lịch sử theo ID."""
    init_db()
    conn = get_connection()
    try:
        cursor = conn.execute("SELECT * FROM analysis_history WHERE id = ?", (item_id,))
        r = cursor.fetchone()
        if not r:
            return None
        return {
            "id": r["id"],
            "created_at": r["created_at"],
            "source_type": r["source_type"],
            "text_content": r["text_content"],
            "char_count": r["char_count"],
            "word_count": r["word_count"],
            "sentence_count": r["sentence_count"],
            "unique_words": r["unique_words"],
            "sentiment_label": r["sentiment_label"],
            "sentiment_score": r["sentiment_score"],
            "sentiment_emoji": r["sentiment_emoji"],
            "keywords": json.loads(r["keywords_json"] or "[]"),
            "summary_text": r["summary_text"],
        }
    finally:
        conn.close()


def delete_history_item(item_id: int) -> bool:
    """
    Xóa một bản ghi lịch sử phân tích cụ thể theo ID.
    Trả về True nếu xóa thành công, False nếu không tìm thấy.
    """
    init_db()
    conn = get_connection()
    try:
        with conn:
            cursor = conn.execute("DELETE FROM analysis_history WHERE id = ?", (item_id,))
            return cursor.rowcount > 0
    finally:
        conn.close()


def delete_all_history() -> bool:
    """Xóa toàn bộ lịch sử phân tích."""
    init_db()
    conn = get_connection()
    try:
        with conn:
            conn.execute("DELETE FROM analysis_history")
            return True
    finally:
        conn.close()


def get_history_stats() -> dict:
    """Thống kê tổng quan số liệu từ database."""
    init_db()
    conn = get_connection()
    try:
        cursor = conn.execute(
            """
            SELECT 
                COUNT(*) as total,
                SUM(CASE WHEN sentiment_label = 'Tích cực' THEN 1 ELSE 0 END) as pos_count,
                SUM(CASE WHEN sentiment_label = 'Tiêu cực' THEN 1 ELSE 0 END) as neg_count,
                SUM(CASE WHEN sentiment_label = 'Trung tính' THEN 1 ELSE 0 END) as neu_count
            FROM analysis_history
            """
        )
        row = cursor.fetchone()
        return {
            "total": row["total"] or 0,
            "pos": row["pos_count"] or 0,
            "neg": row["neg_count"] or 0,
            "neu": row["neu_count"] or 0,
        }
    finally:
        conn.close()
