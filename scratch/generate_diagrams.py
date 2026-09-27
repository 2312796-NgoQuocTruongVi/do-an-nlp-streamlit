import matplotlib.pyplot as plt
import matplotlib.patches as patches
import numpy as np

plt.rcParams['font.sans-serif'] = 'DejaVu Sans'
plt.rcParams['axes.edgecolor'] = '#CCCCCC'
plt.rcParams['axes.linewidth'] = 0.8

# 1. Architecture Diagram NÂNG CẤP
fig, ax = plt.subplots(figsize=(10.5, 7.2), dpi=300)
ax.set_xlim(0, 10.5)
ax.set_ylim(0, 8.2)
ax.axis('off')

# Title
ax.text(5.25, 7.8, "KIẾN TRÚC PHÂN TẦNG HỆ THỐNG VIETNLP STUDIO (4 TẦNG MODULE HÓA)", 
        fontsize=13, fontweight='bold', ha='center', color='#1E3A8A')

layers = [
    ("TẦNG TRÌNH DIỄN (PRESENTATION & UI LAYER - app.py)", 
     ["Giao diện Streamlit Wide Layout, Hero Banner gradient", "Menu Sidebar tùy biến bo góc, không ô tick tròn", 
      "Bộ điều khiển phân đoạn Segmented Control, 3 mẫu văn bản nhanh", "Dashboard KPI thẻ Metric, Tabs Cảm xúc (Breakdown) - Từ khóa - Tóm tắt"],
     6.1, "#1D4ED8"),
    ("TẦNG XỬ LÝ TỆP TIN & PHÂN ĐOẠN (FILE & SEGMENTATION LAYER - file_handler.py)", 
     ["Đọc đa định dạng: .docx (Word), .pdf (pypdf), .txt, .md, .csv", "Giải mã đa encoding fallback (utf-8-sig, utf-16, cp1252)",
      "Bóc tách phân đoạn văn bản (extract_paragraphs_from_text)", "Duyệt đoạn trước/sau & Phân tích đoạn tùy ý không cần tải lại tệp"],
     4.4, "#0D9488"),
    ("TẦNG NGHIỆP VỤ NLP & HỌC SÂU (NLP & DEEP LEARNING LAYER - app.py)", 
     ["Chuẩn hóa Unicode, Tách từ/câu Regex tiếng Việt, Lọc 58 từ dừng", "Trích xuất 10 từ khóa trọng số độ dài & Tóm tắt trích xuất câu",
      "Mô hình PhoBERT-base-vietnamese-sentiment (Transformers)", "Thuật toán Phân đoạn Ngữ nghĩa & Tổng hợp Cảm xúc (Chunking & Majority Voting)"],
     2.7, "#6366F1"),
    ("TẦNG DỮ LIỆU & LƯU TRỮ BỀN VỮNG (PERSISTENT DATA LAYER - database.py)", 
     ["Hệ quản trị CSDL SQLite (history.db), Bảng analysis_history 13 trường", "Đánh chỉ mục idx_history_created_at tối ưu hóa truy vấn",
      "CRUD Lịch sử: Thêm mới, Tìm kiếm LIKE, Lọc cảm xúc, Thống kê tổng hợp", "Chức năng XÓA RIÊNG LẺ TỪNG BẢN GHI cụ thể & Nạp lại văn bản phân tích"],
     1.0, "#B45309")
]

for title, items, y, color in layers:
    rect = patches.FancyBboxPatch((0.5, y), 9.5, 1.35, boxstyle="round,pad=0.1", 
                                  edgecolor=color, facecolor=color, alpha=0.08, linewidth=1.6)
    ax.add_patch(rect)
    ax.text(5.25, y + 1.05, title, fontsize=9.8, fontweight='bold', ha='center', color=color)
    content_str = "  •  ".join(items[:2]) + "\n" + "  •  ".join(items[2:])
    ax.text(5.25, y + 0.45, content_str, fontsize=8.3, ha='center', color='#1E293B', linespacing=1.4)

# Draw arrows between layers
for y_arrow in [6.1, 4.4, 2.7]:
    ax.annotate('', xy=(5.25, y_arrow), xytext=(5.25, y_arrow + 0.35),
                arrowprops=dict(arrowstyle="->,head_width=0.35,head_length=0.5", lw=1.5, color='#64748B'))

plt.tight_layout()
plt.savefig('scratch/figures/architecture_diagram.png', bbox_inches='tight')
plt.close()

# 2. Pipeline Diagram NÂNG CẤP
fig, ax = plt.subplots(figsize=(11.5, 5.0), dpi=300)
ax.set_xlim(0, 11.5)
ax.set_ylim(0, 5.0)
ax.axis('off')

ax.text(5.75, 4.6, "LUỒNG XỬ LÝ DỮ LIỆU ĐA KÊNH & QUY TRÌNH PHÂN TÍCH NLP TOÀN DIỆN", 
        fontsize=12.5, fontweight='bold', ha='center', color='#1E3A8A')

pipeline_steps = [
    ("1. Thu nhận đa kênh\n(Text/Word/PDF/CSV)", 1.0, "#EA580C"),
    ("2. Bóc tách & Phân đoạn\n(Paragraph Chunking)", 3.0, "#0D9488"),
    ("3. Chọn phạm vi\n(Toàn bộ / Đoạn tùy ý)", 5.0, "#2563EB"),
    ("4. Phân tích NLP\n(PhoBERT + Keywords)", 7.1, "#7C3AED"),
    ("5. Lưu SQLite bền vững\n(history.db + ID)", 9.2, "#059669"),
    ("6. Trực quan & CSV\n(Dashboard Metrics)", 10.8, "#D97706")
]

for label, x, col in pipeline_steps:
    rect = patches.FancyBboxPatch((x - 0.85, 2.2), 1.7, 1.3, boxstyle="round,pad=0.08",
                                  edgecolor=col, facecolor=col, alpha=0.12, linewidth=1.5)
    ax.add_patch(rect)
    ax.text(x, 2.85, label, fontsize=8.2, fontweight='bold', ha='center', va='center', color=col)

for i in range(len(pipeline_steps) - 1):
    x1 = pipeline_steps[i][1] + 0.85
    x2 = pipeline_steps[i+1][1] - 0.85
    ax.annotate('', xy=(x2, 2.85), xytext=(x1, 2.85),
                arrowprops=dict(arrowstyle="->,head_width=0.3,head_length=0.4", lw=1.5, color='#94A3B8'))

# Sub details
ax.text(5.75, 0.9, 
        "• Cơ chế Semantic Chunking (max_words=150): Phân tích cảm xúc từng đoạn cho file dài, không bị cắt cụt\n"
        "• Lưu trữ bền vững vào SQLite: Hỗ trợ tìm kiếm, lọc cảm xúc, xóa từng bản ghi riêng lẻ, nạp lại phân tích",
        fontsize=8.2, ha='center', color='#334155', 
        bbox=dict(boxstyle="round,pad=0.5", facecolor="#F8FAFC", edgecolor="#CBD5E1", linewidth=1.2))

plt.tight_layout()
plt.savefig('scratch/figures/nlp_pipeline_diagram.png', bbox_inches='tight')
plt.close()

# 3. Use Case Diagram NÂNG CẤP
fig, ax = plt.subplots(figsize=(10.5, 7.5), dpi=300)
ax.set_xlim(0, 10.5)
ax.set_ylim(0, 8.0)
ax.axis('off')

ax.text(5.25, 7.6, "SƠ ĐỒ USE CASE NÂNG CẤP HỆ THỐNG VIETNLP STUDIO", 
        fontsize=13, fontweight='bold', ha='center', color='#1E3A8A')

# System Boundary
boundary = patches.FancyBboxPatch((2.8, 0.4), 7.2, 6.9, boxstyle="round,pad=0.1",
                                 edgecolor='#2563EB', facecolor='#F8FAFC', linewidth=1.5, linestyle='--')
ax.add_patch(boundary)
ax.text(6.4, 7.0, "Hệ Thống VietNLP Studio (Streamlit + PhoBERT + SQLite)", fontsize=10, fontweight='bold', color='#2563EB', ha='center')

# Actor
ax.scatter([1.2], [4.0], s=400, color='#2563EB', zorder=5) # Head
ax.plot([1.2, 1.2], [3.8, 2.8], color='#2563EB', lw=3) # Body
ax.plot([0.6, 1.8], [3.4, 3.4], color='#2563EB', lw=3) # Arms
ax.plot([1.2, 0.7], [2.8, 2.0], color='#2563EB', lw=3) # Left Leg
ax.plot([1.2, 1.7], [2.8, 2.0], color='#2563EB', lw=3) # Right Leg
ax.text(1.2, 1.6, "Người dùng\n(End User)", fontsize=9.5, fontweight='bold', ha='center', color='#1E3A8A')

use_cases = [
    ("UC01: Nhập văn bản trực tiếp hoặc chọn mẫu gợi ý nhanh", 6.4),
    ("UC02: Tải tệp tin đa định dạng (.docx, .pdf, .txt, .csv)", 5.7),
    ("UC03: Bóc tách phân đoạn & Chọn đoạn tùy ý để phân tích", 5.0),
    ("UC04: Phân tích toàn bộ nội dung tệp (Multi-chunk Aggregation)", 4.3),
    ("UC05: Thống kê định lượng & Trích xuất từ khóa, Tóm tắt", 3.6),
    ("UC06: Tự động lưu trữ kết quả phân tích vào CSDL SQLite", 2.9),
    ("UC07: Tra cứu, tìm kiếm & Lọc lịch sử theo sắc thái cảm xúc", 2.2),
    ("UC08: Xóa riêng lẻ từng bản ghi cụ thể hoặc xóa toàn bộ", 1.5),
    ("UC09: Nạp lại văn bản từ lịch sử & Xuất báo cáo CSV", 0.8),
]

for uc_text, y in use_cases:
    ellipse = patches.Ellipse((6.4, y), 6.2, 0.52, edgecolor='#0D9488', facecolor='#F0FDFA', linewidth=1.2)
    ax.add_patch(ellipse)
    ax.text(6.4, y, uc_text, fontsize=8.2, ha='center', va='center', color='#134E4A', fontweight='bold')
    ax.plot([1.8, 3.3], [3.4, y], color='#94A3B8', lw=1.1, linestyle=':')

plt.tight_layout()
plt.savefig('scratch/figures/usecase_diagram.png', bbox_inches='tight')
plt.close()

# 4. Database ER Diagram (MỚI - Dùng text chuẩn không lỗi emoji)
fig, ax = plt.subplots(figsize=(9.5, 6.2), dpi=300)
ax.set_xlim(0, 9.5)
ax.set_ylim(0, 6.5)
ax.axis('off')

ax.text(4.75, 6.15, "SƠ ĐỒ CƠ SỞ DỮ LIỆU SQLITE (ENTITY RELATIONSHIP & SCHEMA)", 
        fontsize=12.5, fontweight='bold', ha='center', color='#1E3A8A')

# Table Box
rect = patches.FancyBboxPatch((1.0, 0.4), 7.5, 5.4, boxstyle="round,pad=0.08",
                              edgecolor='#0D9488', facecolor='#FFFFFF', linewidth=1.8)
ax.add_patch(rect)

# Table Header
header_rect = patches.Rectangle((1.0, 5.15), 7.5, 0.65, facecolor='#0D9488')
ax.add_patch(header_rect)
ax.text(4.75, 5.45, "BẢNG: analysis_history (history.db)", fontsize=11, fontweight='bold', ha='center', color='#FFFFFF')

schema_fields = [
    ("[PK] id", "INTEGER", "PRIMARY KEY AUTOINCREMENT (Mã định danh duy nhất)"),
    ("created_at", "TEXT", "Thời gian lưu phân tích (dd/mm/yyyy HH:MM:SS)"),
    ("source_type", "TEXT", "Nguồn dữ liệu (Nhập trực tiếp, Tệp Word/PDF/CSV)"),
    ("text_content", "TEXT", "Nội dung văn bản phân tích (UTF-8)"),
    ("char_count", "INTEGER", "Tổng số ký tự trong văn bản"),
    ("word_count", "INTEGER", "Tổng số từ tố tách được"),
    ("sentence_count", "INTEGER", "Tổng số câu trong đoạn văn"),
    ("unique_words", "INTEGER", "Số lượng từ vựng độc nhất"),
    ("sentiment_label", "TEXT", "Nhãn cảm xúc (Tích cực / Tiêu cực / Trung tính)"),
    ("sentiment_score", "REAL", "Độ tin cậy của mô hình (0.0 đến 1.0)"),
    ("sentiment_emoji", "TEXT", "Biểu tượng cảm xúc trực quan"),
    ("keywords_json", "TEXT", "Danh sách Top 10 từ khóa & điểm số (chuỗi JSON)"),
    ("summary_text", "TEXT", "Nội dung bản tóm tắt trích xuất thông minh"),
]

y_pos = 4.8
for field, f_type, desc in schema_fields:
    ax.text(1.3, y_pos, field, fontsize=8.5, fontweight='bold', color='#1E293B')
    ax.text(3.7, y_pos, f_type, fontsize=8.5, color='#2563EB', fontfamily='monospace')
    ax.text(5.1, y_pos, desc, fontsize=8.0, color='#64748B')
    y_pos -= 0.33

plt.tight_layout()
plt.savefig('scratch/figures/database_er_diagram.png', bbox_inches='tight')
plt.close()

# 5. Performance Comparison
fig, ax = plt.subplots(figsize=(8.5, 4.6), dpi=300)
categories = [
    'Tải mô hình lần đầu\n(Cold Start / Download)', 
    'Phân tích đoạn 50 từ\n(Cached Inference)', 
    'Phân tích đoạn 200 từ\n(Cached Inference)', 
    'Phân tích toàn bộ tệp 5 đoạn\n(Multi-chunk Aggregation)'
]
times = [35.0, 0.18, 0.28, 0.85]
colors = ['#EF4444', '#0D9488', '#2563EB', '#7C3AED']

bars = ax.bar(categories, times, color=colors, width=0.52, edgecolor='#1E293B', linewidth=0.8)
ax.set_ylabel('Thời gian thực thi (Giây - log scale)', fontsize=9.5, fontweight='bold', color='#1E3A8A')
ax.set_yscale('log')
ax.set_title('HIỆU NĂNG XỬ LÝ NLP & ĐA PHÂN ĐOẠN (CACHED SINGLETON INFERENCE)', fontsize=11, fontweight='bold', pad=15, color='#1E3A8A')
ax.grid(axis='y', linestyle='--', alpha=0.5)

for bar, val in zip(bars, times):
    y_pos = val * 1.2
    ax.text(bar.get_x() + bar.get_width()/2, y_pos, f'{val}s', ha='center', va='bottom', fontsize=9, fontweight='bold', color='#1E293B')

plt.tight_layout()
plt.savefig('scratch/figures/performance_comparison.png', bbox_inches='tight')
plt.close()

print("All updated diagrams regenerated without warnings!")
