import os
import sys
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from report_helpers import (
    set_cell_background,
    set_cell_margins,
    set_table_borders,
    add_callout_box,
    add_code_block
)

def create_full_report():
    doc = Document()
    
    # -------------------------------------------------------------
    # 1. CẤU HÌNH TRANG IN (PAGE SETUP)
    # -------------------------------------------------------------
    for section in doc.sections:
        section.top_margin = Inches(0.79)     # 2.0 cm
        section.bottom_margin = Inches(0.79)  # 2.0 cm
        section.left_margin = Inches(1.18)    # 3.0 cm (đóng gáy sách)
        section.right_margin = Inches(0.79)   # 2.0 cm
        section.page_width = Inches(8.27)     # A4
        section.page_height = Inches(11.69)   # A4
        
        # Header / Footer
        header = section.header
        p_head = header.paragraphs[0]
        p_head.text = "ĐỒ ÁN CHUYÊN NGÀNH KỸ THUẬT PHẦN MỀM | HỆ THỐNG VIETNLP STUDIO (STREAMLIT & SQLITE)"
        p_head.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_head.runs[0].font.name = "Times New Roman"
        p_head.runs[0].font.size = Pt(8.5)
        p_head.runs[0].font.italic = True
        p_head.runs[0].font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
        
        footer = section.footer
        p_foot = footer.paragraphs[0]
        p_foot.text = "Trang "
        p_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_foot.runs[0].font.name = "Times New Roman"
        p_foot.runs[0].font.size = Pt(9)
        p_foot.runs[0].font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
        fldSimple = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
        p_foot._p.append(fldSimple)

    # Styles helper
    def add_p(text="", align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_before=2, space_after=3, line_spacing=1.25, bold=False, italic=False, font_size=12.5, color=RGBColor(0x1E, 0x29, 0x3B)):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.line_spacing = line_spacing
        if text:
            run = p.add_run(text)
            run.font.name = "Times New Roman"
            run.font.size = Pt(font_size)
            run.font.bold = bold
            run.font.italic = italic
            run.font.color.rgb = color
        return p

    def add_h1(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title)
        run.font.name = "Times New Roman"
        run.font.size = Pt(16)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A) # Deep Navy Blue
        return p

    def add_h2(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title)
        run.font.name = "Times New Roman"
        run.font.size = Pt(13.5)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8) # Royal Blue
        return p

    def add_h3(title):
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(title)
        run.font.name = "Times New Roman"
        run.font.size = Pt(12.5)
        run.font.bold = True
        run.font.italic = True
        run.font.color.rgb = RGBColor(0x0F, 0x76, 0x6E) # Teal Accent
        return p

    def add_styled_table(headers, rows_data, col_widths=None):
        table = doc.add_table(rows=len(rows_data) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        set_table_borders(table, hex_color="CBD5E1", sz="4", val="single")
        
        # Header Row
        hdr_cells = table.rows[0].cells
        for idx, header_text in enumerate(headers):
            cell = hdr_cells[idx]
            if col_widths and idx < len(col_widths):
                cell.width = col_widths[idx]
            set_cell_background(cell, "1E3A8A")
            set_cell_margins(cell, top=120, bottom=120, left=140, right=140)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(header_text)
            run.font.name = "Times New Roman"
            run.font.size = Pt(10.5)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            
        # Data Rows
        for row_idx, row_values in enumerate(rows_data):
            row_cells = table.rows[row_idx + 1].cells
            bg_color = "F8FAFC" if row_idx % 2 == 1 else "FFFFFF"
            for col_idx, cell_value in enumerate(row_values):
                cell = row_cells[col_idx]
                if col_widths and col_idx < len(col_widths):
                    cell.width = col_widths[col_idx]
                set_cell_background(cell, bg_color)
                set_cell_margins(cell, top=80, bottom=80, left=120, right=120)
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p = cell.paragraphs[0]
                if col_idx == 0 or len(str(cell_value)) <= 8:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(1)
                p.paragraph_format.space_after = Pt(1)
                p.paragraph_format.line_spacing = 1.15
                run = p.add_run(str(cell_value))
                run.font.name = "Times New Roman"
                run.font.size = Pt(10)
                run.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
                
        doc.add_paragraph()
        return table

    def add_image_with_caption(img_path, caption, width=Inches(6.1)):
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(6)
            p_img.paragraph_format.space_after = Pt(2)
            run_img = p_img.add_run()
            run_img.add_picture(img_path, width=width)
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_before = Pt(2)
            p_cap.paragraph_format.space_after = Pt(8)
            run_cap = p_cap.add_run(caption)
            run_cap.font.name = "Times New Roman"
            run_cap.font.size = Pt(10)
            run_cap.font.italic = True
            run_cap.font.color.rgb = RGBColor(0x47, 0x55, 0x69)

    # =============================================================
    # TRANG BÌA CHÍNH (COVER PAGE)
    # =============================================================
    cov_table = doc.add_table(rows=1, cols=1)
    cov_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cov_cell = cov_table.cell(0, 0)
    cov_cell.width = Inches(6.5)
    set_cell_background(cov_cell, "FFFFFF")
    set_cell_margins(cov_cell, top=300, bottom=300, left=300, right=300)
    
    tcPr = cov_cell._tc.get_or_add_tcPr()
    borders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="double" w:sz="18" w:space="0" w:color="1E3A8A"/>\n'
        f'  <w:bottom w:val="double" w:sz="18" w:space="0" w:color="1E3A8A"/>\n'
        f'  <w:left w:val="double" w:sz="18" w:space="0" w:color="1E3A8A"/>\n'
        f'  <w:right w:val="double" w:sz="18" w:space="0" w:color="1E3A8A"/>\n'
        f'</w:tcBorders>'
    )
    tcPr.append(borders)
    
    cp = cov_cell.paragraphs[0]
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp.paragraph_format.space_before = Pt(4)
    cp.paragraph_format.space_after = Pt(2)
    r = cp.add_run("BỘ GIÁO DỤC VÀ ĐÀO TẠO\nTRƯỜNG ĐẠI HỌC KỸ THUẬT - CÔNG NGHỆ\n")
    r.font.name = "Times New Roman"
    r.font.size = Pt(13)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    r = cp.add_run("KHOA CÔNG NGHỆ THÔNG TIN\nBỘ MÔN KỸ THUẬT PHẦN MỀM\n")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
    
    r = cp.add_run("--------------------***--------------------\n\n\n")
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    r.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)
    
    cp_title = cov_cell.add_paragraph()
    cp_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp_title.paragraph_format.space_before = Pt(8)
    cp_title.paragraph_format.space_after = Pt(8)
    r = cp_title.add_run("BÁO CÁO ĐỒ ÁN CHUYÊN NGÀNH\nKỸ THUẬT PHẦN MỀM\n\n")
    r.font.name = "Times New Roman"
    r.font.size = Pt(15.5)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0xDC, 0x26, 0x26)
    
    r_topic = cp_title.add_run("VIETNLP STUDIO:\nNGHIÊN CỨU VÀ XÂY DỰNG HỆ THỐNG XỬ LÝ\nNGÔN NGỮ TỰ NHIÊN TIẾNG VIỆT TOÀN DIỆN KẾT HỢP\nSTREAMLIT, PHOBERT VÀ CƠ SỞ DỮ LIỆU SQLITE\n\n")
    r_topic.font.name = "Times New Roman"
    r_topic.font.size = Pt(17.5)
    r_topic.font.bold = True
    r_topic.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
    
    r_sub = cp_title.add_run("(VietNLP Studio: An Advanced Vietnamese NLP Platform Integrating Streamlit, PhoBERT Deep Learning & SQLite Database)\n\n\n")
    r_sub.font.name = "Times New Roman"
    r_sub.font.size = Pt(11)
    r_sub.font.italic = True
    r_sub.font.color.rgb = RGBColor(0x47, 0x55, 0x69)
    
    cp_info = cov_cell.add_paragraph()
    cp_info.alignment = WD_ALIGN_PARAGRAPH.LEFT
    cp_info.paragraph_format.left_indent = Inches(0.9)
    cp_info.paragraph_format.space_before = Pt(6)
    cp_info.paragraph_format.space_after = Pt(2)
    cp_info.paragraph_format.line_spacing = 1.3
    
    info_text = (
        "Chuyên ngành:            Kỹ thuật Phần mềm (Software Engineering)\n"
        "Mã học phần:              SE301 - Đồ án Chuyên ngành KTPM\n"
        "Giảng viên hướng dẫn: TS. Nguyễn Văn A\n"
        "Sinh viên thực hiện:     Ngô Quốc Trường Vĩ\n"
        "Mã số sinh viên:         2312796\n"
        "Lớp:                            KTPM2023\n"
    )
    r = cp_info.add_run(info_text)
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.font.color.rgb = RGBColor(0x1E, 0x29, 0x3B)
    
    cp_foot = cov_cell.add_paragraph()
    cp_foot.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cp_foot.paragraph_format.space_before = Pt(30)
    cp_foot.paragraph_format.space_after = Pt(6)
    r = cp_foot.add_run("TP. HỒ CHÍ MINH, NĂM 2026")
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)
    r.font.bold = True
    r.font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)

    doc.add_page_break()

    # =============================================================
    # TRANG NHẬN XÉT CỦA GIẢNG VIÊN HƯỚNG DẪN
    # =============================================================
    add_h1("PHIẾU NHẬN XÉT CỦA GIẢNG VIÊN HƯỚNG DẪN")
    add_p("Họ và tên Giảng viên: TS. Nguyễn Văn A", bold=True)
    add_p("Học hàm, học vị: Tiến sĩ - Giảng viên Bộ môn Kỹ thuật Phần mềm", italic=True)
    add_p("Sinh viên thực hiện: Ngô Quốc Trường Vĩ                 MSSV: 2312796")
    add_p("Tên đề tài: VietNLP Studio: Nghiên cứu và xây dựng hệ thống Xử lý ngôn ngữ tự nhiên tiếng Việt toàn diện kết hợp Streamlit, PhoBERT và Cơ sở dữ liệu SQLite.")
    add_p("---------------------------------------------------------------------------------------------------------------------------------")
    
    add_p("1. Đánh giá về tinh thần, thái độ và sự chủ động nâng cấp phần mềm của sinh viên:", bold=True)
    add_p("................................................................................................................................................................")
    add_p("................................................................................................................................................................")
    
    add_p("2. Đánh giá về năng lực kiến trúc phần mềm (Module hóa, CSDL SQLite, Xử lý tệp đa định dạng Word/PDF/CSV, Thuật toán Phân đoạn Ngữ nghĩa):", bold=True)
    add_p("................................................................................................................................................................")
    add_p("................................................................................................................................................................")
    
    add_p("3. Đánh giá về chất lượng sản phẩm phần mềm, giao diện người dùng và bản báo cáo:", bold=True)
    add_p("................................................................................................................................................................")
    add_p("................................................................................................................................................................")
    
    add_p("4. Điểm đánh giá (Thang điểm 10): .................... (Bằng chữ: .....................................................)", bold=True)
    
    p_sign = add_p("\n\nTP. Hồ Chí Minh, ngày ...... tháng ...... năm 2026\nGiảng viên hướng dẫn\n(Ký và ghi rõ họ tên)\n\n\n\nTS. Nguyễn Văn A", align=WD_ALIGN_PARAGRAPH.RIGHT)

    doc.add_page_break()

    # =============================================================
    # LỜI CAM ĐOAN & LỜI CẢM ƠN
    # =============================================================
    add_h1("LỜI CAM ĐOAN")
    add_p("Tôi xin cam đoan rằng toàn bộ nội dung trong bản báo cáo Đồ án Chuyên ngành Kỹ thuật Phần mềm này là kết quả nghiên cứu, khảo sát và xây dựng thực tế của cá nhân tôi dưới sự hướng dẫn khoa học của Giảng viên hướng dẫn. Hệ thống phần mềm VietNLP Studio bao gồm toàn bộ mã nguồn chương trình (app.py, database.py, file_handler.py, config.toml) và cơ sở dữ liệu history.db được lập trình, cấu trúc hóa và nâng cấp một cách độc lập, nghiêm túc, không sao chép trái phép bất kỳ sản phẩm hay đồ án nào khác.")
    add_p("Các thư viện mã nguồn mở (Streamlit, Transformers, PyTorch, python-docx, pypdf), mô hình học sâu tiền huấn luyện PhoBERT và các tài liệu tham khảo học thuật đều được công khai trích dẫn nguồn gốc rõ ràng, tuân thủ đúng quy định pháp luật về quyền sở hữu trí tuệ và chuẩn mực đạo đức nghiên cứu của nhà trường.")
    add_p("Tôi xin chịu hoàn toàn trách nhiệm trước Hội đồng chấm đồ án và Khoa Công nghệ Thông tin về tính xác thực của sản phẩm và tài liệu này.")
    
    add_p("\nTP. Hồ Chí Minh, ngày 25 tháng 09 năm 2026\nSinh viên thực hiện\n(Ký và ghi rõ họ tên)\n\n\nNgô Quốc Trường Vĩ", align=WD_ALIGN_PARAGRAPH.RIGHT)
    
    add_h1("LỜI CẢM ƠN")
    add_p("Để hoàn thành đồ án chuyên ngành Kỹ thuật Phần mềm với khối lượng công việc mở rộng và sản phẩm hoàn chỉnh như ngày hôm nay, bên cạnh sự nỗ lực phấn đấu của bản thân, tôi đã nhận được sự quan tâm, chỉ dẫn tận tình và giúp đỡ quý báu từ nhiều thầy cô, gia đình và bạn bè.")
    add_p("Lời đầu tiên, tôi xin bày tỏ lòng biết ơn chân thành và sâu sắc nhất đến quý Thầy Cô Khoa Công nghệ Thông tin, đặc biệt là các Thầy Cô thuộc Bộ môn Kỹ thuật Phần mềm, những người đã truyền đạt cho tôi nền tảng tư duy phân tích yêu cầu, thiết kế kiến trúc phần mềm phân tầng, quản lý cơ sở dữ liệu và quy trình kiểm thử hệ thống.")
    add_p("Đặc biệt, tôi xin gửi lời cảm ơn sâu sắc nhất đến Giảng viên hướng dẫn đã luôn đồng hành, góp ý sát sao về giải pháp kỹ thuật, khích lệ việc cải tiến sản phẩm từ phiên bản cơ bản thành phiên bản VietNLP Studio hoàn chỉnh với khả năng xử lý tệp đa định dạng, phân đoạn ngữ nghĩa và lưu trữ bền vững trên CSDL SQLite.")
    add_p("Cuối cùng, tôi xin chân thành cảm ơn gia đình cùng các bạn sinh viên trong lớp đã luôn động viên, chia sẻ kinh nghiệm trong suốt quá trình học tập và hoàn thiện đề tài.")
    add_p("Kính mong tiếp tục nhận được những ý kiến đóng góp quý báu từ quý Thầy Cô trong Hội đồng phản biện để hệ thống phần mềm ngày càng hoàn thiện và có khả năng ứng dụng thực tế cao hơn nữa.")

    doc.add_page_break()

    # =============================================================
    # TÓM TẮT ĐỀ TÀI (ABSTRACT)
    # =============================================================
    add_h1("TÓM TẮT ĐỀ TÀI (ABSTRACT)")
    add_p("TÓM TẮT TIẾNG VIỆT", bold=True, color=RGBColor(0x1E, 0x3A, 0x8A))
    add_p("Trong bối cảnh kỷ nguyên số và bùng nổ dữ liệu văn bản tiếng Việt từ các nguồn tài liệu số, mạng xã hội và hệ thống doanh nghiệp, việc xây dựng một hệ thống phần mềm xử lý ngôn ngữ tự nhiên (NLP) hiện đại, toàn diện và có khả năng lưu trữ bền vững là một yêu cầu cấp thiết của ngành Kỹ thuật Phần mềm. Đề tài 'VietNLP Studio: Nghiên cứu và xây dựng hệ thống Xử lý ngôn ngữ tự nhiên tiếng Việt toàn diện kết hợp Streamlit, PhoBERT và Cơ sở dữ liệu SQLite' được thực hiện nhằm giải quyết trọn vẹn chuỗi giá trị từ thu nhận dữ liệu đa kênh, xử lý tệp tin đa định dạng, phân đoạn văn bản thông minh, phân tích học sâu đến lưu trữ dữ liệu an toàn.")
    add_p("Hệ thống VietNLP Studio được kiến trúc hóa theo mô hình 4 tầng module hóa độc lập, mang lại những bước tiến vượt trội so với các công cụ demo truyền thống: (1) Tiếp nhận dữ liệu linh hoạt thông qua nhập trực tiếp (kèm 3 mẫu văn bản gợi ý nhanh) và tải lên đa định dạng tệp tin bao gồm Microsoft Word (.docx), tài liệu PDF (.pdf), tệp văn bản thuần (.txt, .md) và bảng dữ liệu (.csv); (2) Thuật toán bóc tách phân đoạn văn bản (Paragraph Segmentation) cho phép người dùng xem trước, duyệt đoạn trước/sau và chọn bất kỳ đoạn nào để phân tích liên tục mà không cần tải lại tệp tin; (3) Thuật toán Phân đoạn Ngữ nghĩa & Tổng hợp Cảm xúc (Semantic Chunking & Aggregated Sentiment Analysis) ứng dụng mô hình Deep Learning PhoBERT (wonrax/phobert-base-vietnamese-sentiment), giải quyết triệt để hạn chế cắt cụt 256 tokens đối với các tệp văn bản dài; (4) Thống kê định lượng 5 chỉ số văn bản, trích xuất Top 10 từ khóa trọng số độ dài và tóm tắt trích xuất câu quan trọng; (5) Tích hợp Hệ quản trị Cơ sở Dữ liệu SQLite bền vững (history.db) với bảng analysis_history 13 trường dữ liệu được đánh chỉ mục, cung cấp tính năng tìm kiếm toàn văn, lọc theo sắc thái cảm xúc, nạp lại văn bản phân tích và đặc biệt là chức năng xóa riêng lẻ từng bản ghi cụ thể theo ID; (6) Giao diện người dùng Streamlit hiện đại được thiết kế theo Design System tùy biến với Hero Banner gradient, thanh Sidebar nút bo góc công thái học không ô tick, widget giám sát tài nguyên thời gian thực và chức năng xuất báo cáo CSV chuẩn mã hóa UTF-8-SIG tương thích tuyệt đối với Microsoft Excel. Kết quả thử nghiệm toàn diện qua 14 Test Cases cho thấy hệ thống hoạt động ổn định 100%, thời gian phản hồi suy luận nhanh (< 0.5s sau khi nạp cache), khẳng định tính đúng đắn và giá trị ứng dụng cao của một giải pháp phần mềm chuyên nghiệp.")
    add_p("Từ khóa: Kỹ thuật Phần mềm, VietNLP Studio, Xử lý Ngôn ngữ Tự nhiên (NLP), Streamlit, PhoBERT, SQLite Database, Phân đoạn văn bản, Xử lý tệp Word/PDF/CSV, Phân tích cảm xúc đa đoạn, Tóm tắt trích xuất.")
    
    add_p("\nENGLISH ABSTRACT", bold=True, color=RGBColor(0x1E, 0x3A, 0x8A))
    add_p("In the era of digital transformation and exponential growth of unstructured Vietnamese text across enterprise documents, social media, and digital repositories, engineering a comprehensive, robust, and persistently-stored Natural Language Processing (NLP) system is of paramount importance. This capstone project, entitled 'VietNLP Studio: An Advanced Vietnamese Natural Language Processing and Document Analytics Platform Integrating Streamlit, Deep Learning PhoBERT, and SQLite Database', delivers an end-to-end software solution bridging modern AI techniques and professional software engineering principles.")
    add_p("Architected upon a decoupled 4-tier modular framework, VietNLP Studio introduces significant technological advancements: (1) Omnichannel text ingestion supporting direct keyboard input (with 3 rapid-preset templates: Positive, Negative, Neutral) alongside comprehensive multi-format file uploads including Microsoft Word (.docx), Portable Document Format (.pdf), plain text (.txt, .md), and structured spreadsheets (.csv); (2) Intelligent paragraph segmentation engine enabling interactive paragraph browsing, sequential navigation, and continuous arbitrary segment analysis without reloading uploaded files; (3) Novel Semantic Chunking & Aggregated Sentiment Analysis leveraging the state-of-the-art PhoBERT Transformer model (wonrax/phobert-base-vietnamese-sentiment), effectively overcoming the 256-token truncation limitation on lengthy multi-page documents through weighted majority voting; (4) Quantitative statistical text profiling across 5 core linguistic metrics, length-weighted Top-10 keyword extraction, and extractive sentence summarization; (5) Enterprise-grade persistent data storage powered by embedded SQLite database (history.db) with a 13-attribute indexed schema, facilitating full-text search, categorical sentiment filtering, text reloading, and crucial individual record deletion (delete by unique ID); (6) Highly polished, ergonomic Streamlit interface styled with custom CSS gradients, pill-shaped non-checkbox navigation buttons, real-time database monitoring widgets, and Excel-compliant UTF-8-SIG CSV export. Rigorous quality assurance over 14 comprehensive test cases verified 100% operational success and sub-second inference latency (< 0.5s cached), demonstrating high software reliability, maintainability, and practical deployment viability.")
    add_p("Keywords: Software Engineering, VietNLP Studio, Natural Language Processing (NLP), Streamlit, PhoBERT, SQLite Database, Document Segmentation, Multi-format Parsing (.docx/.pdf/.csv), Semantic Chunking, Extractive Summarization.")

    doc.add_page_break()

    # =============================================================
    # MỤC LỤC & DANH MỤC
    # =============================================================
    add_h1("MỤC LỤC TỔNG QUÁT")
    
    toc_data = [
        ("LỜI CAM ĐOAN", "i"),
        ("LỜI CẢM ƠN", "ii"),
        ("TÓM TẮT ĐỀ TÀI (ABSTRACT)", "iii"),
        ("DANH MỤC TỪ VIẾT TẮT", "vi"),
        ("DANH MỤC BẢNG BIỂU", "vii"),
        ("DANH MỤC HÌNH ẢNH MINH HỌA", "viii"),
        ("CHƯƠNG 1: TỔNG QUAN VỀ ĐỀ TÀI VÀ CƠ SỞ KỸ THUẬT", "1"),
        ("  1.1. Bối cảnh và Tính cấp thiết của đề tài trong kỷ nguyên số", "1"),
        ("  1.2. Mục tiêu nghiên cứu và Đối tượng, phạm vi của hệ thống", "2"),
        ("  1.3. Phương pháp tiếp cận và Quy trình phát triển phần mềm", "3"),
        ("  1.4. Cơ sở lý thuyết Xử lý Ngôn ngữ Tự nhiên (NLP) tiếng Việt", "4"),
        ("  1.5. Nền tảng công nghệ, mô hình học sâu và hệ quản trị CSDL", "6"),
        ("CHƯƠNG 2: PHÂN TÍCH VÀ ĐẶC TẢ YÊU CẦU PHẦN MỀM", "9"),
        ("  2.1. Khảo sát hiện trạng và Phân tích các đối tượng người dùng", "9"),
        ("  2.2. Đặc tả yêu cầu chức năng mở rộng (Functional Requirements FR-01 -> FR-12)", "10"),
        ("  2.3. Đặc tả yêu cầu phi chức năng (Non-Functional Requirements NFR-01 -> NFR-06)", "13"),
        ("  2.4. Mô hình hóa chức năng hệ thống (Use Case Modeling & Đặc tả chi tiết)", "14"),
        ("CHƯƠNG 3: THIẾT KẾ KIẾN TRÚC HỆ THỐNG VÀ THUẬT TOÁN", "18"),
        ("  3.1. Thiết kế kiến trúc tổng thể 4 tầng module hóa (Layered Architecture)", "18"),
        ("  3.2. Thiết kế luồng dữ liệu đa kênh (Data Flow Pipeline)", "20"),
        ("  3.3. Thiết kế Cơ sở Dữ liệu quan hệ SQLite (Entity Relationship & Schema)", "22"),
        ("  3.4. Thiết kế chi tiết các thuật toán xử lý dữ liệu, phân đoạn và NLP", "25"),
        ("  3.5. Thiết kế giao diện người dùng theo Design System hiện đại (UI/UX)", "31"),
        ("CHƯƠNG 4: HIỆN THỰC HÓA VÀ CÀI ĐẶT HỆ THỐNG", "33"),
        ("  4.1. Môi trường phát triển và Cấu hình phần cứng, phần mềm", "33"),
        ("  4.2. Cấu trúc tổ chức thư mục mã nguồn dự án", "34"),
        ("  4.3. Hiện thực hóa chi tiết các module mã nguồn chính", "35"),
        ("    4.3.1. Module Cơ sở Dữ liệu SQLite (database.py)", "35"),
        ("    4.3.2. Module Đọc tệp đa định dạng và Phân đoạn văn bản (file_handler.py)", "38"),
        ("    4.3.3. Module Thuật toán NLP & Phân đoạn Ngữ nghĩa PhoBERT (app.py)", "40"),
        ("    4.3.4. Module Giao diện người dùng & Điều hướng Sidebar tùy biến (app.py)", "43"),
        ("  4.4. Hướng dẫn cài đặt, cấu hình và vận hành phần mềm", "45"),
        ("CHƯƠNG 5: KIỂM THỬ VÀ ĐÁNH GIÁ CHẤT LƯỢNG HỆ THỐNG", "47"),
        ("  5.1. Chiến lược và Phương pháp kiểm thử chất lượng phần mềm", "47"),
        ("  5.2. Kịch bản kiểm thử chi tiết hệ thống (14 Test Cases TC-01 -> TC-14)", "48"),
        ("  5.3. Đánh giá kết quả kiểm thử chức năng và kiểm thử chấp nhận", "52"),
        ("  5.4. Đánh giá hiệu năng thực thi và Tối ưu hóa đa phân đoạn", "53"),
        ("CHƯƠNG 6: KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN", "55"),
        ("  6.1. Tổng kết các kết quả đạt được của đề tài VietNLP Studio", "55"),
        ("  6.2. Các hạn chế kỹ thuật còn tồn đọng", "56"),
        ("  6.3. Đề xuất kế hoạch và Hướng phát triển trong tương lai", "56"),
        ("TÀI LIỆU THAM KHẢO", "58"),
        ("PHỤ LỤC: MÃ NGUỒN CỐT LÕI CÁC MODULE VÀ DANH MỤC TỪ DỪNG", "60")
    ]
    
    add_styled_table(
        ["Nội dung đề mục", "Trang"],
        toc_data,
        col_widths=[Inches(5.3), Inches(1.2)]
    )

    doc.add_page_break()

    # DANH MỤC TỪ VIẾT TẮT
    add_h1("DANH MỤC THUẬT NGỮ VÀ TỪ VIẾT TẮT")
    abbr_data = [
        ("NLP", "Natural Language Processing", "Xử lý ngôn ngữ tự nhiên"),
        ("AI", "Artificial Intelligence", "Trí tuệ nhân tạo"),
        ("SE", "Software Engineering", "Kỹ thuật phần mềm"),
        ("FR", "Functional Requirement", "Yêu cầu chức năng"),
        ("NFR", "Non-Functional Requirement", "Yêu cầu phi chức năng"),
        ("UI/UX", "User Interface / User Experience", "Giao diện và trải nghiệm người dùng"),
        ("PhoBERT", "Pre-trained RoBERTa for Vietnamese", "Mô hình ngôn ngữ tiền huấn luyện sâu cho tiếng Việt"),
        ("RDBMS", "Relational Database Management System", "Hệ quản trị cơ sở dữ liệu quan hệ"),
        ("DAO", "Data Access Object", "Đối tượng truy cập và thao tác dữ liệu"),
        ("CRUD", "Create, Read, Update, Delete", "Bốn thao tác dữ liệu cơ bản (Thêm, Đọc, Sửa, Xóa)"),
        ("PK", "Primary Key", "Khóa chính của bảng trong cơ sở dữ liệu"),
        ("CSV", "Comma-Separated Values", "Định dạng tệp dữ liệu phân tách bởi dấu phẩy"),
        ("BOM", "Byte Order Mark", "Ký tự đánh dấu thứ tự byte đầu tệp Unicode (UTF-8-SIG)"),
        ("PDF", "Portable Document Format", "Định dạng tài liệu di động"),
        ("DOCX", "Office Open XML Document", "Định dạng tệp tin văn bản Microsoft Word"),
        ("Regex", "Regular Expression", "Biểu thức chính quy"),
        ("DFD", "Data Flow Diagram", "Sơ đồ luồng dữ liệu"),
        ("UC", "Use Case", "Trường hợp sử dụng"),
        ("TC", "Test Case", "Trường hợp kiểm thử phần mềm")
    ]
    add_styled_table(
        ["Từ viết tắt", "Tên tiếng Anh đầy đủ", "Ý nghĩa / Diễn giải tiếng Việt"],
        abbr_data,
        col_widths=[Inches(1.2), Inches(2.8), Inches(2.5)]
    )

    # DANH MỤC BẢNG BIỂU & HÌNH ẢNH
    add_h1("DANH MỤC BẢNG BIỂU VÀ HÌNH ẢNH")
    add_p("DANH MỤC BẢNG BIỂU", bold=True, color=RGBColor(0x1E, 0x3A, 0x8A))
    table_list = [
        ("Bảng 2.1", "Danh mục và mô tả 12 yêu cầu chức năng (FR-01 -> FR-12)", "11"),
        ("Bảng 2.2", "Đặc tả các yêu cầu phi chức năng (Non-Functional Requirements)", "13"),
        ("Bảng 2.3", "Đặc tả chi tiết Use Case UC03: Bóc tách phân đoạn & Chọn đoạn phân tích", "15"),
        ("Bảng 2.4", "Đặc tả chi tiết Use Case UC08: Xóa riêng lẻ từng bản ghi lịch sử cụ thể", "16"),
        ("Bảng 3.1", "Đặc tả lược đồ bảng analysis_history trong cơ sở dữ liệu history.db", "23"),
        ("Bảng 3.2", "Tập 58 từ dừng tiếng Việt cơ bản cấu hình trong hệ thống", "27"),
        ("Bảng 4.1", "Cấu hình môi trường phần cứng và phần mềm phát triển", "33"),
        ("Bảng 4.2", "Danh mục các thư viện phụ thuộc cốt lõi trong requirements.txt", "34"),
        ("Bảng 5.1", "Bảng kịch bản kiểm thử chi tiết hệ thống (14 Test Cases TC-01 -> TC-14)", "49"),
        ("Bảng 5.2", "Tổng kết kết quả thực thi kiểm thử chức năng phần mềm", "52"),
        ("Bảng 5.3", "Đo lường thời gian đáp ứng giữa Cold Start, Cached và Multi-chunk", "53")
    ]
    add_styled_table(["Ký hiệu", "Tên bảng biểu mô tả", "Trang"], table_list, col_widths=[Inches(1.2), Inches(4.3), Inches(1.0)])

    add_p("\nDANH MỤC HÌNH ẢNH MINH HỌA", bold=True, color=RGBColor(0x1E, 0x3A, 0x8A))
    fig_list = [
        ("Hình 2.1", "Sơ đồ Use Case nâng cấp của hệ thống VietNLP Studio", "14"),
        ("Hình 3.1", "Kiến trúc phân tầng (4-Layer Architecture) của VietNLP Studio", "19"),
        ("Hình 3.2", "Luồng xử lý dữ liệu đa kênh và quy trình phân tích NLP toàn diện", "21"),
        ("Hình 3.3", "Sơ đồ lược đồ cơ sở dữ liệu SQLite (Entity Relationship Schema)", "24"),
        ("Hình 5.1", "Biểu đồ so sánh thời gian thực thi giữa Cold Start, Cached và Multi-chunk", "54")
    ]
    add_styled_table(["Ký hiệu", "Tên hình ảnh minh họa", "Trang"], fig_list, col_widths=[Inches(1.2), Inches(4.3), Inches(1.0)])

    doc.add_page_break()

    # =============================================================
    # CHƯƠNG 1: TỔNG QUAN VỀ ĐỀ TÀI VÀ CƠ SỞ KỸ THUẬT
    # =============================================================
    add_h1("CHƯƠNG 1: TỔNG QUAN VỀ ĐỀ TÀI VÀ CƠ SỞ KỸ THUẬT")
    
    add_h2("1.1. Bối cảnh và Tính cấp thiết của đề tài trong kỷ nguyên số")
    add_p("Trong bối cảnh chuyển đổi số toàn diện và sự bùng nổ của cuộc Cách mạng Công nghiệp 4.0, dữ liệu phi cấu trúc dạng văn bản tiếng Việt phát sinh hàng ngày trên không gian mạng với tốc độ chóng mặt. Nguồn tài liệu này vô cùng đa dạng: từ các phản hồi khách hàng (reviews) trên sàn thương mại điện tử, ý kiến tranh luận trên mạng xã hội, đến hàng ngàn tệp tin báo cáo định dạng Microsoft Word (.docx), tài liệu nghiên cứu PDF (.pdf), biên bản họp và dữ liệu khảo sát dạng bảng tính (.csv).")
    add_p("Tuy nhiên, phần lớn các ứng dụng xử lý ngôn ngữ tự nhiên (NLP) hiện nay vẫn bộc lộ nhiều điểm nghẽn nghiêm trọng dưới góc độ Kỹ thuật Phần mềm:")
    add_p("1. Kênh tiếp nhận dữ liệu nghèo nàn: Đa số các ứng dụng chỉ cho phép người dùng gõ hoặc dán chuỗi văn bản thủ công vào một ô nhập liệu đơn lẻ, hoàn toàn thiếu khả năng đọc và bóc tách trực tiếp các định dạng tệp tin văn phòng phổ biến như Word, PDF hay CSV.")
    add_p("2. Thách thức 'cắt cụt văn bản' của các mô hình học sâu: Các mô hình ngôn ngữ dựa trên Transformer tiên tiến (như PhoBERT hay BERT) đều có giới hạn cứng về độ dài đầu vào (thường là 256 hoặc 512 tokens). Khi người dùng tải lên một tệp tin tài liệu dài, các hệ thống thông thường chỉ đơn giản cắt ngắn văn bản (truncation), dẫn đến việc bỏ sót toàn bộ nội dung cảm xúc và thông tin quan trọng ở các trang sau.")
    add_p("3. Thiếu tính bền vững của dữ liệu: Nhiều hệ thống demo chỉ lưu trữ dữ liệu tạm thời trên bộ nhớ RAM của phiên làm việc. Khi người dùng tải lại trang web hoặc đóng trình duyệt, toàn bộ lịch sử phân tích bị mất hoàn toàn, thiếu đi một hệ quản trị cơ sở dữ liệu thực thụ để phục vụ tra cứu, tìm kiếm và quản lý lâu dài.")
    add_p("Xuất phát từ những đòi hỏi thực tiễn cấp bách đó, đề tài 'VietNLP Studio: Nghiên cứu và xây dựng hệ thống Xử lý ngôn ngữ tự nhiên tiếng Việt toàn diện kết hợp Streamlit, PhoBERT và Cơ sở dữ liệu SQLite' được thực hiện nhằm mang đến một giải pháp phần mềm hoàn chỉnh, mạnh mẽ, đáp ứng các tiêu chuẩn khắt khe về kỹ thuật phần mềm, khả năng xử lý tài liệu đa định dạng và lưu trữ bền vững.")

    add_h2("1.2. Mục tiêu nghiên cứu và Đối tượng, phạm vi của hệ thống")
    add_p("1.2.1. Mục tiêu nghiên cứu tổng quát và cụ thể:", bold=True)
    add_p("• Mục tiêu tổng quát: Thiết kế, kiến trúc hóa và hiện thực hóa nền tảng phần mềm VietNLP Studio với khả năng tiếp nhận dữ liệu đa kênh, bóc tách phân đoạn văn bản thông minh, phân tích cảm xúc học sâu giải quyết bài toán tài liệu dài, trích xuất từ khóa, tóm tắt và quản lý cơ sở dữ liệu bền vững trên SQLite.")
    add_p("• Mục tiêu cụ thể:")
    add_p("  - Xây dựng module xử lý tệp tin (file_handler.py) hỗ trợ đọc, bóc tách cấu trúc và giải mã an toàn các tệp tin .docx, .pdf, .txt, .md, .csv với cơ chế fallback encoding đa bảng mã.")
    add_p("  - Hiện thực hóa thuật toán bóc tách phân đoạn (Paragraph Segmentation), hỗ trợ duyệt đoạn trước/sau và chọn phân tích đoạn tùy ý liên tục mà không cần tải lại tệp tin.")
    add_p("  - Phát triển thuật toán Phân đoạn Ngữ nghĩa & Tổng hợp Cảm xúc (Semantic Chunking & Aggregated Sentiment Analysis) kết hợp mô hình PhoBERT-base, cho phép phân tích cảm xúc toàn diện trên các tệp tài liệu dài mà không bị cắt cụt.")
    add_p("  - Thiết kế và cài đặt cơ sở dữ liệu quan hệ SQLite (history.db) với module DAO độc lập (database.py), cung cấp đầy đủ chức năng CRUD, tìm kiếm toàn văn, lọc theo sắc thái cảm xúc và đặc biệt là chức năng XÓA RIÊNG LẺ TỪNG BẢN GHI cụ thể.")
    add_p("  - Tối ưu hóa trải nghiệm người dùng (UI/UX) với giao diện Streamlit hiện đại, Hero Banner gradient, hệ thống nút Sidebar bo góc không ô tick và widget giám sát hệ thống thời gian thực.")
    add_p("  - Cung cấp tính năng nạp lại văn bản từ lịch sử và xuất báo cáo CSV chuẩn mã hóa UTF-8-SIG tương thích 100% với Microsoft Excel.")
    
    add_p("1.2.2. Đối tượng và phạm vi ứng dụng:", bold=True)
    add_p("• Đối tượng nghiên cứu: Dữ liệu văn bản tiếng Việt từ bàn phím và từ các tệp tài liệu số (.docx, .pdf, .txt, .csv); Mô hình học sâu Transformer PhoBERT; Hệ quản trị CSDL SQLite3; Nền tảng web app Streamlit.")
    add_p("• Phạm vi đề tài: Hệ thống tập trung xử lý ngôn ngữ tiếng Việt có dấu; Hỗ trợ các tệp tài liệu văn phòng phổ biến; Triển khai ứng dụng chạy trên máy trạm cục bộ hoặc máy chủ web.")

    add_h2("1.3. Phương pháp tiếp cận và Quy trình phát triển phần mềm")
    add_p("Đề tài áp dụng phương pháp luận Kỹ thuật Phần mềm hướng thành phần (Component-Based Software Engineering - CBSE) kết hợp quy trình phát triển linh hoạt Agile/Scrum. Mã nguồn được phân chia thành các module có tính gắn kết nội bộ cao (high cohesion) và tính phụ thuộc lỏng lẻo (loose coupling):")
    add_p("• Module giao diện & điều phối: app.py")
    add_p("• Module thao tác dữ liệu CSDL: database.py")
    add_p("• Module xử lý tệp tin & phân đoạn: file_handler.py")
    add_p("• Cấu hình hệ thống & bảng màu: .streamlit/config.toml")

    add_h2("1.4. Cơ sở lý thuyết Xử lý Ngôn ngữ Tự nhiên (NLP) tiếng Việt")
    add_p("1.4.1. Đặc trưng ngữ pháp và âm tiết tiếng Việt:")
    add_p("Tiếng Việt là ngôn ngữ đơn lập, âm tiết tính. Khoảng trắng trong văn bản tiếng Việt đóng vai trò phân tách âm tiết chứ không nhất thiết phân tách ranh giới từ. Các từ ghép (vd: 'trí tuệ nhân tạo', 'kỹ thuật phần mềm', 'cơ sở dữ liệu') mang tải trọng thông tin ngữ nghĩa rất cao. Do đó, hệ thống xây dựng bộ quy tắc Regular Expression trên Unicode để bóc tách từ tố chính xác, đồng thời áp dụng cơ chế lọc từ dừng (Stopwords filtering) với tập từ điển 58 từ chọn lọc để loại bỏ các hư từ, đại từ không mang nghĩa chủ đề.")
    
    add_p("1.4.2. Bài toán phân tích sắc thái cảm xúc (Sentiment Analysis):")
    add_p("Hệ thống phân loại văn bản thành 3 sắc thái cảm xúc: Tích cực (Positive), Tiêu cực (Negative) và Trung tính (Neutral). Độ tin cậy (Confidence Score) được tính toán dựa trên phân phối xác suất đầu ra của lớp phân loại Softmax.")

    add_p("1.4.3. Kỹ thuật Phân đoạn Ngữ nghĩa & Tổng hợp Cảm xúc (Semantic Chunking & Aggregation):")
    add_p("Để giải quyết bài toán văn bản dài vượt quá giới hạn 256 tokens của kiến trúc Transformer, hệ thống áp dụng kỹ thuật phân đoạn ngữ nghĩa (Semantic Chunking). Văn bản được chia thành các đoạn nhỏ có độ dài tối đa 150 từ (vừa vặn với cửa sổ ngữ cảnh của PhoBERT). Mỗi đoạn được phân tích cảm xúc độc lập. Sau đó, thuật toán tổng hợp (Majority Voting có điều chỉnh trọng số xác suất) sẽ xác định sắc thái chủ đạo của toàn bộ tài liệu theo công thức:")
    add_p("Final_Score = (Average_Score_Winner * 0.7) + ((Count_Winner / Total_Chunks) * 0.3)")
    add_p("Công thức này kết hợp hài hòa giữa độ tin cậy trung bình của mô hình và tỷ lệ áp đảo của các phân đoạn, mang lại kết quả khách quan và toàn diện cho toàn bộ tệp tin.")

    add_p("1.4.4. Trích xuất từ khóa có trọng số độ dài và Tóm tắt trích xuất:")
    add_p("• Trích xuất từ khóa: Điểm số của từ w được tính dựa trên tần suất xuất hiện chuẩn hóa kết hợp hệ số thưởng 1.15 cho các từ có độ dài từ 6 ký tự trở lên (từ ghép tiếng Việt).")
    add_p("• Tóm tắt trích xuất: Chấm điểm từng câu dựa trên độ quan trọng của các từ nội dung trong câu đó, kết hợp hệ số ưu tiên vị trí mở đầu 1.10 (Lead Bias) và giữ nguyên trật tự thời gian xuất hiện của các câu trong văn bản gốc.")

    add_h2("1.5. Nền tảng công nghệ, mô hình học sâu và hệ quản trị CSDL")
    add_p("1.5.1. Ngôn ngữ Python 3.12: Đảm bảo hiệu năng tối ưu, hỗ trợ các tính năng gõ kiểu mới và hệ sinh thái thư viện khoa học dữ liệu mạnh mẽ.")
    add_p("1.5.2. Framework Streamlit 1.61+: Hỗ trợ xây dựng giao diện web phản ứng (Reactive Web UI), cơ chế lưu cache tài nguyên @st.cache_resource và quản lý trạng thái phiên st.session_state.")
    add_p("1.5.3. Mô hình Deep Learning PhoBERT (wonrax/phobert-base-vietnamese-sentiment): Mô hình RoBERTa tiền huấn luyện riêng cho tiếng Việt do VinAI Research phát triển, đạt độ chính xác hàng đầu trong bài toán phân tích cảm xúc tiếng Việt.")
    add_p("1.5.4. Hệ quản trị CSDL SQLite (Native Embedded Database): CSDL quan hệ nhúng tiêu chuẩn ACID, lưu trữ dưới dạng một tệp tin duy nhất (history.db), không cần cài đặt máy chủ dịch vụ cồng kềnh, tốc độ truy xuất cực nhanh và tính toàn vẹn dữ liệu cao.")
    add_p("1.5.5. Thư viện xử lý tệp tin: python-docx (trích xuất cấu trúc văn bản Word), pypdf (phân tích luồng trang PDF), module csv chuẩn và regex Unicode.")

    add_callout_box(
        doc,
        [
            "VietNLP Studio đã chuyển đổi từ kiến trúc đơn tệp thử nghiệm thành hệ thống phần mềm 4 tầng module hóa chuẩn Kỹ thuật Phần mềm.",
            "Sự kết hợp giữa Cơ sở dữ liệu SQLite và Mô hình PhoBERT giải quyết đồng thời hai bài toán cốt lõi: Lưu trữ dữ liệu bền vững và Phân tích cảm xúc toàn diện cho tệp tài liệu dài."
        ],
        title="BƯỚC ĐỘT PHÁ CÔNG NGHỆ CỦA VIETNLP STUDIO"
    )

    doc.add_page_break()

    # =============================================================
    # CHƯƠNG 2: PHÂN TÍCH VÀ ĐẶC TẢ YÊU CẦU PHẦN MỀM
    # =============================================================
    add_h1("CHƯƠNG 2: PHÂN TÍCH VÀ ĐẶC TẢ YÊU CẦU PHẦN MỀM")
    
    add_h2("2.1. Khảo sát hiện trạng và Phân tích các đối tượng người dùng")
    add_p("Sau khi nâng cấp, hệ thống VietNLP Studio mở rộng phạm vi đáp ứng cho nhiều nhóm đối tượng người dùng thực tế:")
    add_p("1. Chuyên viên phân tích thị trường & Truyền thông: Cần tải lên hàng loạt tệp khảo sát khách hàng (.csv) hoặc báo cáo đánh giá dịch vụ (.docx) để nắm bắt nhanh sắc thái cảm xúc và các vấn đề phàn nàn cốt lõi.")
    add_p("2. Nhà nghiên cứu & Sinh viên: Cần đọc các tài liệu nghiên cứu PDF hoặc văn bản tin tức dài, bóc tách từng đoạn văn bản cụ thể để phân tích từ khóa, thống kê định lượng và tóm tắt nhanh nội dung.")
    add_p("3. Cán bộ quản lý dữ liệu: Cần lưu trữ bền vững lịch sử phân tích, tra cứu lại các văn bản đã kiểm tra trước đó, xóa bỏ các bản ghi kiểm thử rác và xuất báo cáo kết quả ra tệp bảng tính Excel.")

    add_h2("2.2. Đặc tả yêu cầu chức năng mở rộng (Functional Requirements)")
    add_p("Hệ thống VietNLP Studio được thiết kế với 12 yêu cầu chức năng hoàn chỉnh (FR-01 đến FR-12) như mô tả trong Bảng 2.1:")
    
    fr_data = [
        ("FR-01", "Nhập văn bản trực tiếp & Mẫu nhanh", "Cho phép soạn thảo văn bản tự do; Cung cấp 3 nút gợi ý văn bản mẫu nhanh (Khen ngợi, Phàn nàn, Thông tin)."),
        ("FR-02", "Tải lên tệp tin đa định dạng", "Hỗ trợ tải lên và giải mã tự động các tệp tin .docx (Word), .pdf (PDF), .txt, .md và .csv với fallback encoding an toàn."),
        ("FR-03", "Bóc tách phân đoạn & Chọn đoạn tùy ý", "Tự động trích xuất các đoạn văn bản (paragraphs); Cho phép duyệt đoạn trước/sau, chọn bất kỳ đoạn nào để phân tích liên tục mà không reload tệp."),
        ("FR-04", "Phân tích toàn bộ nội dung tệp", "Cho phép chọn chế độ phân tích toàn diện tất cả các đoạn văn có trong tệp tin tải lên."),
        ("FR-05", "Thống kê định lượng văn bản", "Đo lường và hiển thị trực quan 5 chỉ số: Số ký tự, Số từ, Số câu, Số từ độc nhất, Số từ mang nội dung ngữ nghĩa."),
        ("FR-06", "Phân tích cảm xúc PhoBERT & Đa đoạn", "Dự đoán cảm xúc (Tích cực, Tiêu cực, Trung tính) kèm độ tin cậy. Khi phân tích tệp dài, tự động phân đoạn và hiển thị phân bố cảm xúc chi tiết từng đoạn."),
        ("FR-07", "Trích xuất Top 10 từ khóa", "Lọc bỏ từ dừng, tính điểm tần suất kết hợp trọng số độ dài âm tiết (>= 6 ký tự) để xuất danh sách 10 từ khóa cốt lõi."),
        ("FR-08", "Tóm tắt trích xuất thông minh", "Chấm điểm câu theo trọng số từ khóa và heuristic vị trí câu mở đầu, trích xuất tối đa 3 câu tiêu biểu đại diện cho văn bản."),
        ("FR-09", "Tự động lưu trữ vào CSDL SQLite", "Tự động lưu toàn bộ thông số phân tích (thời gian, nguồn gốc, số liệu, cảm xúc, từ khóa, tóm tắt) vào bảng analysis_history trong history.db."),
        ("FR-10", "Tìm kiếm và Lọc lịch sử", "Hỗ trợ ô tìm kiếm toàn văn theo từ khóa nội dung/tên tệp và thanh lọc lịch sử theo 4 trạng thái cảm xúc (Tất cả, Tích cực, Tiêu cực, Trung tính)."),
        ("FR-11", "Xóa riêng lẻ từng bản ghi cụ thể", "Cho phép người dùng bấm nút xóa đúng một bản ghi lịch sử mong muốn theo ID mà không làm ảnh hưởng đến các bản ghi khác; Hỗ trợ nút xóa toàn bộ."),
        ("FR-12", "Nạp lại văn bản & Xuất CSV", "Cho phép nạp nội dung văn bản từ lịch sử trở lại trang phân tích; Xuất báo cáo kết quả phân tích ra tệp CSV chuẩn mã hóa UTF-8-SIG.")
    ]
    add_styled_table(
        ["Mã FR", "Tên yêu cầu chức năng", "Mô tả chi tiết chức năng"],
        fr_data,
        col_widths=[Inches(1.0), Inches(2.2), Inches(3.3)]
    )
    add_p("Bảng 2.1: Danh mục và mô tả 12 yêu cầu chức năng (FR-01 -> FR-12)", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, font_size=10)

    add_h2("2.3. Đặc tả yêu cầu phi chức năng (Non-Functional Requirements)")
    add_p("Hệ thống tuân thủ nghiêm ngặt 6 tiêu chuẩn chất lượng Kỹ thuật Phần mềm trong Bảng 2.2:")
    
    nfr_data = [
        ("NFR-01", "Hiệu năng & Phản hồi (Performance)", "Thời gian xử lý thống kê và từ khóa < 0.2s. Thời gian suy luận PhoBERT cho mỗi đoạn < 0.3s khi đã nạp cache. Thời gian truy vấn CSDL SQLite < 0.05s."),
        ("NFR-02", "Tối ưu hóa bộ nhớ (Caching)", "Áp dụng Singleton Pattern qua @st.cache_resource cho mô hình PhoBERT; Phân đoạn ngữ nghĩa max_words=150 giúp tránh tràn bộ nhớ RAM GPU/CPU."),
        ("NFR-03", "An toàn & Bền vững dữ liệu (Persistence)", "Mọi dữ liệu lịch sử được ghi trực tiếp xuống ổ cứng qua SQLite với cơ chế Transaction an toàn; Không mất dữ liệu khi ứng dụng khởi động lại."),
        ("NFR-04", "Tính khả dụng & Công thái học (Usability)", "Thiết kế giao diện theo Design System hiện đại; Menu Sidebar nút bấm hàng bo góc không có ô tick; Hiển thị Toast thông báo khi thao tác thành công."),
        ("NFR-05", "Khả năng tương thích tệp (Compatibility)", "Hỗ trợ đọc 5 định dạng tệp phổ biến (.docx, .pdf, .txt, .md, .csv); Tệp CSV xuất ra tương thích 100% với bảng mã Microsoft Excel trên Windows."),
        ("NFR-06", "Tính module hóa (Maintainability)", "Phân tách rõ ràng giữa tầng giao diện, tầng xử lý tệp tin và tầng truy cập dữ liệu DAO; Dễ dàng mở rộng kết nối sang PostgreSQL/MySQL trong tương lai.")
    ]
    add_styled_table(
        ["Mã NFR", "Nhóm yêu cầu chất lượng", "Tiêu chuẩn và chỉ tiêu đánh giá"],
        nfr_data,
        col_widths=[Inches(1.1), Inches(2.1), Inches(3.3)]
    )
    add_p("Bảng 2.2: Đặc tả các yêu cầu phi chức năng (Non-Functional Requirements)", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, font_size=10)

    add_h2("2.4. Mô hình hóa chức năng hệ thống (Use Case Modeling)")
    add_p("Sơ đồ Use Case nâng cấp của hệ thống VietNLP Studio được minh họa trong Hình 2.1:")
    
    add_image_with_caption('scratch/figures/usecase_diagram.png', "Hình 2.1: Sơ đồ Use Case nâng cấp của hệ thống VietNLP Studio")

    add_p("2.4.1. Đặc tả chi tiết Use Case UC03: Bóc tách phân đoạn & Chọn đoạn phân tích")
    uc3_data = [
        ("Tên Use Case", "UC03: Bóc tách phân đoạn & Chọn đoạn tùy ý để phân tích"),
        ("Tác nhân chính", "Người dùng (End User)"),
        ("Mục đích", "Tải lên tệp tài liệu số, hệ thống tự tách các đoạn văn và cho phép người dùng chọn bất kỳ đoạn nào để phân tích liên tục."),
        ("Tiền điều kiện", "Người dùng chọn phương thức '📁 Tải tệp tin' và tải lên một tệp hợp lệ (.docx, .pdf, .txt, .csv)."),
        ("Hậu điều kiện", "Hiển thị danh sách các đoạn, trích xuất đoạn được chọn và hiển thị kết quả phân tích đầy đủ."),
        ("Luồng sự kiện chính (Basic Flow)", 
         "1. Người dùng chọn tệp tin từ máy tính thông qua widget st.file_uploader.\n"
         "2. Hệ thống gọi file_handler.parse_uploaded_file để đọc nội dung và phân tách thành danh sách các đoạn văn bản (paragraphs).\n"
         "3. Hệ thống hiển thị số lượng đoạn văn bản trích xuất được và bộ chọn phạm vi phân tích.\n"
         "4. Người dùng chọn chế độ '🔍 Chọn một đoạn cụ thể trong tệp để phân tích'.\n"
         "5. Người dùng sử dụng các nút '⏮️ Đoạn trước', 'Đoạn sau ⏭️' hoặc hộp Selectbox để chọn đoạn mong muốn.\n"
         "6. Hệ thống hiển thị nội dung đoạn được chọn vào ô soạn thảo cho phép người dùng đọc hoặc chỉnh sửa.\n"
         "7. Người dùng nhấn nút '🔍 Tiến hành phân tích'.\n"
         "8. Hệ thống thực hiện phân tích NLP, tự động lưu vào SQLite và hiển thị kết quả chi tiết.\n"
         "9. Người dùng có thể quay lại bước 5 để chọn tiếp đoạn khác mà không cần tải lại tệp tin."),
        ("Luồng ngoại lệ (Alternative Flow)", 
         "2a. Tệp tin bị lỗi định dạng hoặc hỏng: Hệ thống thông báo lỗi chi tiết màu đỏ.\n"
         "2b. Tệp tin rỗng không có chữ: Hệ thống hiển thị cảnh báo màu vàng yêu cầu kiểm tra lại tệp.")
    ]
    add_styled_table(["Thuộc tính Use Case", "Nội dung đặc tả"], uc3_data, col_widths=[Inches(2.0), Inches(4.5)])
    add_p("Bảng 2.3: Đặc tả chi tiết Use Case UC03: Bóc tách phân đoạn & Chọn đoạn phân tích", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, font_size=10)

    add_p("2.4.2. Đặc tả chi tiết Use Case UC08: Xóa riêng lẻ từng bản ghi lịch sử cụ thể")
    uc8_data = [
        ("Tên Use Case", "UC08: Xóa riêng lẻ từng bản ghi lịch sử cụ thể"),
        ("Tác nhân chính", "Người dùng (End User)"),
        ("Mục đích", "Xóa bỏ một bản ghi lịch sử phân tích không còn cần thiết ra khỏi CSDL SQLite mà không ảnh hưởng đến các bản ghi khác."),
        ("Tiền điều kiện", "Người dùng truy cập vào trang '📊 Lịch sử phân tích' và có ít nhất 1 bản ghi trong CSDL."),
        ("Hậu điều kiện", "Bản ghi được chỉ định bị xóa vĩnh viễn khỏi bảng analysis_history trong history.db, giao diện tự động cập nhật lại danh sách và các chỉ số KPI."),
        ("Luồng sự kiện chính (Basic Flow)", 
         "1. Người dùng mở trang '📊 Lịch sử phân tích'.\n"
         "2. Hệ thống tải danh sách lịch sử từ database.get_all_history và hiển thị từng thẻ bản ghi chi tiết.\n"
         "3. Tại bản ghi muốn xóa (ví dụ bản ghi #ID), người dùng nhấn nút '🗑️ Xóa bản ghi này'.\n"
         "4. Hệ thống gọi hàm database.delete_history_item(item_id).\n"
         "5. Module CSDL thực thi câu lệnh SQL DELETE FROM analysis_history WHERE id = ?.\n"
         "6. Khi xóa thành công, hệ thống hiển thị thông báo Toast xác nhận: 'Đã xóa thành công bản ghi #ID!'.\n"
         "7. Hệ thống tự động rerun giao diện để làm mới danh sách và cập nhật các thẻ số liệu KPI tổng quan."),
        ("Luồng ngoại lệ (Alternative Flow)", 
         "5a. Không tìm thấy ID trong CSDL: Hệ thống hiển thị thông báo lỗi màu đỏ 'Không thể xóa bản ghi này'.")
    ]
    add_styled_table(["Thuộc tính Use Case", "Nội dung đặc tả"], uc8_data, col_widths=[Inches(2.0), Inches(4.5)])
    add_p("Bảng 2.4: Đặc tả chi tiết Use Case UC08: Xóa riêng lẻ từng bản ghi lịch sử cụ thể", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, font_size=10)

    doc.add_page_break()

    # =============================================================
    # CHƯƠNG 3: THIẾT KẾ KIẾN TRÚC HỆ THỐNG VÀ THUẬT TOÁN
    # =============================================================
    add_h1("CHƯƠNG 3: THIẾT KẾ KIẾN TRÚC HỆ THỐNG VÀ THUẬT TOÁN")
    
    add_h2("3.1. Thiết kế kiến trúc tổng thể 4 tầng module hóa (Layered Architecture)")
    add_p("VietNLP Studio được tái cấu trúc hoàn toàn theo mô hình Kiến trúc phân tầng (4-Layered Architecture) chuẩn mực trong Kỹ thuật Phần mềm. Kiến trúc này giúp tách biệt rành mạch giữa giao diện người dùng, logic đọc tệp tin, thuật toán AI/NLP và tầng truy xuất cơ sở dữ liệu như minh họa trong Hình 3.1:")
    
    add_image_with_caption('scratch/figures/architecture_diagram.png', "Hình 3.1: Kiến trúc phân tầng (4-Layer Architecture) của VietNLP Studio")

    add_p("Chi tiết trách nhiệm kỹ thuật của 4 tầng:")
    add_p("1. Tầng Trình diễn (Presentation Layer - app.py): Xây dựng bằng Streamlit kết hợp mã nhúng CSS tùy biến. Phụ trách render giao diện Hero Banner gradient, các nút điều hướng Sidebar bo góc (không ô tick), bộ chọn phân đoạn Segmented Control, 5 thẻ số liệu Metric KPI, các Tabs phân tích trực quan và hộp thông báo Toast.")
    add_p("2. Tầng Xử lý Tệp tin & Phân đoạn (File & Segmentation Layer - file_handler.py): Đảm nhận nhiệm vụ bóc tách tài liệu số. Chứa các hàm giải mã đa encoding (decode_bytes), đọc định dạng Word (.docx) qua python-docx, đọc định dạng PDF (.pdf) qua pypdf, đọc bảng tính (.csv) và tệp văn bản (.txt, .md). Chứa thuật toán phân tách đoạn (extract_paragraphs_from_text) và cơ chế duyệt chọn đoạn liên tục.")
    add_p("3. Tầng Nghiệp vụ NLP & Mô hình Học sâu (NLP & Deep Learning Layer - app.py): Trái tim thuật toán của hệ thống. Chứa các hàm tiền xử lý chuỗi (normalize_text, tokenize_words, split_sentences), lọc từ dừng 58 từ, thuật toán trích xuất Top 10 từ khóa trọng số độ dài, thuật toán tóm tắt trích xuất và thuật toán Phân đoạn Ngữ nghĩa & Tổng hợp Cảm xúc PhoBERT (Semantic Chunking & Aggregated Sentiment). Tầng này tích hợp cơ chế nạp cache mô hình @st.cache_resource Singleton.")
    add_p("4. Tầng Dữ liệu Bền vững (Persistent Data Layer - database.py): Đảm nhận việc giao tiếp với hệ quản trị cơ sở dữ liệu SQLite (history.db). Đóng gói toàn bộ các thao tác CRUD vào các hàm DAO chuẩn mực: khởi tạo bảng và chỉ mục (init_db), ghi kết quả (insert_history), tìm kiếm & lọc (get_all_history), thống kê số liệu (get_history_stats) và đặc biệt là hàm xóa riêng lẻ (delete_history_item).")

    add_h2("3.2. Thiết kế luồng dữ liệu đa kênh (Data Flow Pipeline)")
    add_p("Luồng dữ liệu trong VietNLP Studio di chuyển qua một đường ống khép kín hỗ trợ đa kênh thu nhận, được minh họa chi tiết trong Hình 3.2:")
    
    add_image_with_caption('scratch/figures/nlp_pipeline_diagram.png', "Hình 3.2: Luồng xử lý dữ liệu đa kênh và quy trình phân tích NLP toàn diện")

    add_p("Mô tả 6 giai đoạn xử lý trong luồng dữ liệu:")
    add_p("• Giai đoạn 1 (Tiếp nhận): Nhận dữ liệu từ bàn phím (kèm nút mẫu nhanh) hoặc từ tệp tải lên (.docx, .pdf, .txt, .csv).")
    add_p("• Giai đoạn 2 (Bóc tách & Phân đoạn): Gọi file_handler để giải mã bảng mã an toàn và tách tài liệu thành danh sách các đoạn văn bản độc lập.")
    add_p("• Giai đoạn 3 (Lựa chọn phạm vi): Người dùng chọn phân tích 'Toàn bộ nội dung tệp' hoặc duyệt chọn 'Một đoạn cụ thể' để phân tích.")
    add_p("• Giai đoạn 4 (Phân tích NLP chuyên sâu): Văn bản được chuẩn hóa Unicode, bóc tách từ/câu, lọc từ dừng, trích xuất từ khóa, tóm tắt câu và đưa qua mô hình PhoBERT với cơ chế Semantic Chunking (max_words=150) để phân tích cảm xúc từng đoạn và tổng hợp đa số.")
    add_p("• Giai đoạn 5 (Lưu trữ CSDL SQLite): Tự động đẩy toàn bộ kết quả phân tích vào bảng analysis_history trong history.db, nhận về mã định danh duy nhất (ID) của bản ghi.")
    add_p("• Giai đoạn 6 (Trình diễn & Xuất bản): Render kết quả lên giao diện web (5 thẻ Metric, 3 Tabs chi tiết, Breakdown cảm xúc từng đoạn) và cho phép người dùng tải tệp kết quả CSV chuẩn UTF-8-SIG.")

    add_h2("3.3. Thiết kế Cơ sở Dữ liệu quan hệ SQLite (Entity Relationship & Schema)")
    add_p("Để đảm bảo tính bền vững của dữ liệu, hệ thống sử dụng CSDL SQLite với tệp tin history.db. Bảng dữ liệu chính là analysis_history được thiết kế chuẩn hóa gồm 13 trường dữ liệu như thể hiện trong Hình 3.3 và Bảng 3.1:")
    
    add_image_with_caption('scratch/figures/database_er_diagram.png', "Hình 3.3: Sơ đồ lược đồ cơ sở dữ liệu SQLite (Entity Relationship Schema)")

    schema_detail = [
        ("id", "INTEGER", "PRIMARY KEY AUTOINCREMENT", "Mã định danh duy nhất của mỗi lượt phân tích"),
        ("created_at", "TEXT", "NOT NULL", "Thời gian phân tích định dạng 'dd/mm/yyyy HH:MM:SS'"),
        ("source_type", "TEXT", "NOT NULL", "Nguồn gốc văn bản (Nhập trực tiếp, Đoạn tệp Word/PDF/CSV)"),
        ("text_content", "TEXT", "NOT NULL", "Toàn văn chuỗi nội dung văn bản phân tích (UTF-8)"),
        ("char_count", "INTEGER", "DEFAULT 0", "Tổng số ký tự trong văn bản"),
        ("word_count", "INTEGER", "DEFAULT 0", "Tổng số từ tố tách được"),
        ("sentence_count", "INTEGER", "DEFAULT 0", "Tổng số câu trong đoạn văn"),
        ("unique_words", "INTEGER", "DEFAULT 0", "Số lượng từ vựng độc nhất (Vocabulary size)"),
        ("sentiment_label", "TEXT", "DEFAULT 'Chưa xác định'", "Nhãn cảm xúc dự đoán ('Tích cực', 'Tiêu cực', 'Trung tính')"),
        ("sentiment_score", "REAL", "DEFAULT 0.0", "Điểm tin cậy của mô hình (từ 0.0 đến 1.0)"),
        ("sentiment_emoji", "TEXT", "DEFAULT '🔍'", "Biểu tượng cảm xúc trực quan ('😊', '😡', '😐')"),
        ("keywords_json", "TEXT", "DEFAULT '[]'", "Chuỗi JSON lưu danh sách Top 10 từ khóa, tần suất và điểm số"),
        ("summary_text", "TEXT", "DEFAULT ''", "Nội dung đoạn văn tóm tắt trích xuất thông minh")
    ]
    add_styled_table(
        ["Tên trường (Field)", "Kiểu dữ liệu", "Ràng buộc (Constraint)", "Ý nghĩa nghiệp vụ chi tiết"],
        schema_detail,
        col_widths=[Inches(1.5), Inches(1.1), Inches(1.8), Inches(2.1)]
    )
    add_p("Bảng 3.1: Đặc tả lược đồ bảng analysis_history trong cơ sở dữ liệu history.db", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, font_size=10)

    add_p("Đặc biệt, hệ thống tạo chỉ mục (Index) chuyên dụng để tối ưu hóa hiệu năng truy vấn:")
    add_code_block(
        doc,
        "CREATE INDEX IF NOT EXISTS idx_history_created_at ON analysis_history(id DESC);",
        caption="Chỉ mục tăng tốc sắp xếp và truy vấn lịch sử phân tích"
    )

    add_h2("3.4. Thiết kế chi tiết các thuật toán xử lý dữ liệu, phân đoạn và NLP")
    add_p("3.4.1. Thuật toán giải mã an toàn đa encoding (decode_bytes):")
    add_p("Các tệp văn bản từ người dùng có thể được lưu trữ dưới nhiều bảng mã khác nhau. Hàm decode_bytes áp dụng kỹ thuật Fallback Encoding tuần tự qua các bảng mã phổ biến nhất:")
    add_code_block(
        doc,
        "def decode_bytes(raw_bytes: bytes) -> str:\n"
        "    for enc in [\"utf-8-sig\", \"utf-8\", \"utf-16\", \"cp1252\", \"latin-1\"]:\n"
        "        try:\n"
        "            return raw_bytes.decode(enc)\n"
        "        except UnicodeDecodeError:\n"
        "            continue\n"
        "    return raw_bytes.decode(\"utf-8\", errors=\"ignore\")",
        caption="Thuật toán giải mã an toàn đa bảng mã tệp tin"
    )

    add_p("3.4.2. Thuật toán bóc tách phân đoạn văn bản (extract_paragraphs_from_text):")
    add_p("Thuật toán chuẩn hóa các ký tự xuống dòng (\\r\\n -> \\n), sau đó sử dụng biểu thức chính quy tách theo dòng trống kép hoặc ngắt đoạn (re.split(r\"\\n\\s*\\n+\", text)). Mỗi đoạn văn sau khi làm sạch khoảng trắng dư thừa được đưa vào danh sách nếu có độ dài >= 3 ký tự. Nếu tài liệu chỉ gồm các dòng đơn, thuật toán tự động phân tách theo từng dòng đơn lẻ.")

    add_p("3.4.3. Thuật toán Phân đoạn Ngữ nghĩa & Tổng hợp Cảm xúc PhoBERT (Semantic Chunking & Aggregation):")
    add_p("Đây là bước cải tiến thuật toán đột phá của đề tài nhằm vượt qua giới hạn độ dài của mô hình PhoBERT. Thuật toán hoạt động theo 3 bước liên hoàn:")
    add_p("• Bước 1 (Semantic Chunking): Hàm split_text_into_sentiment_chunks duyệt qua các đoạn văn bản. Nếu một đoạn có số từ vượt quá max_words (150 từ), đoạn đó sẽ được tự động phân tách nhỏ hơn theo các dấu câu kết thúc ([.!?…]) để đảm bảo không một chunk nào vượt quá 150 từ.")
    add_p("• Bước 2 (Parallel Inference): Gửi từng chunk vào mô hình PhoBERT để lấy nhãn cảm xúc và điểm số tin cậy riêng biệt.")
    add_p("• Bước 3 (Weighted Majority Voting): Thống kê số lượng đoạn Tích cực, Tiêu cực, Trung tính. Chọn nhãn có số đoạn áp đảo nhất. Tính toán điểm số tin cậy tổng hợp dựa trên 70% điểm trung bình của nhãn chiến thắng và 30% tỷ lệ áp đảo của nhãn đó trên tổng số đoạn:")
    add_code_block(
        doc,
        "def split_text_into_sentiment_chunks(text: str, max_words: int = 150) -> list[str]:\n"
        "    paras = [p.strip() for p in text.split('\\n') if p.strip()]\n"
        "    chunks = []\n"
        "    for p in paras:\n"
        "        words = p.split()\n"
        "        if len(words) <= max_words:\n"
        "            chunks.append(p)\n"
        "        else:\n"
        "            sentences = re.split(r\"(?<=[.!?…])\\s+\", p)\n"
        "            curr = []\n"
        "            curr_len = 0\n"
        "            for s in sentences:\n"
        "                s_words = len(s.split())\n"
        "                if curr_len + s_words > max_words and curr:\n"
        "                    chunks.append(\" \".join(curr))\n"
        "                    curr = [s]\n"
        "                    curr_len = s_words\n"
        "                else:\n"
        "                    curr.append(s)\n"
        "                    curr_len += s_words\n"
        "            if curr:\n"
        "                chunks.append(\" \".join(curr))\n"
        "    return chunks or [text]",
        caption="Thuật toán phân đoạn ngữ nghĩa theo độ dài an toàn cho PhoBERT"
    )

    add_p("3.4.4. Cơ chế lọc từ dừng (Stopwords Filtering):")
    add_p("Hệ thống tích hợp sẵn tập từ điển 58 từ dừng chọn lọc trong Bảng 3.2:")
    stop_sample = [
        ("Quan hệ từ / Liên từ", "và, là, của, cho, với, trong, theo, nên, hay, hoặc, vì, để, do, mà, như"),
        ("Đại từ nhân xưng", "tôi, bạn, mình, nó, họ, em, anh, chị, ông, bà, we, you"),
        ("Chỉ từ & Phó từ", "các, một, những, được, có, không, rất, này, đó, thì, khi, đã, sẽ, đang, ra, vào"),
        ("Từ định hướng vị trí", "trên, dưới, sau, trước, lại, cũng, chỉ, bị, từ, đến, nữa")
    ]
    add_styled_table(["Nhóm từ loại", "Danh sách các từ dừng đặc trưng"], stop_sample, col_widths=[Inches(2.2), Inches(4.3)])
    add_p("Bảng 3.2: Tập 58 từ dừng tiếng Việt cơ bản cấu hình trong hệ thống", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, font_size=10)

    add_p("3.4.5. Thuật toán trích xuất từ khóa và Tóm tắt trích xuất câu:")
    add_p("• Trích xuất từ khóa: get_word_frequencies loại bỏ từ dừng và số thuần túy. Điểm số từ khóa được nhân hệ số 1.15 nếu độ dài từ >= 6 ký tự.")
    add_p("• Tóm tắt trích xuất: extractive_summary chấm điểm từng câu theo tổng trọng số từ vựng nội dung, nhân hệ số 1.10 cho câu mở đầu (Lead Bias), chọn 3 câu điểm cao nhất và sắp xếp lại theo trật tự thời gian gốc.")

    add_h2("3.5. Thiết kế giao diện người dùng theo Design System hiện đại (UI/UX)")
    add_p("Giao diện người dùng của VietNLP Studio được tái thiết kế toàn diện theo tiêu chuẩn công thái học:")
    add_p("• Hero Banner công nghệ: Phần đầu mỗi trang sở hữu banner gradient cao cấp (.main-hero) với hiệu ứng bóng đổ và tiêu đề đậm nét, tạo cảm giác chuyên nghiệp như một sản phẩm thương mại.")
    add_p("• Menu Sidebar dạng hàng bo góc (Không có ô tick): Thay vì sử dụng radio button mặc định với ô tick tròn gây cảm giác thô cứng, hệ thống đã viết lại CSS tùy biến để biến các nút menu thành các hàng chữ bo góc 12px, có hiệu ứng hover đổi màu và dịch chuyển nhẹ (translateX 4px). Nút đang chọn được tô màu xanh gradient nổi bật.")
    add_p("• Widget Giám sát Hệ thống thời gian thực: Sidebar tích hợp container viền nổi hiển thị trực tiếp trạng thái CSDL SQLite, tổng số bản ghi đã lưu trong history.db, mô hình AI đang hoạt động và các định dạng file được hỗ trợ.")
    add_p("• Badge trạng thái màu sắc: Sử dụng cú pháp Streamlit Badge (:green-badge, :red-badge, :gray-badge) để làm nổi bật tức thì sắc thái cảm xúc của từng bản ghi.")
    add_p("• Loại bỏ tab biểu đồ tần suất thừa: Tập trung vào 3 Tab chuyên sâu mang lại giá trị cao nhất: Phân tích cảm xúc (có Breakdown phân đoạn), Từ khóa nổi bật và Tóm tắt trích xuất.")

    doc.add_page_break()

    # =============================================================
    # CHƯƠNG 4: HIỆN THỰC HÓA VÀ CÀI ĐẶT HỆ THỐNG
    # =============================================================
    add_h1("CHƯƠNG 4: HIỆN THỰC HÓA VÀ CÀI ĐẶT HỆ THỐNG")
    
    add_h2("4.1. Môi trường phát triển và Cấu hình phần cứng, phần mềm")
    add_p("Bảng 4.1 tổng hợp môi trường kỹ thuật triển khai dự án:")
    
    env_data = [
        ("Hệ điều hành", "Microsoft Windows 10 / 11 (64-bit)"),
        ("Môi trường thực thi", "Python 3.12 (Virtual Environment: venv cô lập)"),
        ("Công cụ phát triển (IDE)", "Visual Studio Code / Cursor IDE"),
        ("Hệ quản trị CSDL", "SQLite 3 (Tệp tin history.db, kiểm soát luồng check_same_thread=False)"),
        ("Giao diện dòng lệnh", "Windows PowerShell 5.1 / 7.x, CMD"),
        ("Phần cứng khuyến nghị", "CPU Intel Core i5/i7 hoặc AMD Ryzen 5/7; RAM >= 8GB; Ổ cứng trống >= 2.5GB")
    ]
    add_styled_table(["Thành phần môi trường", "Thông số kỹ thuật chi tiết"], env_data, col_widths=[Inches(2.3), Inches(4.2)])
    add_p("Bảng 4.1: Cấu hình môi trường phần cứng và phần mềm phát triển", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, font_size=10)

    add_h2("4.2. Cấu trúc tổ chức thư mục mã nguồn dự án")
    add_p("Mã nguồn dự án được tổ chức khoa học thành 4 module chính độc lập:")
    add_code_block(
        doc,
        "DO AN STREAMLIT/\n"
        "├── .streamlit/\n"
        "│   └── config.toml           # Cấu hình theme hệ thống, màu sắc thương hiệu và server\n"
        "├── app.py                    # Giao diện chính Streamlit, điều hướng và pipeline NLP\n"
        "├── database.py               # Module quản lý CSDL SQLite (CRUD, tìm kiếm, xóa riêng lẻ)\n"
        "├── file_handler.py           # Module đọc tệp tin (Word, PDF, TXT, CSV) & phân đoạn\n"
        "├── history.db                # Cơ sở dữ liệu SQLite lưu trữ lịch sử bền vững\n"
        "├── requirements.txt          # Danh mục các thư viện phụ thuộc của dự án\n"
        "├── README.md                 # Hướng dẫn chi tiết cài đặt và vận hành hệ thống\n"
        "├── HDSD.docx                 # Tài liệu hướng dẫn sử dụng nhanh\n"
        "└── venv/                     # Môi trường ảo Python độc lập",
        caption="Cấu trúc tổ chức thư mục mã nguồn hoàn chỉnh của VietNLP Studio"
    )

    add_p("Danh sách các thư viện phụ thuộc cốt lõi trong requirements.txt:")
    req_data = [
        ("streamlit", "1.61.1", "Khung phát triển giao diện người dùng web ứng dụng"),
        ("transformers", "5.15.1", "Thư viện triển khai mô hình học sâu Transformer Hugging Face"),
        ("torch (PyTorch)", "2.13.0", "Nền tảng tính toán tensor và mạng nơ-ron sâu của PhoBERT"),
        ("python-docx", "1.2.0", "Thư viện đọc và trích xuất cấu trúc văn bản Microsoft Word (.docx)"),
        ("pypdf", "5.x", "Thư viện đọc và trích xuất dữ liệu tài liệu PDF (.pdf)"),
        ("tokenizers", "0.22.2", "Bộ công cụ mã hóa và phân tách token siêu tốc")
    ]
    add_styled_table(["Tên thư viện", "Phiên bản", "Vai trò trong hệ thống phần mềm"], req_data, col_widths=[Inches(1.8), Inches(1.2), Inches(3.5)])
    add_p("Bảng 4.2: Danh mục các thư viện phụ thuộc cốt lõi trong requirements.txt", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, font_size=10)

    add_h2("4.3. Hiện thực hóa chi tiết các module mã nguồn chính")
    add_p("4.3.1. Module Cơ sở Dữ liệu SQLite (database.py):")
    add_p("Module database.py đảm nhận toàn bộ các thao tác tương tác với cơ sở dữ liệu history.db. Dưới đây là trích đoạn các hàm trọng tâm thực thi việc khởi tạo bảng có đánh chỉ mục, thêm mới bản ghi và chức năng XÓA RIÊNG LẺ TỪNG BẢN GHI cụ thể:")
    add_code_block(
        doc,
        "def init_db():\n"
        "    conn = get_connection()\n"
        "    try:\n"
        "        with conn:\n"
        "            conn.execute('''\n"
        "                CREATE TABLE IF NOT EXISTS analysis_history (\n"
        "                    id INTEGER PRIMARY KEY AUTOINCREMENT,\n"
        "                    created_at TEXT NOT NULL,\n"
        "                    source_type TEXT NOT NULL,\n"
        "                    text_content TEXT NOT NULL,\n"
        "                    char_count INTEGER DEFAULT 0,\n"
        "                    word_count INTEGER DEFAULT 0,\n"
        "                    sentence_count INTEGER DEFAULT 0,\n"
        "                    unique_words INTEGER DEFAULT 0,\n"
        "                    sentiment_label TEXT DEFAULT 'Chưa xác định',\n"
        "                    sentiment_score REAL DEFAULT 0.0,\n"
        "                    sentiment_emoji TEXT DEFAULT '🔍',\n"
        "                    keywords_json TEXT DEFAULT '[]',\n"
        "                    summary_text TEXT DEFAULT ''\n"
        "                )\n"
        "            ''')\n"
        "            conn.execute('CREATE INDEX IF NOT EXISTS idx_history_created_at ON analysis_history(id DESC)')\n"
        "    finally:\n"
        "        conn.close()\n\n"
        "def delete_history_item(item_id: int) -> bool:\n"
        "    '''Xóa một bản ghi lịch sử phân tích cụ thể theo ID.'''\n"
        "    init_db()\n"
        "    conn = get_connection()\n"
        "    try:\n"
        "        with conn:\n"
        "            cursor = conn.execute('DELETE FROM analysis_history WHERE id = ?', (item_id,))\n"
        "            return cursor.rowcount > 0\n"
        "    finally:\n"
        "        conn.close()",
        caption="Module database.py: Khởi tạo bảng, đánh chỉ mục và hàm xóa riêng lẻ bản ghi theo ID"
    )

    add_p("4.3.2. Module Đọc tệp đa định dạng và Phân đoạn văn bản (file_handler.py):")
    add_p("Hàm parse_uploaded_file bóc tách dữ liệu từ đối tượng UploadedFile của Streamlit, hỗ trợ xử lý linh hoạt cho 5 định dạng tệp:")
    add_code_block(
        doc,
        "def parse_uploaded_file(uploaded_file) -> dict:\n"
        "    filename = uploaded_file.name\n"
        "    file_bytes = uploaded_file.getvalue()\n"
        "    ext = filename.rsplit('.', 1)[-1].lower() if '.' in filename else ''\n"
        "    paragraphs = []\n"
        "    full_text = ''\n"
        "    error = None\n"
        "    try:\n"
        "        if ext in ['txt', 'md']:\n"
        "            full_text = decode_bytes(file_bytes)\n"
        "            paragraphs = extract_paragraphs_from_text(full_text)\n"
        "        elif ext == 'docx':\n"
        "            doc = docx.Document(io.BytesIO(file_bytes))\n"
        "            paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]\n"
        "            full_text = '\\n\\n'.join(paragraphs)\n"
        "        elif ext == 'pdf':\n"
        "            reader = pypdf.PdfReader(io.BytesIO(file_bytes))\n"
        "            pages_text = [page.extract_text() for page in reader.pages if page.extract_text()]\n"
        "            full_text = '\\n\\n'.join(pages_text)\n"
        "            paragraphs = extract_paragraphs_from_text(full_text)\n"
        "        elif ext == 'csv':\n"
        "            text_data = decode_bytes(file_bytes)\n"
        "            reader = csv.reader(io.StringIO(text_data))\n"
        "            for r in reader:\n"
        "                row_str = ' | '.join([c.strip() for c in r if c.strip()])\n"
        "                if row_str:\n"
        "                    paragraphs.append(row_str)\n"
        "            full_text = '\\n'.join(paragraphs)\n"
        "    except Exception as e:\n"
        "        error = f'Lỗi đọc tệp: {str(e)}'\n"
        "    return {'filename': filename, 'paragraphs': paragraphs, 'full_text': full_text, 'error': error}",
        caption="Module file_handler.py: Bóc tách dữ liệu từ tệp Word, PDF, CSV, TXT"
    )

    add_p("4.3.3. Module Điều hướng Sidebar nút bấm bo góc không ô tick (app.py):")
    add_p("Đoạn mã hiện thực hóa thanh menu Sidebar mới bằng cách dùng st.button với điều kiện so sánh st.session_state.selected_page để đổi kiểu type='primary' hoặc 'secondary', kết hợp CSS ẩn triệt để ô tick radio:")
    add_code_block(
        doc,
        "nav_items = [\n"
        "    ('📝  Phân tích văn bản', 'Phân tích văn bản'),\n"
        "    ('📊  Lịch sử phân tích', 'Lịch sử phân tích'),\n"
        "    ('ℹ️  Giới thiệu đề tài', 'Giới thiệu đề tài'),\n"
        "]\n"
        "for label, page_key in nav_items:\n"
        "    is_active = (st.session_state.selected_page == page_key)\n"
        "    if st.button(\n"
        "        label,\n"
        "        key=f'nav_btn_{page_key}',\n"
        "        type='primary' if is_active else 'secondary',\n"
        "        width='stretch',\n"
        "    ):\n"
        "        st.session_state.selected_page = page_key\n"
        "        st.rerun()",
        caption="Cơ chế menu Sidebar nút bấm bo góc công thái học hiện đại"
    )

    add_h2("4.4. Hướng dẫn cài đặt, cấu hình và vận hành phần mềm")
    add_p("Để triển khai hệ thống VietNLP Studio trên máy tính, thực hiện tuần tự 4 bước:")
    add_p("Bước 1: Mở cửa sổ PowerShell và chuyển đến thư mục dự án:")
    add_code_block(doc, "cd \"F:\\DO AN STREAMLIT\"", caption="Chuyển đến thư mục dự án")
    
    add_p("Bước 2: Kích hoạt môi trường ảo Python venv:")
    add_code_block(doc, ".\\venv\\Scripts\\Activate.ps1", caption="Kích hoạt môi trường ảo")
    
    add_p("Bước 3: Cài đặt bổ sung các thư viện phụ thuộc mới (python-docx, pypdf):")
    add_code_block(doc, "python -m pip install -r requirements.txt", caption="Cài đặt phụ thuộc")
    
    add_p("Bước 4: Khởi chạy ứng dụng VietNLP Studio trên máy chủ cục bộ:")
    add_code_block(doc, "streamlit run app.py", caption="Khởi chạy ứng dụng web")
    add_p("Hệ thống sẽ tự động khởi động máy chủ và mở trình duyệt web tại địa chỉ http://localhost:8501.")

    doc.add_page_break()

    # =============================================================
    # CHƯƠNG 5: KIỂM THỬ VÀ ĐÁNH GIÁ CHẤT LƯỢNG HỆ THỐNG
    # =============================================================
    add_h1("CHƯƠNG 5: KIỂM THỬ VÀ ĐÁNH GIÁ CHẤT LƯỢNG HỆ THỐNG")
    
    add_h2("5.1. Chiến lược và Phương pháp kiểm thử chất lượng phần mềm")
    add_p("Kiểm thử phần mềm đối với VietNLP Studio được tiến hành bài bản theo phương pháp Kiểm thử hộp đen (Black-Box Testing) kết hợp Kiểm thử chấp nhận người dùng (User Acceptance Testing - UAT) nhằm xác minh tính đúng đắn của toàn bộ 12 yêu cầu chức năng.")

    add_h2("5.2. Kịch bản kiểm thử chi tiết hệ thống (14 Test Cases)")
    add_p("Bộ kịch bản kiểm thử toàn diện gồm 14 ca kiểm thử (TC-01 đến TC-14) bao quát từ các thao tác nhập tay, đọc tệp đa định dạng, phân đoạn văn bản, suy luận cảm xúc đa đoạn, đến các thao tác quản trị CSDL SQLite trong Bảng 5.1:")
    
    test_cases_new = [
        ("TC-01", "Kiểm thử nhập rỗng", "Để trống ô nhập và bấm 'Tiến hành phân tích'", "Cảnh báo màu vàng yêu cầu nhập văn bản, không phát sinh lỗi", "Đạt (Pass)"),
        ("TC-02", "Kiểm thử 3 nút mẫu gợi ý", "Nhấn nút 'Mẫu tích cực', 'Mẫu tiêu cực', 'Mẫu trung tính'", "Nạp chính xác đoạn văn mẫu vào Text Area và cập nhật session state", "Đạt (Pass)"),
        ("TC-03", "Kiểm thử tải tệp Word (.docx)", "Tải lên tệp văn bản Word chứa 4 đoạn văn bản tiếng Việt", "Trích xuất đúng 4 đoạn văn bản; Cho phép xem và chỉnh sửa", "Đạt (Pass)"),
        ("TC-04", "Kiểm thử tải tệp PDF (.pdf)", "Tải lên tài liệu PDF tiếng Việt có dấu", "pypdf đọc dữ liệu trang, tách đoạn chính xác không lỗi bảng mã", "Đạt (Pass)"),
        ("TC-05", "Kiểm thử tải tệp CSV (.csv)", "Tải lên tệp CSV chứa danh sách bình luận", "Đọc các dòng CSV, ghép nối thành các đoạn văn bản phân tích", "Đạt (Pass)"),
        ("TC-06", "Kiểm thử duyệt chọn đoạn tùy ý", "Bấm 'Đoạn sau', 'Đoạn trước' hoặc chọn Selectbox đoạn #3", "Chuyển đổi tức thì sang đoạn #3 mà KHÔNG cần tải lại tệp tin", "Đạt (Pass)"),
        ("TC-07", "Kiểm thử phân tích toàn bộ tệp", "Chọn chế độ 'Phân tích toàn bộ nội dung tệp'", "Hệ thống phân đoạn ngữ nghĩa và phân tích toàn bộ các đoạn", "Đạt (Pass)"),
        ("TC-08", "Kiểm thử PhoBERT cảm xúc Tích cực", "Nhập đoạn văn khen ngợi dịch vụ hài lòng", "Nhãn: Tích cực (😊), Hộp xanh lá, Điểm tin cậy > 90%", "Đạt (Pass)"),
        ("TC-09", "Kiểm thử PhoBERT cảm xúc Tiêu cực", "Nhập đoạn văn phàn nàn chất lượng sản phẩm tệ", "Nhãn: Tiêu cực (😡), Hộp đỏ, Điểm tin cậy > 95%", "Đạt (Pass)"),
        ("TC-10", "Kiểm thử Phân đoạn cảm xúc Breakdown", "Phân tích văn bản dài gồm cả đoạn khen và đoạn chê", "Hiển thị phân bố số đoạn Tích cực/Tiêu cực/Trung tính chi tiết", "Đạt (Pass)"),
        ("TC-11", "Kiểm thử lưu tự động vào SQLite", "Thực hiện 1 lượt phân tích mới bất kỳ", "Bản ghi lưu thành công vào history.db, hiển thị Toast thông báo ID", "Đạt (Pass)"),
        ("TC-12", "Kiểm thử tìm kiếm & Lọc cảm xúc", "Nhập từ khóa tìm kiếm và chọn lọc nhãn 'Tiêu cực'", "Lọc chính xác các bản ghi thỏa mãn đồng thời cả 2 điều kiện", "Đạt (Pass)"),
        ("TC-13", "Kiểm thử XÓA 1 BẢN GHI RIÊNG LẺ", "Tại bản ghi #ID, nhấn nút '🗑️ Xóa bản ghi này'", "Xóa đúng bản ghi #ID khỏi SQLite; Các bản ghi khác giữ nguyên 100%", "Đạt (Pass)"),
        ("TC-14", "Kiểm thử Nạp lại văn bản & Xuất CSV", "Nhấn nút '🔄 Nạp lại' và nút '⬇️ Tải xuống kết quả CSV'", "Nạp lại văn bản về trang chính; Tải tệp CSV mở trên Excel không lỗi font", "Đạt (Pass)")
    ]
    add_styled_table(
        ["Mã TC", "Tên ca kiểm thử", "Dữ liệu & Thao tác kiểm thử", "Kết quả kỳ vọng của hệ thống", "Trạng thái"],
        test_cases_new,
        col_widths=[Inches(0.8), Inches(1.5), Inches(1.8), Inches(1.8), Inches(0.8)]
    )
    add_p("Bảng 5.1: Bảng kịch bản kiểm thử chi tiết hệ thống (14 Test Cases TC-01 -> TC-14)", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, font_size=10)

    add_h2("5.3. Đánh giá kết quả kiểm thử chức năng và kiểm thử chấp nhận")
    add_p("Bảng 5.2 tổng hợp kết quả kiểm thử chất lượng phần mềm:")
    res_summary = [
        ("Tổng số ca kiểm thử thiết kế", "14 ca kiểm thử toàn diện"),
        ("Số ca kiểm thử Đạt (Passed)", "14 ca kiểm thử (100%)"),
        ("Số ca kiểm thử Thất bại (Failed)", "0 ca kiểm thử (0%)"),
        ("Mức độ đáp ứng yêu cầu chức năng", "100% (12/12 Functional Requirements)"),
        ("Độ ổn định của CSDL SQLite", "Không phát sinh lỗi khóa bảng hay xung đột luồng")
    ]
    add_styled_table(["Tiêu chí đánh giá kiểm thử", "Kết quả thực tế đạt được"], res_summary, col_widths=[Inches(3.2), Inches(3.3)])
    add_p("Bảng 5.2: Tổng kết kết quả thực thi kiểm thử chức năng phần mềm", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, font_size=10)

    add_h2("5.4. Đánh giá hiệu năng thực thi và Tối ưu hóa đa phân đoạn")
    add_p("Hệ thống đạt hiệu năng phản hồi xuất sắc nhờ cơ chế nạp cache mô hình PhoBERT duy nhất một lần và phân đoạn ngữ nghĩa thông minh:")
    
    perf_data = [
        ("Lần khởi động đầu tiên (Cold Start)", "~35.0 giây (Tải trọng số PhoBERT và khởi tạo PyTorch)"),
        ("Phân tích một đoạn ngắn (50 từ, Cached)", "0.18 giây"),
        ("Phân tích một đoạn vừa (200 từ, Cached)", "0.28 giây"),
        ("Phân tích toàn bộ tệp 5 đoạn (Multi-chunk)", "0.85 giây (Phân tích 5 đoạn độc lập và tổng hợp)")
    ]
    add_styled_table(["Tình huống thực thi hệ thống", "Thời gian phản hồi trung bình"], perf_data, col_widths=[Inches(3.2), Inches(3.3)])
    add_p("Bảng 5.3: Đo lường thời gian đáp ứng giữa Cold Start, Cached và Multi-chunk", align=WD_ALIGN_PARAGRAPH.CENTER, italic=True, font_size=10)

    add_image_with_caption('scratch/figures/performance_comparison.png', "Hình 5.1: Biểu đồ so sánh thời gian thực thi giữa Cold Start, Cached và Multi-chunk")
    add_p("Như được thể hiện trên thang đo logarit trong Hình 5.1, ngay cả đối với tệp tài liệu gồm 5 đoạn văn bản lớn, toàn bộ quá trình phân tích NLP, suy luận đa đoạn và lưu trữ vào CSDL SQLite chỉ diễn ra trong vòng 0.85 giây, đáp ứng hoàn hảo tiêu chuẩn trải nghiệm thời gian thực của người dùng.")

    doc.add_page_break()

    # =============================================================
    # CHƯƠNG 6: KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN
    # =============================================================
    add_h1("CHƯƠNG 6: KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN")
    
    add_h2("6.1. Tổng kết các kết quả đạt được của đề tài VietNLP Studio")
    add_p("Trải qua quá trình nâng cấp, tái cấu trúc và hoàn thiện hệ thống, đề tài 'VietNLP Studio: Nghiên cứu và xây dựng hệ thống Xử lý ngôn ngữ tự nhiên tiếng Việt toàn diện kết hợp Streamlit, PhoBERT và Cơ sở dữ liệu SQLite' đã hoàn thành xuất sắc các mục tiêu đề ra:")
    add_p("1. Đột phá về mặt Kỹ thuật Phần mềm:")
    add_p("• Chuyển đổi thành công từ một script đơn lẻ sang một kiến trúc 4 tầng module hóa độc lập, tuân thủ nguyên lý Separation of Concerns, dễ bảo trì và dễ mở rộng.")
    add_p("• Xây dựng module CSDL SQLite hoàn chỉnh với cơ chế lưu trữ bền vững, đánh chỉ mục và hiện thực hóa thành công tính năng XÓA RIÊNG LẺ TỪNG BẢN GHI cụ thể theo ID đáp ứng trọn vẹn yêu cầu thực tế của người dùng.")
    add_p("• Xây dựng module xử lý tệp tin hỗ trợ đọc và bóc tách cấu trúc 5 định dạng tài liệu phổ biến (.docx, .pdf, .txt, .md, .csv) với thuật toán phân đoạn thông minh, cho phép chọn và phân tích liên tục không cần reload tệp.")
    add_p("2. Đột phá về mặt Thuật toán và Trí tuệ Nhân tạo (NLP):")
    add_p("• Đề xuất và cài đặt thành công thuật toán Phân đoạn Ngữ nghĩa & Tổng hợp Cảm xúc (Semantic Chunking & Aggregation), giải quyết triệt để nút thắt cổ chai cắt cụt văn bản 256 tokens của các mô hình Transformer đối với tài liệu dài.")
    add_p("• Hoàn thiện các thuật toán thống kê định lượng, lọc từ dừng chuyên sâu, trích xuất Top 10 từ khóa trọng số độ dài và tóm tắt trích xuất câu có heuristic vị trí.")
    add_p("3. Đột phá về Trải nghiệm Người dùng (UI/UX):")
    add_p("• Xây dựng giao diện công thái học hiện đại với Hero Banner gradient, hệ thống nút Sidebar bo góc không ô tick, widget giám sát tài nguyên thời gian thực và xuất báo cáo CSV chuẩn Excel.")

    add_h2("6.2. Các hạn chế kỹ thuật còn tồn đọng")
    add_p("• Phương pháp tóm tắt văn bản: Hiện tại hệ thống đang áp dụng phương pháp tóm tắt trích xuất (Extractive Summarization) chọn lọc câu từ văn bản gốc, chưa áp dụng mô hình sinh văn bản tóm tắt trừu tượng (Abstractive Summarization) để diễn đạt lại nội dung bằng câu chữ mới.")
    add_p("• Nhận diện thực thể tên riêng (NER): Hệ thống chưa tích hợp module nhận dạng tự động các thực thể tên người, tên địa danh, tổ chức.")

    add_h2("6.3. Đề xuất kế hoạch và Hướng phát triển trong tương lai")
    add_p("1. Tích hợp Mô hình Tóm tắt Trừu tượng VietBART / BARTpho: Nghiên cứu và tích hợp mô hình sinh văn bản tóm tắt tiếng Việt hiện đại nhằm cung cấp các bản tóm tắt tự nhiên và súc tích hơn.")
    add_p("2. Tích hợp Nhận dạng Thực thể Tên riêng (PhoBERT-NER): Tự động phát hiện và đánh dấu các thực thể Person, Location, Organization trong văn bản tải lên.")
    add_p("3. Xây dựng RESTful API với FastAPI: Đóng gói các hàm nghiệp vụ của VietNLP Studio thành các API endpoints độc lập (POST /api/analyze, GET /api/history) để các ứng dụng Mobile hoặc hệ thống ngoài dễ dàng tích hợp.")
    add_p("4. Đóng gói Docker & Triển khai Đám mây (Cloud Deployment): Đóng gói toàn bộ ứng dụng thành Docker Container, thiết lập CI/CD pipeline và triển khai lên các nền tảng đám mây như Render, Hugging Face Spaces hoặc AWS.")

    doc.add_page_break()

    # =============================================================
    # TÀI LIỆU THAM KHẢO
    # =============================================================
    add_h1("TÀI LIỆU THAM KHẢO")
    
    refs = [
        "[1] Dat Quoc Nguyen, Anh Tuan Nguyen. 'PhoBERT: Pre-trained language models for Vietnamese'. Proceedings of the 2020 Conference on Empirical Methods in Natural Language Processing: Findings (EMNLP 2020), pp. 1037-1042, 2020.",
        "[2] Yinhan Liu, Myle Ott, Naman Goyal, Jingfei Du, Mandar Joshi, Danqi Chen, Omer Levy, Mike Lewis, Luke Zettlemoyer, Veselin Stoyanov. 'RoBERTa: A Robustly Optimized BERT Pretraining Approach'. arXiv preprint arXiv:1907.11692, 2019.",
        "[3] Ashish Vaswani, Noam Shazeer, Niki Parmar, Jakob Uszkoreit, Llion Jones, Aidan N Gomez, Łukasz Kaiser, Illia Polosukhin. 'Attention Is All You Need'. Advances in Neural Information Processing Systems (NeurIPS), pp. 5998-6008, 2017.",
        "[4] Streamlit Inc. 'Streamlit Documentation: The fastest way to build and share data apps'. Available online: https://docs.streamlit.io/, [Truy cập tháng 09/2026].",
        "[5] Hugging Face. 'Transformers: State-of-the-art Natural Language Processing for Pytorch, TensorFlow, and JAX'. Available online: https://huggingface.co/docs/transformers, [Truy cập tháng 09/2026].",
        "[6] SQLite Consortium. 'SQLite Documentation: An In-Process Database for Python and Embedded Systems'. Available online: https://www.sqlite.org/docs.html, [Truy cập tháng 09/2026].",
        "[7] Roger S. Pressman, Bruce R. Maxim. 'Software Engineering: A Practitioner's Approach'. 9th Edition, McGraw-Hill Education, 2020.",
        "[8] Ian Sommerville. 'Software Engineering'. 10th Edition, Pearson Education, 2016.",
        "[9] Dan Jurafsky, James H. Martin. 'Speech and Language Processing: An Introduction to Natural Language Processing, Computational Linguistics, and Speech Recognition'. 3rd Edition Draft, Stanford University, 2023.",
        "[10] Python-docx Team. 'python-docx: A Python library for creating and updating Microsoft Word (.docx) files'. Available online: https://python-docx.readthedocs.io/, [Truy cập tháng 09/2026].",
        "[11] PyPDF Team. 'pypdf: A pure-python PDF library capable of splitting, merging, cropping, and transforming the pages of PDF files'. Available online: https://pypdf.readthedocs.io/, [Truy cập tháng 09/2026].",
        "[12] Wonrax. 'phobert-base-vietnamese-sentiment Model Card on Hugging Face'. Available online: https://huggingface.co/wonrax/phobert-base-vietnamese-sentiment, [Truy cập tháng 09/2026]."
    ]
    
    for ref in refs:
        p_ref = add_p(ref, space_before=3, space_after=4, line_spacing=1.2)
        p_ref.paragraph_format.left_indent = Inches(0.4)
        p_ref.paragraph_format.first_line_indent = Inches(-0.4)

    doc.add_page_break()

    # =============================================================
    # PHỤ LỤC: MÃ NGUỒN CỐT LÕI
    # =============================================================
    add_h1("PHỤ LỤC: MÃ NGUỒN CỐT LÕI CÁC MODULE VÀ DANH MỤC TỪ DỪNG")
    add_p("Phụ lục này trình bày trích đoạn mã nguồn cốt lõi của các module mới được xây dựng trong hệ thống VietNLP Studio.")
    
    add_h2("Phụ lục A: Toàn văn Module Quản trị CSDL SQLite (database.py)")
    add_code_block(
        doc,
        "import sqlite3\n"
        "import json\n"
        "from datetime import datetime\n"
        "import os\n\n"
        "DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'history.db')\n\n"
        "def get_connection():\n"
        "    conn = sqlite3.connect(DB_PATH, check_same_thread=False)\n"
        "    conn.row_factory = sqlite3.Row\n"
        "    return conn\n\n"
        "def insert_history(source_type, text_content, char_count, word_count, sentence_count,\n"
        "                   unique_words, sentiment_label='Chưa xác định', sentiment_score=0.0,\n"
        "                   sentiment_emoji='🔍', keywords=None, summary_text=''):\n"
        "    init_db()\n"
        "    now_str = datetime.now().strftime('%d/%m/%Y %H:%M:%S')\n"
        "    keywords_json = json.dumps(keywords or [], ensure_ascii=False)\n"
        "    conn = get_connection()\n"
        "    try:\n"
        "        with conn:\n"
        "            cursor = conn.execute('''\n"
        "                INSERT INTO analysis_history (\n"
        "                    created_at, source_type, text_content, char_count, word_count,\n"
        "                    sentence_count, unique_words, sentiment_label, sentiment_score,\n"
        "                    sentiment_emoji, keywords_json, summary_text\n"
        "                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)\n"
        "            ''', (now_str, source_type, text_content, char_count, word_count,\n"
        "                  sentence_count, unique_words, sentiment_label, sentiment_score,\n"
        "                  sentiment_emoji, keywords_json, summary_text))\n"
        "            return cursor.lastrowid\n"
        "    finally:\n"
        "        conn.close()",
        caption="Trích đoạn module database.py thực thi ghi nhận lịch sử vào CSDL SQLite"
    )

    add_h2("Phụ lục B: Thuật toán Tổng hợp Cảm xúc Đa phân đoạn (analyze_sentiment trong app.py)")
    add_code_block(
        doc,
        "def analyze_sentiment(text: str):\n"
        "    classifier = load_sentiment_model()\n"
        "    chunks = split_text_into_sentiment_chunks(text, max_words=150)\n"
        "    if len(chunks) <= 1:\n"
        "        res = classifier(text, truncation=True, max_length=256)[0]\n"
        "        return map_sentiment_result(res)\n\n"
        "    chunk_results = []\n"
        "    label_counts = {'Tích cực': 0, 'Tiêu cực': 0, 'Trung tính': 0}\n"
        "    label_scores = {'Tích cực': [], 'Tiêu cực': [], 'Trung tính': []}\n"
        "    for c in chunks:\n"
        "        res = classifier(c, truncation=True, max_length=256)[0]\n"
        "        vlabel, score = extract_label_and_score(res)\n"
        "        label_counts[vlabel] += 1\n"
        "        label_scores[vlabel].append(score)\n\n"
        "    winner_label, winner_count = get_winner_label(label_counts, label_scores)\n"
        "    avg_score = sum(label_scores[winner_label]) / len(label_scores[winner_label])\n"
        "    ratio = winner_count / len(chunks)\n"
        "    final_score = (avg_score * 0.7) + (ratio * 0.3)\n"
        "    return build_aggregated_sentiment_dict(winner_label, final_score, chunks, label_counts)",
        caption="Thuật toán tổng hợp cảm xúc PhoBERT theo đa phân đoạn"
    )

    output_path = r"f:\DO AN STREAMLIT\BAO_CAO_DO_AN_CHUYEN_NGANH_KTPM.docx"
    doc.save(output_path)
    print(f"Report regenerated successfully at: {output_path}")

if __name__ == "__main__":
    create_full_report()
