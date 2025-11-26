import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_LINE_SPACING
from docx.enum.text import WD_TAB_ALIGNMENT
import io
import os
import re
import random

# --- Helper Functions and Constants ---

# Hàm tạo 150 màu (giữ nguyên)
def generate_vibrant_rgb_colors(count=150):
    colors = set()
    while len(colors) < count:
        h = random.random()
        s, v = 0.8, 0.9
        
        if s == 0.0: r = g = b = v
        else:
            i = int(h * 6.0); f = h * 6.0 - i; p = v * (1.0 - s); q = v * (1.0 - s * f); t = v * (1.0 - s * (1.0 - f))
            if i % 6 == 0: r, g, b = v, t, p
            elif i % 6 == 1: r, g, b = q, v, p
            elif i % 6 == 2: r, g, b = p, v, t
            elif i % 6 == 3: r, g, b = p, q, v
            elif i % 6 == 4: r, g, b = t, p, v
            else: r, g, b = v, p, q
        
        r, g, b = int(r * 255), int(g * 255), int(b * 255)
        if (r < 50 and g < 50 and b < 50) or (r > 200 and g > 200 and b > 200): continue 
        colors.add((r, g, b))
    
    return list(colors)

FONT_COLORS_RGB_150 = generate_vibrant_rgb_colors(150)
speaker_color_map = {}
used_colors = []

def get_speaker_color(speaker_name):
    global used_colors
    global speaker_color_map
    
    if speaker_name not in speaker_color_map:
        if used_colors:
            color_object = used_colors.pop()
        else:
            r, g, b = random.choice(FONT_COLORS_RGB_150)
            color_object = RGBColor(r, g, b)
            
        speaker_color_map[speaker_name] = color_object
        
    return speaker_color_map[speaker_name]

# FIX: Regex để tìm kiếm TẤT CẢ các tên người nói trong một đoạn
SPEAKER_REGEX_GLOBAL = re.compile(r"([A-Z][a-z\s&]+):\s*", re.IGNORECASE)

TIMECODE_REGEX = re.compile(r"^\d{2}:\d{2}:\d{2},\d{3}\s+-->\s+\d{2}:\d{2}:\d{2},\d{3}$")
HTML_CONTENT_REGEX = re.compile(r"((?:</?[ibu]>)+)(.*?)(?:</?[ibu]>)+", re.IGNORECASE | re.DOTALL)

# Hàm định dạng chung
def set_all_text_formatting(doc):
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
        
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(6)

# FIX: Hàm xử lý nội dung đa người nói (được gọi bên trong process_docx)
def process_dialogue_with_speakers(paragraph, text, document):
    """Xử lý nội dung đối thoại (có thể có nhiều người nói hoặc thẻ HTML)."""
    
    # 1. Áp dụng căn lề/dãn đoạn cho đoạn nội dung
    paragraph.style = document.styles['Normal']
    paragraph.paragraph_format.space_after = Pt(6) 
    paragraph.paragraph_format.space_before = Pt(0)
    
    # 2. Tìm tất cả người nói trong text
    matches = list(SPEAKER_REGEX_GLOBAL.finditer(text))
    
    if not matches:
        # Trường hợp không có người nói (chỉ là nội dung tiếp tục/nội dung đơn thuần)
        paragraph.paragraph_format.left_indent = None
        paragraph.paragraph_format.first_line_indent = None
        paragraph.text = text
        return # Thoát khỏi hàm xử lý speaker

    # 3. FIX: Xử lý ĐA NGƯỜI NÓI (Multi-Speaker)
    
    # Thiết lập căn lề treo cho đoạn văn
    paragraph.paragraph_format.left_indent = Inches(1.0)
    paragraph.paragraph_format.first_line_indent = Inches(-1.0)
    paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(1.0), WD_TAB_ALIGNMENT.LEFT)
    
    paragraph.text = "" # Xóa nội dung để xây dựng lại
    
    last_end = 0
    for match in matches:
        speaker_full = match.group(0) # e.g., "Coby: "
        speaker_name = match.group(1).strip() # e.g., "Coby"
        start, end = match.span()
        
        # A. Thêm text KHÔNG PHẢI người nói (text trước người nói hiện tại)
        text_before = text[last_end:start].strip()
        if text_before:
            paragraph.add_run(text_before)
        
        # B. Thêm NGƯỜI NÓI (Bold và Color)
        font_color_object = get_speaker_color(speaker_name) 
        run_speaker = paragraph.add_run(speaker_full)
        run_speaker.font.bold = True
        run_speaker.font.color.rgb = font_color_object 
        
        # C. Insert Tab sau tên người nói
        paragraph.add_run('\t') 
        
        last_end = end
        
    # D. Thêm nội dung cuối cùng sau người nói cuối cùng
    current_text = text[last_end:]
    
    # E. Xử lý các thẻ HTML còn lại trong nội dung cuối cùng
    matches_html = list(HTML_CONTENT_REGEX.finditer(current_text))
    last_end_html = 0
    
    if not matches_html:
        # Nếu không có thẻ HTML, thêm toàn bộ nội dung còn lại
        paragraph.add_run(current_text)
    else:
        # Nếu có thẻ HTML, xử lý từng phần
        for match in matches_html:
            tag_text = match.group(2) 
            start, end = match.span()

            # Thêm text TRƯỚC tag (nếu có)
            if start > last_end_html:
                paragraph.add_run(current_text[last_end_html:start])
            
            # Thêm nội dung HTML (Bold và Italic)
            run_html = paragraph.add_run(tag_text)
            run_html.font.bold = True
            run_html.font.italic = True
            
            last_end_html = end

        # Thêm nội dung sau tag cuối cùng
        if last_end_html < len(current_text):
            paragraph.add_run(current_text[last_end_html:])

# --- Hàm xử lý chính ---

def process_docx(uploaded_file, file_name_without_ext):
    
    global speaker_color_map
    global used_colors
    speaker_color_map = {}
    used_colors = [RGBColor(r, g, b) for r, g, b in FONT_COLORS_RGB_150]
    random.shuffle(used_colors)
    
    original_document = Document(io.BytesIO(uploaded_file.getvalue()))
    raw_paragraphs = [p for p in original_document.paragraphs]
    
    document = Document()
    
    # --- A. Set Main Title ---
    title_paragraph = document.add_paragraph(file_name_without_ext.upper())
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.paragraph_format.space_before = Pt(0)
    title_paragraph.paragraph_format.space_after = Pt(0) 
    
    title_run = title_paragraph.runs[0]
    title_run.font.name = 'Times New Roman'
    title_run.font.size = Pt(25) 
    title_run.bold = True
    
    document.add_paragraph().paragraph_format.space_after = Pt(0)
    document.add_paragraph().paragraph_format.space_after = Pt(0)

    # --- B. Process raw paragraphs and add to new document ---
    
    # FIX: Vùng gộp đoạn văn
    temp_content_block = []
    
    for paragraph in raw_paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # 1. Nếu là Timecode hoặc Index (dòng riêng biệt) -> Xử lý khối nội dung tạm
        if TIMECODE_REGEX.match(text) or re.fullmatch(r"^\s*\d+\s*$", text):
            
            # Xử lý khối nội dung đối thoại (nếu có)
            if temp_content_block:
                merged_content = " ".join(temp_content_block)
                new_paragraph = document.add_paragraph()
                process_dialogue_with_speakers(new_paragraph, merged_content, document)
                temp_content_block = [] # Reset khối
            
            # Bỏ Index
            if re.fullmatch(r"^\s*\d+\s*$", text):
                continue

            # Thêm Timecode
            new_paragraph = document.add_paragraph(text)
            for run in new_paragraph.runs:
                run.font.bold = True
            new_paragraph.paragraph_format.space_after = Pt(0) # Timecode không có dãn đoạn
            
        # 2. Nếu là nội dung đối thoại -> Thêm vào khối tạm
        else:
            temp_content_block.append(text)
            
    # Xử lý khối nội dung cuối cùng (nếu còn sót)
    if temp_content_block:
        merged_content = " ".join(temp_content_block)
        new_paragraph = document.add_paragraph()
        process_dialogue_with_speakers(new_paragraph, merged_content, document)

    # C. Apply General Font/Size and Spacing (Global settings)
    set_all_text_formatting(document)
    
    # Save the file
    modified_file = io.BytesIO()
    document.save(modified_file)
    modified_file.seek(0)
    
    return modified_file

# --- GIAO DIỆN STREAMLIT ---
# (Phần giao diện không đổi)

st.set_page_config(page_title="Automatic Word Script Editor", layout="wide")

st.markdown("## 📄 Automatic Subtitle Script (.docx) Converter")
st.markdown("A Python/Streamlit tool to automatically format subtitle scripts based on specific requirements.")
st.markdown("---")

uploaded_file = st.file_uploader(
    "1. Upload your Word file (.docx)",
    type=['docx'],
    help="Chỉ chấp nhận định dạng .docx của Microsoft Word."
)

if uploaded_file is not None:
    original_filename = uploaded_file.name
    # FIX TÊN FILE: Bỏ tiền tố và thêm hậu tố "_edit"
    file_name_without_ext = os.path.splitext(original_filename)[0]
    
    st.info(f"File received: **{original_filename}**.")
    
    if st.button("2. RUN AUTOMATIC FORMATTING"):
        with st.spinner('Đang xử lý và định dạng file...'):
            try:
                modified_file_io = process_docx(uploaded_file, file_name_without_ext)
                
                # FIX TÊN FILE: Tên_gốc_edit.docx
                new_filename = f"{file_name_without_ext}_edit.docx"

                st.success("✅ Định dạng hoàn tất! Bạn có thể tải file về.")
                
                # Nút tải file
                st.download_button(
                    label="3. Tải File Word Đã Định Dạng Về",
                    data=modified_file_io,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                # Đã loại bỏ phần xem trước thành phẩm theo yêu cầu cuối cùng.
                
                st.markdown("---")
                st.balloons()

            except Exception as e:
                st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")
                st.warning("Vui lòng kiểm tra lại định dạng file đầu vào.")
