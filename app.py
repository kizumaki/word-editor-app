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

# --- Helper Functions and Constants (Giữ nguyên) ---

def generate_vibrant_rgb_colors(count=150):
    """Generates a list of highly saturated, distinct RGB colors."""
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

# Regexes remain the same
SPEAKER_REGEX_DELIMITER = re.compile(r"([A-Z][a-z\s&]+):\s*", re.IGNORECASE)
TIMECODE_REGEX = re.compile(r"^\d{2}:\d{2}:\d{2},\d{3}\s+-->\s+\d{2}:\d{2}:\d{2},\d{3}$")
HTML_CONTENT_REGEX = re.compile(r"((?:</?[ibu]>)+)(.*?)(?:</?[ibu]>)+", re.IGNORECASE | re.DOTALL)

def set_all_text_formatting(doc):
    """Áp dụng định dạng chung cho toàn bộ văn bản."""
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
        
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(6)

def apply_html_formatting_to_run(paragraph, current_text):
    """Thêm nội dung văn bản, xử lý các thẻ HTML <i>, <b>, <u>."""
    if not current_text.strip():
        return
        
    matches = list(HTML_CONTENT_REGEX.finditer(current_text))
    last_end = 0
    
    for match in matches:
        tag_text = match.group(2) 
        start, end = match.span()

        if start > last_end:
            paragraph.add_run(current_text[last_end:start])
        
        run_html = paragraph.add_run(tag_text)
        run_html.font.bold = True
        run_html.font.italic = True
        
        last_end = end

    if last_end < len(current_text):
        paragraph.add_run(current_text[last_end:])

# Logic xử lý căn Tab triệt để
def format_and_split_dialogue(document, text):
    """
    Tách một dòng text thô (có thể chứa nhiều người nói) thành các đoạn văn bản 
    riêng biệt và áp dụng định dạng căn lề/Tab chính xác.
    """
    
    # Tách văn bản thành các phần dựa trên sự xuất hiện của tên người nói
    parts = SPEAKER_REGEX_DELIMITER.split(text)
    
    # --- CÁC THIẾT LẬP CĂN LỀ CHUNG ---
    TAB_STOP_POSITION = Inches(1.0) # Vị trí căn thẳng lời thoại
    
    # ---------------------------------------------
    # CASE 1: NO SPEAKER FOUND (Continuation Line)
    # ---------------------------------------------
    if len(parts) == 1:
        new_paragraph = document.add_paragraph()
        
        # Áp dụng cấu trúc Hanging Indent
        new_paragraph.paragraph_format.left_indent = TAB_STOP_POSITION
        new_paragraph.paragraph_format.first_line_indent = Inches(-1.0) 
        new_paragraph.paragraph_format.tab_stops.add_tab_stop(TAB_STOP_POSITION, WD_TAB_ALIGNMENT.LEFT)
        
        new_paragraph.add_run('\t') # Luôn chỉ dùng 1 Tab cho nội dung tiếp tục
        
        # BỎ DÒNG TRẮNG SAU KHI XỬ LÝ (Áp dụng Pt(0))
        new_paragraph.paragraph_format.space_after = Pt(0) 
        new_paragraph.paragraph_format.space_before = Pt(0)
        
        apply_html_formatting_to_run(new_paragraph, text)
        return
    
    # ---------------------------------------------
    # CASE 2: ONE OR MORE SPEAKERS FOUND
    # ---------------------------------------------

    # parts[0] là nội dung TRƯỚC người nói đầu tiên (thường là continuation)
    leading_content = parts[0].strip()
    if leading_content:
        # Tạo một đoạn continuation cho nội dung dẫn đầu này
        continuation_paragraph = document.add_paragraph()
        
        # Áp dụng cấu trúc Hanging Indent
        continuation_paragraph.paragraph_format.left_indent = TAB_STOP_POSITION
        continuation_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
        continuation_paragraph.paragraph_format.tab_stops.add_tab_stop(TAB_STOP_POSITION, WD_TAB_ALIGNMENT.LEFT)
        
        continuation_paragraph.add_run('\t') # Luôn dùng 1 Tab cho continuation
        continuation_paragraph.paragraph_format.space_after = Pt(0) # BỎ DÒNG TRẮNG SAU KHI XỬ LÝ
        continuation_paragraph.paragraph_format.space_before = Pt(0)
        apply_html_formatting_to_run(continuation_paragraph, leading_content)
    
    
    # Lặp qua các cặp (Tên người nói + Nội dung)
    speaker_matches = list(SPEAKER_REGEX_DELIMITER.finditer(text))
    
    for i, match in enumerate(speaker_matches):
        speaker_full = match.group(0) # e.g., "Coby: "
        speaker_name = match.group(1).strip() # e.g., "Coby"
        start, end = match.span()
        
        # Xác định nội dung của người nói hiện tại
        if i + 1 < len(speaker_matches):
            next_match_start = speaker_matches[i+1].start()
        else:
            next_match_start = len(text)
            
        content = text[end:next_match_start].strip()

        new_paragraph = document.add_paragraph()
        
        # Áp dụng cấu trúc Hanging Indent cho tất cả các dòng đối thoại
        new_paragraph.paragraph_format.left_indent = TAB_STOP_POSITION
        new_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
        
        # Đặt Tab Stop ở vị trí 1.0 inch
        new_paragraph.paragraph_format.tab_stops.add_tab_stop(TAB_STOP_POSITION, WD_TAB_ALIGNMENT.LEFT)
        
        # 1. Run cho tên người nói (Bold và Color)
        font_color_object = get_speaker_color(speaker_name) 
        run_speaker = new_paragraph.add_run(speaker_full)
        run_speaker.font.bold = True
        run_speaker.font.color.rgb = font_color_object 
        
        # 2. Xử lý Tab Linh hoạt (1 Tab hoặc 2 Tab) - YÊU CẦU CUỐI CÙNG
        # Nếu tên người nói (đã bao gồm ": ") dài hơn 10 ký tự, cần 2 Tabs
        if len(speaker_full) > 10:
             new_paragraph.add_run('\t\t') 
        else:
             new_paragraph.add_run('\t') 

        # 3. Thêm nội dung (NẰM TRÊN CÙNG DÒNG VỚI TÊN NGƯỜI NÓI)
        if content:
            apply_html_formatting_to_run(new_paragraph, content)

        # BỎ DÒNG TRẮNG SAU KHI XỬ LÝ
        new_paragraph.paragraph_format.space_after = Pt(0)
        new_paragraph.paragraph_format.space_before = Pt(0)
        
    return 

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
    
    # --- A. Set Main Title (FIX: Size 60, Thêm Dòng liệt kê Tên người nói) ---
    
    # 1. Làm sạch tên file để làm Tiêu đề
    title_text_raw = file_name_without_ext.upper()
    title_text = title_text_raw.replace("CONVERTED_", "").replace("FORMATTED_", "").replace("_EDIT", "").replace(" (GỐC)", "").strip()
    
    title_paragraph = document.add_paragraph(title_text)
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.paragraph_format.space_before = Pt(0)
    title_paragraph.paragraph_format.space_after = Pt(0) 
    
    title_run = title_paragraph.runs[0]
    title_run.font.name = 'Times New Roman'
    title_run.font.size = Pt(60) # FIX: Gấp 3 lần (20pt * 3 = 60pt)
    title_run.bold = True
    
    # 2. Thu thập tất cả tên người nói duy nhất
    unique_speakers = set()
    for paragraph in original_document.paragraphs:
        text = paragraph.text
        for match in SPEAKER_REGEX_DELIMITER.finditer(text):
            unique_speakers.add(match.group(1).strip())
            
    sorted_speakers = sorted(list(unique_speakers))
    
    # 3. Thêm Dòng liệt kê Tên người nói (Size 12, Normal)
    if sorted_speakers:
        speaker_list_text = "NGƯỜI NÓI: " + ", ".join(sorted_speakers)
        speaker_list_paragraph = document.add_paragraph(speaker_list_text)
        
        # Áp dụng định dạng Size 12, không in đậm
        for run in speaker_list_paragraph.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            run.font.bold = False
        
        # Dãn đoạn 6pt sau dòng liệt kê
        speaker_list_paragraph.paragraph_format.space_after = Pt(6) 
        speaker_list_paragraph.paragraph_format.space_before = Pt(0)
    
    # Thêm 2 dòng trắng sau tiêu đề (từ yêu cầu trước)
    document.add_paragraph().paragraph_format.space_after = Pt(0)
    document.add_paragraph().paragraph_format.space_after = Pt(0)

    # --- B. Process raw paragraphs ---
    
    for paragraph in raw_paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        
        # FIX: BỎ dòng "SRT Conversion:..." hoàn toàn
        if text.lower().startswith("srt conversion:"):
            continue 
            
        # B.1 Remove SRT Line Numbers
        if re.fullmatch(r"^\s*\d+\s*$", text):
            continue 
            
        # B.2 Timecode (Có dãn đoạn 6pt sau Timecode)
        if TIMECODE_REGEX.match(text):
            new_paragraph = document.add_paragraph(text)
            for run in new_paragraph.runs:
                run.font.bold = True
            new_paragraph.paragraph_format.space_after = Pt(6) # FIX: Dãn đoạn 6pt sau timecode
            new_paragraph.paragraph_format.space_before = Pt(0) 
            
        # B.3 Dialogue Content (Không có dãn đoạn sau)
        else:
            format_and_split_dialogue(document, text)
            
    # C. Apply General Font/Size and Spacing (Global settings)
    set_all_text_formatting(document)
    
    # Save the file
    modified_file = io.BytesIO()
    document.save(modified_file)
    modified_file.seek(0)
    
    return modified_file

# --- FIX Đặt Tên File (Giữ nguyên) ---
def clean_file_name_for_output(original_filename):
    """Xóa tiền tố/hậu tố không mong muốn và thêm '_edit'."""
    name_without_ext = os.path.splitext(original_filename)[0]
    
    cleaned_name = name_without_ext.replace("CONVERTED_", "").replace("FORMATTED_", "").strip()
    cleaned_name = re.sub(r'\s*\(.*\)$', '', cleaned_name).strip() 
    
    if cleaned_name.lower().endswith("_edit"):
         cleaned_name = cleaned_name[:-5].strip()

    return f"{cleaned_name}_edit.docx"

# --- GIAO DIỆN STREAMLIT ---

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
    file_name_without_ext = os.path.splitext(original_filename)[0] 
    
    st.info(f"File received: **{original_filename}**.")
    
    if st.button("2. RUN AUTOMATIC FORMATTING"):
        with st.spinner('Đang xử lý và định dạng file...'):
            try:
                modified_file_io = process_docx(uploaded_file, file_name_without_ext)
                
                # Sử dụng hàm làm sạch tên file cho output
                new_filename = clean_file_name_for_output(original_filename)

                st.success("✅ Định dạng hoàn tất! Bạn có thể tải file về.")
                
                # Nút tải file
                st.download_button(
                    label="3. Tải File Word Đã Định Dạng Về",
                    data=modified_file_io,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                st.markdown("---")
                st.balloons()

            except Exception as e:
                st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")
                st.warning("Vui lòng kiểm tra lại định dạng file đầu vào.")
