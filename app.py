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
# Bỏ import base64

# --- Helper Functions and Constants ---

# Colors remain the same
FONT_COLORS_RGB = [
    (192, 0, 0), (0, 51, 153), (0, 102, 0), (102, 0, 102), (255, 128, 0), 
    (0, 153, 153), (204, 102, 0), (153, 153, 0), (255, 0, 127), (51, 51, 255), 
    (153, 51, 255), (0, 204, 0), (255, 165, 0), (255, 51, 51), (0, 204, 204), 
    (255, 204, 0), (102, 51, 0), (0, 128, 0), (153, 0, 76), (255, 255, 102)
]

speaker_color_map = {}
used_colors = [RGBColor(r, g, b) for r, g, b in FONT_COLORS_RGB]
random.shuffle(used_colors)

def get_speaker_color(speaker_name):
    # Logic to assign persistent random color
    if speaker_name not in speaker_color_map:
        if used_colors:
            color_object = used_colors.pop()
        else:
            r, g, b = random.choice(FONT_COLORS_RGB)
            color_object = RGBColor(r, g, b)
            
        speaker_color_map[speaker_name] = color_object
        
    return speaker_color_map[speaker_name]

# Regexes remain the same
SPEAKER_REGEX = re.compile(r"^([A-Z][a-z\s&]+):\s*", re.IGNORECASE)
TIMECODE_REGEX = re.compile(r"^\d{2}:\d{2}:\d{2},\d{3}\s+-->\s+\d{2}:\d{2}:\d{2},\d{3}$")
HTML_CONTENT_REGEX = re.compile(r"((?:</?[ibu]>)+)(.*?)(?:</?[ibu]>)+", re.IGNORECASE | re.DOTALL)

def set_all_text_formatting(doc):
    """Applies Times New Roman 12pt and specific Spacing (Before: 0pt, After: 6pt, Single Line) to all runs/paragraphs."""
    for paragraph in doc.paragraphs:
        # Áp dụng Font và Size
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
        
        # Thiết lập dãn đoạn chung cho tất cả các đoạn (sẽ được ghi đè bên dưới)
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE


def process_docx(uploaded_file, file_name_without_ext):
    """Performs all required document modifications by rebuilding the document to ensure correct formatting."""
    
    global speaker_color_map
    global used_colors
    speaker_color_map = {}
    used_colors = [RGBColor(r, g, b) for r, g, b in FONT_COLORS_RGB]
    random.shuffle(used_colors)
    
    original_document = Document(io.BytesIO(uploaded_file.getvalue()))
    raw_paragraphs = [p for p in original_document.paragraphs if p.text.strip()]
    
    document = Document()
    
    # --- A. Set Main Title (25pt, 2 blank lines after) ---
    title_paragraph = document.add_paragraph(file_name_without_ext.upper())
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_paragraph.paragraph_format.space_before = Pt(0)
    title_paragraph.paragraph_format.space_after = Pt(0) 
    
    title_run = title_paragraph.runs[0]
    title_run.font.name = 'Times New Roman'
    title_run.font.size = Pt(25) 
    title_run.bold = True
    
    # Add two blank paragraphs
    document.add_paragraph().paragraph_format.space_after = Pt(0)
    document.add_paragraph().paragraph_format.space_after = Pt(0)

    # --- B. Process raw paragraphs and add to new document ---
    
    for paragraph in raw_paragraphs:
        text = paragraph.text.strip()
        
        # B.1 Remove SRT Line Numbers
        if re.fullmatch(r"^\s*\d+\s*$", text):
            continue 
            
        new_paragraph = document.add_paragraph()
        new_paragraph.style = document.styles['Normal']
        new_paragraph.paragraph_format.space_before = Pt(0)
        
        # B.2 Bold Timecode (Không dãn đoạn)
        if TIMECODE_REGEX.match(text):
            new_paragraph.text = text
            for run in new_paragraph.runs:
                run.font.bold = True
            new_paragraph.paragraph_format.space_after = Pt(0) 

        # B.3 Nội dung (Speaker/Content)
        else:
            # FIX: Áp dụng dãn đoạn After 6pt (Cho các đoạn nội dung)
            new_paragraph.paragraph_format.space_after = Pt(6) 
            
            speaker_match = SPEAKER_REGEX.match(text)
            
            if speaker_match:
                # FIX CĂN LỀ: Dùng Tab Stop và Thụt lề treo 
                
                # 1. Thiết lập Thụt lề treo (Hanging Indent) 
                # Lề trái: 1 inch (tổng khối văn bản bắt đầu từ đây)
                new_paragraph.paragraph_format.left_indent = Inches(1.0)
                # Thụt lề dòng đầu: -1 inch (đưa tên người nói về vị trí 0)
                new_paragraph.paragraph_format.first_line_indent = Inches(-1.0)
                
                # 2. Đặt Tab Stop ở vị trí 1.0 inch để căn chỉnh nội dung đối thoại
                new_paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(1.0), WD_TAB_ALIGNMENT.LEFT)
                
                speaker_full = speaker_match.group(0) 
                speaker_name = speaker_match.group(1).strip()
                
                font_color_object = get_speaker_color(speaker_name) 
                rest_of_text = text[len(speaker_full):]
                
                # 1. Run for the speaker name (Bold and Font Color)
                run_speaker = new_paragraph.add_run(speaker_full)
                run_speaker.font.bold = True
                run_speaker.font.color.rgb = font_color_object 
                
                # 2. Insert Tab character to align the dialogue text (Bắt đầu khối căn đều)
                new_paragraph.add_run('\t') 
                
                current_text = rest_of_text
                
            else:
                # Nếu không có người nói, đảm bảo không có thụt lề
                new_paragraph.paragraph_format.left_indent = None
                new_paragraph.paragraph_format.first_line_indent = None
                current_text = text


            # --- B.4 Process HTML tags within the current_text (cho cả 2 trường hợp) ---
            
            matches = list(HTML_CONTENT_REGEX.finditer(current_text))
            last_end = 0
            
            # Xóa text cũ nếu có speaker để chỉ giữ lại nội dung đã định dạng
            if speaker_match:
                # Đảm bảo nội dung sau tab được thêm vào.
                pass 
            else:
                new_paragraph.text = "" # Xóa nội dung gốc để định dạng lại

            # Logic thêm text đã được định dạng
            for match in matches:
                tag_text = match.group(2) 
                start, end = match.span()

                # Add text BEFORE the tag (if any)
                if start > last_end:
                    new_paragraph.add_run(current_text[last_end:start])
                
                # Add the HTML content (Bold and Italic)
                run_html = new_paragraph.add_run(tag_text)
                run_html.font.bold = True
                run_html.font.italic = True
                
                last_end = end

            # Add remaining text AFTER the last tag (or the whole text if no tags found)
            if last_end < len(current_text):
                new_paragraph.add_run(current_text[last_end:])
            
            # Xử lý trường hợp không có tag và không có speaker (nội dung đơn thuần)
            elif not speaker_match and not matches:
                # Nếu không có tag và không có speaker, gán lại nội dung
                new_paragraph.add_run(current_text)
            
            # Xử lý trường hợp có speaker nhưng không có tag (nội dung đơn thuần sau tab)
            elif speaker_match and not matches:
                new_paragraph.add_run(current_text)

    # C. Apply General Font/Size and Spacing (Global settings)
    set_all_text_formatting(document)
    
    # Save the file
    modified_file = io.BytesIO()
    document.save(modified_file)
    modified_file.seek(0)
    
    return modified_file

# Bỏ hoàn toàn hàm get_base64_html_preview

# --- GIAO DIỆN STREAMLIT (Đã loại bỏ phần Preview) ---
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
                
                new_filename = f"FORMATTED_{original_filename}"

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
