import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_LINE_SPACING
import io
import os
import re
import random
import base64

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
                # FIX: Xử lý căn lề/Tab theo Ảnh 2: Thụt lề đầu dòng và Dùng Tab
                # Đặt tab stop ở vị trí mong muốn (ví dụ: 1 inch)
                new_paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(1.0), WD_TAB_ALIGNMENT.LEFT)
                
                speaker_full = speaker_match.group(0) 
                speaker_name = speaker_match.group(1).strip()
                
                font_color_object = get_speaker_color(speaker_name) 
                rest_of_text = text[len(speaker_full):]
                
                # 1. Run for the speaker name (Bold and Font Color)
                run_speaker = new_paragraph.add_run(speaker_full)
                run_speaker.font.bold = True
                run_speaker.font.color.rgb = font_color_object 
                
                # 2. Insert Tab character to align the dialogue text
                new_paragraph.add_run('\t') 
                
                current_text = rest_of_text
                
            else:
                current_text = text

            # --- B.4 Process HTML tags within the current_text ---
            
            # Nếu có người nói, tiếp tục thêm nội dung sau tab
            if speaker_match:
                matches = list(HTML_CONTENT_REGEX.finditer(current_text))
                last_end = 0
                
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

                # Add remaining text AFTER the last tag
                if last_end < len(current_text):
                    new_paragraph.add_run(current_text[last_end:])
            
            # Nếu không có người nói (chỉ là nội dung tiếp theo), thêm nội dung
            else:
                new_paragraph.text = current_text # Gán lại nội dung nếu không có speaker


    # C. Apply General Font/Size and Spacing (Global settings)
    set_all_text_formatting(document)
    
    # Save the file
    modified_file = io.Bytesिओ(document.save)
    modified_file.seek(0)
    
    return modified_file

# --- Streamlit Preview Helper ---
def get_base64_html_preview(docx_io):
    # Tạo base64 string từ file Word để nhúng vào HTML
    base64_docx = base64.b64encode(docx_io.read()).decode('utf-8')
    docx_io.seek(0)
    
    # HTML/JavaScript để tạo nút download nhanh
    html = f"""
    <div style="border: 1px solid #ccc; padding: 10px; text-align: center;">
        <p>⚠️ TÍNH NĂNG PREVIEW TRỰC TIẾP KHÔNG THỂ THỰC HIỆN ĐƯỢC.</p>
        <p>Vui lòng tải xuống file Word để xem thành phẩm cuối cùng.</p>
        <a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{base64_docx}" download="preview.docx" style="text-decoration: none;">
            <button style="padding: 10px 20px; background-color: #4CAF50; color: white; border: none; border-radius: 5px; cursor: pointer;">
                Tải xuống bản Preview để xem
            </button>
        </a>
    </div>
    """
    return html

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
                
                new_filename = f"FORMATTED_{original_filename}"

                st.success("✅ Định dạng hoàn tất! Bạn có thể xem và tải file về.")
                
                # Nút tải file
                st.download_button(
                    label="3. Tải File Word Đã Định Dạng Về",
                    data=modified_file_io,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                # Thêm Preview
                st.subheader("Xem trước thành phẩm")
                modified_file_io.seek(0) # Đặt lại con trỏ file trước khi dùng cho preview
                st.markdown(get_base64_html_preview(modified_file_io), unsafe_allow_html=True)
                
                st.markdown("---")
                st.balloons()

            except Exception as e:
                st.error(f"Đã xảy ra lỗi trong quá trình xử lý: {e}")
                st.warning("Vui lòng kiểm tra lại định dạng file đầu vào.")
