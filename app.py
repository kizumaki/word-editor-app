import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.text import WD_LINE_SPACING
import io
import os
import re
import random

# --- Stable RGB Colors for Font (Text) Color (20 distinct options) ---
FONT_COLORS_RGB = [
    (192, 0, 0),      # Dark Red
    (0, 51, 153),     # Dark Blue
    (0, 102, 0),      # Dark Green
    (102, 0, 102),    # Purple
    (255, 128, 0),    # Orange
    (0, 153, 153),    # Teal
    (204, 102, 0),    # Brown
    (153, 153, 0),    # Olive
    (255, 0, 127),    # Bright Pink
    (51, 51, 255),    # Medium Blue
    (153, 51, 255),   # Lavender
    (0, 204, 0),      # Bright Green
    (255, 165, 0),    # Orange-Gold
    (255, 51, 51),    # Light Red
    (0, 204, 204),    # Cyan
    (255, 204, 0),    # Gold
    (102, 51, 0),     # Dark Brown
    (0, 128, 0),      # Standard Green
    (153, 0, 76),     # Wine
    (255, 255, 102)   # Light Yellow
]

# Dictionary to store speaker names and their assigned font color (RGBColor object)
speaker_color_map = {}
used_colors = [RGBColor(r, g, b) for r, g, b in FONT_COLORS_RGB]
random.shuffle(used_colors)

def get_speaker_color(speaker_name):
    """Assigns a persistent random RGBColor object (for font color) to a speaker."""
    if speaker_name not in speaker_color_map:
        if used_colors:
            color_object = used_colors.pop()
        else:
            r, g, b = random.choice(FONT_COLORS_RGB)
            color_object = RGBColor(r, g, b)
            
        speaker_color_map[speaker_name] = color_object
        
    return speaker_color_map[speaker_name]

# Define regex patterns for speakers and timecodes and HTML tags
SPEAKER_REGEX = re.compile(r"^([A-Z][a-z\s&]+):\s*", re.IGNORECASE)
TIMECODE_REGEX = re.compile(r"^\d{2}:\d{2}:\d{2},\d{3}\s+-->\s+\d{2}:\d{2}:\d{2},\d{3}$")
HTML_TAG_REGEX = re.compile(r"(</?[ibu]>)+", re.IGNORECASE)
HTML_CONTENT_REGEX = re.compile(r"((?:</?[ibu]>)+)(.*?)(?:</?[ibu]>)+", re.IGNORECASE | re.DOTALL)


def set_all_text_formatting(doc):
    """Applies Times New Roman 12pt and standard line spacing to all runs/paragraphs."""
    for paragraph in doc.paragraphs:
        # Set line spacing to Single
        paragraph.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
        
        # Ensure all paragraphs start with 0 space after, we will manually add space for content paragraphs
        paragraph.paragraph_format.space_after = Pt(0) 
        
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            if run.font.size is None:
                run.font.size = Pt(12)


def process_docx(uploaded_file, file_name_without_ext):
    """Performs all required document modifications."""
    
    global speaker_color_map
    global used_colors
    speaker_color_map = {}
    used_colors = [RGBColor(r, g, b) for r, g, b in FONT_COLORS_RGB]
    random.shuffle(used_colors)
    
    document = Document(io.BytesIO(uploaded_file.read()))
    
    # --- A. Set Main Title (Size 25, 2 blank lines after) ---
    if not document.paragraphs:
        document.add_paragraph()
        
    title_paragraph = document.paragraphs[0]
        
    title_paragraph.text = file_name_without_ext.upper()
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    if title_paragraph.runs:
        title_run = title_paragraph.runs[0]
    else:
        title_run = title_paragraph.add_run(title_paragraph.text)
        
    title_run.font.name = 'Times New Roman'
    title_run.font.size = Pt(25) 
    title_run.bold = True
    
    # Add two blank paragraphs *after* the title paragraph to ensure two empty lines
    document.add_paragraph() # Dòng trống 1
    document.add_paragraph() # Dòng trống 2

    # --- B. Process other paragraphs ---
    
    paragraphs_to_remove = []
    all_paragraphs = list(document.paragraphs)

    # Start iteration from the 4th paragraph (index 3) which is the first potential content line.
    for i, paragraph in enumerate(all_paragraphs):
        
        if i <= 2: # Skip the title (0) and the two blank paragraphs (1, 2)
            continue
            
        paragraph.style = document.styles['Normal']
            
        text = paragraph.text.strip()
        
        # --- B.1 Remove SRT Line Numbers ---
        if re.fullmatch(r"^\s*\d+\s*$", text):
            paragraphs_to_remove.append(paragraph)
            continue
            
        # --- B.2 Bold Timecode ---
        if TIMECODE_REGEX.match(text):
            for run in paragraph.runs:
                run.font.bold = True
            
        # --- B.3 Bold Speaker Name and Random Font Color ---
        else:
            # FIX: Apply Pt(6) space after for content paragraphs (Close distance)
            paragraph.paragraph_format.space_after = Pt(6)
            
            speaker_match = SPEAKER_REGEX.match(text)
            if speaker_match:
                speaker_full = speaker_match.group(0) 
                speaker_name = speaker_match.group(1).strip()
                
                font_color_object = get_speaker_color(speaker_name) 
                rest_of_text = text[len(speaker_full):]
                
                # Rebuild paragraph
                paragraph.text = "" 
                
                # Run for the speaker name (Bold and Font Color)
                run_speaker = paragraph.add_run(speaker_full)
                run_speaker.font.bold = True
                run_speaker.font.color.rgb = font_color_object 
                
                # --- B.4 Process HTML tags within the rest of the text ---
                
                current_text = rest_of_text
                
                # Find all HTML content (e.g., <i>(Vọng lại:)</i>)
                matches = list(HTML_CONTENT_REGEX.finditer(current_text))

                last_end = 0
                for match in matches:
                    full_match = match.group(0)
                    tag_text = match.group(2)
                    start, end = match.span()

                    # Add text BEFORE the tag (if any)
                    if start > last_end:
                        paragraph.add_run(current_text[last_end:start])
                    
                    # Add the HTML content (Bold and Italic)
                    run_html = paragraph.add_run(tag_text)
                    run_html.font.bold = True
                    run_html.font.italic = True
                    
                    last_end = end

                # Add remaining text AFTER the last tag
                if last_end < len(current_text):
                    paragraph.add_run(current_text[last_end:])

            # If no speaker is found, just process the text for HTML tags (if it's not a timecode)
            else:
                # The paragraph already contains the text, now we re-process runs
                current_text = paragraph.text.strip()
                paragraph.text = ""
                
                matches = list(HTML_CONTENT_REGEX.finditer(current_text))
                
                last_end = 0
                for match in matches:
                    full_match = match.group(0)
                    tag_text = match.group(2)
                    start, end = match.span()

                    # Add text BEFORE the tag (if any)
                    if start > last_end:
                        paragraph.add_run(current_text[last_end:start])
                    
                    # Add the HTML content (Bold and Italic)
                    run_html = paragraph.add_run(tag_text)
                    run_html.font.bold = True
                    run_html.font.italic = True
                    
                    last_end = end

                # Add remaining text AFTER the last tag
                if last_end < len(current_text):
                    paragraph.add_run(current_text[last_end:])


    # Delete the content of paragraphs identified as line numbers
    for paragraph in paragraphs_to_remove:
        paragraph.clear()

    # --- C. Apply General Font/Size and Spacing ---
    set_all_text_formatting(document)
    
    # Save the modified file to an in-memory buffer
    modified_file = io.BytesIO()
    document.save(modified_file)
    modified_file.seek(0)
    
    return modified_file

# --- STREAMLIT INTERFACE ---
st.set_page_config(page_title="Automatic Word Script Editor", layout="wide")

st.markdown("## 📄 Automatic Subtitle Script (.docx) Converter")
st.markdown("A Python/Streamlit tool to automatically format subtitle scripts based on specific requirements.")
st.markdown("---")

uploaded_file = st.file_uploader(
    "1. Upload your Word file (.docx)",
    type=['docx'],
    help="Only Microsoft Word .docx format is accepted. Max size 200MB."
)

if uploaded_file is not None:
    original_filename = uploaded_file.name
    file_name_without_ext = os.path.splitext(original_filename)[0]
    
    st.info(f"File received: **{original_filename}**. The main title will be: **{file_name_without_ext.upper()}**")
    
    if st.button("2. RUN AUTOMATIC FORMATTING"):
        with st.spinner('Processing and formatting the Word file...'):
            try:
                modified_file_io = process_docx(uploaded_file, file_name_without_ext)
                
                new_filename = f"FORMATTED_{original_filename}"

                st.success("✅ Formatting complete! You can download the file.")
                
                # Download button
                st.download_button(
                    label="3. Download Formatted Word File",
                    data=modified_file_io,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
                st.markdown("---")
                st.balloons()

            except Exception as e:
                # This catches any remaining error and displays the specific message
                st.error(f"An error occurred during processing: {e}")
                st.warning("Please check the formatting of your input file (e.g., timecodes and speaker names must match the expected pattern).")
