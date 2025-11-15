import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_HEADER_FOOTER
from docx.enum.text import WD_COLOR_INDEX
import io
import os
import re
import random

# --- FINAL FIX: Use integer values for stable WD_COLOR_INDEX ---
# These integers map directly to the color constants (e.g., 6 is Yellow)
# This array is the stable solution for highlighting text across different python-docx versions.
HIGHLIGHT_COLORS_INDEX = [
    6,  # YELLOW
    11, # BRIGHT_GREEN
    3,  # TURQUOISE
    13, # VIOLET
    14, # PINK
    9,  # RED
    10, # DARK_BLUE
    15, # TEAL
    16, # GRAY_25
    17, # GRAY_50
    12, # LIME
    7,  # GOLD
    5,  # LIGHT_ORANGE
    1,  # PALE_BLUE
    18, # SEA_GREEN
    8,  # BLUE
    4,  # DARK_RED
    19, # DARK_YELLOW
    0,  # AUTO (No Color - fallback)
    1,  # WHITE (PALE_BLUE is 1, using it as a low-contrast color)
]

# Dictionary to store speaker names and their assigned highlight color (integer index)
speaker_color_map = {}
# Use a shuffled list of WD_COLOR_INDEX integers
used_colors = list(HIGHLIGHT_COLORS_INDEX)
random.shuffle(used_colors)

def get_speaker_color(speaker_name):
    """Assigns a persistent random WD_COLOR_INDEX integer to a speaker."""
    if speaker_name not in speaker_color_map:
        if used_colors:
            color_index = used_colors.pop()
        else:
            color_index = random.choice(HIGHLIGHT_COLORS_INDEX)
            
        speaker_color_map[speaker_name] = color_index
        
    return speaker_color_map[speaker_name]

# Define regex patterns for speakers and timecodes
SPEAKER_REGEX = re.compile(r"^([A-Z][a-z\s&]+):\s*", re.IGNORECASE)
TIMECODE_REGEX = re.compile(r"^\d{2}:\d{2}:\d{2},\d{3}\s+-->\s+\d{2}:\d{2}:\d{2},\d{3}$")


def set_page_number(section):
    """Adds a page number to the right corner of the footer."""
    footer = section.footer
    
    if not footer.paragraphs:
        # If footer is empty, add a new paragraph
        footer.add_paragraph()
        
    footer_paragraph = footer.paragraphs[0]
    
    # FIX: add_field must be called on the Paragraph object
    footer_paragraph.add_field('PAGE') 
    
    # Set right alignment
    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # Apply general formatting to the page number field
    if footer_paragraph.runs:
        run_page = footer_paragraph.runs[-1]
        run_page.font.name = 'Times New Roman'
        run_page.font.size = Pt(12)

def set_all_text_formatting(doc):
    """Applies Times New Roman 12pt to all runs in the document."""
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            if run.font.size is None:
                run.font.size = Pt(12)


def process_docx(uploaded_file, file_name_without_ext):
    """Performs all required document modifications."""
    
    # Reset color map and used colors for each processing run
    global speaker_color_map
    global used_colors
    speaker_color_map = {}
    used_colors = list(HIGHLIGHT_COLORS_INDEX)
    random.shuffle(used_colors)
    
    document = Document(io.BytesIO(uploaded_file.read()))
    
    # --- A. Set Main Title ---
    if document.paragraphs:
        title_paragraph = document.paragraphs[0]
    else:
        title_paragraph = document.add_paragraph()
        
    title_paragraph.text = file_name_without_ext.upper()
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    if title_paragraph.runs:
        title_run = title_paragraph.runs[0]
    else:
        title_run = title_paragraph.add_run(title_paragraph.text)
        
    title_run.font.name = 'Times New Roman'
    title_run.font.size = Pt(16)
    title_run.bold = True
    
    # --- B. Process other paragraphs ---
    
    paragraphs_to_remove = []
    
    all_paragraphs = list(document.paragraphs)

    for i, paragraph in enumerate(all_paragraphs):
        
        if i == 0:
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
            
        # --- B.3 Bold Speaker Name and Random Highlight ---
        else:
            speaker_match = SPEAKER_REGEX.match(text)
            if speaker_match:
                speaker_full = speaker_match.group(0) 
                speaker_name = speaker_match.group(1).strip()
                
                highlight_color_index = get_speaker_color(speaker_name) 
                rest_of_text = text[len(speaker_full):]
                
                # Rebuild paragraph
                paragraph.text = "" 
                
                # Run for the speaker name (Bold and Highlight)
                run_speaker = paragraph.add_run(speaker_full)
                run_speaker.font.bold = True
                
                # Apply WD_COLOR_INDEX integer 
                run_speaker.font.highlight_color = highlight_color_index 
                
                # Run for the rest of the text
                paragraph.add_run(rest_of_text)

    # Delete the content of paragraphs identified as line numbers
    for paragraph in paragraphs_to_remove:
        paragraph.clear()

    # --- C. Apply General Font/Size ---
    set_all_text_formatting(document)
    
    # --- D. Add Page Numbering ---
    for section in document.sections:
        set_page_number(section)

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
                st.error(f"An error occurred during processing: {e}")
                st.warning("Please check the formatting of your input file (e.g., timecodes and speaker names must match the expected pattern).")
