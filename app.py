import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.dml import MSO_THEME_COLOR
from docx.enum.section import WD_HEADER_FOOTER
from docx.oxml.ns import qn
import io
import os
import re
import random

# --- 20 Highlight Colors (RGB) for better speaker differentiation ---
# These are basic, recognizable, yet non-harsh colors
HIGHLIGHT_COLORS = [
    (255, 255, 0),    # Yellow
    (153, 204, 255),  # Light Blue
    (152, 251, 152),  # Pale Green
    (255, 192, 203),  # Pink
    (192, 192, 192),  # Silver/Gray
    (255, 204, 153),  # Peach
    (204, 204, 0),    # Olive
    (255, 153, 204),  # Medium Pink
    (170, 170, 255),  # Lavender
    (255, 228, 181),  # Moccasin
    (128, 0, 128),    # Purple (Darker, ensure visibility)
    (0, 128, 128),    # Teal (Darker, ensure visibility)
    (255, 165, 0),    # Orange
    (0, 255, 255),    # Cyan
    (255, 0, 255),    # Magenta
    (100, 149, 237),  # Cornflower Blue
    (255, 99, 71),    # Tomato Red (Less harsh than pure red)
    (60, 179, 113),   # Medium Sea Green
    (218, 112, 214),  # Orchid
    (240, 230, 140)   # Khaki
]

# Dictionary to store speaker names and their assigned highlight color
# This ensures the same speaker always gets the same color within one run
speaker_color_map = {}
used_colors = list(HIGHLIGHT_COLORS)
random.shuffle(used_colors)

def get_speaker_color(speaker_name):
    """Assigns a persistent random color to a speaker."""
    if speaker_name not in speaker_color_map:
        if used_colors:
            # Pop a color from the randomized list
            color_rgb = used_colors.pop()
        else:
            # If all 20 colors are used, wrap around or use a default
            color_rgb = random.choice(HIGHLIGHT_COLORS)
            
        r, g, b = color_rgb
        speaker_color_map[speaker_name] = RGBColor(r, g, b)
        
    return speaker_color_map[speaker_name]

# Define regex patterns for speakers and timecodes
SPEAKER_REGEX = re.compile(r"^([A-Z][a-z\s&]+):\s*", re.IGNORECASE)
TIMECODE_REGEX = re.compile(r"^\d{2}:\d{2}:\d{2},\d{3}\s+-->\s+\d{2}:\d{2}:\d{2},\d{3}$")


def set_page_number(section):
    """Adds a page number to the right corner of the footer."""
    footer = section.footer
    # Use the first paragraph in the footer
    if footer.paragraphs:
        footer_paragraph = footer.paragraphs[0]
    else:
        footer_paragraph = footer.add_paragraph()

    # 1. Add Field for the current page number
    footer_paragraph.add_run().add_field('PAGE')
    
    # 2. Set right alignment
    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # 3. Apply general formatting to the page number run
    for run in footer_paragraph.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)

def set_all_text_formatting(doc):
    """Applies Times New Roman 12pt to all runs in the document."""
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            # Ensure any runs without explicit size also get 12pt
            if run.font.size is None:
                run.font.size = Pt(12)


def process_docx(uploaded_file, file_name_without_ext):
    """Performs all required document modifications."""
    
    # Reset color map for each processing run
    global speaker_color_map
    speaker_color_map = {}
    
    document = Document(io.BytesIO(uploaded_file.read()))
    
    # --- A. Set Main Title (Uppercase, Center, Bold, Size 16) ---
    # Ensure the first paragraph exists or create it
    if document.paragraphs:
        title_paragraph = document.paragraphs[0]
    else:
        title_paragraph = document.add_paragraph()
        
    title_paragraph.text = file_name_without_ext.upper()
    title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    title_run = title_paragraph.runs[0]
    title_run.font.name = 'Times New Roman'
    title_run.font.size = Pt(16)
    title_run.bold = True
    
    # --- B. Process other paragraphs ---
    
    paragraphs_to_remove = []
    
    # Iterate through paragraphs, starting from the second one (index 1) 
    # as the first one is now the title.
    # Note: We still iterate over ALL paragraphs but exclude the first one from text processing logic.
    for i, paragraph in enumerate(document.paragraphs):
        
        # Reset to Normal style for consistency and remove list formatting
        paragraph.style = document.styles['Normal']
        
        # Skip the newly created title paragraph (index 0) from the cleaning logic
        if i == 0:
            continue
            
        text = paragraph.text.strip()
        
        # --- B.1 Remove SRT Line Numbers ---
        # Detect paragraphs only containing a number (e.g., '1', '23')
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
                # Group 1 is the speaker name, Group 0 is the full match (e.g., "Ethan:")
                speaker_full = speaker_match.group(0)
                speaker_name = speaker_match.group(1).strip() # Name without colon
                
                highlight_color = get_speaker_color(speaker_name)
                rest_of_text = text[len(speaker_full):]
                
                # Rebuild paragraph with correct formatting
                paragraph.text = "" # Clear old content
                
                # Run for the speaker name (Bold and Highlight)
                run_speaker = paragraph.add_run(speaker_full)
                run_speaker.font.bold = True
                # Set highlight color using the custom RGB function
                r, g, b = highlight_color.rgb
                run_speaker.font.highlight_color = highlight_color
                
                # Run for the rest of the text
                paragraph.add_run(rest_of_text)


    # Delete the paragraphs identified as line numbers
    for paragraph in paragraphs_to_remove:
        # Clear the content instead of removing the paragraph object itself 
        # (simpler and safer implementation with python-docx)
        paragraph.clear()


    # --- C. Apply General Font/Size (Times New Roman 12) ---
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
    help="Only Microsoft Word .docx format is accepted."
)

if uploaded_file is not None:
    # Get the file name for the title and the new saved file name
    original_filename = uploaded_file.name
    file_name_without_ext = os.path.splitext(original_filename)[0]
    
    st.info(f"File received: **{original_filename}**. The main title will be: **{file_name_without_ext.upper()}**")
    
    if st.button("2. RUN AUTOMATIC FORMATTING"):
        with st.spinner('Processing and formatting the Word file...'):
            try:
                # Call the main processing function
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
