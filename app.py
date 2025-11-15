import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_HEADER_FOOTER
import io
import os
import re
import random

# --- 20 Highlight Colors (RGB) for better speaker differentiation ---
HIGHLIGHT_COLORS_RGB = [
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
    (128, 0, 128),    # Purple 
    (0, 128, 128),    # Teal 
    (255, 165, 0),    # Orange
    (0, 255, 255),    # Cyan
    (255, 0, 255),    # Magenta
    (100, 149, 237),  # Cornflower Blue
    (255, 99, 71),    # Tomato Red 
    (60, 179, 113),   # Medium Sea Green
    (218, 112, 214),  # Orchid
    (240, 230, 140)   # Khaki
]

# Dictionary to store speaker names and their assigned highlight color (RGBColor object)
speaker_color_map = {}
# Create a list of RGBColor objects once
used_colors = [RGBColor(r, g, b) for r, g, b in HIGHLIGHT_COLORS_RGB]
random.shuffle(used_colors)

def get_speaker_color(speaker_name):
    """Assigns a persistent random RGBColor object to a speaker."""
    if speaker_name not in speaker_color_map:
        if used_colors:
            # Pop a color from the randomized list
            color_object = used_colors.pop()
        else:
            # If all 20 colors are used, wrap around
            r, g, b = random.choice(HIGHLIGHT_COLORS_RGB)
            color_object = RGBColor(r, g, b)
            
        speaker_color_map[speaker_name] = color_object
        
    return speaker_color_map[speaker_name]

# Define regex patterns for speakers and timecodes
# Speaker: Starts with a capitalized name/names followed by a colon (e.g., Ethan:, Ryan:, Ethan & Leo:)
SPEAKER_REGEX = re.compile(r"^([A-Z][a-z\s&]+):\s*", re.IGNORECASE)
# Timecode: e.g., 00:00:00,913 --> 00:00:04,520
TIMECODE_REGEX = re.compile(r"^\d{2}:\d{2}:\d{2},\d{3}\s+-->\s+\d{2}:\d{2}:\d{2},\d{3}$")


def set_page_number(section):
    """Adds a page number to the right corner of the footer."""
    footer = section.footer
    
    # Ensure footer has at least one paragraph
    if not footer.paragraphs:
        footer.add_paragraph()
        
    footer_paragraph = footer.paragraphs[0]
    
    # 1. Add Field for the current page number
    run_page = footer_paragraph.add_run()
    run_page.add_field('PAGE')
    
    # 2. Set right alignment
    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # 3. Apply general formatting to the page number run
    # Ensure it only applies to the newly added run for consistency
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
    
    # Reset color map for each processing run
    global speaker_color_map
    global used_colors
    speaker_color_map = {}
    used_colors = [RGBColor(r, g, b) for r, g, b in HIGHLIGHT_COLORS_RGB]
    random.shuffle(used_colors)
    
    document = Document(io.BytesIO(uploaded_file.read()))
    
    # --- A. Set Main Title (Uppercase, Center, Bold, Size 16) ---
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
    
    # Use a list comprehension to get paragraphs to avoid issues with modification during iteration
    all_paragraphs = list(document.paragraphs)

    for i, paragraph in enumerate(all_paragraphs):
        
        # Skip the title paragraph (index 0) from the cleaning logic
        if i == 0:
            continue
            
        # Reset to Normal style for consistency and remove list formatting
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
                speaker_full = speaker_match.group(0) # e.g., "Ethan: " (with colon and space)
                speaker_name = speaker_match.group(1).strip() # e.g., "Ethan"
                
                highlight_color_object = get_speaker_color(speaker_name)
                rest_of_text = text[len(speaker_full):]
                
                # Rebuild paragraph with correct formatting
                paragraph.text = "" # Clear old content
                
                # Run for the speaker name (Bold and Highlight)
                run_speaker = paragraph.add_run(speaker_full)
                run_speaker.font.bold = True
                
                # Apply the RGBColor object directly to the highlight attribute (FIXED)
                run_speaker.font.highlight_color = highlight_color_object
                
                # Run for the rest of the text
                paragraph.add_run(rest_of_text)

    # Delete the content of paragraphs identified as line numbers
    for paragraph in paragraphs_to_remove:
        # Clear the content instead of removing the paragraph object itself
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
    help="Only Microsoft Word .docx format is accepted. Max size 200MB."
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
