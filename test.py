import streamlit as st
import fitz
import os
from pathlib import Path
import tempfile
from pptx import Presentation
from pptx.util import Inches, Pt
import io

def merge_four_slides_pptx(input_path, output_path):
    """Merge four slides into one in PowerPoint"""
    # Open the presentation
    prs = Presentation(input_path)
    new_prs = Presentation()
    
    # Set slide width and height (standard 16:9)
    new_prs.slide_width = Inches(13.333)
    new_prs.slide_height = Inches(7.5)
    
    # Calculate dimensions for each quadrant
    quad_width = new_prs.slide_width / 2
    quad_height = new_prs.slide_height / 2
    
    # Process slides in groups of 4
    for i in range(0, len(prs.slides), 4):
        # Add a blank slide
        blank_slide_layout = new_prs.slide_layouts[6]  # Usually layout 6 is blank
        new_slide = new_prs.slides.add_slide(blank_slide_layout)
        
        # Process up to 4 slides for this group
        for j in range(4):
            if i + j < len(prs.slides):
                src_slide = prs.slides[i + j]
                
                # Calculate position for this quadrant
                left = quad_width * (j % 2)
                top = quad_height * (j // 2)
                
                # Copy all shapes from source slide
                for shape in src_slide.shapes:
                    # Scale factor for the shape
                    scale_x = 0.5  # Because we're fitting into half width
                    scale_y = 0.5  # Because we're fitting into half height
                    
                    if shape.shape_type == 17:  # If shape is a connector
                        continue  # Skip connectors as they can cause issues
                        
                    # Get the original dimensions
                    orig_left = shape.left
                    orig_top = shape.top
                    orig_width = shape.width
                    orig_height = shape.height
                    
                    # Calculate new position and size
                    new_left = left + (orig_left * scale_x)
                    new_top = top + (orig_top * scale_y)
                    new_width = orig_width * scale_x
                    new_height = orig_height * scale_y
                    
                    # Copy shape to new position
                    if hasattr(shape, 'text'):
                        text_box = new_slide.shapes.add_textbox(
                            new_left, new_top, new_width, new_height
                        )
                        text_frame = text_box.text_frame
                        text_frame.text = shape.text
                        
                        # Copy text formatting
                        for para_idx, paragraph in enumerate(shape.text_frame.paragraphs):
                            if para_idx < len(text_frame.paragraphs):
                                new_para = text_frame.paragraphs[para_idx]
                                new_para.text = paragraph.text
                                
                                # Handle text formatting
                                if hasattr(paragraph, 'runs') and paragraph.runs:
                                    for run_idx, run in enumerate(paragraph.runs):
                                        if run_idx < len(new_para.runs):
                                            new_run = new_para.runs[run_idx]
                                            if hasattr(run.font, 'size') and run.font.size is not None:
                                                try:
                                                    new_size = Pt(int(run.font.size.pt * scale_y))
                                                    new_run.font.size = new_size
                                                except AttributeError:
                                                    # Skip if size cannot be determined
                                                    pass
                    
                # Add slide number
                slide_num = new_slide.shapes.add_textbox(
                    left + Inches(0.1),
                    top + Inches(0.1),
                    Inches(1),
                    Inches(0.3)
                )
                slide_num.text_frame.text = f"Slide {i + j + 1}"
    
    # Save the presentation
    new_prs.save(output_path)

def merge_four_pages_pdf(input_pdf_path, output_pdf_path):
    """Merge four pages into one in PDF"""
    doc = fitz.open(input_pdf_path)
    num_pages = len(doc)
    output_doc = fitz.open()
    
    # Page dimensions
    page_width = doc[0].rect.width
    page_height = doc[0].rect.height
    
    # New page dimensions
    new_width = page_width * 2
    new_height = page_height * 2
    
    for i in range(0, num_pages, 4):
        # Create a new blank page
        new_page = output_doc.new_page(width=new_width, height=new_height)
        
        for j in range(4):
            if i + j < num_pages:
                src_page = doc[i + j]
                x_offset = (j % 2) * page_width
                y_offset = (j // 2) * page_height
                
                new_page.show_pdf_page(
                    fitz.Rect(x_offset, y_offset, x_offset + page_width, y_offset + page_height),
                    doc,
                    i + j
                )
                
                # Add page number
                page_number = i + j + 1
                new_page.insert_text(
                    (x_offset + 10, y_offset + 20),
                    f"Page {page_number}",
                    fontsize=12,
                    color=(0, 0, 0)
                )
    
    output_doc.save(output_pdf_path)
    output_doc.close()

# Streamlit UI
st.title("Document Merger (4-in-1)")
st.write("Upload a PDF or PowerPoint file, and we will merge every 4 pages into one.")

uploaded_file = st.file_uploader("Upload a document", type=["pdf", "pptx"])

if uploaded_file:
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            # Get original file info
            original_name = Path(uploaded_file.name).stem
            file_extension = Path(uploaded_file.name).suffix.lower()
            
            # Save the uploaded file temporarily
            temp_input_path = os.path.join(temp_dir, f"input{file_extension}")
            with open(temp_input_path, "wb") as f:
                f.write(uploaded_file.read())
            
            # Output filename with same extension as input
            output_filename = f"{original_name}_merged{file_extension}"
            output_path = os.path.join(temp_dir, output_filename)
            
            # Process based on file type
            with st.spinner("Merging pages..."):
                if file_extension == '.pptx':
                    merge_four_slides_pptx(temp_input_path, output_path)
                else:  # PDF
                    merge_four_pages_pdf(temp_input_path, output_path)
            
            # Provide download button
            with open(output_path, "rb") as f:
                mime_type = 'application/vnd.openxmlformats-officedocument.presentationml.presentation' \
                    if file_extension == '.pptx' else 'application/pdf'
                st.download_button(
                    label="Download Processed Document",
                    data=f,
                    file_name=output_filename,
                    mime=mime_type
                )
            
            st.success(f"Processing complete! Download '{output_filename}' above.")
    
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
