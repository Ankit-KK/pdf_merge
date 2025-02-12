import streamlit as st
import fitz
import os
from pathlib import Path
import tempfile
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE
from copy import deepcopy

def merge_four_slides_pptx(input_path, output_path):
    """Merge four slides into one in PowerPoint while preserving all content"""
    # Open the presentation
    prs = Presentation(input_path)
    new_prs = Presentation()
    
    # Set slide width and height (standard 16:9)
    new_prs.slide_width = prs.slide_width
    new_prs.slide_height = prs.slide_height
    
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
                x_offset = 0 if j % 2 == 0 else quad_width
                y_offset = 0 if j < 2 else quad_height
                
                # Copy all shapes from source slide
                for shape in src_slide.shapes:
                    # Get element position and size
                    left = shape.left
                    top = shape.top
                    width = shape.width
                    height = shape.height
                    
                    # Calculate new position and size
                    new_left = x_offset + (left * 0.5)
                    new_top = y_offset + (top * 0.5)
                    new_width = width * 0.5
                    new_height = height * 0.5
                    
                    # Copy shape based on its type
                    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                        # For pictures
                        image = shape.image
                        new_picture = new_slide.shapes.add_picture(
                            image.blob,
                            new_left,
                            new_top,
                            new_width,
                            new_height
                        )
                    elif shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                        # For basic shapes
                        new_shape = new_slide.shapes.add_shape(
                            shape.auto_shape_type,
                            new_left,
                            new_top,
                            new_width,
                            new_height
                        )
                        # Copy fill
                        if hasattr(shape.fill, 'fore_color'):
                            new_shape.fill.fore_color.rgb = shape.fill.fore_color.rgb
                    elif hasattr(shape, 'text'):
                        # For text boxes
                        text_box = new_slide.shapes.add_textbox(
                            new_left,
                            new_top,
                            new_width,
                            new_height
                        )
                        text_frame = text_box.text_frame
                        text_frame.text = shape.text
                        
                        # Copy text formatting
                        for p_idx, p in enumerate(shape.text_frame.paragraphs):
                            if p_idx < len(text_frame.paragraphs):
                                new_p = text_frame.paragraphs[p_idx]
                                new_p.text = p.text
                                if p.font:
                                    if p.font.size:
                                        new_p.font.size = Pt(int(p.font.size.pt * 0.5))
                
                # Add slide number
                slide_num = new_slide.shapes.add_textbox(
                    x_offset + Inches(0.1),
                    y_offset + Inches(0.1),
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
        st.error("Error details for debugging:", str(e.__class__.__name__))
