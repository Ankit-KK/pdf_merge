import streamlit as st
import fitz
import os
from pathlib import Path
import tempfile
from pptx import Presentation
from docx import Document
from PIL import Image
import io

def convert_pptx_to_images(pptx_path):
    """Convert PowerPoint slides to images"""
    prs = Presentation(pptx_path)
    images = []
    
    for slide in prs.slides:
        # Save slide as PNG
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            # You would need to implement the actual slide-to-image conversion here
            # This is a placeholder for the conversion logic
            img = Image.new('RGB', (1920, 1080), 'white')
            img.save(tmp.name)
            images.append(tmp.name)
    
    return images

def convert_docx_to_images(docx_path):
    """Convert Word document pages to images"""
    doc = Document(docx_path)
    images = []
    
    # Since python-docx doesn't directly support page rendering,
    # we'll need to use a PDF intermediate step with a different library
    # This is a placeholder for the actual conversion logic
    img = Image.new('RGB', (1920, 1080), 'white')
    with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
        img.save(tmp.name)
        images.append(tmp.name)
    
    return images

def images_to_pdf(image_paths, output_path):
    """Convert images to PDF"""
    doc = fitz.open()
    for img_path in image_paths:
        img_doc = fitz.open(img_path)
        pdfbytes = img_doc.convert_to_pdf()
        img_pdf = fitz.open("pdf", pdfbytes)
        doc.insert_pdf(img_pdf)
    doc.save(output_path)
    doc.close()

def merge_four_pages_with_numbers(input_pdf_path, output_pdf_path):
    """Merge four pages into one with page numbers"""
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
st.title("Document Merger (4-in-1) with Page Numbers")
st.write("Upload a PDF, PowerPoint, or Word document, and we will merge every 4 pages into one with page numbers.")

uploaded_file = st.file_uploader("Upload a document", type=["pdf", "pptx", "docx"])

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
            
            # Convert to PDF if necessary
            temp_pdf_path = os.path.join(temp_dir, "temp.pdf")
            
            if file_extension == '.pptx':
                st.info("Converting PowerPoint to PDF...")
                images = convert_pptx_to_images(temp_input_path)
                images_to_pdf(images, temp_pdf_path)
                input_pdf_path = temp_pdf_path
                
                # Clean up temporary image files
                for img_path in images:
                    os.unlink(img_path)
            
            elif file_extension == '.docx':
                st.info("Converting Word document to PDF...")
                images = convert_docx_to_images(temp_input_path)
                images_to_pdf(images, temp_pdf_path)
                input_pdf_path = temp_pdf_path
                
                # Clean up temporary image files
                for img_path in images:
                    os.unlink(img_path)
            
            else:  # PDF
                input_pdf_path = temp_input_path
            
            # Output filename
            output_filename = f"{original_name}_merged.pdf"
            output_path = os.path.join(temp_dir, output_filename)
            
            # Process the PDF
            with st.spinner("Merging pages..."):
                merge_four_pages_with_numbers(input_pdf_path, output_path)
            
            # Provide download button
            with open(output_path, "rb") as f:
                st.download_button(
                    label="Download Processed Document",
                    data=f,
                    file_name=output_filename,
                    mime="application/pdf"
                )
            
            st.success(f"Processing complete! Download '{output_filename}' above.")
    
    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
