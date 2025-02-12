import streamlit as st
import fitz
import os
from pathlib import Path
import tempfile
from pdf2image import convert_from_path
import comtypes.client
import win32com.client
import pythoncom

def convert_ppt_to_pdf(input_path, output_path):
    """Convert PowerPoint to PDF"""
    pythoncom.CoInitialize()
    powerpoint = win32com.client.Dispatch("Powerpoint.Application")
    deck = powerpoint.Presentations.Open(input_path)
    deck.SaveAs(output_path, 32)  # 32 is the PDF format code
    deck.Close()
    powerpoint.Quit()

def convert_doc_to_pdf(input_path, output_path):
    """Convert Word document to PDF"""
    pythoncom.CoInitialize()
    word = win32com.client.Dispatch("Word.Application")
    doc = word.Documents.Open(input_path)
    doc.SaveAs(output_path, FileFormat=17)  # 17 is the PDF format code
    doc.Close()
    word.Quit()

def merge_four_pages_with_numbers(input_pdf_path, output_pdf_path):
    """Merge four pages into one with page numbers"""
    doc = fitz.open(input_pdf_path)
    num_pages = len(doc)
    output_doc = fitz.open()
    
    # Page dimensions
    page_width = doc[0].rect.width
    page_height = doc[0].rect.height
    
    # New page dimensions (same width, but 2x height)
    new_width = page_width * 2
    new_height = page_height * 2
    
    for i in range(0, num_pages, 4):
        # Create a new blank page
        new_page = output_doc.new_page(width=new_width, height=new_height)
        
        for j in range(4):
            if i + j < num_pages:  # Check if page exists
                src_page = doc[i + j]
                x_offset = (j % 2) * page_width  # Left or right
                y_offset = (j // 2) * page_height  # Top or bottom
                
                # Paste the source page onto the new page at the correct position
                new_page.show_pdf_page(
                    fitz.Rect(x_offset, y_offset, x_offset + page_width, y_offset + page_height),
                    doc,
                    i + j
                )
                
                # Add page number
                page_number = i + j + 1
                new_page.insert_text(
                    (x_offset + 10, y_offset + 20),  # Position near top-left
                    f"Page {page_number}",
                    fontsize=12,
                    color=(0, 0, 0)  # Black color
                )
    
    output_doc.save(output_pdf_path)
    output_doc.close()

# Streamlit UI
st.title("Document Merger (4-in-1) with Page Numbers")
st.write("Upload a PDF, PowerPoint, or Word document, and we will merge every 4 pages into one with page numbers.")

# File uploader that accepts multiple formats
uploaded_file = st.file_uploader("Upload a document", type=["pdf", "ppt", "pptx", "doc", "docx"])

if uploaded_file:
    try:
        # Create temporary directory
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
            
            if file_extension in ['.ppt', '.pptx']:
                st.info("Converting PowerPoint to PDF...")
                convert_ppt_to_pdf(temp_input_path, temp_pdf_path)
                input_pdf_path = temp_pdf_path
            
            elif file_extension in ['.doc', '.docx']:
                st.info("Converting Word document to PDF...")
                convert_doc_to_pdf(temp_input_path, temp_pdf_path)
                input_pdf_path = temp_pdf_path
            
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
        st.error("Please make sure you have Microsoft Office installed if processing PowerPoint or Word documents.")
