import streamlit as st
import fitz  # PyMuPDF
import os
from datetime import datetime
from io import BytesIO
from pptx import Presentation
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PIL import Image
import ebooklib
from ebooklib import epub
from bs4 import BeautifulSoup

# Function to merge pages with numbers
def merge_pages_with_numbers(input_pdf_path, output_pdf_path, rows=2, cols=2, 
                            font_size=12, page_num_pos="top-left", font_color=(0, 0, 0)):
    doc = fitz.open(input_pdf_path)
    num_pages = len(doc)
    output_doc = fitz.open()

    # Calculate layout dimensions
    page_width = doc[0].rect.width
    page_height = doc[0].rect.height
    new_width = page_width * cols
    new_height = page_height * rows

    # Page number position mapping
    pos_offsets = {
        "top-left": (10, 20),
        "top-right": (page_width - 50, 20),
        "bottom-left": (10, page_height - 20),
        "bottom-right": (page_width - 50, page_height - 20)
    }

    per_sheet = rows * cols
    total_sheets = (num_pages + per_sheet - 1) // per_sheet

    progress_bar = st.progress(0)
    status_text = st.empty()

    for sheet_num in range(total_sheets):
        status_text.text(f"Processing sheet {sheet_num + 1}/{total_sheets}")
        progress_bar.progress((sheet_num + 1) / total_sheets)
        
        new_page = output_doc.new_page(width=new_width, height=new_height)
        
        for pos_in_sheet in range(per_sheet):
            page_index = sheet_num * per_sheet + pos_in_sheet
            if page_index >= num_pages:
                break

            # Calculate position
            col = pos_in_sheet % cols
            row = pos_in_sheet // cols
            
            x_offset = col * page_width
            y_offset = row * page_height

            # Add page content
            src_page = doc[page_index]
            new_page.show_pdf_page(
                fitz.Rect(x_offset, y_offset, x_offset + page_width, y_offset + page_height),
                doc,
                page_index
            )

            # Add page number
            if page_num_pos in pos_offsets:
                num_x = x_offset + pos_offsets[page_num_pos][0]
                num_y = y_offset + pos_offsets[page_num_pos][1]
                
                new_page.insert_text(
                    (num_x, num_y),
                    f"{page_index + 1}",
                    fontsize=font_size,
                    color=font_color
                )

    progress_bar.empty()
    status_text.empty()
    output_doc.save(output_pdf_path)
    output_doc.close()

# Function to convert PPTX to PDF
def pptx_to_pdf(pptx_path, pdf_path):
    prs = Presentation(pptx_path)
    pdf_doc = fitz.open()
    
    for slide in prs.slides:
        img_path = f"temp_slide.png"
        slide.save(img_path)
        img = Image.open(img_path)
        img_page = pdf_doc.new_page(width=img.width, height=img.height)
        img_page.insert_image(img_page.rect, filename=img_path)
        os.remove(img_path)
    
    pdf_doc.save(pdf_path)
    pdf_doc.close()

# Function to convert DOCX to PDF
def docx_to_pdf(docx_path, pdf_path):
    doc = Document(docx_path)
    pdf_buffer = BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=letter)
    width, height = letter

    for para in doc.paragraphs:
        c.drawString(72, height - 72, para.text)
        c.showPage()

    c.save()
    pdf_buffer.seek(0)
    with open(pdf_path, "wb") as f:
        f.write(pdf_buffer.read())

# Function to convert images to PDF
def images_to_pdf(image_paths, pdf_path):
    pdf_doc = fitz.open()
    for img_path in image_paths:
        img = Image.open(img_path)
        img_page = pdf_doc.new_page(width=img.width, height=img.height)
        img_page.insert_image(img_page.rect, filename=img_path)
    pdf_doc.save(pdf_path)
    pdf_doc.close()

# Function to convert EPUB to PDF
def epub_to_pdf(epub_path, pdf_path):
    book = epub.read_epub(epub_path)
    pdf_doc = fitz.open()
    
    for item in book.get_items():
        if item.get_type() == ebooklib.ITEM_DOCUMENT:
            soup = BeautifulSoup(item.get_content(), 'html.parser')
            text = soup.get_text()
            pdf_page = pdf_doc.new_page()
            pdf_page.insert_text((72, 72), text)
    
    pdf_doc.save(pdf_path)
    pdf_doc.close()

# Function to convert TXT to PDF
def txt_to_pdf(txt_path, pdf_path):
    with open(txt_path, "r") as f:
        text = f.read()
    
    pdf_buffer = BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=letter)
    width, height = letter
    c.drawString(72, height - 72, text)
    c.save()
    
    pdf_buffer.seek(0)
    with open(pdf_path, "wb") as f:
        f.write(pdf_buffer.read())

# Streamlit UI
st.title("Multi-File Merger with Page Numbers")
st.markdown("Merge PDF, PPTX, DOCX, Images, EPUB, or TXT files into a single PDF with customizable page numbers")

with st.sidebar:
    st.header("Settings")
    cols = st.number_input("Columns per sheet", min_value=1, max_value=4, value=2)
    rows = st.number_input("Rows per sheet", min_value=1, max_value=4, value=2)
    font_size = st.slider("Page number size", 8, 24, 12)
    page_num_pos = st.selectbox("Page number position", 
                              ["top-left", "top-right", "bottom-left", "bottom-right"])
    font_color = st.color_picker("Page number color", "#000000")

uploaded_file = st.file_uploader("Upload file", type=["pdf", "pptx", "docx", "png", "jpg", "jpeg", "tiff", "bmp", "gif", "epub", "txt"])

if uploaded_file:
    try:
        # Generate unique filenames
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        temp_input = f"temp_input_{timestamp}.{uploaded_file.name.split('.')[-1]}"
        temp_pdf = f"temp_pdf_{timestamp}.pdf"
        output_file = f"merged_output_{timestamp}.pdf"

        # Save uploaded file
        with open(temp_input, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Convert to PDF if necessary
        if uploaded_file.type == "application/pdf":
            temp_pdf = temp_input
        elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.presentationml.presentation":
            pptx_to_pdf(temp_input, temp_pdf)
        elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
            docx_to_pdf(temp_input, temp_pdf)
        elif uploaded_file.type in ["image/png", "image/jpeg", "image/tiff", "image/bmp", "image/gif"]:
            images_to_pdf([temp_input], temp_pdf)
        elif uploaded_file.type == "application/epub+zip":
            epub_to_pdf(temp_input, temp_pdf)
        elif uploaded_file.type == "text/plain":
            txt_to_pdf(temp_input, temp_pdf)

        # Process PDF
        with st.spinner("Processing file..."):
            merge_pages_with_numbers(
                temp_pdf,
                output_file,
                rows=rows,
                cols=cols,
                font_size=font_size,
                page_num_pos=page_num_pos,
                font_color=tuple(int(font_color.lstrip('#')[i:i+2], 16) for i in (0, 2, 4))
            )

        # Show preview
        with fitz.open(output_file) as doc:
            if len(doc) > 0:
                pix = doc[0].get_pixmap()
                preview_img = pix.tobytes("png")
                st.image(preview_img, caption="First Page Preview", use_column_width=True)

        # Download button
        with open(output_file, "rb") as f:
            st.download_button(
                "Download Merged PDF",
                data=f,
                file_name=output_file,
                mime="application/pdf"
            )

    except Exception as e:
        st.error(f"Error processing file: {str(e)}")
    finally:
        # Cleanup temporary files
        for file in [temp_input, temp_pdf, output_file]:
            if os.path.exists(file):
                os.remove(file)
