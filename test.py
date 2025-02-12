import streamlit as st
import fitz  # PyMuPDF
import os
import tempfile
from pptx import Presentation
from docx import Document
from PIL import Image
import ebooklib
from ebooklib import epub
import io

# Function to merge four pages into one
# (Same as before, used for PDFs and converted formats)
def merge_four_pages_with_numbers(input_pdf_path, output_pdf_path):
    doc = fitz.open(input_pdf_path)
    num_pages = len(doc)
    output_doc = fitz.open()

    page_width, page_height = doc[0].rect.width, doc[0].rect.height
    new_width, new_height = page_width * 2, page_height * 2

    for i in range(0, num_pages, 4):
        new_page = output_doc.new_page(width=new_width, height=new_height)
        for j in range(4):
            if i + j < num_pages:
                src_page = doc[i + j]
                x_offset, y_offset = (j % 2) * page_width, (j // 2) * page_height
                new_page.show_pdf_page(
                    fitz.Rect(x_offset, y_offset, x_offset + page_width, y_offset + page_height),
                    doc,
                    i + j
                )
                new_page.insert_text((x_offset + 10, y_offset + 20), f"Page {i + j + 1}", fontsize=12, color=(0, 0, 0))

    output_doc.save(output_pdf_path)
    output_doc.close()

# Convert PPTX to PDF
def convert_pptx_to_pdf(input_pptx_path, output_pdf_path):
    prs = Presentation(input_pptx_path)
    pdf_doc = fitz.open()
    
    for slide in prs.slides:
        img = slide_to_image(slide)
        pdf_bytes = image_to_pdf(img)
        pdf_doc.insert_pdf(fitz.open(stream=pdf_bytes, filetype="pdf"))
    
    pdf_doc.save(output_pdf_path)
    pdf_doc.close()

# Convert DOCX to PDF
def convert_docx_to_pdf(input_docx_path, output_pdf_path):
    doc = Document(input_docx_path)
    pdf_doc = fitz.open()
    
    for para in doc.paragraphs:
        text_page = fitz.open()
        text_page.insert_page(0, text=para.text, fontsize=12)
        pdf_doc.insert_pdf(text_page)
    
    pdf_doc.save(output_pdf_path)
    pdf_doc.close()

# Convert image to PDF
def image_to_pdf(image):
    img_bytes = io.BytesIO()
    image.save(img_bytes, format='PDF')
    return img_bytes.getvalue()

# Convert EPUB to PDF
def convert_epub_to_pdf(input_epub_path, output_pdf_path):
    book = epub.read_epub(input_epub_path)
    pdf_doc = fitz.open()
    
    for item in book.get_items():
        if item.get_type() == ebooklib.ITEM_DOCUMENT:
            text_page = fitz.open()
            text_page.insert_page(0, text=item.get_content().decode("utf-8"), fontsize=12)
            pdf_doc.insert_pdf(text_page)
    
    pdf_doc.save(output_pdf_path)
    pdf_doc.close()

# Streamlit UI
st.title("Student File Merger (4-in-1) with Page Numbers")
st.write("Upload a file, and we will merge every 4 pages/slides/images into one with page numbers.")

uploaded_file = st.file_uploader("Upload a file", type=["pdf", "pptx", "docx", "png", "jpg", "jpeg", "epub", "txt"])

if uploaded_file:
    temp_dir = tempfile.mkdtemp()
    temp_input_path = os.path.join(temp_dir, uploaded_file.name)
    with open(temp_input_path, "wb") as f:
        f.write(uploaded_file.read())

    original_name = os.path.splitext(uploaded_file.name)[0]
    output_pdf_path = os.path.join(temp_dir, f"{original_name}_merged.pdf")

    # Process based on file type
    if uploaded_file.type == "application/pdf":
        merge_four_pages_with_numbers(temp_input_path, output_pdf_path)
    elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.presentationml.presentation":
        ppt_pdf_path = os.path.join(temp_dir, "converted_ppt.pdf")
        convert_pptx_to_pdf(temp_input_path, ppt_pdf_path)
        merge_four_pages_with_numbers(ppt_pdf_path, output_pdf_path)
    elif uploaded_file.type == "application/vnd.openxmlformats-officedocument.wordprocessingml.document":
        doc_pdf_path = os.path.join(temp_dir, "converted_doc.pdf")
        convert_docx_to_pdf(temp_input_path, doc_pdf_path)
        merge_four_pages_with_numbers(doc_pdf_path, output_pdf_path)
    elif uploaded_file.type in ["image/png", "image/jpeg", "image/jpg"]:
        img = Image.open(temp_input_path).convert("RGB")
        img_pdf_path = os.path.join(temp_dir, "converted_img.pdf")
        with open(img_pdf_path, "wb") as f:
            f.write(image_to_pdf(img))
        merge_four_pages_with_numbers(img_pdf_path, output_pdf_path)
    elif uploaded_file.type == "application/epub+zip":
        epub_pdf_path = os.path.join(temp_dir, "converted_epub.pdf")
        convert_epub_to_pdf(temp_input_path, epub_pdf_path)
        merge_four_pages_with_numbers(epub_pdf_path, output_pdf_path)
    elif uploaded_file.type == "text/plain":
        txt_pdf_path = os.path.join(temp_dir, "converted_txt.pdf")
        with open(temp_input_path, "r", encoding="utf-8") as txt_file:
            text = txt_file.read()
        text_page = fitz.open()
        text_page.insert_page(0, text=text, fontsize=12)
        text_page.save(txt_pdf_path)
        merge_four_pages_with_numbers(txt_pdf_path, output_pdf_path)
    
    # Provide download button
    with open(output_pdf_path, "rb") as f:
        st.download_button(
            label="Download Processed PDF",
            data=f,
            file_name=f"{original_name}_merged.pdf",
            mime="application/pdf"
        )

    st.success(f"Processing complete! Download '{original_name}_merged.pdf' above.")
