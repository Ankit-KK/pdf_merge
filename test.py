import streamlit as st
import fitz  # PyMuPDF for PDFs
import os
from pptx import Presentation
from docx import Document
from PIL import Image
import io
from pdf2image import convert_from_path
from ebooklib import epub
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import comtypes.client  # For converting PPTX to PDF on Windows

# Function to merge every 4 pages into one with numbers
def merge_four_pages_with_numbers(input_pdf_path, output_pdf_path):
    doc = fitz.open(input_pdf_path)
    num_pages = len(doc)
    output_doc = fitz.open()

    page_width = doc[0].rect.width
    page_height = doc[0].rect.height

    new_width = page_width * 2
    new_height = page_height * 2

    for i in range(0, num_pages, 4):
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

                page_number = i + j + 1
                new_page.insert_text(
                    (x_offset + 10, y_offset + 20),
                    f"Page {page_number}",
                    fontsize=12,
                    color=(0, 0, 0)
                )

    output_doc.save(output_pdf_path)
    output_doc.close()

# Function to convert PPTX slides to PDF using PowerPoint (Windows only)
def convert_pptx_to_pdf(input_path, output_path):
    powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
    powerpoint.Visible = 1
    deck = powerpoint.Presentations.Open(input_path)
    deck.SaveAs(output_path, 32)  # 32 is the constant for PDF format
    deck.Close()
    powerpoint.Quit()

# Function to convert Word DOCX to PDF
def convert_docx_to_pdf(input_path, output_path):
    doc = Document(input_path)
    pdf_canvas = canvas.Canvas(output_path, pagesize=letter)
    
    y_position = 750  # Start position for text
    for para in doc.paragraphs:
        pdf_canvas.drawString(100, y_position, para.text)
        y_position -= 20  # Move down for next line

        if y_position < 50:
            pdf_canvas.showPage()
            y_position = 750

    pdf_canvas.save()

# Function to convert images to PDF
def convert_images_to_pdf(image_paths, output_path):
    images = [Image.open(img_path).convert("RGB") for img_path in image_paths]
    images[0].save(output_path, save_all=True, append_images=images[1:])

# Function to convert EPUB to PDF
def convert_epub_to_pdf(input_path, output_path):
    book = epub.read_epub(input_path)
    pdf_canvas = canvas.Canvas(output_path, pagesize=letter)

    y_position = 750  # Start position
    for item in book.get_items():
        if item.get_type() == 9:  # 9 means it's text
            content = item.get_content().decode('utf-8')
            for line in content.split("\n"):
                pdf_canvas.drawString(100, y_position, line[:100])  # Trim to fit
                y_position -= 20
                if y_position < 50:
                    pdf_canvas.showPage()
                    y_position = 750

    pdf_canvas.save()

# Function to convert TXT to PDF
def convert_txt_to_pdf(input_path, output_path):
    pdf_canvas = canvas.Canvas(output_path, pagesize=letter)

    with open(input_path, "r", encoding="utf-8") as file:
        lines = file.readlines()

    y_position = 750
    for line in lines:
        pdf_canvas.drawString(100, y_position, line.strip())
        y_position -= 20
        if y_position < 50:
            pdf_canvas.showPage()
            y_position = 750

    pdf_canvas.save()

# Streamlit UI
st.title("📚 Multi-Format File Merger (4-in-1) with Page Numbers")
st.write("Upload a file, and we'll process it into a 4-in-1 formatted PDF.")

uploaded_file = st.file_uploader("Upload a file", type=["pdf", "pptx", "docx", "png", "jpg", "jpeg", "tiff", "bmp", "gif", "epub", "txt"])

if uploaded_file:
    file_extension = uploaded_file.name.split(".")[-1].lower()
    original_name = os.path.splitext(uploaded_file.name)[0]
    temp_input_path = f"{original_name}_input.{file_extension}"
    
    with open(temp_input_path, "wb") as f:
        f.write(uploaded_file.read())

    output_pdf_path = f"{original_name}_converted.pdf"
    
    # Process based on file type
    if file_extension == "pdf":
        merge_four_pages_with_numbers(temp_input_path, output_pdf_path)

    elif file_extension == "pptx":
        convert_pptx_to_pdf(temp_input_path, output_pdf_path)
        merge_four_pages_with_numbers(output_pdf_path, output_pdf_path)

    elif file_extension == "docx":
        convert_docx_to_pdf(temp_input_path, output_pdf_path)
        merge_four_pages_with_numbers(output_pdf_path, output_pdf_path)

    elif file_extension in ["png", "jpg", "jpeg", "tiff", "bmp", "gif"]:
        convert_images_to_pdf([temp_input_path], output_pdf_path)
        merge_four_pages_with_numbers(output_pdf_path, output_pdf_path)

    elif file_extension == "epub":
        convert_epub_to_pdf(temp_input_path, output_pdf_path)
        merge_four_pages_with_numbers(output_pdf_path, output_pdf_path)

    elif file_extension == "txt":
        convert_txt_to_pdf(temp_input_path, output_pdf_path)
        merge_four_pages_with_numbers(output_pdf_path, output_pdf_path)

    else:
        st.error("Unsupported file format. Please upload a supported format.")

    # Provide download button
    with open(output_pdf_path, "rb") as f:
        st.download_button(
            label="📥 Download Processed PDF",
            data=f,
            file_name=output_pdf_path,
            mime="application/pdf"
        )

    st.success(f"✅ Processing complete! Download '{output_pdf_path}' above.")
