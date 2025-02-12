import streamlit as st
import os
from pptx import Presentation
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from io import BytesIO
from PIL import Image

# Function to merge every 4 pages into one with numbers
def merge_four_pages_with_numbers(input_pdf_path, output_pdf_path):
    import fitz  # PyMuPDF for PDFs
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

# Function to convert PPTX slides to PDF with images
def convert_pptx_to_pdf(input_path, output_path):
    presentation = Presentation(input_path)
    pdf_canvas = canvas.Canvas(output_path, pagesize=letter)
    
    slide_width = presentation.slide_width
    slide_height = presentation.slide_height

    # Convert slide width/height from PowerPoint's EMU to PDF's points
    slide_width = slide_width / 12700  # 1 point = 1/72 inch, PowerPoint uses EMU (English Metric Units)
    slide_height = slide_height / 12700

    # Loop through each slide in the presentation
    for slide_number, slide in enumerate(presentation.slides):
        pdf_canvas.setPageSize((slide_width, slide_height))
        pdf_canvas.showPage()

        # Extract and draw text from slide shapes
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                text = shape.text.strip()
                # You can adjust the positioning based on your needs (e.g., prevent overlapping)
                pdf_canvas.drawString(100, slide_height - 100, text)

            # Extract and draw images from slide shapes
            if hasattr(shape, "image") and shape.image:
                image_stream = BytesIO(shape.image.blob)
                img = Image.open(image_stream)
                # Save the image as a temporary file to use with reportlab
                img_path = "/tmp/temp_image.png"
                img.save(img_path)
                pdf_canvas.drawImage(img_path, 50, slide_height - 300, width=300, height=200)

        # Move to the next page for the next slide
        pdf_canvas.showPage()

    pdf_canvas.save()

# Streamlit UI
st.title("📚 Multi-Format File Merger (4-in-1) with Page Numbers")
st.write("Upload a file, and we'll process it into a 4-in-1 formatted PDF.")

uploaded_file = st.file_uploader("Upload a file", type=["pptx"])

if uploaded_file:
    file_extension = uploaded_file.name.split(".")[-1].lower()
    original_name = os.path.splitext(uploaded_file.name)[0]
    temp_input_path = f"{original_name}_input.{file_extension}"
    
    with open(temp_input_path, "wb") as f:
        f.write(uploaded_file.read())

    output_pdf_path = f"{original_name}_converted.pdf"
    
    # Process PPTX
    if file_extension == "pptx":
        convert_pptx_to_pdf(temp_input_path, output_pdf_path)
        merge_four_pages_with_numbers(output_pdf_path, output_pdf_path)

    else:
        st.error("Unsupported file format. Please upload a PPTX file.")

    # Provide download button
    with open(output_pdf_path, "rb") as f:
        st.download_button(
            label="📥 Download Processed PDF",
            data=f,
            file_name=output_pdf_path,
            mime="application/pdf"
        )

    st.success(f"✅ Processing complete! Download '{output_pdf_path}' above.")
