import streamlit as st
import fitz
import os

def merge_four_pages_with_numbers(input_pdf_path, output_pdf_path):
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
st.title("PDF Page Merger (4-in-1) with Page Numbers")
st.write("Upload a PDF, and we will merge every 4 pages into one.")

uploaded_file = st.file_uploader("Upload a PDF", type=["pdf"])

if uploaded_file:
    with open("temp_input.pdf", "wb") as f:
        f.write(uploaded_file.read())

    output_pdf_path = "merged_output.pdf"
    merge_four_pages_with_numbers("temp_input.pdf", output_pdf_path)

    with open(output_pdf_path, "rb") as f:
        st.download_button(
            label="Download Processed PDF",
            data=f,
            file_name="merged_output.pdf",
            mime="application/pdf"
        )

    st.success("Processing complete! Click above to download the merged PDF.")

