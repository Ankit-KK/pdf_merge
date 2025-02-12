import subprocess
import streamlit as st
import os

# Function to convert PPTX to PDF using LibreOffice
def convert_pptx_to_pdf(input_path, output_path):
    # Using LibreOffice in headless mode (no GUI)
    subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', input_path], check=True)
    # Move the generated file to the desired location
    os.rename(input_path.replace('.pptx', '.pdf'), output_path)

# Streamlit UI
st.title("📚 PowerPoint to PDF Converter")
st.write("Upload a PowerPoint file, and we'll convert it to PDF.")

uploaded_file = st.file_uploader("Upload a PPTX file", type=["pptx"])

if uploaded_file:
    original_name = os.path.splitext(uploaded_file.name)[0]
    temp_input_path = f"{original_name}_input.pptx"
    
    with open(temp_input_path, "wb") as f:
        f.write(uploaded_file.read())

    output_pdf_path = f"{original_name}_converted.pdf"

    # Convert PPTX to PDF
    convert_pptx_to_pdf(temp_input_path, output_pdf_path)

    # Provide download button
    with open(output_pdf_path, "rb") as f:
        st.download_button(
            label="📥 Download Processed PDF",
            data=f,
            file_name=output_pdf_path,
            mime="application/pdf"
        )

    st.success(f"✅ Processing complete! Download '{output_pdf_path}' above.")
