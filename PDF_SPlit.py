import streamlit as st
from PyPDF2 import PdfReader, PdfWriter
import io
import zipfile

st.title("📄 PDF Splitter (Page by Page)")

# 1. Soft Code: Ask user to upload instead of hardcoded path
uploaded_file = st.file_uploader("Upload a PDF to split", type="pdf")

if uploaded_file is not None:
    # Read PDF from memory
    reader = PdfReader(uploaded_file)
    total_pages = len(reader.pages)
    
    st.info(f"PDF Loaded. Total Pages: {total_pages}")
    
    if st.button("Split All Pages"):
        # Create a progress bar
        progress_bar = st.progress(0)
        
        # Create an in-memory ZIP file to hold all the PDFs
        zip_buffer = io.BytesIO()
        
        with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
            for i, page in enumerate(reader.pages):
                # Create a new PDF writer for this single page
                writer = PdfWriter()
                writer.add_page(page)
                
                # Write the single page PDF to a memory buffer
                single_page_pdf = io.BytesIO()
                writer.write(single_page_pdf)
                
                # Define the filename (e.g., 0001.pdf, 0002.pdf)
                filename = f"{i+1:04d}.pdf"
                
                # Add this PDF to the ZIP file
                zip_file.writestr(filename, single_page_pdf.getvalue())
                
                # Update progress bar
                progress_bar.progress((i + 1) / total_pages)

        # Move pointer to the beginning of the ZIP file so it can be downloaded
        zip_buffer.seek(0)
        
        st.success("✅ Done! Files are ready.")
        
        # 2. Soft Code: Download button instead of saving to local folder
        st.download_button(
            label="Download All Split Pages (.zip)",
            data=zip_buffer,
            file_name="split_pages.zip",
            mime="application/zip"
        )
