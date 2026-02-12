#!/usr/bin/env python
"""
Mail Merge + QR Code + PDF Generator (Streamlit App)
"""

import streamlit as st
import pandas as pd
from mailmerge import MailMerge
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from docx import Document
from docxcompose.composer import Composer
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import qrcode
import tempfile
import os
import zipfile
import io
import subprocess

# ============ CONFIG ============
REQUIRED_COL = "Property_Account_No"
QR_URL_COL = "URL"

# ============ HELPER FUNCTIONS ============

def sanitize_filename(name: str) -> str:
    """Remove invalid filename characters."""
    return (name.replace("/", "-").replace("\\", "-").replace(":", "-")
                .replace("*", "-").replace("?", "-").replace("\"", "-")
                .replace("<", "-").replace(">", "-").replace("|", "-"))

def convert_docx_to_pdf(docx_path: str, output_folder: str) -> str:
    """Convert DOCX to PDF using LibreOffice."""
    try:
        subprocess.run([
            'soffice',
            '--headless',
            '--convert-to', 'pdf',
            '--outdir', output_folder,
            docx_path
        ], check=True, capture_output=True, timeout=30)
        
        pdf_path = docx_path.replace('.docx', '.pdf')
        if os.path.exists(pdf_path):
            return pdf_path
        return None
    except Exception as e:
        st.warning(f"⚠️ PDF conversion failed for {os.path.basename(docx_path)}: {e}")
        return None

def add_qr_to_pdf(in_pdf_path: str, out_pdf_path: str, url: str, qr_temp_folder: str):
    """Add QR code to first page of PDF."""
    try:
        reader = PdfReader(in_pdf_path)
        writer = PdfWriter()

        # Create QR image
        qr = qrcode.make(url)
        qr_png = os.path.join(qr_temp_folder, f"qr_{os.path.basename(in_pdf_path)}.png")
        qr.save(qr_png)

        # Create overlay
        overlay_pdf = os.path.join(qr_temp_folder, f"overlay_{os.path.basename(in_pdf_path)}")
        c = canvas.Canvas(overlay_pdf, pagesize=letter)

        page_width = float(reader.pages[0].mediabox.width)
        qr_size = 70
        x = page_width - qr_size - 40
        y = 78

        c.drawImage(qr_png, x, y, width=qr_size, height=qr_size)
        c.setFont("Helvetica", 7)
        c.drawCentredString(x + qr_size / 2, y - 3, "Scan your custom QR code to enroll")
        c.save()

        overlay_page = PdfReader(overlay_pdf).pages[0]

        # Merge overlay onto first page
        base_page = reader.pages[0]
        base_page.merge_page(overlay_page)
        writer.add_page(base_page)

        # Remaining pages unchanged
        for p in range(1, len(reader.pages)):
            writer.add_page(reader.pages[p])

        with open(out_pdf_path, "wb") as f_out:
            writer.write(f_out)

        return True
    except Exception as e:
        st.warning(f"⚠️ QR failed for {os.path.basename(in_pdf_path)}: {e}")
        return False

# ============ STREAMLIT UI ============

st.set_page_config(page_title="Mail Merge + QR Generator", page_icon="📧", layout="wide")

st.title("📧 Mail Merge + QR Code + PDF Generator")
st.markdown("Upload your Excel data and Word template to generate personalized mailouts with optional QR codes.")

# Sidebar
with st.sidebar:
    st.header("⚙️ Settings")
    qr_mode = st.radio(
        "QR Code Mode:",
        ["Without QR", "With QR"],
        help="Choose whether to add QR codes to PDFs (requires 'URL' column in Excel)"
    )
    
    output_format = st.radio(
        "Output Format:",
        ["DOCX Only", "PDF Only", "Both DOCX and PDF"],
        help="Choose output file format"
    )
    
    st.info(f"**Required Excel column:** `{REQUIRED_COL}`")
    if qr_mode == "With QR":
        st.info(f"**QR URL column:** `{QR_URL_COL}`")

# File uploads
col1, col2 = st.columns(2)

with col1:
    uploaded_excel = st.file_uploader(
        "📊 Upload Excel File (.xlsx)",
        type=["xlsx"],
        help=f"Excel must contain '{REQUIRED_COL}' column"
    )

with col2:
    uploaded_template = st.file_uploader(
        "📄 Upload Word Template (.docx)",
        type=["docx"],
        help="Word template with merge fields matching Excel columns"
    )

# Main process button
if st.button("🚀 Run Mail Merge", type="primary", use_container_width=True):
    if not uploaded_excel or not uploaded_template:
        st.error("❌ Please upload both Excel and Word template files.")
        st.stop()

    with st.spinner("Processing mail merge..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                # Save uploads
                excel_path = os.path.join(tmpdir, "data.xlsx")
                template_path = os.path.join(tmpdir, "template.docx")
                output_folder = os.path.join(tmpdir, "output")
                qr_temp_folder = os.path.join(tmpdir, "qr_temp")
                os.makedirs(output_folder, exist_ok=True)
                os.makedirs(qr_temp_folder, exist_ok=True)

                with open(excel_path, "wb") as f:
                    f.write(uploaded_excel.read())
                with open(template_path, "wb") as f:
                    f.write(uploaded_template.read())

                # Read Excel
                df = pd.read_excel(excel_path).fillna("")
                df.columns = [c.strip() for c in df.columns]

                for col in df.columns:
                    if df[col].dtype != "object":
                        df[col] = df[col].astype(str)

                # Validate required column
                if REQUIRED_COL not in df.columns:
                    st.error(f"❌ Excel file missing required column: `{REQUIRED_COL}`")
                    st.stop()

                # Check QR column
                has_qr_col = QR_URL_COL in df.columns
                if qr_mode == "With QR" and not has_qr_col:
                    st.warning(f"⚠️ Column '{QR_URL_COL}' not found. QR codes will be skipped.")

                # Process frequency/occurrence
                df["Account_Frequency"] = df.groupby(REQUIRED_COL)[REQUIRED_COL].transform("count")
                df["Occurrence_Number"] = df.groupby(REQUIRED_COL).cumcount() + 1

                generated_docx_list = []
                generated_pdf_list = []
                error_count = 0

                # Progress bar
                progress_bar = st.progress(0)
                status_text = st.empty()

                # ========== STEP 1: MAIL MERGE (DOCX) ==========
                status_text.text("📝 Generating DOCX files...")
                for index, row in df.iterrows():
                    account = str(row.get(REQUIRED_COL, "Unknown")).strip()
                    if not account or account.lower() == "nan":
                        continue

                    county = str(row.get("Property County", "Unknown")).strip().upper()
                    occurrence = int(row["Occurrence_Number"])
                    frequency = int(row["Account_Frequency"])

                    if frequency > 1 and occurrence > 1:
                        base_name = f"{account}_{county}_Mailout"
                    else:
                        base_name = f"{account}_Mailout"

                    base_name = sanitize_filename(base_name)
                    docx_abs = os.path.join(output_folder, f"{base_name}.docx")

                    try:
                        document = MailMerge(template_path)
                        merge_dict = row.to_dict()
                        document.merge(**merge_dict)
                        document.write(docx_abs)
                        document.close()
                        generated_docx_list.append(docx_abs)
                    except Exception as e:
                        error_count += 1
                        st.warning(f"⚠️ Error for {account}: {str(e)[:100]}")

                    progress = (index + 1) / len(df) * 0.3  # 30% for DOCX
                    progress_bar.progress(progress)

                st.success(f"✅ Generated {len(generated_docx_list)} DOCX files ({error_count} errors)")

                # ========== STEP 2: CONVERT TO PDF ==========
                if output_format in ["PDF Only", "Both DOCX and PDF"]:
                    status_text.text("📄 Converting DOCX to PDF...")
                    for idx, docx_path in enumerate(generated_docx_list):
                        pdf_path = convert_docx_to_pdf(docx_path, output_folder)
                        if pdf_path:
                            generated_pdf_list.append(pdf_path)
                        
                        progress = 0.3 + ((idx + 1) / len(generated_docx_list)) * 0.4  # 30-70%
                        progress_bar.progress(progress)
                    
                    st.success(f"✅ Converted {len(generated_pdf_list)} PDFs")

                # ========== STEP 3: ADD QR CODES ==========
                if qr_mode == "With QR" and has_qr_col and len(generated_pdf_list) > 0:
                    status_text.text("🔲 Adding QR codes to PDFs...")
                    qr_count = 0
                    
                    for idx, row in df.iterrows():
                        account = str(row.get(REQUIRED_COL, "Unknown")).strip()
                        if not account or account.lower() == "nan":
                            continue

                        county = str(row.get("Property County", "Unknown")).strip().upper()
                        occurrence = int(row["Occurrence_Number"])
                        frequency = int(row["Account_Frequency"])

                        if frequency > 1 and occurrence > 1:
                            base_name = f"{account}_{county}_Mailout"
                        else:
                            base_name = f"{account}_Mailout"

                        base_name = sanitize_filename(base_name)
                        pdf_path = os.path.join(output_folder, f"{base_name}.pdf")

                        url = row.get(QR_URL_COL, "").strip() if QR_URL_COL in row.index else ""
                        if url and os.path.exists(pdf_path):
                            if add_qr_to_pdf(pdf_path, pdf_path, url, qr_temp_folder):
                                qr_count += 1
                        
                        progress = 0.7 + ((idx + 1) / len(df)) * 0.2  # 70-90%
                        progress_bar.progress(progress)
                    
                    st.success(f"✅ Added QR codes to {qr_count} PDFs")

                # ========== STEP 4: CREATE COMBINED FILES ==========
                status_text.text("📦 Creating combined files...")
                
                # Combined DOCX
                if output_format in ["DOCX Only", "Both DOCX and PDF"] and len(generated_docx_list) > 0:
                    master_doc = Document(generated_docx_list[0])
                    composer = Composer(master_doc)

                    for doc_path in generated_docx_list[1:]:
                        if os.path.exists(doc_path):
                            try:
                                doc_to_append = Document(doc_path)
                                master_doc.add_page_break()
                                composer.append(doc_to_append)
                            except:
                                pass

                    master_docx_path = os.path.join(output_folder, "All_Mailouts_Combined.docx")
                    composer.save(master_docx_path)
                    generated_docx_list.append(master_docx_path)

                # Combined PDF
                if output_format in ["PDF Only", "Both DOCX and PDF"] and len(generated_pdf_list) > 0:
                    merger = PdfMerger()
                    for pdf_file in generated_pdf_list:
                        if os.path.exists(pdf_file):
                            try:
                                merger.append(pdf_file)
                            except Exception as e:
                                st.warning(f"Could not add {os.path.basename(pdf_file)}: {e}")
                    
                    master_pdf_path = os.path.join(output_folder, "All_Mailouts_Combined.pdf")
                    with open(master_pdf_path, "wb") as f:
                        merger.write(f)
                    merger.close()
                    generated_pdf_list.append(master_pdf_path)

                progress_bar.progress(1.0)
                status_text.empty()
                progress_bar.empty()

                # ========== STEP 5: CREATE ZIP ==========
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    if output_format in ["DOCX Only", "Both DOCX and PDF"]:
                        for path in generated_docx_list:
                            zf.write(path, arcname=os.path.basename(path))
                    
                    if output_format in ["PDF Only", "Both DOCX and PDF"]:
                        for path in generated_pdf_list:
                            zf.write(path, arcname=os.path.basename(path))

                zip_buffer.seek(0)

                # ========== DOWNLOAD SECTION ==========
                st.markdown("---")
                st.subheader("📥 Download Results")

                col1, col2, col3 = st.columns(3)
                
                with col1:
                    st.download_button(
                        label="📦 Download ZIP (All Files)",
                        data=zip_buffer,
                        file_name="Mailouts.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

                with col2:
                    if output_format in ["DOCX Only", "Both DOCX and PDF"]:
                        with open(master_docx_path, "rb") as f:
                            st.download_button(
                                label="📄 Combined DOCX",
                                data=f,
                                file_name="All_Mailouts_Combined.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )

                with col3:
                    if output_format in ["PDF Only", "Both DOCX and PDF"]:
                        with open(master_pdf_path, "rb") as f:
                            st.download_button(
                                label="📕 Combined PDF",
                                data=f,
                                file_name="All_Mailouts_Combined.pdf",
                                mime="application/pdf",
                                use_container_width=True
                            )

                # Stats
                st.info(f"**Generated:** {len(generated_docx_list)-1 if output_format != 'PDF Only' else 0} DOCX + {len(generated_pdf_list)-1 if output_format != 'DOCX Only' else 0} PDFs + Combined files")

        except Exception as e:
            st.error(f"❌ Error during processing: {e}")
            import traceback
            st.code(traceback.format_exc())

# Footer
st.markdown("---")
st.caption("Built with Streamlit • Uses LibreOffice for PDF conversion")
