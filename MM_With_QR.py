#!/usr/bin/env python
"""
Mail Merge + QR Code (DOCX only) – Streamlit App
"""

import streamlit as st
import pandas as pd
from mailmerge import MailMerge
from docx import Document
from docxcompose.composer import Composer
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import qrcode
import tempfile
import os
import zipfile
import io

# ============ CONFIG ============
REQUIRED_COL = "Property_Account_No"
QR_URL_COL = "URL"

# ============ HELPER FUNCTIONS ============

def sanitize_filename(name: str) -> str:
    """Remove invalid filename characters."""
    return (name.replace("/", "-").replace("\\", "-").replace(":", "-")
                .replace("*", "-").replace("?", "-").replace("\"", "-")
                .replace("<", "-").replace(">", "-").replace("|", "-"))

def add_qr_to_docx_bottom_right(docx_path: str, url: str, qr_temp_folder: str):
    """Add QR code image at bottom-right area of last page in DOCX."""
    try:
        # Generate QR
        qr = qrcode.make(url)
        qr_png = os.path.join(qr_temp_folder, f"qr_{os.path.basename(docx_path)}.png")
        qr.save(qr_png)

        # Open DOCX
        doc = Document(docx_path)

        # 1-cell table used as a positioning container
        table = doc.add_table(rows=1, cols=1)
        table.autofit = False
        cell = table.rows[0].cells[0]

        # Right-align paragraph inside cell
        paragraph = cell.paragraphs[0]
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = paragraph.add_run()
        run.add_picture(qr_png, width=Inches(0.9))

        # Caption under QR
        caption_para = cell.add_paragraph("Scan your custom QR code to enroll")
        caption_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in caption_para.runs:
            run.font.size = Pt(7)

        # Optional: remove visible borders (simple way)
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcBorders = OxmlElement('w:tcBorders')
        for edge in ('top', 'left', 'bottom', 'right'):
            edge_el = OxmlElement(f'w:{edge}')
            edge_el.set(qn('w:val'), 'nil')
            tcBorders.append(edge_el)
        tcPr.append(tcBorders)

        doc.save(docx_path)
        return True
    except Exception as e:
        st.warning(f"⚠️ QR in DOCX failed for {os.path.basename(docx_path)}: {e}")
        return False

# ============ STREAMLIT UI ============

st.set_page_config(page_title="Mail Merge + QR (DOCX)", page_icon="📧", layout="wide")

st.title("📧 Mail Merge + QR Code (DOCX Only)")
st.markdown("Upload your Excel data and Word template to generate personalized DOCX letters, with optional QR codes inserted into each document.")

# Sidebar
with st.sidebar:
    st.header("⚙️ Settings")
    qr_mode = st.radio(
        "QR Code Mode:",
        ["Without QR", "With QR"],
        help="If 'With QR', app uses the URL column to insert QR images into each DOCX."
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
                # Paths
                excel_path = os.path.join(tmpdir, "data.xlsx")
                template_path = os.path.join(tmpdir, "template.docx")
                output_folder = os.path.join(tmpdir, "output")
                qr_temp_folder = os.path.join(tmpdir, "qr_temp")
                os.makedirs(output_folder, exist_ok=True)
                os.makedirs(qr_temp_folder, exist_ok=True)

                # Save uploads
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
                    qr_mode = "Without QR"

                # Frequency / occurrence
                df["Account_Frequency"] = df.groupby(REQUIRED_COL)[REQUIRED_COL].transform("count")
                df["Occurrence_Number"] = df.groupby(REQUIRED_COL).cumcount() + 1

                generated_docx_list = []
                error_count = 0

                # Progress
                progress_bar = st.progress(0)
                status_text = st.empty()

                # ========== STEP 1: MAIL MERGE + QR IN DOCX ==========
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
                        # Mail merge
                        document = MailMerge(template_path)
                        merge_dict = row.to_dict()
                        document.merge(**merge_dict)
                        document.write(docx_abs)
                        document.close()

                        # Add QR into DOCX if needed
                        if qr_mode == "With QR" and has_qr_col:
                            url = row.get(QR_URL_COL, "").strip()
                            if url:
                                add_qr_to_docx_bottom_right(docx_abs, url, qr_temp_folder)

                        generated_docx_list.append(docx_abs)

                    except Exception as e:
                        error_count += 1
                        st.warning(f"⚠️ Error for {account}: {str(e)[:120]}")

                    progress = (index + 1) / len(df) * 0.8  # 0–80%
                    progress_bar.progress(progress)
                    status_text.text(f"Processing {index + 1}/{len(df)}...")

                st.success(f"✅ Generated {len(generated_docx_list)} DOCX files ({error_count} errors)")

                # ========== STEP 2: CREATE COMBINED DOCX ==========
                status_text.text("📦 Creating combined DOCX...")
                if len(generated_docx_list) > 0:
                    master_doc = Document(generated_docx_list[0])
                    composer = Composer(master_doc)

                    for doc_path in generated_docx_list[1:]:
                        if os.path.exists(doc_path):
                            try:
                                doc_to_append = Document(doc_path)
                                master_doc.add_page_break()
                                composer.append(doc_to_append)
                            except Exception:
                                pass

                    master_docx_path = os.path.join(output_folder, "All_Mailouts_Combined.docx")
                    composer.save(master_docx_path)
                    generated_docx_list.append(master_docx_path)

                progress_bar.progress(1.0)
                status_text.empty()
                progress_bar.empty()

                # ========== STEP 3: CREATE ZIP ==========
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for path in generated_docx_list:
                        zf.write(path, arcname=os.path.basename(path))
                zip_buffer.seek(0)

                # ========== DOWNLOAD SECTION ==========
                st.markdown("---")
                st.subheader("📥 Download Results")

                col1, col2 = st.columns(2)

                with col1:
                    st.download_button(
                        label="📦 Download ZIP (All DOCX Files)",
                        data=zip_buffer,
                        file_name="Mailouts_DOCX.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

                with col2:
                    if len(generated_docx_list) > 0:
                        with open(master_docx_path, "rb") as f:
                            st.download_button(
                                label="📄 Download Combined DOCX",
                                data=f,
                                file_name="All_Mailouts_Combined.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True
                            )

                st.info(f"**Files generated:** {len(generated_docx_list)-1} individual + 1 combined DOCX")

        except Exception as e:
            st.error(f"❌ Error during processing: {e}")
            import traceback
            st.code(traceback.format_exc())

# Footer
st.markdown("---")
st.caption("Built with Streamlit • DOCX output with optional in-document QR codes")
