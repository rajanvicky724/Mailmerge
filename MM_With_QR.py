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

def add_qr_simple_bottom_right(docx_path: str, url: str, qr_temp_folder: str):
    """
    Add QR code near bottom-right by using a 1x2 table:
    [empty][QR+caption].
    This avoids complex XML and works well on Streamlit Cloud.
    """
    try:
        # Generate QR image
        qr = qrcode.make(url)
        qr_png = os.path.join(qr_temp_folder, f"qr_{os.path.basename(docx_path)}.png")
        qr.save(qr_png)

        doc = Document(docx_path)

        # Add some spacing before QR block
        doc.add_paragraph()

        # Two-cell table: left empty, right has QR
        table = doc.add_table(rows=1, cols=2)
        table.autofit = True

        left_cell = table.rows[0].cells[0]
        right_cell = table.rows[0].cells[1]

        # Right cell content
        p = right_cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = p.add_run()
        run.add_picture(qr_png, width=Inches(0.9))

        cap_para = right_cell.add_paragraph("Scan your custom QR code to enroll")
        cap_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for r in cap_para.runs:
            r.font.size = Pt(7)

        doc.save(docx_path)
        return True
    except Exception as e:
        st.warning(f"⚠️ QR insertion failed for {os.path.basename(docx_path)}: {e}")
        return False

# ============ STREAMLIT UI ============

st.set_page_config(page_title="Mail Merge + QR (DOCX)", page_icon="📧", layout="wide")

st.title("📧 Mail Merge + QR Code (DOCX Only)")
st.markdown(
    "Upload your Excel data and Word template to generate personalized DOCX letters, "
    "with optional QR codes added near the bottom-right."
)

with st.sidebar:
    st.header("⚙️ Settings")
    qr_mode = st.radio(
        "QR Code Mode:",
        ["Without QR", "With QR"],
        help="If 'With QR', the app uses the URL column to insert QR images into each DOCX."
    )

    st.info(f"**Required Excel column:** `{REQUIRED_COL}`")
    if qr_mode == "With QR":
        st.info(f"**QR URL column:** `{QR_URL_COL}`")

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

if st.button("🚀 Run Mail Merge", type="primary", use_container_width=True):
    if not uploaded_excel or not uploaded_template:
        st.error("❌ Please upload both Excel and Word template files.")
        st.stop()

    with st.spinner("Processing mail merge..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
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

                # Check required column
                if REQUIRED_COL not in df.columns:
                    st.error(f"❌ Excel file missing required column: `{REQUIRED_COL}`")
                    st.stop()

                has_qr_col = QR_URL_COL in df.columns
                if qr_mode == "With QR" and not has_qr_col:
                    st.warning(f"⚠️ Column '{QR_URL_COL}' not found. QR codes will be skipped.")
                    qr_mode = "Without QR"

                # Frequency / occurrence
                df["Account_Frequency"] = df.groupby(REQUIRED_COL)[REQUIRED_COL].transform("count")
                df["Occurrence_Number"] = df.groupby(REQUIRED_COL).cumcount() + 1

                generated_docx_list = []
                error_count = 0

                progress_bar = st.progress(0)
                status_text = st.empty()

                # MAIL MERGE + QR
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
                    docx_abs = os.path.join(output_folder, f"{base_name}.docx")

                    try:
                        doc = MailMerge(template_path)
                        merge_dict = row.to_dict()
                        doc.merge(**merge_dict)
                        doc.write(docx_abs)
                        doc.close()

                        if qr_mode == "With QR" and has_qr_col:
                            url = row.get(QR_URL_COL, "").strip()
                            if url:
                                add_qr_simple_bottom_right(docx_abs, url, qr_temp_folder)

                        generated_docx_list.append(docx_abs)

                    except Exception as e:
                        error_count += 1
                        st.warning(f"⚠️ Error for {account}: {str(e)[:120]}")

                    progress = (idx + 1) / len(df) * 0.8
                    progress_bar.progress(progress)
                    status_text.text(f"Processing {idx + 1}/{len(df)}...")

                st.success(f"✅ Generated {len(generated_docx_list)} DOCX files ({error_count} errors)")

                # COMBINED DOCX
                status_text.text("📦 Creating combined DOCX...")
                master_docx_path = None
                if generated_docx_list:
                    master_doc = Document(generated_docx_list[0])
                    composer = Composer(master_doc)
                    for p in generated_docx_list[1:]:
                        if os.path.exists(p):
                            try:
                                d = Document(p)
                                master_doc.add_page_break()
                                composer.append(d)
                            except Exception:
                                pass
                    master_docx_path = os.path.join(output_folder, "All_Mailouts_Combined.docx")
                    composer.save(master_docx_path)

                progress_bar.progress(1.0)
                status_text.empty()
                progress_bar.empty()

                # ZIP
                zip_buffer = io.BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                    for p in generated_docx_list:
                        zf.write(p, arcname=os.path.basename(p))
                    if master_docx_path:
                        zf.write(master_docx_path, arcname=os.path.basename(master_docx_path))
                zip_buffer.seek(0)

                st.markdown("---")
                st.subheader("📥 Download Results")

                col_a, col_b = st.columns(2)

                with col_a:
                    st.download_button(
                        label="📦 Download ZIP (All DOCX Files)",
                        data=zip_buffer,
                        file_name="Mailouts_DOCX.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )

                with col_b:
                    if master_docx_path and os.path.exists(master_docx_path):
                        with open(master_docx_path, "rb") as f:
                            st.download_button(
                                label="📄 Download Combined DOCX",
                                data=f,
                                file_name="All_Mailouts_Combined.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                            )

                st.info(f"**Files generated:** {len(generated_docx_list)} individual DOCX"
                        + (", plus combined DOCX" if master_docx_path else ""))

        except Exception as e:
            st.error(f"❌ Error during processing: {e}")

st.markdown("---")
st.caption("Built with Streamlit • DOCX output with optional QR codes")
