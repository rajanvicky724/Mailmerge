#!/usr/bin/env python
"""
Mail Merge + QR Code (DOCX, X/Y positioning) – Streamlit App
"""

import streamlit as st
import pandas as pd
from mailmerge import MailMerge
from docx import Document
from docxcompose.composer import Composer
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import qrcode
import tempfile
import os
import zipfile
import io

# ============ CONFIG ============
REQUIRED_COL = "Property_Account_No"
QR_URL_COL = "URL"
EMU_PER_INCH = 914400

# ============ HELPER FUNCTIONS ============

def sanitize_filename(name: str) -> str:
    return (name.replace("/", "-").replace("\\", "-").replace(":", "-")
                .replace("*", "-").replace("?", "-").replace("\"", "-")
                .replace("<", "-").replace(">", "-").replace("|", "-"))

from io import BytesIO
from docx.opc.constants import RELATIONSHIP_TYPE as RT

def _new_anchor(run, image_path, width_inches, height_inches, pos_x_inches, pos_y_inches):
    """Create wp:anchor element for floating image at absolute page coords."""
    part = run.part

    # Read image and wrap in BytesIO so python-docx can seek()
    with open(image_path, "rb") as f:
        image_bytes = f.read()
    image_stream = BytesIO(image_bytes)

    # Add image part correctly
    image_part = part.package.image_parts.get_or_add_image_part(image_stream)
    rId = part.relate_to(image_part, RT.IMAGE)

    cx = int(width_inches * EMU_PER_INCH)
    cy = int(height_inches * EMU_PER_INCH)
    pos_x = int(pos_x_inches * EMU_PER_INCH)
    pos_y = int(pos_y_inches * EMU_PER_INCH)

    anchor_xml = f"""
    <wp:anchor xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
               xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
               xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
               distT="0" distB="0" distL="0" distR="0"
               simplePos="0" relativeHeight="251658240"
               behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0" />
      <wp:positionH relativeFrom="page">
        <wp:posOffset>{pos_x}</wp:posOffset>
      </wp:positionH>
      <wp:positionV relativeFrom="page">
        <wp:posOffset>{pos_y}</wp:posOffset>
      </wp:positionV>
      <wp:extent cx="{cx}" cy="{cy}" />
      <wp:effectExtent l="0" t="0" r="0" b="0" />
      <wp:wrapSquare wrapText="bothSides" />
      <wp:docPr id="1" name="QR Code"/>
      <wp:cNvGraphicFramePr/>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic>
            <pic:nvPicPr>
              <pic:cNvPr id="0" name="QR Code"/>
              <pic:cNvPicPr/>
            </pic:nvPicPr>
            <pic:blipFill>
              <a:blip r:embed="{rId}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>
              <a:stretch><a:fillRect/></a:stretch>
            </pic:blipFill>
            <pic:spPr>
              <a:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="{cx}" cy="{cy}"/>
              </a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
    """
    return parse_xml(anchor_xml)

def add_qr_xy_to_docx(docx_path: str, url: str, qr_temp_folder: str,
                      x_inches: float, y_inches: float,
                      qr_size_inches: float):
    """Add QR as floating image at absolute (x_inches, y_inches) from top-left of page."""
    try:
        qr = qrcode.make(url)
        qr_png = os.path.join(qr_temp_folder, f"qr_{os.path.basename(docx_path)}.png")
        qr.save(qr_png)

        doc = Document(docx_path)

        # Use the last existing paragraph (still on page 1 for your template)
        # so we don't create an extra page.
 if doc.paragraphs:
   p = doc.paragraphs[-1]
 else:
   p = doc.add_paragraph()

run = p.add_run()

        anchor = _new_anchor(
            run,
            qr_png,
            width_inches=qr_size_inches,
            height_inches=qr_size_inches,
            pos_x_inches=x_inches,
            pos_y_inches=y_inches,
        )

        drawing = OxmlElement("w:drawing")
        drawing.append(anchor)
        run._r.append(drawing)

        doc.save(docx_path)
        return True
    except Exception as e:
        st.warning(f"⚠️ QR XY failed for {os.path.basename(docx_path)}: {e}")
        return False

# ============ STREAMLIT UI ============

st.set_page_config(page_title="stampaunioneqr – Mail Merge + QR", page_icon="📧", layout="wide")

st.title("📧 stampaunioneqr – Mail Merge + QR (DOCX, X/Y)")

with st.sidebar:
    st.header("⚙️ Settings")
    qr_mode = st.radio(
        "QR Code Mode:",
        ["Without QR", "With QR"],
        help="If 'With QR', uses the URL column to insert QR into each DOCX."
    )

    x_pos = st.number_input(
        "QR X position (inches from left)",
        min_value=0.0, max_value=8.5, value=6.5, step=0.1
    )
    y_pos = st.number_input(
        "QR Y position (inches from top)",
        min_value=0.0, max_value=11.0, value=9.5, step=0.1
    )
    qr_size = st.number_input(
        "QR size (inches)",
        min_value=0.3, max_value=2.0, value=0.9, step=0.1
    )

    st.info(f"Required Excel column: `{REQUIRED_COL}`")
    if qr_mode == "With QR":
        st.info(f"QR URL column: `{QR_URL_COL}`")

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

    with st.spinner("Processing..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
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

                df = pd.read_excel(excel_path).fillna("")
                df.columns = [c.strip() for c in df.columns]
                for col in df.columns:
                    if df[col].dtype != "object":
                        df[col] = df[col].astype(str)

                if REQUIRED_COL not in df.columns:
                    st.error(f"❌ Excel file missing required column: `{REQUIRED_COL}`")
                    st.stop()

                has_qr_col = QR_URL_COL in df.columns
                if qr_mode == "With QR" and not has_qr_col:
                    st.warning(f"⚠️ Column '{QR_URL_COL}' not found. QR codes will be skipped.")
                    qr_mode = "Without QR"

                df["Account_Frequency"] = df.groupby(REQUIRED_COL)[REQUIRED_COL].transform("count")
                df["Occurrence_Number"] = df.groupby(REQUIRED_COL).cumcount() + 1

                generated_docx_list = []
                progress_bar = st.progress(0)
                status_text = st.empty()
                error_count = 0

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
                                add_qr_xy_to_docx(
                                    docx_abs, url, qr_temp_folder,
                                    x_inches=x_pos, y_inches=y_pos,
                                    qr_size_inches=qr_size
                                )

                        generated_docx_list.append(docx_abs)
                    except Exception as e:
                        error_count += 1
                        st.warning(f"⚠️ Error for {account}: {str(e)[:120]}")

                    progress = (idx + 1) / len(df) * 0.8
                    progress_bar.progress(progress)
                    status_text.text(f"Processing {idx + 1}/{len(df)}...")

                # Combine DOCX
                status_text.text("📦 Creating combined DOCX...")
                master_docx_path = None
                if generated_docx_list:
                    master_doc = Document(generated_docx_list[0])
                    composer = Composer(master_doc)
                    for pth in generated_docx_list[1:]:
                        if os.path.exists(pth):
                            try:
                                d = Document(pth)
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
                    for pth in generated_docx_list:
                        zf.write(pth, arcname=os.path.basename(pth))
                    if master_docx_path:
                        zf.write(master_docx_path, arcname=os.path.basename(master_docx_path))
                zip_buffer.seek(0)

                st.markdown("---")
                st.subheader("📥 Download Results")

                col_a, col_b = st.columns(2)
                with col_a:
                    st.download_button(
                        "📦 Download ZIP (All DOCX)",
                        data=zip_buffer,
                        file_name="Mailouts_DOCX.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )
                with col_b:
                    if master_docx_path and os.path.exists(master_docx_path):
                        with open(master_docx_path, "rb") as f:
                            st.download_button(
                                "📄 Download Combined DOCX",
                                data=f,
                                file_name="All_Mailouts_Combined.docx",
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                use_container_width=True,
                            )

                st.info(f"Done. Generated {len(generated_docx_list)} DOCX files"
                        + (f" with {error_count} errors." if error_count else "."))

        except Exception as e:
            st.error(f"❌ Error during processing: {e}")

st.markdown("---")
st.caption("stampaunioneqr – DOCX mail merge with X/Y QR positioning")




