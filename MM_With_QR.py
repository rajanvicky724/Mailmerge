#!/usr/bin/env python
"""
Mail Merge + QR Code + PDF Generator
- Generates DOCX via mail merge
- Converts all DOCX to PDF
- Adds record-specific QR codes on each PDF (page 1)
- Merges all PDFs into a single combined PDF
"""

import pandas as pd
from mailmerge import MailMerge
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from docx import Document
from docxcompose.composer import Composer
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
import qrcode
import os
import sys
import logging

# ============ HARD-CODED PATHS ============
EXCEL_FILE = r"C:\\Users\\vigneshn\\OneDrive - O`Connor and Associates\\Desktop\\XMAX Tree\\Mailout\\Sakthi-working\\NY Bronx Apartment_Updated_Mailout Sheet.xlsx"
TEMPLATE_FILE = r"C:\\Users\\vigneshn\\OneDrive - O`Connor and Associates\\Desktop\\XMAX Tree\\Mailout\\Sakthi-working\\Apartment_NY_New York City 2026.docx"
OUTPUT_FOLDER = r"C:\\Users\\vigneshn\\OneDrive - O`Connor and Associates\\Desktop\\XMAX Tree\\Mailout\\Sakthi-working\\Test_Result"
QR_TEMP_FOLDER = os.path.join(OUTPUT_FOLDER, "qr_temp")
QR_URL_COL = "URL"                                        # column in Excel with QR URL

# ============ LOGGING SETUP ============
LOG_FILE = os.path.join(OUTPUT_FOLDER, "mailmerge_log.txt")
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(QR_TEMP_FOLDER, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ============ VALIDATION ============
if not os.path.exists(EXCEL_FILE):
    logger.error(f"Excel file not found: {EXCEL_FILE}")
    sys.exit(1)

if not os.path.exists(TEMPLATE_FILE):
    logger.error(f"Template file not found: {TEMPLATE_FILE}")
    sys.exit(1)

logger.info("Starting mail merge + QR process")
logger.info(f"Excel: {EXCEL_FILE}")
logger.info(f"Template: {TEMPLATE_FILE}")

# ============ 1. READ EXCEL ============
logger.info("Reading Excel file...")
try:
    df = pd.read_excel(EXCEL_FILE)
    df = df.fillna("")
    df.columns = [c.strip() for c in df.columns]

    for col in df.columns:
        if df[col].dtype != "object":
            df[col] = df[col].astype(str)

    logger.info(f"Loaded {len(df)} records")
except Exception as e:
    logger.error(f"Error reading Excel: {e}")
    sys.exit(1)

required_col = "Property_Account_No"
if required_col not in df.columns:
    logger.error(f"Column '{required_col}' not found")
    sys.exit(1)

if QR_URL_COL not in df.columns:
    logger.warning(f"Column '{QR_URL_COL}' not found – QR codes will not be created")

df["Account_Frequency"] = df.groupby(required_col)[required_col].transform("count")
df["Occurrence_Number"] = df.groupby(required_col).cumcount() + 1

# ============ 2. GENERATE DOCX (MAIL MERGE) ============
logger.info("Generating DOCX files...")

generated_docx_list = []
generated_pdf_list = []
error_count = 0

for index, row in df.iterrows():
    account = str(row.get(required_col, "Unknown")).strip()
    if not account or account.lower() == "nan":
        continue

    county = str(row.get("Property County", "Unknown")).strip().upper()
    occurrence = int(row["Occurrence_Number"])
    frequency = int(row["Account_Frequency"])

    if frequency > 1 and occurrence > 1:
        base_name = f"{account}_{county}_Mailout"
    else:
        base_name = f"{account}_Mailout"

    base_name = (base_name.replace("/", "-").replace("\\", "-").replace(":", "-")
                 .replace("*", "-").replace("?", "-").replace("\"", "-")
                 .replace("<", "-").replace(">", "-").replace("|", "-"))

    docx_abs = os.path.join(OUTPUT_FOLDER, f"{base_name}.docx")
    pdf_abs = os.path.join(OUTPUT_FOLDER, f"{base_name}.pdf")

    try:
        document = MailMerge(TEMPLATE_FILE)
        merge_dict = row.to_dict()
        document.merge(**merge_dict)
        document.write(docx_abs)
        document.close()

        generated_docx_list.append(docx_abs)
        generated_pdf_list.append(pdf_abs)

    except Exception as e:
        error_count += 1
        logger.warning(f"Error for {account}: {str(e)[:200]}")

logger.info(f"Generated {len(generated_docx_list)} DOCX files ({error_count} errors)")

# ============ 3. CONVERT DOCX TO PDF ============
logger.info("Converting DOCX to PDF...")
try:
    from docx2pdf import convert
    convert(OUTPUT_FOLDER)
    logger.info("PDF conversion complete")
except Exception as e:
    logger.warning(f"PDF conversion warning: {str(e)[:200]}")

# ============ 4. ADD QR CODES TO EACH PDF ============
def add_qr_to_pdf(in_pdf_path: str, out_pdf_path: str, url: str):
    """
    Add a QR code for `url` to the first page of `in_pdf_path`
    and save as `out_pdf_path`.
    """
    try:
        reader = PdfReader(in_pdf_path)
        writer = PdfWriter()

        # create QR image
        qr = qrcode.make(url)
        qr_png = os.path.join(QR_TEMP_FOLDER, "qr_temp.png")
        qr.save(qr_png)

        # create overlay for first page
        overlay_pdf = os.path.join(QR_TEMP_FOLDER, "overlay_temp.pdf")
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

        # merge overlay onto first page
        base_page = reader.pages[0]
        base_page.merge_page(overlay_page)
        writer.add_page(base_page)

        # remaining pages (if any) unchanged
        for p in range(1, len(reader.pages)):
            writer.add_page(reader.pages[p])

        with open(out_pdf_path, "wb") as f_out:
            writer.write(f_out)

        logger.info(f"QR inserted: {os.path.basename(out_pdf_path)}")

    except Exception as e:
        logger.warning(f"QR failed for {in_pdf_path}: {e}")

logger.info("Adding QR codes to PDFs...")

qr_pdf_list = []  # PDFs with QR applied, to merge later

for index, row in df.iterrows():
    account = str(row.get(required_col, "Unknown")).strip()
    if not account or account.lower() == "nan":
        continue

    county = str(row.get("Property County", "Unknown")).strip().upper()
    occurrence = int(row["Occurrence_Number"])
    frequency = int(row["Account_Frequency"])

    if frequency > 1 and occurrence > 1:
        base_name = f"{account}_{county}_Mailout"
    else:
        base_name = f"{account}_Mailout"

    base_name = (base_name.replace("/", "-").replace("\\", "-").replace(":", "-")
                 .replace("*", "-").replace("?", "-").replace("\"", "-")
                 .replace("<", "-").replace(">", "-").replace("|", "-"))

    original_pdf = os.path.join(OUTPUT_FOLDER, f"{base_name}.pdf")

    url = row.get(QR_URL_COL, "").strip() if QR_URL_COL in row.index else ""
    if url and os.path.exists(original_pdf):
        # write QR version back onto the same filename
        add_qr_to_pdf(original_pdf, original_pdf, url)
        qr_pdf_list.append(original_pdf)
    else:
        if os.path.exists(original_pdf):
            qr_pdf_list.append(original_pdf)
            logger.info(f"No URL/QR for {base_name}, using original PDF")

# ============ 5. MERGE WORD DOCUMENTS (DOCX) ============
logger.info("Creating master DOCX document...")
if len(generated_docx_list) > 0:
    try:
        if os.path.exists(generated_docx_list[0]):
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

            master_path = os.path.join(OUTPUT_FOLDER, "All_Mailouts_Combined.docx")
            composer.save(master_path)
            logger.info(f"Master DOCX created: {master_path}")
    except Exception as e:
        logger.error(f"Master DOCX error: {e}")

# ============ 6. MERGE QR PDFs IN EXCEL ORDER ============
logger.info("Merging PDFs in Excel order (with QR where available)...")

if len(qr_pdf_list) > 0:
    try:
        merger = PdfMerger()
        for pdf_file in qr_pdf_list:
            if os.path.exists(pdf_file):
                try:
                    merger.append(pdf_file)
                    logger.info(f"Added: {os.path.basename(pdf_file)}")
                except Exception as e:
                    logger.warning(f"Could not add PDF {pdf_file}: {e}")
        combined_pdf_path = os.path.join(OUTPUT_FOLDER, "All_Mailouts_Combined_QR.pdf")
        with open(combined_pdf_path, "wb") as f:
            merger.write(f)
        merger.close()
        logger.info(f"Combined QR PDF created: {combined_pdf_path}")
    except Exception as e:
        logger.error(f"PDF merge error: {e}")
else:
    logger.warning("No PDFs found to merge")

# ============ 7. CLEANUP ============
logger.info("Cleaning up temporary DOCX files...")
for f in generated_docx_list:
    if os.path.exists(f):
        try:
            os.remove(f)
        except Exception:
            pass

logger.info("Process COMPLETED SUCCESSFULLY")
