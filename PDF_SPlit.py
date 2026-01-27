import os
from PyPDF2 import PdfReader, PdfWriter

INPUT_PDF = r"C:\Users\vigneshn\OneDrive - O`Connor and Associates\Desktop\XMAX Tree\PDF\IN\Retail_Merge.pdf"
OUTPUT_FOLDER = r"C:\Users\vigneshn\OneDrive - O`Connor and Associates\Desktop\XMAX Tree\PDF\Out"

def split_pdf(input_path, output_folder):
    """Split PDF into individual pages"""
    
    # Create output folder
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    # Read PDF
    reader = PdfReader(input_path)
    total_pages = len(reader.pages)
    print(f"Total Pages: {total_pages}\n")
    
    # Process each page
    for page_num, page in enumerate(reader.pages, 1):
        try:
            # Create new PDF with single page
            writer = PdfWriter()
            writer.add_page(page)
            
            # Save as 0001.pdf, 0002.pdf, etc (4-digit padding)
            filename = f"{page_num:04d}.pdf"
            output_path = os.path.join(output_folder, filename)
            
            with open(output_path, 'wb') as f:
                writer.write(f)
            
            print(f"Page {page_num:3d}/{total_pages} ✓ {filename}")
            
        except Exception as e:
            print(f"Page {page_num:3d} ✗ ERROR: {e}")
    
    print(f"\n✓ Done! Check '{output_folder}' folder")

if __name__ == "__main__":
    split_pdf(INPUT_PDF, OUTPUT_FOLDER)
