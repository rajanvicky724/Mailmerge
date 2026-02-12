from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import qrcode
import os

def add_qr_to_docx_xy(docx_path: str, url: str, qr_temp_folder: str, 
                       x_inches=6.5, y_inches=9.8, qr_size_inches=0.9):
    """
    Add QR code at absolute X,Y position in DOCX (like PDF).
    
    Args:
        x_inches: Horizontal position from left edge (default: 6.5")
        y_inches: Vertical position from top edge (default: 9.8")
        qr_size_inches: QR code width/height (default: 0.9")
    """
    try:
        # Generate QR
        qr = qrcode.make(url)
        qr_png = os.path.join(qr_temp_folder, f"qr_{os.path.basename(docx_path)}.png")
        qr.save(qr_png)

        doc = Document(docx_path)
        
        # Access the last section
        section = doc.sections[-1]
        
        # Add the QR as an inline shape first
        # (We need to add it to get a reference, then convert to floating)
        paragraph = doc.add_paragraph()
        run = paragraph.add_run()
        inline_shape = run.add_picture(qr_png, width=Inches(qr_size_inches))
        
        # Get the drawing element
        drawing = inline_shape._inline
        
        # Convert inline to anchor (floating/absolute positioning)
        # Get parent <w:drawing>
        drawing_parent = drawing.getparent()
        
        # Create anchor element
        anchor = OxmlElement('wp:anchor')
        
        # Copy attributes from inline
        anchor.set('distT', '0')
        anchor.set('distB', '0')
        anchor.set('distL', '114300')
        anchor.set('distR', '114300')
        anchor.set('simplePos', '0')
        anchor.set('relativeHeight', '251658240')
        anchor.set('behindDoc', '0')
        anchor.set('locked', '0')
        anchor.set('layoutInCell', '1')
        anchor.set('allowOverlap', '1')
        
        # Simple position
        simple_pos = OxmlElement('wp:simplePos')
        simple_pos.set('x', '0')
        simple_pos.set('y', '0')
        anchor.append(simple_pos)
        
        # Position horizontal (from left margin)
        pos_h = OxmlElement('wp:positionH')
        pos_h.set('relativeFrom', 'page')
        pos_offset_h = OxmlElement('wp:posOffset')
        # Convert inches to EMUs (English Metric Units: 1 inch = 914400 EMUs)
        pos_offset_h.text = str(int(x_inches * 914400))
        pos_h.append(pos_offset_h)
        anchor.append(pos_h)
        
        # Position vertical (from top margin)
        pos_v = OxmlElement('wp:positionV')
        pos_v.set('relativeFrom', 'page')
        pos_offset_v = OxmlElement('wp:posOffset')
        pos_offset_v.text = str(int(y_inches * 914400))
        pos_v.append(pos_offset_v)
        anchor.append(pos_v)
        
        # Extent (size)
        extent = OxmlElement('wp:extent')
        cx = int(qr_size_inches * 914400)
        cy = int(qr_size_inches * 914400)
        extent.set('cx', str(cx))
        extent.set('cy', str(cy))
        anchor.append(extent)
        
        # Effect extent
        effect_extent = OxmlElement('wp:effectExtent')
        effect_extent.set('l', '0')
        effect_extent.set('t', '0')
        effect_extent.set('r', '0')
        effect_extent.set('b', '0')
        anchor.append(effect_extent)
        
        # Wrap type (square wrapping)
        wrap_square = OxmlElement('wp:wrapSquare')
        wrap_square.set('wrapText', 'bothSides')
        anchor.append(wrap_square)
        
        # Document properties
        doc_pr = OxmlElement('wp:docPr')
        doc_pr.set('id', '1')
        doc_pr.set('name', 'QR Code')
        anchor.append(doc_pr)
        
        # Non-visual properties
        c_nv_graphic_frame_pr = OxmlElement('wp:cNvGraphicFramePr')
        anchor.append(c_nv_graphic_frame_pr)
        
        # Copy graphic data from inline
        graphic = drawing.find(qn('a:graphic'))
        if graphic is not None:
            anchor.append(graphic)
        
        # Replace inline with anchor in the drawing parent
        drawing_parent.replace(drawing, anchor)
        
        # Optional: Add caption below (as regular text)
        caption_para = doc.add_paragraph("Scan your custom QR code to enroll")
        caption_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        caption_para.runs[0].font.size = Pt(7)
        
        doc.save(docx_path)
        return True
        
    except Exception as e:
        print(f"⚠️ QR positioning failed for {os.path.basename(docx_path)}: {e}")
        return False
