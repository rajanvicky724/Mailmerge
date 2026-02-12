from docx import Document
from docx.shared import Inches
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import qrcode
import os

EMU_PER_INCH = 914400

def _new_anchor(run, image_path, width_inches, height_inches, pos_x_inches, pos_y_inches):
    """Create a wp:anchor element for a floating picture at absolute page coords."""
    part = run.part
    # add image to document media and get relation id
    with open(image_path, "rb") as f:
        image_bytes = f.read()
    image = part.package.image_parts.get_or_add_image_part(image_bytes)
    rId = part.relate_to(image, RT.IMAGE)

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
              <a:stretch>
                <a:fillRect/>
              </a:stretch>
            </pic:blipFill>
            <pic:spPr>
              <a:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="{cx}" cy="{cy}"/>
              </a:xfrm>
              <a:prstGeom prst="rect">
                <a:avLst/>
              </a:prstGeom>
            </pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
    """
    anchor = parse_xml(anchor_xml)
    return anchor

def add_qr_xy_to_docx(docx_path: str, url: str, qr_temp_folder: str,
                      x_inches: float, y_inches: float,
                      qr_size_inches: float = 0.9):
    """
    Add QR as floating image at absolute (x_inches, y_inches) from top-left of page.
    Page is usually 8.5 x 11 inches (letter).
    """
    try:
        # Generate QR image
        qr = qrcode.make(url)
        qr_png = os.path.join(qr_temp_folder, f"qr_{os.path.basename(docx_path)}.png")
        qr.save(qr_png)

        doc = Document(docx_path)

        # Anchor attached to an otherwise empty paragraph at end
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
        print("QR XY error:", e)
        return False
