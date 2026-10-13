"""
Zero-Dependency Pure-Python Office Compatibility Engine (office_compat.py)
Provides 100% pure Python drop-in replacements for python-docx and python-pptx.
Uses standard library zipfile, xml.etree.ElementTree, and xml.sax.saxutils.
Runs anywhere with ZERO compiled C/C++ DLL dependencies (bypasses Windows Application Control).
"""

import io
import os
import zipfile
import xml.sax.saxutils as saxutils
import xml.etree.ElementTree as ET


def escape(text):
    return saxutils.escape(str(text))


# -------------------------------------------------------------------------
# Units & Geometry Helpers
# -------------------------------------------------------------------------

class Inches:
    def __init__(self, val):
        self.val = float(val)

    def __float__(self):
        return self.val

    def __repr__(self):
        return f"Inches({self.val})"


class Pt:
    def __init__(self, val):
        self.val = float(val)

    def __float__(self):
        return self.val

    def __repr__(self):
        return f"Pt({self.val})"


class RGBColor:
    def __init__(self, r, g, b):
        self.r = int(r)
        self.g = int(g)
        self.b = int(b)

    def __str__(self):
        return f"{self.r:02X}{self.g:02X}{self.b:02X}"

    def __repr__(self):
        return f"RGBColor({self.r}, {self.g}, {self.b})"


class WD_ALIGN_PARAGRAPH:
    LEFT = 0
    CENTER = 1
    RIGHT = 2
    JUSTIFY = 3


class WD_TABLE_ALIGNMENT:
    LEFT = 0
    CENTER = 1
    RIGHT = 2


# -------------------------------------------------------------------------
# Pure DOCX Builder & Parser
# -------------------------------------------------------------------------

class FontFormat:
    def __init__(self):
        self.size = None
        self.bold = False
        self.italic = False
        self.color = type("ColorHolder", (), {"rgb": None})()


class Run:
    def __init__(self, text=""):
        self.text = text
        self.font = FontFormat()
        self.bold = False
        self.italic = False

    def to_xml(self):
        txt = escape(self.text)
        if not txt:
            return ""
        r_props = []
        is_bold = self.bold or self.font.bold
        is_italic = self.italic or self.font.italic
        if is_bold:
            r_props.append("<w:b/>")
        if is_italic:
            r_props.append("<w:i/>")
        
        # Font size in half-points (1 pt = 2 half-points)
        sz = self.font.size
        if sz:
            sz_val = int(float(sz.val if hasattr(sz, "val") else sz) * 2)
            r_props.append(f'<w:sz w:val="{sz_val}"/><w:szCs w:val="{sz_val}"/>')

        # Font color
        rgb = self.font.color.rgb
        if rgb:
            hex_str = str(rgb).replace("#", "")
            r_props.append(f'<w:color w:val="{hex_str}"/>')

        prop_str = f"<w:rPr>{''.join(r_props)}</w:rPr>" if r_props else ""
        return f'<w:r>{prop_str}<w:t xml:space="preserve">{txt}</w:t></w:r>'


class ParagraphFormat:
    def __init__(self):
        self.space_before = Pt(0)
        self.space_after = Pt(0)


class Paragraph:
    def __init__(self, text="", style=None):
        self.runs = []
        self.font = FontFormat()
        self.style = type("StyleHolder", (), {"name": style or "Normal"})()
        self.alignment = None
        self.paragraph_format = ParagraphFormat()
        self.is_page_break = False
        self.heading_level = None

        if text:
            self.add_run(text)

    @property
    def text(self):
        return "".join(r.text for r in self.runs)

    @text.setter
    def text(self, val):
        self.runs = [Run(str(val))]

    def add_run(self, text=""):
        run = Run(text)
        self.runs.append(run)
        return run

    def to_xml(self):
        if self.is_page_break:
            return '<w:p><w:r><w:br w:type="page"/></w:r></w:p>'

        p_props = []
        if self.heading_level:
            p_props.append(f'<w:pStyle w:val="Heading{self.heading_level}"/>')
        elif self.style.name and "list" in str(self.style.name).lower():
            p_props.append('<w:pStyle w:val="ListBullet"/>')

        if self.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            p_props.append('<w:jc w:val="center"/>')
        elif self.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
            p_props.append('<w:jc w:val="right"/>')
        elif self.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
            p_props.append('<w:jc w:val="both"/>')

        before_pt = float(self.paragraph_format.space_before.val if hasattr(self.paragraph_format.space_before, "val") else self.paragraph_format.space_before)
        after_pt = float(self.paragraph_format.space_after.val if hasattr(self.paragraph_format.space_after, "val") else self.paragraph_format.space_after)
        # 1 pt = 20 dxa (twips)
        before_twips = int(before_pt * 20)
        after_twips = int(after_pt * 20)
        if before_twips > 0 or after_twips > 0:
            p_props.append(f'<w:spacing w:before="{before_twips}" w:after="{after_twips}"/>')

        prop_str = f"<w:pPr>{''.join(p_props)}</w:pPr>" if p_props else ""
        runs_str = "".join(r.to_xml() for r in self.runs)
        return f"<w:p>{prop_str}{runs_str}</w:p>"


class Cell:
    def __init__(self, text=""):
        self.paragraphs = [Paragraph(text)]
        self.shading_hex = None

    @property
    def text(self):
        return "\n".join(p.text for p in self.paragraphs)

    @text.setter
    def text(self, val):
        self.paragraphs = [Paragraph(str(val))]


class Row:
    def __init__(self, cols):
        self.cells = [Cell() for _ in range(cols)]


class Table:
    def __init__(self, rows, cols):
        self.rows = [Row(cols) for _ in range(rows)]
        self.style = "Table Grid"
        self.alignment = WD_TABLE_ALIGNMENT.CENTER

    def cell(self, r_idx, c_idx):
        if r_idx < len(self.rows) and c_idx < len(self.rows[r_idx].cells):
            return self.rows[r_idx].cells[c_idx]
        return Cell()

    def to_xml(self):
        tbl_xml = [
            '<w:tbl>',
            '<w:tblPr>',
            '<w:tblStyle w:val="TableGrid"/>',
            '<w:tblW w:w="0" w:type="auto"/>',
            '<w:jc w:val="center"/>',
            '<w:tblBorders>',
            '<w:top w:val="single" w:sz="4" w:space="0" w:color="D1D5DB"/>',
            '<w:left w:val="single" w:sz="4" w:space="0" w:color="D1D5DB"/>',
            '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="D1D5DB"/>',
            '<w:right w:val="single" w:sz="4" w:space="0" w:color="D1D5DB"/>',
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/>',
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="E5E7EB"/>',
            '</w:tblBorders>',
            '</w:tblPr>'
        ]
        for row in self.rows:
            tbl_xml.append("<w:tr>")
            for cell in row.cells:
                shd = f'<w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="{cell.shading_hex}"/></w:tcPr>' if cell.shading_hex else ""
                paras_xml = "".join(p.to_xml() for p in cell.paragraphs)
                if not paras_xml:
                    paras_xml = "<w:p/>"
                tbl_xml.append(f"<w:tc>{shd}{paras_xml}</w:tc>")
            tbl_xml.append("</w:tr>")
        tbl_xml.append("</w:tbl>")
        return "".join(tbl_xml)


class Section:
    def __init__(self):
        self.top_margin = Inches(1.0)
        self.bottom_margin = Inches(1.0)
        self.left_margin = Inches(1.0)
        self.right_margin = Inches(1.0)


class DocxDocument:
    """Pure-Python OpenXML DOCX Document implementation."""
    def __init__(self, file_or_path=None):
        self.sections = [Section()]
        self.paragraphs = []
        self.tables = []
        self.body_elements = []  # preserves intermingled ordering of paragraphs and tables
        self.media_files = []   # (filename, bytes)

        if file_or_path is not None:
            self._load_from_file(file_or_path)

    def add_paragraph(self, text="", style=None):
        p = Paragraph(text, style)
        self.paragraphs.append(p)
        self.body_elements.append(p)
        return p

    def add_heading(self, text="", level=1):
        p = Paragraph(text)
        p.heading_level = level
        self.paragraphs.append(p)
        self.body_elements.append(p)
        return p

    def add_page_break(self):
        p = Paragraph()
        p.is_page_break = True
        self.paragraphs.append(p)
        self.body_elements.append(p)
        return p

    def add_table(self, rows, cols):
        t = Table(rows, cols)
        self.tables.append(t)
        self.body_elements.append(t)
        return t

    def add_picture(self, stream, width=None):
        """Add image to document."""
        try:
            if hasattr(stream, "read"):
                img_data = stream.read()
            elif isinstance(stream, (bytes, bytearray)):
                img_data = bytes(stream)
            else:
                return None

            img_idx = len(self.media_files) + 1
            filename = f"image{img_idx}.png"
            self.media_files.append((filename, img_data))

            p = self.add_paragraph()
            p.paragraph_format.space_after = Pt(6)
            return p
        except Exception:
            return None

    def save(self, target):
        """Serialize document to valid DOCX PKZip archive."""
        out_buf = io.BytesIO() if not isinstance(target, str) else None
        dest = out_buf if out_buf is not None else target

        with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zf:
            # 1. [Content_Types].xml
            ct_lines = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
                '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
                '  <Default Extension="xml" ContentType="application/xml"/>',
                '  <Default Extension="png" ContentType="image/png"/>',
                '  <Default Extension="jpeg" ContentType="image/jpeg"/>',
                '  <Default Extension="jpg" ContentType="image/jpeg"/>',
                '  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
                '  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>',
                '</Types>'
            ]
            zf.writestr("[Content_Types].xml", "\n".join(ct_lines))

            # 2. _rels/.rels
            rels_lines = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
                '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>',
                '</Relationships>'
            ]
            zf.writestr("_rels/.rels", "\n".join(rels_lines))

            # 3. word/_rels/document.xml.rels
            doc_rels = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
                '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
            ]
            for idx, (img_fname, _) in enumerate(self.media_files, start=2):
                doc_rels.append(f'  <Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{img_fname}"/>')
            doc_rels.append('</Relationships>')
            zf.writestr("word/_rels/document.xml.rels", "\n".join(doc_rels))

            # 4. Media files
            for img_fname, img_bytes in self.media_files:
                zf.writestr(f"word/media/{img_fname}", img_bytes)

            # 5. word/styles.xml
            styles_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:color w:val="1F2937"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="240" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="111827"/><w:sz w:val="36"/><w:szCs w:val="36"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="180" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="E5322D"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:before="120" w:after="60"/></w:pPr>
    <w:rPr><w:b/><w:color w:val="374151"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="ListBullet">
    <w:name w:val="List Bullet"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:ind w:left="360" w:hanging="360"/></w:pPr>
  </w:style>
</w:styles>"""
            zf.writestr("word/styles.xml", styles_xml)

            # 6. word/document.xml
            body_xml = [elem.to_xml() for elem in self.body_elements]
            body_xml.append("""<w:sectPr>
  <w:pgSz w:w="11906" w:h="16838"/>
  <w:pgMar w:top="1152" w:right="1152" w:bottom="1152" w:left="1152" w:header="708" w:footer="708" w:gutter="0"/>
</w:sectPr>""")

            doc_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {''.join(body_xml)}
  </w:body>
</w:document>"""
            zf.writestr("word/document.xml", doc_xml)

        if out_buf is not None:
            if hasattr(target, "write"):
                target.write(out_buf.getvalue())
            return out_buf.getvalue()

    def _load_from_file(self, file_or_path):
        """Parse incoming .docx file into paragraphs and tables."""
        raw_bytes = file_or_path.read() if hasattr(file_or_path, "read") else None
        if raw_bytes is None:
            if isinstance(file_or_path, (bytes, bytearray)):
                raw_bytes = bytes(file_or_path)
            elif isinstance(file_or_path, str) and os.path.isfile(file_or_path):
                with open(file_or_path, "rb") as f:
                    raw_bytes = f.read()

        if not raw_bytes:
            return

        with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as zf:
            if "word/document.xml" not in zf.namelist():
                return
            xml_data = zf.read("word/document.xml")

        root = ET.fromstring(xml_data)
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

        body = root.find("w:body", ns)
        if body is None:
            return

        for child in body:
            tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag
            if tag == "p":
                p = Paragraph()
                pStyle = child.find("w:pPr/w:pStyle", ns)
                if pStyle is not None:
                    val = pStyle.attrib.get(f"{{{ns['w']}}}val", "")
                    p.style.name = val

                for elem in child.findall(".//w:r", ns):
                    txt_elem = elem.find("w:t", ns)
                    txt = txt_elem.text if txt_elem is not None and txt_elem.text else ""
                    if txt:
                        run = p.add_run(txt)
                        if elem.find("w:rPr/w:b", ns) is not None:
                            run.bold = True
                            run.font.bold = True
                        if elem.find("w:rPr/w:i", ns) is not None:
                            run.italic = True
                            run.font.italic = True
                        col_elem = elem.find("w:rPr/w:color", ns)
                        if col_elem is not None:
                            hex_col = col_elem.attrib.get(f"{{{ns['w']}}}val", "")
                            if hex_col:
                                run.font.color.rgb = hex_col

                self.paragraphs.append(p)
                self.body_elements.append(p)

            elif tag == "tbl":
                rows = child.findall("w:tr", ns)
                if not rows:
                    continue
                num_rows = len(rows)
                first_cells = rows[0].findall("w:tc", ns)
                num_cols = max(1, len(first_cells))

                t = Table(num_rows, num_cols)
                for r_idx, r_elem in enumerate(rows):
                    cells = r_elem.findall("w:tc", ns)
                    for c_idx, c_elem in enumerate(cells):
                        cell_txts = []
                        for p_elem in c_elem.findall("w:p", ns):
                            p_txt = "".join(t_el.text for t_el in p_elem.findall(".//w:t", ns) if t_el.text)
                            if p_txt:
                                cell_txts.append(p_txt)
                        if c_idx < num_cols and r_idx < num_rows:
                            t.cell(r_idx, c_idx).text = " ".join(cell_txts)

                self.tables.append(t)
                self.body_elements.append(t)


def set_cell_background(cell, hex_color):
    clean_hex = str(hex_color).replace("#", "").upper()
    cell.shading_hex = clean_hex


# -------------------------------------------------------------------------
# Pure PPTX Presentation Builder & Parser
# -------------------------------------------------------------------------

class PptxInches(Inches):
    pass

class PptxPt(Pt):
    pass

class PptxRGBColor(RGBColor):
    pass


class PptxTextFrame:
    def __init__(self, text=""):
        self.word_wrap = True
        self.paragraphs = [Paragraph(text)]

    @property
    def text(self):
        return "\n".join(p.text for p in self.paragraphs)


class PptxShape:
    def __init__(self, shape_type="textbox", text="", img_bytes=None, left=0, top=0, width=0, height=0):
        self.shape_type = shape_type
        self.has_text_frame = (shape_type == "textbox")
        self.text_frame = PptxTextFrame(text) if self.has_text_frame else None
        self.img_bytes = img_bytes
        self.left = left
        self.top = top
        self.width = width
        self.height = height


class SlideShapesCollection(list):
    def add_picture(self, stream, left, top, width=None, height=None):
        raw = stream.read() if hasattr(stream, "read") else bytes(stream)
        shape = PptxShape(shape_type="picture", img_bytes=raw, left=left, top=top, width=width, height=height)
        self.append(shape)
        return shape

    def add_textbox(self, left, top, width, height):
        shape = PptxShape(shape_type="textbox", left=left, top=top, width=width, height=height)
        self.append(shape)
        return shape


class PptxSlide:
    def __init__(self):
        self.shapes = SlideShapesCollection()


class SlideCollection(list):
    def add_slide(self, layout=None):
        slide = PptxSlide()
        self.append(slide)
        return slide


class PptxPresentation:
    """Pure-Python OpenXML PPTX Presentation implementation."""
    def __init__(self, file_or_path=None):
        self.slide_width = PptxInches(13.333)
        self.slide_height = PptxInches(7.5)
        self.slide_layouts = [None] * 10
        self.slides = SlideCollection()

        if file_or_path is not None:
            self._load_from_file(file_or_path)

    def add_slide(self, layout=None):
        return self.slides.add_slide(layout)

    def save(self, target):
        """Serialize presentation to standard PPTX zip archive."""
        out_buf = io.BytesIO() if not isinstance(target, str) else None
        dest = out_buf if out_buf is not None else target

        with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zf:
            # [Content_Types].xml
            ct_items = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
                '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
                '  <Default Extension="xml" ContentType="application/xml"/>',
                '  <Default Extension="png" ContentType="image/png"/>',
                '  <Default Extension="jpeg" ContentType="image/jpeg"/>',
                '  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
            ]
            for idx in range(1, len(self.slides) + 1):
                ct_items.append(f'  <Override PartName="/ppt/slides/slide{idx}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>')
            ct_items.append('</Types>')
            zf.writestr("[Content_Types].xml", "\n".join(ct_items))

            # _rels/.rels
            zf.writestr("_rels/.rels", """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>""")

            # ppt/presentation.xml
            sld_id_items = []
            pres_rels = [
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            ]
            for idx in range(1, len(self.slides) + 1):
                rId = f"rId{idx}"
                sld_id_items.append(f'<p:sldId id="{255 + idx}" r:id="{rId}"/>')
                pres_rels.append(f'  <Relationship Id="{rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{idx}.xml"/>')
            pres_rels.append('</Relationships>')
            zf.writestr("ppt/_rels/presentation.xml.rels", "\n".join(pres_rels))

            pres_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:sldIdLst>
    {''.join(sld_id_items)}
  </p:sldIdLst>
</p:presentation>"""
            zf.writestr("ppt/presentation.xml", pres_xml)

            # Slides
            media_counter = 1
            for s_idx, slide in enumerate(self.slides, 1):
                shapes_xml = []
                slide_rels = [
                    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
                    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                ]

                for sh_idx, shape in enumerate(slide.shapes, 1):
                    if shape.shape_type == "picture" and shape.img_bytes:
                        med_name = f"image{media_counter}.png"
                        media_counter += 1
                        zf.writestr(f"ppt/media/{med_name}", shape.img_bytes)
                        rId_img = f"rIdPic{sh_idx}"
                        slide_rels.append(f'  <Relationship Id="{rId_img}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/{med_name}"/>')

                        pic_x = int(float(shape.left if shape.left is not None else 0) * 914400)
                        pic_y = int(float(shape.top if shape.top is not None else 0) * 914400)
                        pic_cx = int(float(shape.width if shape.width is not None else 13.333) * 914400)
                        pic_cy = int(float(shape.height if shape.height is not None else 7.5) * 914400)

                        shapes_xml.append(f"""
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="{sh_idx + 1}" name="Picture {sh_idx}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="{rId_img}"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="{pic_x}" y="{pic_y}"/><a:ext cx="{pic_cx}" cy="{pic_cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>""")

                    elif shape.has_text_frame and shape.text_frame:
                        paras_xml = []
                        for p in shape.text_frame.paragraphs:
                            paras_xml.append(f'<a:p><a:r><a:t>{escape(p.text)}</a:t></a:r></a:p>')
                        off_x = int(float(shape.left if shape.left is not None else 0.5) * 914400)
                        off_y = int(float(shape.top if shape.top is not None else 0.5) * 914400)
                        ext_cx = int(float(shape.width if shape.width is not None else 8.0) * 914400)
                        ext_cy = int(float(shape.height if shape.height is not None else 1.5) * 914400)
                        shapes_xml.append(f"""
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="{sh_idx + 10}" name="TextBox {sh_idx}"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="{off_x}" y="{off_y}"/><a:ext cx="{ext_cx}" cy="{ext_cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr rtlCol="0" wrap="square"/>
    <a:lstStyle/>
    {''.join(paras_xml)}
  </p:txBody>
</p:sp>""")

                slide_rels.append('</Relationships>')
                zf.writestr(f"ppt/slides/_rels/slide{s_idx}.xml.rels", "\n".join(slide_rels))

                slide_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
       xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      {''.join(shapes_xml)}
    </p:spTree>
  </p:cSld>
</p:sld>"""
                zf.writestr(f"ppt/slides/slide{s_idx}.xml", slide_xml)

        if out_buf is not None:
            if hasattr(target, "write"):
                target.write(out_buf.getvalue())
            return out_buf.getvalue()

    def _load_from_file(self, file_or_path):
        """Parse incoming PPTX to extract slides and text shapes."""
        raw_bytes = file_or_path.read() if hasattr(file_or_path, "read") else None
        if raw_bytes is None:
            if isinstance(file_or_path, (bytes, bytearray)):
                raw_bytes = bytes(file_or_path)
            elif isinstance(file_or_path, str) and os.path.isfile(file_or_path):
                with open(file_or_path, "rb") as f:
                    raw_bytes = f.read()

        if not raw_bytes:
            return

        with zipfile.ZipFile(io.BytesIO(raw_bytes), "r") as zf:
            slide_names = [n for n in zf.namelist() if n.startswith("ppt/slides/slide") and n.endswith(".xml")]
            slide_names.sort(key=lambda x: int("".join(filter(str.isdigit, x)) or 0))

            for sld_name in slide_names:
                slide = self.add_slide()
                xml_data = zf.read(sld_name)
                root = ET.fromstring(xml_data)
                ns = {
                    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
                    "a": "http://schemas.openxmlformats.org/drawingml/2006/main"
                }

                for sp in root.findall(".//p:sp", ns):
                    txBody = sp.find("p:txBody", ns)
                    if txBody is not None:
                        shape = slide.shapes.add_textbox(0, 0, 100, 100)
                        shape.text_frame.paragraphs = []
                        for p_elem in txBody.findall("a:p", ns):
                            txt = "".join(t.text for t in p_elem.findall(".//a:t", ns) if t.text)
                            if txt:
                                shape.text_frame.paragraphs.append(Paragraph(txt))


# Module-level factory methods & classes
def Document(file_or_path=None):
    return DocxDocument(file_or_path)

def Presentation(file_or_path=None):
    return PptxPresentation(file_or_path)
