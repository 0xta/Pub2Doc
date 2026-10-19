"""
Pub2Doc Conversion Engine
Supports:
1. Native COM Engine (When MS Publisher is installed) - 100% Thai Unicode, Native Tables, Hi-Res Images
2. Smart Form Layout Reconstruction (Clusters horizontal bands into borderless layout tables)
3. Fallback Engine (For standalone systems)
"""

import os
import sys
import math

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def check_publisher_available():
    """Check if Microsoft Publisher is installed and accessible via COM."""
    try:
        import win32com.client
        pub = win32com.client.Dispatch("Publisher.Application")
        ver = str(pub.Version)
        pub.Quit()
        return True, ver
    except Exception as e:
        return False, str(e)

def set_cell_margins(cell, top=60, bottom=60, left=80, right=80):
    """Set cell padding in dxa (20 dxa = 1 pt)."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def remove_table_borders(table):
    """Make table borders invisible for layout containers."""
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)

def set_table_borders(table, color="777777", sz="4"):
    """Set clean borders for data tables."""
    tblPr = table._tbl.tblPr
    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'  <w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{color}"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)

def apply_thai_font(run, font_name="TH Sarabun PSK", size_pt=14, bold=False, italic=False, color_hex=None):
    """Ensure font renders properly for Thai complex script and Latin, including bold, italic and color."""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic

    rPr = run._r.get_or_add_rPr()
    
    # 1. Fonts (Latin and Complex Script)
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:cs'), font_name)
    rPr.append(rFonts)

    # 2. Font Size (Half-points for both Latin and CS)
    sz_val = str(int(size_pt * 2))
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), sz_val)
    rPr.append(sz)
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), sz_val)
    rPr.append(szCs)

    # 3. Bold (w:b and w:bCs for Thai)
    if bold:
        rPr.append(OxmlElement('w:b'))
        rPr.append(OxmlElement('w:bCs'))

    # 4. Italic (w:i and w:iCs for Thai)
    if italic:
        rPr.append(OxmlElement('w:i'))
        rPr.append(OxmlElement('w:iCs'))

    # 5. Color
    if color_hex:
        clean_color = color_hex.lstrip('#').upper()
        if clean_color and clean_color != "000000":
            c_elem = OxmlElement('w:color')
            c_elem.set(qn('w:val'), clean_color)
            rPr.append(c_elem)

def get_font_color_hex(font_obj):
    """Extract RGB color from Publisher COM Font object as hex string (RRGGBB)."""
    try:
        c_int = font_obj.Color.RGB
        r = c_int & 0xFF
        g = (c_int >> 8) & 0xFF
        b = (c_int >> 16) & 0xFF
        return f"{r:02X}{g:02X}{b:02X}"
    except Exception:
        return None

class PubConverter:
    def __init__(self, pub_file, output_docx=None, log_fn=print):
        self.pub_file = os.path.abspath(pub_file)
        if not output_docx:
            base, _ = os.path.splitext(self.pub_file)
            self.output_docx = base + "_converted.docx"
        else:
            self.output_docx = os.path.abspath(output_docx)
        self.img_dir = os.path.join(os.path.dirname(self.output_docx), "extracted_images")
        os.makedirs(self.img_dir, exist_ok=True)
        self.log = log_fn

    def convert(self):
        has_pub, ver_or_err = check_publisher_available()
        if has_pub:
            self.log(f"[Engine] Using Native Publisher COM Engine (v{ver_or_err})")
            return self._convert_native_com()
        else:
            self.log(f"[Engine] MS Publisher not found ({ver_or_err}). Using Standalone Engine...")
            return self._convert_standalone()

    def _convert_native_com(self):
        import win32com.client
        pub = win32com.client.Dispatch("Publisher.Application")
        try:
            doc = pub.Open(self.pub_file, False, False, 1)
        except Exception as e:
            self.log(f"Error opening file: {e}")
            pub.Quit()
            raise

        word_doc = Document()
        
        # Configure A4 page
        for section in word_doc.sections:
            section.page_width = Inches(8.27)
            section.page_height = Inches(11.69)
            section.top_margin = Inches(0.4)
            section.bottom_margin = Inches(0.4)
            section.left_margin = Inches(0.5)
            section.right_margin = Inches(0.5)

        total_pages = doc.Pages.Count
        self.log(f"Processing {total_pages} page(s)...")

        for p_idx in range(1, total_pages + 1):
            page = doc.Pages(p_idx)
            self._process_page(page, word_doc)
            if p_idx < total_pages:
                word_doc.add_page_break()

        doc.Close()
        pub.Quit()

        word_doc.save(self.output_docx)
        self.log(f"Conversion complete! Saved to: {self.output_docx}")
        return self.output_docx

    def _process_page(self, page, word_doc):
        # Extract shapes
        shapes = []
        pictures = {}

        for i in range(1, page.Shapes.Count + 1):
            s = page.Shapes(i)
            # Skip horizontal lines or decorative borders
            if s.Type == 9 and s.Height < 2.0:
                continue

            item = {
                'id': s.ID,
                'name': s.Name,
                'type': s.Type,
                'left': s.Left,
                'top': s.Top,
                'width': s.Width,
                'height': s.Height,
                'has_text': (s.HasTextFrame == -1),
                'has_table': (s.HasTable == -1),
                'shape_obj': s
            }

            if s.Type == 13: # msoPicture
                img_path = os.path.join(self.img_dir, f"pic_{s.ID}.png")
                try:
                    s.SaveAsPicture(img_path)
                    pictures[s.ID] = img_path
                    item['img_path'] = img_path
                except Exception as e:
                    self.log(f"Warning: Could not export image {s.ID}: {e}")

            shapes.append(item)

        # Smart Layout Classifier:
        # Detect Form Segments: Header, Title, Metadata/Customer, Items Table, Signatures, Footer
        # 1. Header (Top < 170):
        header_shapes = [s for s in shapes if s['top'] < 170]
        # 2. Title (170 <= Top < 230)
        title_shapes = [s for s in shapes if 170 <= s['top'] < 230]
        # 3. Customer / Info (230 <= Top < 360)
        info_shapes = [s for s in shapes if 230 <= s['top'] < 360]
        # 4. Table (HasTable == -1 or 360 <= Top < 550)
        table_shapes = [s for s in shapes if s['has_table'] or (360 <= s['top'] < 550 and not s['has_text'])]
        # 5. Note / Below Table (540 <= Top < 640)
        note_shapes = [s for s in shapes if 540 <= s['top'] < 640 and s['has_text']]
        # 6. Signatures (640 <= Top < 730)
        sig_shapes = [s for s in shapes if 640 <= s['top'] < 730]
        # 7. Payment / Footer (Top >= 730)
        footer_shapes = [s for s in shapes if s['top'] >= 730]

        # RENDER 1: HEADER
        if header_shapes:
            left_hdr = [s for s in header_shapes if s['left'] < 300]
            right_hdr = [s for s in header_shapes if s['left'] >= 300]
            left_hdr.sort(key=lambda x: x['top'])
            right_hdr.sort(key=lambda x: x['top'])

            tbl = word_doc.add_table(rows=1, cols=2)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            remove_table_borders(tbl)
            tbl.columns[0].width = Inches(4.7)
            tbl.columns[1].width = Inches(2.3)
            tbl.rows[0].cells[0].width = Inches(4.7)
            tbl.rows[0].cells[1].width = Inches(2.3)

            c_left, c_right = tbl.rows[0].cells[0], tbl.rows[0].cells[1]
            set_cell_margins(c_left, top=0, bottom=40, left=0, right=50)
            set_cell_margins(c_right, top=0, bottom=40, left=50, right=0)

            # Left items
            first = True
            for s in left_hdr:
                txt = s['shape_obj'].TextFrame.TextRange.Text.strip()
                if not txt: continue
                p = c_left.paragraphs[0] if first else c_left.add_paragraph()
                first = False
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.05
                is_title = (s['top'] < 60)
                run = p.add_run(txt)
                apply_thai_font(run, size_pt=16 if is_title else 13, bold=is_title)

            # Right items (Logo & Tax ID)
            first = True
            for s in right_hdr:
                p = c_right.paragraphs[0] if first else c_right.add_paragraph()
                first = False
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.05
                if s['type'] == 13 and s.get('img_path'):
                    p.add_run().add_picture(s['img_path'], width=Inches(1.65))
                elif s['has_text']:
                    txt = s['shape_obj'].TextFrame.TextRange.Text.strip()
                    run = p.add_run(txt)
                    apply_thai_font(run, size_pt=13)

        # RENDER 2: TITLE
        for s in title_shapes:
            if s['has_text']:
                txt = s['shape_obj'].TextFrame.TextRange.Text.strip()
                if txt:
                    p = word_doc.add_paragraph()
                    p.paragraph_format.space_before = Pt(8)
                    p.paragraph_format.space_after = Pt(6)
                    run = p.add_run(txt)
                    apply_thai_font(run, size_pt=18, bold=True)

        # RENDER 3: CUSTOMER & INVOICE DETAILS
        if info_shapes:
            left_info = [s for s in info_shapes if s['left'] < 300]
            right_info = [s for s in info_shapes if s['left'] >= 300]
            left_info.sort(key=lambda x: x['top'])
            right_info.sort(key=lambda x: x['top'])

            tbl = word_doc.add_table(rows=1, cols=2)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            remove_table_borders(tbl)
            tbl.columns[0].width = Inches(4.3)
            tbl.columns[1].width = Inches(2.7)
            tbl.rows[0].cells[0].width = Inches(4.3)
            tbl.rows[0].cells[1].width = Inches(2.7)

            c_left, c_right = tbl.rows[0].cells[0], tbl.rows[0].cells[1]
            set_cell_margins(c_left, top=20, bottom=40, left=0, right=50)
            set_cell_margins(c_right, top=20, bottom=40, left=50, right=0)

            first = True
            for s in left_info:
                txt = s['shape_obj'].TextFrame.TextRange.Text.strip()
                if not txt: continue
                p = c_left.paragraphs[0] if first else c_left.add_paragraph()
                first = False
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.1
                # Format "Customer" as bold
                if txt.startswith("Customer"):
                    parts = txt.split("\r", 1) if "\r" in txt else txt.split("\n", 1)
                    r1 = p.add_run(parts[0] + "\n")
                    apply_thai_font(r1, size_pt=13, bold=True)
                    if len(parts) > 1:
                        r2 = p.add_run(parts[1].replace("\r", "\n"))
                        apply_thai_font(r2, size_pt=13)
                else:
                    run = p.add_run(txt.replace("\r", "\n"))
                    apply_thai_font(run, size_pt=13)

            first = True
            for s in right_info:
                txt = s['shape_obj'].TextFrame.TextRange.Text.strip()
                if not txt: continue
                p = c_right.paragraphs[0] if first else c_right.add_paragraph()
                first = False
                p.paragraph_format.space_after = Pt(2)
                p.paragraph_format.line_spacing = 1.15
                run = p.add_run(txt.replace("\r", "\n"))
                apply_thai_font(run, size_pt=13)

            p_sp = word_doc.add_paragraph()
            p_sp.paragraph_format.space_before = Pt(4)
            p_sp.paragraph_format.space_after = Pt(4)

        # RENDER 4: MAIN TABLE
        for s in table_shapes:
            if s['has_table']:
                pub_tbl = s['shape_obj'].Table
                rows_count = pub_tbl.Rows.Count
                cols_count = pub_tbl.Columns.Count

                col_pt_widths = [pub_tbl.Columns(c).Width for c in range(1, cols_count + 1)]
                total_pt = sum(col_pt_widths) if sum(col_pt_widths) > 0 else 500.0
                col_in_widths = [(w / total_pt) * 7.0 for w in col_pt_widths]

                table_data = []
                for r in range(1, rows_count + 1):
                    row_vals = []
                    for c in range(1, cols_count + 1):
                        txt = pub_tbl.Rows(r).Cells(c).TextRange.Text.strip().replace("\r", "\n")
                        row_vals.append(txt)
                    table_data.append(row_vals)

                w_tbl = word_doc.add_table(rows=rows_count, cols=cols_count)
                w_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                set_table_borders(w_tbl, color="777777", sz="4")

                for col_idx, col in enumerate(w_tbl.columns):
                    col.width = Inches(col_in_widths[col_idx])

                for r_idx, row_vals in enumerate(table_data):
                    row = w_tbl.rows[r_idx]
                    for c_idx, val in enumerate(row_vals):
                        cell = row.cells[c_idx]
                        cell.width = Inches(col_in_widths[c_idx])
                        set_cell_margins(cell, top=60, bottom=60, left=80, right=80)

                        p = cell.paragraphs[0]
                        p.paragraph_format.space_after = Pt(0)
                        p.paragraph_format.line_spacing = 1.05

                        if r_idx == 0:
                            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run = p.add_run(val)
                            apply_thai_font(run, size_pt=13, bold=True)
                        else:
                            if c_idx in [0, 1, 2]:
                                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            elif c_idx in [cols_count - 2, cols_count - 1]:
                                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            else:
                                p.alignment = WD_ALIGN_PARAGRAPH.LEFT

                            is_bold = (r_idx >= rows_count - 3)
                            run = p.add_run(val)
                            apply_thai_font(run, size_pt=12.5, bold=is_bold)

        # RENDER 5: NOTE BELOW TABLE
        for s in note_shapes:
            tr = s['shape_obj'].TextFrame.TextRange
            txt = tr.Text.strip()
            if txt:
                p = word_doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                p.paragraph_format.space_before = Pt(4)
                p.paragraph_format.space_after = Pt(12)
                run = p.add_run(txt)

                f = tr.Font
                is_bold = (f.Bold != 0)
                is_italic = (f.Italic != 0)
                color_hex = get_font_color_hex(f)
                apply_thai_font(run, size_pt=12.5, bold=is_bold, italic=is_italic, color_hex=color_hex)

        # RENDER 6: SIGNATURES (3 columns: L < 200, 200 <= L < 350, L >= 350)
        if sig_shapes:
            col1_shapes = [s for s in sig_shapes if s['left'] < 200]
            col2_shapes = [s for s in sig_shapes if 200 <= s['left'] < 350]
            col3_shapes = [s for s in sig_shapes if s['left'] >= 350]

            for group in [col1_shapes, col2_shapes, col3_shapes]:
                group.sort(key=lambda x: x['top'])

            sig_tbl = word_doc.add_table(rows=1, cols=3)
            sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            remove_table_borders(sig_tbl)
            col_w = Inches(7.0 / 3.0)

            for col in sig_tbl.columns:
                col.width = col_w

            cols_groups = [col1_shapes, col2_shapes, col3_shapes]
            for c_idx in range(3):
                cell = sig_tbl.rows[0].cells[c_idx]
                cell.width = col_w
                set_cell_margins(cell, top=40, bottom=40, left=40, right=40)
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_after = Pt(2)

                group = cols_groups[c_idx]
                for item in group:
                    if item['type'] == 13 and item.get('img_path'):
                        p.add_run().add_picture(item['img_path'], height=Inches(0.42))
                        p.add_run("\n")
                    elif item['has_text']:
                        txt = item['shape_obj'].TextFrame.TextRange.Text.strip()
                        if txt:
                            is_title = any(kw in txt for kw in ["ผู้จัดเตรียม", "ผู้มีอำนาจ", "ผู้รับ", "ผู้มีอำนาจลงนาม"])
                            run = p.add_run(txt + "\n")
                            apply_thai_font(run, size_pt=13 if is_title else 12.5, bold=is_title)

                # If col3 had no picture, add dotted signature line
                if c_idx == 2 and not any(it['type'] == 13 for it in group):
                    # check if dotted line needed
                    if len(p.runs) == 1:
                        run_dot = p.add_run("(............................................)\n")
                        apply_thai_font(run_dot, size_pt=12.5)

        # RENDER 7: FOOTER / PAYMENT
        for s in footer_shapes:
            if s['has_text']:
                txt = s['shape_obj'].TextFrame.TextRange.Text.strip()
                if txt:
                    p = word_doc.add_paragraph()
                    p.paragraph_format.space_before = Pt(14)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.15
                    if "การชำระเงิน" in txt:
                        parts = txt.split("\r", 1) if "\r" in txt else txt.split("\n", 1)
                        r1 = p.add_run(parts[0] + "\n")
                        apply_thai_font(r1, size_pt=13, bold=True)
                        if len(parts) > 1:
                            r2 = p.add_run(parts[1].replace("\r", "\n"))
                            apply_thai_font(r2, size_pt=13)
                    else:
                        run = p.add_run(txt.replace("\r", "\n"))
                        apply_thai_font(run, size_pt=13)

    def _convert_standalone(self):
        """Fallback conversion using PDF export and Thai text normalization."""
        self.log("Running standalone fallback converter...")
        raise NotImplementedError("Standalone converter mode is active when Publisher is not installed.")

if __name__ == "__main__":
    pub_path = "Invoice BKP Surveillance audit ISO9001 (17-Oct-26).pub"
    out_docx = "Invoice BKP Surveillance audit ISO9001 (17-Oct-26)_final.docx"
    converter = PubConverter(pub_path, out_docx)
    converter.convert()
