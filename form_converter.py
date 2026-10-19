import os
import sys
import math

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import win32com.client
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

def set_cell_margins(cell, top=50, bottom=50, left=100, right=100):
    """Set cell margins in dxa (1 pt = 20 dxa)"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('w:top', top), ('w:bottom', bottom), ('w:left', left), ('w:right', right)]:
        node = OxmlElement(m)
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def remove_table_borders(table):
    """Make table borders completely invisible for layout purposes"""
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

def set_table_borders(table, color="CCCCCC", sz="4"):
    """Set clean light borders for grid tables"""
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

def convert_invoice_smart(pub_path, docx_path):
    pub_path = os.path.abspath(pub_path)
    docx_path = os.path.abspath(docx_path)
    img_dir = os.path.join(os.path.dirname(docx_path), "extracted_images")
    os.makedirs(img_dir, exist_ok=True)

    print(f"Reading Publisher file: {pub_path}")
    pub = win32com.client.Dispatch("Publisher.Application")
    doc = pub.Open(pub_path, False, False, 1)
    page = doc.Pages(1)

    # Export pictures
    pictures = {}
    for i in range(1, page.Shapes.Count + 1):
        s = page.Shapes(i)
        if s.Type == 13: # msoPicture
            img_file = os.path.join(img_dir, f"pic_{s.ID}.png")
            s.SaveAsPicture(img_file)
            pictures[s.ID] = {
                'path': img_file,
                'width': s.Width,
                'height': s.Height,
                'left': s.Left,
                'top': s.Top
            }

    word_doc = Document()
    
    # Standard A4 margins: 0.5 in (approx 36 pt)
    for section in word_doc.sections:
        section.page_width = Inches(8.27)  # A4 width
        section.page_height = Inches(11.69) # A4 height
        section.top_margin = Inches(0.4)
        section.bottom_margin = Inches(0.4)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    # 1. HEADER SECTION (Table: 1 row, 2 cols)
    # Left: Company name, address, phone
    # Right: Logo image, Tax ID
    header_tbl = word_doc.add_table(rows=1, cols=2)
    header_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(header_tbl)
    
    # Left cell width ~4.6 in, Right cell width ~2.4 in (Total 7.0 in)
    header_tbl.columns[0].width = Inches(4.7)
    header_tbl.columns[1].width = Inches(2.3)
    header_tbl.rows[0].cells[0].width = Inches(4.7)
    header_tbl.rows[0].cells[1].width = Inches(2.3)

    cell_l = header_tbl.rows[0].cells[0]
    cell_r = header_tbl.rows[0].cells[1]
    set_cell_margins(cell_l, top=0, bottom=40, left=0, right=50)
    set_cell_margins(cell_r, top=0, bottom=40, left=50, right=0)

    # Company Name
    p_comp = cell_l.paragraphs[0]
    p_comp.paragraph_format.space_after = Pt(2)
    p_comp.paragraph_format.line_spacing = 1.05
    run = p_comp.add_run("บริษัท คิวเอ็มแอล คอร์ปอเรชั่น จำกัด (สำนักงานใหญ่)")
    apply_thai_font(run, size_pt=16, bold=True)

    # Address
    p_addr = cell_l.add_paragraph()
    p_addr.paragraph_format.space_after = Pt(2)
    p_addr.paragraph_format.line_spacing = 1.05
    run = p_addr.add_run("116 ซอยลาดพร้าว 96 ถนนลาดพร้าว\nแขวงพลับพลา  เขตวังทองหลาง\nกรุงเทพมหานคร  10310")
    apply_thai_font(run, size_pt=13)

    # Phone
    p_phone = cell_l.add_paragraph()
    p_phone.paragraph_format.space_after = Pt(2)
    run = p_phone.add_run("โทร 095-8824173, 081-6893090")
    apply_thai_font(run, size_pt=13)

    # Right: Logo
    p_logo = cell_r.paragraphs[0]
    p_logo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_logo.paragraph_format.space_after = Pt(4)
    if 310 in pictures:
        p_logo.add_run().add_picture(pictures[310]['path'], width=Inches(1.65))

    # Right: Tax ID
    p_tax = cell_r.add_paragraph()
    p_tax.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_tax.paragraph_format.space_after = Pt(2)
    p_tax.paragraph_format.line_spacing = 1.05
    run = p_tax.add_run("เลขประจำตัวผู้เสียภาษีอากร\n0105546154518")
    apply_thai_font(run, size_pt=13)

    # 2. INVOICE TITLE (ใบแจ้งหนี้ (Invoice))
    p_title = word_doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(8)
    p_title.paragraph_format.space_after = Pt(6)
    run_title = p_title.add_run("ใบแจ้งหนี้  (Invoice)")
    apply_thai_font(run_title, size_pt=18, bold=True)

    # 3. CUSTOMER & INVOICE DETAILS SECTION (Table: 1 row, 2 cols)
    info_tbl = word_doc.add_table(rows=1, cols=2)
    info_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(info_tbl)
    info_tbl.columns[0].width = Inches(4.3)
    info_tbl.columns[1].width = Inches(2.7)
    info_tbl.rows[0].cells[0].width = Inches(4.3)
    info_tbl.rows[0].cells[1].width = Inches(2.7)

    c_cust = info_tbl.rows[0].cells[0]
    c_inv = info_tbl.rows[0].cells[1]
    set_cell_margins(c_cust, top=20, bottom=40, left=0, right=50)
    set_cell_margins(c_inv, top=20, bottom=40, left=50, right=0)

    # Customer
    p_cust = c_cust.paragraphs[0]
    p_cust.paragraph_format.space_after = Pt(2)
    p_cust.paragraph_format.line_spacing = 1.1
    r_cust_hdr = p_cust.add_run("Customer\n")
    apply_thai_font(r_cust_hdr, size_pt=13, bold=True)
    r_cust_body = p_cust.add_run(
        "บริษัท บี เค พี อุตสาหกรรมพลาสติก จำกัด \n"
        "166-166/1 หมู่ 8 ซ.กังวาล 2 ถ.เพชรเกษม\n"
        "ต.อ้อมใหญ่ อ.สามพราน จ.นครปฐม 73160\n"
        "เลขประจำตัวผู้เสียภาษีอากร 0735531000051"
    )
    apply_thai_font(r_cust_body, size_pt=13)

    # Invoice Details (Date, Due Date, Quotation)
    p_inv = c_inv.paragraphs[0]
    p_inv.paragraph_format.space_after = Pt(2)
    p_inv.paragraph_format.line_spacing = 1.15
    r_inv = p_inv.add_run(
        "วันที่ 17 ตุลาคม พ.ศ. 2569\n"
        "กำหนดวันชำระ:  ภายในวันที่ 20 ตุลาคม พ.ศ. 2569\n"
        "อ้างอิงใบเสนอราคา\n"
        "Quotation No. 25-BKP-CB-019"
    )
    apply_thai_font(r_inv, size_pt=13)

    # Spacer
    p_sp = word_doc.add_paragraph()
    p_sp.paragraph_format.space_before = Pt(4)
    p_sp.paragraph_format.space_after = Pt(4)

    # 4. ITEMS TABLE (6 rows x 6 cols)
    # Col widths in pub: [31.0, 34.0, 34.2, 227.3, 97.8, 82.2] pt
    # Total = 506.5 pt (~7.03 in)
    col_pt_widths = [31.0, 34.0, 34.2, 227.3, 97.8, 82.2]
    total_pt = sum(col_pt_widths)
    # Map to 7.0 inches
    col_in_widths = [(w / total_pt) * 7.0 for w in col_pt_widths]

    # Read table from publisher
    pub_table = None
    for i in range(1, page.Shapes.Count + 1):
        if page.Shapes(i).HasTable == -1:
            pub_table = page.Shapes(i).Table
            break

    table_data = []
    if pub_table:
        for r in range(1, pub_table.Rows.Count + 1):
            row_vals = []
            for c in range(1, pub_table.Columns.Count + 1):
                txt = pub_table.Rows(r).Cells(c).TextRange.Text.strip().replace("\r", "\n")
                row_vals.append(txt)
            table_data.append(row_vals)

    items_tbl = word_doc.add_table(rows=len(table_data), cols=6)
    items_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_borders(items_tbl, color="777777", sz="4")

    for col_idx, col in enumerate(items_tbl.columns):
        col.width = Inches(col_in_widths[col_idx])

    for r_idx, row_data in enumerate(table_data):
        row = items_tbl.rows[r_idx]
        for c_idx, val in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.width = Inches(col_in_widths[c_idx])
            set_cell_margins(cell, top=60, bottom=60, left=80, right=80)
            
            p = cell.paragraphs[0]
            p.paragraph_format.space_after = Pt(0)
            p.paragraph_format.line_spacing = 1.05
            
            # Text alignment logic:
            # Header row (r=0): Center
            # Numbers/Quantities (c=0, 1, 2): Center
            # Descriptions (c=3): Left
            # Amounts (c=4, 5): Right
            if r_idx == 0:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(val)
                apply_thai_font(run, size_pt=13, bold=True)
            else:
                if c_idx in [0, 1, 2]:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                elif c_idx in [4, 5]:
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                
                # Highlight total summary
                is_bold = (r_idx >= 3)
                run = p.add_run(val)
                apply_thai_font(run, size_pt=12.5, bold=is_bold)

    # 5. NOTE (Right aligned under table)
    p_note = word_doc.add_paragraph()
    p_note.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_note.paragraph_format.space_before = Pt(4)
    p_note.paragraph_format.space_after = Pt(12)
    run_note = p_note.add_run("หมายเหตุ: ลูกค้าสามารถหักภาษี ณ ที่จ่าย 3% ได้")
    apply_thai_font(run_note, size_pt=12.5, italic=True, color_hex="CC0000")

    # 6. SIGNATURES (Table: 1 row, 3 cols)
    sig_tbl = word_doc.add_table(rows=1, cols=3)
    sig_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(sig_tbl)
    
    col_w_sig = Inches(7.0 / 3.0) # ~2.33 in each
    for col in sig_tbl.columns:
        col.width = col_w_sig

    for c_idx in range(3):
        cell = sig_tbl.rows[0].cells[c_idx]
        cell.width = col_w_sig
        set_cell_margins(cell, top=40, bottom=40, left=40, right=40)

    # Col 1: ผู้จัดเตรียม
    c1 = sig_tbl.rows[0].cells[0]
    p1 = c1.paragraphs[0]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_after = Pt(2)
    run1 = p1.add_run("ผู้จัดเตรียม\n")
    apply_thai_font(run1, size_pt=13, bold=True)
    if 334 in pictures:
        p1.add_run().add_picture(pictures[334]['path'], height=Inches(0.42))
    p1.add_run("\n")
    run1_dt = p1.add_run("17 ตุลาคม พ.ศ. 2569")
    apply_thai_font(run1_dt, size_pt=12.5)

    # Col 2: ผู้มีอำนาจลงนาม
    c2 = sig_tbl.rows[0].cells[1]
    p2 = c2.paragraphs[0]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after = Pt(2)
    run2 = p2.add_run("ผู้มีอำนาจลงนาม\n")
    apply_thai_font(run2, size_pt=13, bold=True)
    if 313 in pictures:
        p2.add_run().add_picture(pictures[313]['path'], height=Inches(0.42))
    p2.add_run("\n")
    run2_dt = p2.add_run("17 ตุลาคม พ.ศ. 2569")
    apply_thai_font(run2_dt, size_pt=12.5)

    # Col 3: ผู้รับเอกสาร
    c3 = sig_tbl.rows[0].cells[2]
    p3 = c3.paragraphs[0]
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_after = Pt(2)
    run3 = p3.add_run("ผู้รับเอกสาร\n\n(............................................)\n")
    apply_thai_font(run3, size_pt=13, bold=True)
    run3_dt = p3.add_run("17 ตุลาคม พ.ศ. 2569")
    apply_thai_font(run3_dt, size_pt=12.5)

    # 7. PAYMENT INSTRUCTIONS
    p_pay = word_doc.add_paragraph()
    p_pay.paragraph_format.space_before = Pt(14)
    p_pay.paragraph_format.space_after = Pt(0)
    p_pay.paragraph_format.line_spacing = 1.15
    run_pay_hdr = p_pay.add_run("การชำระเงิน\n")
    apply_thai_font(run_pay_hdr, size_pt=13, bold=True)
    run_pay_body = p_pay.add_run(
        "     บัญชีธนาคารกรุงเทพฯ จำกัด มหาชน บัญชีเลขที่ 187-4-08684-4  สะสมทรัพย์  สาขารามอินทรา ในนาม \n"
        "      “บริษัทคิวเอ็มแอล คอร์ปอเรชั่น จำกัด”"
    )
    apply_thai_font(run_pay_body, size_pt=13)

    doc.Close()
    pub.Quit()

    word_doc.save(docx_path)
    print(f"\n[DONE] High-Fidelity Form DOCX saved: {docx_path}")

if __name__ == "__main__":
    src_pub = "Invoice BKP Surveillance audit ISO9001 (17-Oct-26).pub"
    dst_docx = "Invoice BKP Surveillance audit ISO9001 (17-Oct-26)_formatted.docx"
    convert_invoice_smart(src_pub, dst_docx)
