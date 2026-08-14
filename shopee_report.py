# -*- coding: utf-8 -*-
r"""
Shopee order -> SKU summary report generator
=================================================
Reads 3 input files:
    1) sku_shopee.xlsx
    2) shopee_order.pdf
    3) Seller_Centre_shopee.pdf

Produces:
    - deliver_shopee_order.pdf
    - shopee_order_numbered.pdf

Archive behavior on success:
    Moves `shopee_order.pdf` and `Seller_Centre_shopee.pdf` to an `Archive`
    folder, appended with today's date (DD-MM-YYYY).
"""

import sys, os, re, glob, platform, shutil, argparse
from datetime import datetime
from collections import defaultdict
import io

# --- [ส่วนที่เพิ่มใหม่] บังคับให้เปลี่ยนโฟลเดอร์ทำงานมาที่อยู่ปัจจุบันเสมอ (แก้ปัญหาดับเบิลคลิกบน Mac) ---
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
    # เผื่อกรณีรันเป็นไฟล์ .app
    if sys.platform == 'darwin' and application_path.endswith('MacOS'):
        application_path = os.path.abspath(os.path.join(application_path, '../../..'))
else:
    application_path = os.path.dirname(os.path.abspath(__file__))
os.chdir(application_path)
# -----------------------------------------------------------------------------------------

import openpyxl
import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.colors import red, white, black
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io


# ---------------------------------------------------------------------------
# 1) Load SKU reference table
# ---------------------------------------------------------------------------
def load_sku_reference(xlsx_path):
    """Returns dict: sku_code -> {'myanmar': str, 'thai': str, 'unit': int}"""
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.worksheets[0]

    ref = {}
    header_seen = False
    for row in ws.iter_rows(values_only=True):
        if not row or row[0] is None:
            continue
        sku = str(row[0]).strip()
        if sku.lower() == "sku":
            header_seen = True
            continue
        if sku.startswith("ตาราง"):
            if header_seen:
                break
            continue

        thai = row[1] if len(row) > 1 else None
        seq = row[3] if len(row) > 3 else None
        myanmar = row[5] if len(row) > 5 else None

        unit = 1
        if seq:
            m = re.search(r"(\d+)", str(seq))
            if m:
                unit = int(m.group(1))
        m2 = re.search(r"-(\d{2})$", sku)
        if m2:
            unit = max(unit, int(m2.group(1)))

        ref[sku.upper()] = {
            "myanmar": (str(myanmar).strip() if myanmar else ""),
            "thai": (str(thai).strip() if thai else ""),
            "unit": unit,
        }
    return ref


# ---------------------------------------------------------------------------
# 2) Parse the delivery-label PDF -> ordered list of Shopee Order No.
# ---------------------------------------------------------------------------
ORDER_NO_RE = re.compile(r"Shopee Order No\.\s*([A-Za-z0-9]+)")

def parse_order_labels(pdf_path):
    order_sequence = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            m = ORDER_NO_RE.search(text)
            order_sequence.append(m.group(1) if m else None)
    return order_sequence


def stamp_running_numbers(pdf_path, out_path, order_sequence):
    reader = PdfReader(pdf_path)
    writer = PdfWriter()

    for i, page in enumerate(reader.pages):
        num = i + 1
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)

        buf = io.BytesIO()
        c = canvas.Canvas(buf, pagesize=(w, h))

        x = w * (100.0 / 612.0)
        y_from_top = h * (425.3 / 792.0) + 90.0
        y = h - y_from_top

        box_w, box_h = 56, 56
        c.setFillColor(white)
        c.setStrokeColor(black)
        c.setLineWidth(1.6)
        c.roundRect(x - box_w / 2, y - box_h * 0.32, box_w, box_h, 4,
                    fill=1, stroke=1)

        c.setFillColor(red)
        c.setFont("Helvetica-Bold", 40)
        c.drawCentredString(x, y, str(num))
        c.save()
        buf.seek(0)

        overlay_reader = PdfReader(buf)
        page.merge_page(overlay_reader.pages[0])
        writer.add_page(page)

    with open(out_path, "wb") as f:
        writer.write(f)


# ---------------------------------------------------------------------------
# 3) Parse the Picklist PDF -> order_no -> [(sku, qty), ...]
# ---------------------------------------------------------------------------
def parse_picklist(pdf_path, known_skus):
    result = defaultdict(list)
    known_skus_upper = {s.upper() for s in known_skus}
    sku_list_sorted = sorted(known_skus_upper, key=len, reverse=True)
    known_find_re = re.compile("(" + "|".join(re.escape(s) for s in sku_list_sorted) + ")") \
        if sku_list_sorted else None
    generic_sku_re = re.compile(r"[A-Z0-9]{1,4}-[A-Z0-9]{1,3}-\d{2}")

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            words = page.extract_words()
            if not words:
                continue

            anchor_tops = [w["top"] for w in words if "หมายเลขคําสั่งซ้ือ" in w["text"]
                           or "หมายเลขคาสงซอ" in re.sub(r"[\u0300-\u036f\uf700-\uf7ff]", "", w["text"])]

            headers = []
            for atop in anchor_tops:
                same_line = [w for w in words if abs(w["top"] - atop) < 3]
                pkg = [w for w in same_line if w["text"] == "package"]
                if not pkg:
                    continue
                cands = [w for w in same_line
                         if re.fullmatch(r"[A-Za-z0-9]{8,}", w["text"])]
                if cands:
                    headers.append((atop, cands[0]["text"]))

            if not headers:
                continue
            headers.sort(key=lambda t: t[0])

            def order_for_top(top):
                current = None
                for htop, oid in headers:
                    if htop <= top + 0.5:
                        current = oid
                    else:
                        break
                return current

            by_top = defaultdict(list)
            for w in words:
                by_top[round(w["top"], 1)].append(w)

            for w in words:
                token = w["text"].upper()
                m_sku = known_find_re.search(token) if known_find_re else None
                if not m_sku:
                    m_sku = generic_sku_re.search(token)
                if m_sku:
                    sku = m_sku.group(0) if m_sku.re is generic_sku_re else m_sku.group(1)
                    line_words = sorted(by_top[round(w["top"], 1)], key=lambda x: x["x0"])
                    nums = [x["text"] for x in line_words if re.fullmatch(r"\d+", x["text"])]
                    qty = None
                    if len(nums) >= 2:
                        price, total = int(nums[-2]), int(nums[-1])
                        if price > 0 and total % price == 0:
                            qty = total // price
                        elif price > 0:
                            qty = round(total / price)
                    if qty is None:
                        qty = int(nums[-1]) if nums else 1

                    order_no = order_for_top(w["top"])
                    if order_no:
                        result[order_no].append((sku, qty))
    return result


# ---------------------------------------------------------------------------
# 4) Build final report
# ---------------------------------------------------------------------------
def _family_name_map(sku_ref):
    fam = defaultdict(list)
    for sku, info in sku_ref.items():
        prefix = sku.split("-")[0]
        if info["myanmar"]:
            fam[prefix].append(info["myanmar"])
    return {p: max(set(v), key=v.count) for p, v in fam.items()}


def build_report(order_sequence, picklist_map, sku_ref):
    rows = []
    missing_orders = []
    missing_skus = set()
    fam_map = _family_name_map(sku_ref)

    for i, order_no in enumerate(order_sequence, start=1):
        if not order_no:
            continue
        items = picklist_map.get(order_no)
        if not items:
            missing_orders.append((i, order_no))
            continue
        for sku, qty in items:
            info = sku_ref.get(sku.upper())
            if info:
                unit = info["unit"]
                name = info["myanmar"]
            else:
                missing_skus.add(sku)
                m = re.search(r"-(\d{2})$", sku)
                unit = int(m.group(1)) if m else 1
                name = fam_map.get(sku.split("-")[0], "")
            total_qty = qty * unit
            rows.append((i, name, sku, total_qty))

    return rows, missing_orders, missing_skus


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def register_fonts():
    thai_name, thai_bold_name = "Tahoma", "Tahoma-Bold"
    candidates = []
    if platform.system() == "Windows":
        windir = os.environ.get("WINDIR", r"C:\Windows")
        candidates.append((os.path.join(windir, "Fonts", "tahoma.ttf"),
                            os.path.join(windir, "Fonts", "tahomabd.ttf")))

    tahoma_ok = False
    for reg_path, bold_path in candidates:
        if os.path.exists(reg_path):
            pdfmetrics.registerFont(TTFont("Tahoma", reg_path))
            if os.path.exists(bold_path):
                pdfmetrics.registerFont(TTFont("Tahoma-Bold", bold_path))
            else:
                pdfmetrics.registerFont(TTFont("Tahoma-Bold", reg_path))
            tahoma_ok = True
            break

    if not tahoma_ok:
        fallback = os.path.join(SCRIPT_DIR, "fonts", "NotoSansThai-Regular.ttf")
        pdfmetrics.registerFont(TTFont("Tahoma", fallback))
        pdfmetrics.registerFont(TTFont("Tahoma-Bold", fallback))

    myanmar_path = os.path.join(SCRIPT_DIR, "fonts", "NotoSansMyanmar-Regular.ttf")
    pdfmetrics.registerFont(TTFont("Myanmar", myanmar_path))

    return thai_name, thai_bold_name, "Myanmar"


# ---------------------------------------------------------------------------
# Write PDF Report (SHOPEE)
# ---------------------------------------------------------------------------
def write_pdf_report(rows, out_path, rows_per_page=20, font_size=16):
    thai_font, thai_bold, myanmar_font = register_fonts()

    headers = ["order_number", "ชื่อต้นไม้(ภาษา)", "sku", "จำนวนรวมต้นไม้"]
    col_widths = [140, 155, 115, 95]

    # สไตล์สำหรับหัวกระดาษ (Shopee สีส้ม)
    title_style = ParagraphStyle(
        "title", fontName=thai_bold, fontSize=22,
        leading=28, alignment=1, spaceAfter=20, textColor=colors.HexColor("#EE4D2D")
    )
    
    header_style = ParagraphStyle(
        "header", fontName=thai_bold, fontSize=font_size,
        leading=font_size * 1.25, alignment=1, textColor=colors.white,
    )
    cell_style = ParagraphStyle(
        "cell", fontName=thai_font, fontSize=font_size,
        leading=font_size * 1.25, alignment=1,
    )
    myanmar_style = ParagraphStyle(
        "myanmar", fontName=myanmar_font, fontSize=font_size,
        leading=font_size * 1.25, alignment=1,
    )
    total_style = ParagraphStyle(
        "total", fontName=thai_bold, fontSize=font_size,
        leading=font_size * 1.25, alignment=1,
    )

    doc = SimpleDocTemplate(
        out_path, pagesize=A4,
        leftMargin=15 * mm, rightMargin=15 * mm,
        topMargin=15 * mm, bottomMargin=15 * mm,
    )

    grand_total = sum(int(qty) for _, _, _, qty in rows)

    elements = []
    
    # เพิ่มหัวกระดาษของระบบ Shopee ลงไปเป็นสิ่งแรก
    elements.append(Paragraph("ใบสรุปรายการจัดเตรียมสินค้า - SHOPEE", title_style))

    chunks = [rows[i:i + rows_per_page] for i in range(0, len(rows), rows_per_page)] or [[]]

    for chunk_i, chunk in enumerate(chunks):
        is_last_chunk = (chunk_i == len(chunks) - 1)

        table_data = [[Paragraph(h, header_style) for h in headers]]
        for order_number, name, sku, qty in chunk:
            table_data.append([
                Paragraph(str(order_number), cell_style),
                Paragraph(name or "", myanmar_style),
                Paragraph(str(sku), cell_style),
                Paragraph(str(qty), cell_style),
            ])

        style_cmds = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EE4D2D")), # เปลี่ยนสีหัวตารางเป็นส้ม
            ("GRID", (0, 0), (-1, -1), 0.75, colors.black),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F2F2F2")]),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]

        if is_last_chunk:
            total_row_idx = len(table_data)
            table_data.append([
                Paragraph("รวมจำนวนต้นไม้ทั้งหมด", total_style),
                "", "",
                Paragraph(str(grand_total), total_style),
            ])
            style_cmds += [
                ("SPAN", (0, total_row_idx), (2, total_row_idx)),
                ("BACKGROUND", (0, total_row_idx), (-1, total_row_idx),
                 colors.HexColor("#FFCCB6")), # สีพื้นหลังช่องสรุปยอด
                ("LINEABOVE", (0, total_row_idx), (-1, total_row_idx), 1.5, colors.black),
            ]

        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle(style_cmds))
        elements.append(table)
        
        if not is_last_chunk:
            elements.append(PageBreak())

    doc.build(elements)

def show_popup(message, title="Shopee Order Report", icon="info"):
    print(f"\n[{title}] {message}\n")
    if platform.system() == "Windows":
        import ctypes
        MB_ICONINFORMATION = 0x40
        MB_ICONERROR = 0x10
        MB_ICONWARNING = 0x30
        icon_flag = {"info": MB_ICONINFORMATION, "error": MB_ICONERROR,
                     "warning": MB_ICONWARNING}.get(icon, MB_ICONINFORMATION)
        ctypes.windll.user32.MessageBoxW(0, message, title, icon_flag)


class MissingInputFileError(Exception):
    pass


# ---------------------------------------------------------------------------
# File Helper & Archiving
# ---------------------------------------------------------------------------
def find_input_files(folder):
    sku_xlsx = glob.glob(os.path.join(folder, "sku_shopee*.xlsx*")) or \
               glob.glob(os.path.join(folder, "sku_shoppee*.xlsx*"))
    order_pdf = glob.glob(os.path.join(folder, "shopee_order*.pdf*")) or \
                glob.glob(os.path.join(folder, "shoppee_order*.pdf*"))
    picklist_pdf = glob.glob(os.path.join(folder, "Seller_Centre_shopee*.pdf*")) or \
                   glob.glob(os.path.join(folder, "Seller_Centre_shoppee*.pdf*"))

    missing = []
    if not sku_xlsx:
        missing.append("sku_shopee.xlsx")
    if not order_pdf:
        missing.append("shopee_order.pdf")
    if not picklist_pdf:
        missing.append("Seller_Centre_shopee.pdf")

    if missing:
        files_list = "\n".join(f"  - {m}" for m in missing)
        raise MissingInputFileError(
            f"ไม่พบไฟล์ต่อไปนี้ในโฟลเดอร์:\n{files_list}\n\n"
            f"กรุณานำไฟล์ที่ขาดไปวางไว้ในโฟลเดอร์เดียวกับโปรแกรมนี้ แล้วลองรันใหม่อีกครั้ง"
        )

    return sku_xlsx[0], order_pdf[0], picklist_pdf[0]


def archive_files(folder, order_pdf, picklist_pdf):
    archive_dir = os.path.join(folder, "Archive")
    if not os.path.exists(archive_dir):
        os.makedirs(archive_dir)

    today_str = datetime.now().strftime("%d-%m-%Y")

    def get_dest_path(original_file, default_prefix):
        ext = os.path.splitext(original_file)[1]
        dest_name = f"{default_prefix}_{today_str}{ext}"
        dest_path = os.path.join(archive_dir, dest_name)
        
        # Handle duplicates if run multiple times on same day
        counter = 1
        while os.path.exists(dest_path):
            dest_name = f"{default_prefix}_{today_str}_{counter}{ext}"
            dest_path = os.path.join(archive_dir, dest_name)
            counter += 1
        return dest_path

    order_dest = get_dest_path(order_pdf, "shopee_order")
    picklist_dest = get_dest_path(picklist_pdf, "Seller_Centre_shopee")

    shutil.move(order_pdf, order_dest)
    shutil.move(picklist_pdf, picklist_dest)
    
    print(f"[Archive] ย้ายไฟล์ไปยังโฟลเดอร์ Archive เรียบร้อยแล้ว:")
    print(f"      - {os.path.basename(order_dest)}")
    print(f"      - {os.path.basename(picklist_dest)}")


# ---------------------------------------------------------------------------
# main
# ---------------------------------------------------------------------------
def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--folder", default=".", help="โฟลเดอร์ที่เก็บไฟล์ต้นฉบับทั้ง 3 ไฟล์")
    args = parser.parse_args()

    folder = os.path.abspath(args.folder)
    outdir = folder

    sku_xlsx, order_pdf, picklist_pdf = find_input_files(folder)

    print(f"[1/5] อ่านตาราง SKU: {os.path.basename(sku_xlsx)}")
    sku_ref = load_sku_reference(sku_xlsx)
    print(f"      พบ {len(sku_ref)} รายการ SKU")

    print(f"[2/5] อ่านใบออเดอร์ (label): {os.path.basename(order_pdf)}")
    order_sequence = parse_order_labels(order_pdf)
    print(f"      พบ {len(order_sequence)} ออเดอร์ (หน้า)")

    numbered_pdf_out = os.path.join(outdir, "shopee_order_numbered.pdf")
    print(f"[3/5] ประทับหมายเลขรันนิ่งลงบนใบออเดอร์ -> {os.path.basename(numbered_pdf_out)}")
    stamp_running_numbers(order_pdf, numbered_pdf_out, order_sequence)

    print(f"[4/5] อ่าน Picklist: {os.path.basename(picklist_pdf)}")
    picklist_map = parse_picklist(picklist_pdf, sku_ref.keys())
    print(f"      พบข้อมูล SKU ของ {len(picklist_map)} ออเดอร์")

    rows, missing_orders, missing_skus = build_report(order_sequence, picklist_map, sku_ref)

    pdf_report_out = os.path.join(outdir, "deliver_shopee_order.pdf")
    print(f"[5/5] สร้างไฟล์สรุป PDF -> {os.path.basename(pdf_report_out)}")
    write_pdf_report(rows, pdf_report_out, rows_per_page=20, font_size=16)

    # Archive files after successful report creation
    archive_files(folder, order_pdf, picklist_pdf)

    print()
    print(f"เสร็จสิ้น! สร้างไฟล์ {len(rows)} แถวข้อมูล จาก {len(order_sequence)} ออเดอร์")
    if missing_orders:
        print(f"หมายเหตุ: ไม่พบข้อมูล SKU ใน Picklist สำหรับออเดอร์ {len(missing_orders)} รายการ:")
        for num, oid in missing_orders[:20]:
            print(f"   - order_number {num}: {oid}")
    if missing_skus:
        print(f"หมายเหตุ: พบรหัส SKU ที่ไม่อยู่ในตารางอ้างอิง: {sorted(missing_skus)}")

    return pdf_report_out, numbered_pdf_out


if __name__ == "__main__":
    try:
        report_pdf_out, numbered_pdf_out = main()
        print("\n=== SUCCESS ===")
        show_popup(
            "ทำงานสำเร็จ!\n\n"
            f"ไฟล์ผลลัพธ์:\n"
            f"  - {os.path.basename(report_pdf_out)}\n"
            f"  - {os.path.basename(numbered_pdf_out)}\n\n"
            "ย้ายไฟล์ต้นฉบับไปยังโฟลเดอร์ Archive เรียบร้อยแล้ว",
            title="สำเร็จ", icon="info",
        )
    except MissingInputFileError as e:
        show_popup(str(e), title="ไม่พบไฟล์ต้นฉบับ", icon="warning")
        sys.exit(1)
    except Exception as e:
        import traceback
        traceback.print_exc()
        show_popup(f"เกิดข้อผิดพลาดระหว่างประมวลผล:\n\n{e}",
                    title="เกิดข้อผิดพลาด", icon="error")
        sys.exit(1)