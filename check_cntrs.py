# -*- coding: utf-8 -*-
"""
ตรวจจับและเปรียบเทียบข้อมูลที่ไม่ตรงกันระหว่าง MANIFEST.xls และ CNTRS.xls
หัวข้อที่ตรวจ: SHED NUMBER, STATUS, TEMP, DG (CLASS/UN), REMARK

กฎ SHED NUMBER อ้างอิง: DG=2826, SCT=0302, BMT=0110, ลาดกระบัง=0332
กฎ REMARK (เฉพาะ SHED 0332 - ลาดกระบัง): ปกติ = BY TRAIN, ถ้ามี TEMP หรือ DG = BY TRUCK

Input:  input/MANIFEST.xls, input/CNTRS.xls
Output: output/CNTRS_EDI.xlsx  (สีเขียว ✓ = ตรงกัน, สีแดง ⚠ = ไม่ตรงกัน)
"""
import os
import re
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

INPUT_DIR = "input"
OUTPUT_DIR = "output"
MANIFEST_PATH = f"{INPUT_DIR}/MANIFEST.xls"
CNTRS_PATH = f"{INPUT_DIR}/CNTRS.xls"
OUTPUT_PATH = f"{OUTPUT_DIR}/CNTRS_EDI.xlsx"

# Expected SHED NUMBER reference table (per user's business rule)
# เก็บทั้งรูปแบบมี 0 นำหน้าและไม่มี เพราะ MANIFEST/CNTRS อาจพิมพ์เลขต่างรูปแบบกัน
SHED_RULES = {
    "2826": "DG (สินค้าอันตราย)",
    "0302": "SCT", "302": "SCT",
    "0110": "BMT", "110": "BMT",
    "0332": "ลาดกระบัง (ราง/รถไฟ)", "332": "ลาดกระบัง (ราง/รถไฟ)",
}

PKG_RE = re.compile(r"^[\d,]+\s+[A-Z/]+\s*\(")
NUM_RE = re.compile(r"([+-]?\d+)")
CONTAINER_RE = re.compile(r"^[A-Z]{4}\d{6,7}$")

# DG (CLASS/UN) shows up in MANIFEST free text in many different spellings,
# often with CLASS and UN mentioned in separate cells/rows of the same block:
# "CLASS:9  UN:3082", "UN NO.: 3082" + "CLASS: 9" on the next row, "(CLASS
# NO:9/UN NO:3082/PG:III)", "UN 2717 CLASS 4.1", the bare shorthand
# "9/3082/III" (class/UN/packing-group), or squashed run-on text where the
# line-wrap space got lost on export e.g. "5 PALLETSUN 3077 CL9". These are
# matched independently against the whole block's text rather than requiring
# one fixed phrasing. The trailing 3-4 digit requirement on UN is what keeps
# this safe even without a word-boundary check in front of it.
CLASS_RE = re.compile(r"CL(?:ASS)?\s*(?:NO)?[\s:#]*([0-9]+(?:\.[0-9]+)?(?:\+[0-9.]+)?)", re.IGNORECASE)
UN_RE = re.compile(r"UN\s*(?:NO)?[\s:#.]*([0-9]{3,4})", re.IGNORECASE)
SHORTHAND_DG_RE = re.compile(r"(?<!\d)([0-9]{1,2}(?:\.[0-9])?)\s*/\s*([0-9]{3,4})\s*/\s*(I{1,3}|N/?A)", re.IGNORECASE)


def norm_shed(s):
    """เทียบ SHED NUMBER ตามที่ปรากฏจริงในเอกสาร ไม่ตัดเลข 0 นำหน้า"""
    if s is None:
        return None
    s = str(s).strip()
    if s == "" or s.lower() == "nan":
        return None
    return s


def format_shed_code(raw):
    """CNTRS.xls เก็บ SHED NUMBER เป็นตัวเลข (เช่น 332) แต่ Excel แสดงผลด้วย
    format 4 หลัก (0332) - แปลงกลับให้ตรงกับที่ตามองเห็นจริงในไฟล์ CNTRS"""
    if raw is None:
        return None
    s = str(raw).strip()
    if s.endswith(".0"):
        s = s[:-2]
    if s.isdigit() and len(s) < 4:
        s = s.zfill(4)
    return s


def norm_status(raw):
    """รวม CY/FCL เป็น FCL และรวม LCL(AGENT)/LCL/CFS ให้เป็น LCL เดียวกัน
    เพื่อเทียบข้ามรูปแบบคำที่ MANIFEST กับ CNTRS ใช้ไม่เหมือนกัน"""
    if not raw:
        return None
    r = str(raw).upper().strip()
    if "=" in r:
        r = r.split("=", 1)[1]
    if r == "CY" or "FCL" in r:
        return "FCL"
    if "LCL" in r:
        return "LCL"
    return r


TEMP_AFTER_KEYWORD_RE = re.compile(r"TEMP\w*\s*:?\s*([+-]?\d+)", re.IGNORECASE)


def extract_manifest_temp(temps_list):
    """ดึงตัวเลขอุณหภูมิ - ต้องหาตัวเลขที่อยู่ "หลังคำว่า TEMP" ก่อนเสมอ เพราะบาง
    เซลล์มีตัวเลขอื่นปนอยู่ก่อนหน้า (เช่น "HS CODE:080830TEMP:+1'C" ไม่งั้นจะไป
    หยิบเลข HS CODE 080830 มาใส่แทนอุณหภูมิจริง +1). ถ้าไม่เจอคำว่า TEMP เลย (เช่น
    ประโยคที่ใช้คำว่า DEGREES/CELSIUS แทน) ค่อย fallback ไปหาตัวเลขตัวแรกในข้อความ"""
    text = " ".join(temps_list).strip()
    if not text:
        return None
    m = TEMP_AFTER_KEYWORD_RE.search(text)
    if not m:
        m = NUM_RE.search(text)
    if not m:
        return None
    val = m.group(1)
    if not val.startswith(("+", "-")):
        # ข้อความบางแบบเขียนว่า "MINUS 18 DEGREES" แทนเครื่องหมาย "-"
        val = ("-" if re.search(r"\bMINUS\b", text.upper()) else "+") + val
    return val + "C"


def extract_cntrs_temp(remark):
    if not remark or "REEFER" not in remark.upper():
        return None
    m = NUM_RE.search(remark)
    if not m:
        return None
    val = m.group(1)
    if not val.startswith(("+", "-")):
        val = "+" + val
    return val + "C"


def extract_dg_parts(text):
    """คืนค่า (class, un) แยกส่วน สำหรับเทียบข้อมูล - ใช้คนละหน้าที่กับ extract_dg()
    ซึ่งคืนค่าข้อความสวยๆ สำหรับแสดงผล"""
    if not text:
        return None
    b = text.upper()
    cls_m = CLASS_RE.search(b)
    un_m = UN_RE.search(b)
    if cls_m and un_m:
        return (cls_m.group(1), un_m.group(1))
    short_m = SHORTHAND_DG_RE.search(b)
    if short_m:
        return (short_m.group(1), short_m.group(2))
    if "HAZARDOUS" in b or "DANGEROUS" in b:
        return ("?", None)
    return None


def extract_dg(text):
    parts = extract_dg_parts(text)
    if not parts:
        return None
    cls, un = parts
    if un is None:
        return "DG (unspecified class/UN)"
    return f"CLASS {cls} UN{un}"


def dg_matches(manifest_text, cntrs_text):
    """เทียบ DG โดยยึด UN NUMBER เป็นหลัก (เลขนี้ระบุชนิดสารเคมีตัวเดียวตายตัว)
    ส่วน CLASS บางทีฝั่ง MANIFEST เขียนรวม subsidiary risk ต่อท้ายด้วย '+' เช่น
    '3+6.1' ขณะที่ CNTRS เขียนแค่คลาสหลัก '3' เฉยๆ - ถือว่าตรงกันถ้าเลขคลาสหลัก
    (ส่วนก่อน '+') เท่ากัน ไม่ต้องเป๊ะทั้งสตริง"""
    m_parts = extract_dg_parts(manifest_text)
    c_parts = extract_dg_parts(cntrs_text)
    if not m_parts or not c_parts:
        return m_parts == c_parts
    m_class, m_un = m_parts
    c_class, c_un = c_parts
    if m_un != c_un:
        return False
    if m_un is None:
        return m_class == c_class
    return m_class.split("+")[0].strip() == c_class.split("+")[0].strip()


# ---------------- MANIFEST PARSING ----------------
PAGE_PORT_RE = re.compile(r"PortOfDischarge:\s*([A-Z]{5})")
REEFER_TYPE_RE = re.compile(r"^\d{2}R\d")
VENT_RE = re.compile(r"(?<!PRE)(?<!E)VENT[A-Z]*\s*[:.]?\s*([A-Z0-9]\S*)", re.IGNORECASE)


EXCLUDED_PORTS = ("BKK", "UCT")  # ตรวจทุก PORT ยกเว้นพอร์ตที่รหัสลงท้ายด้วยชุดนี้ (THBKK, THUCT)


def parse_manifest(path, excluded_ports=EXCLUDED_PORTS):
    """MANIFEST.xls is a printed report exported to Excel: the merged-cell
    column boundaries drift from shipment to shipment (a container list that
    sits in column 5 in one file can land in column 6 in another, etc.), so
    every field below is located by scanning ALL columns of a row for its
    pattern instead of trusting a fixed column index. Only column 0 (B/L no. /
    'S :' / 'C :' / 'N :' prefix) has stayed stable across the files seen.

    excluded_ports: B/L blocks printed on pages whose footer says
    'PortOfDischarge: <port>' with a port code ending in one of these (BKK,
    UCT) are skipped; every other port is checked."""
    mdf = pd.ExcelFile(path).parse("Sheet1", header=None)
    n = len(mdf)
    ncols = mdf.shape[1]

    def row_cells(idx):
        for c in range(ncols):
            v = mdf.iat[idx, c]
            if pd.notna(v):
                yield c, str(v).strip()

    # each page ends with a "PortOfDischarge: THxxx ..." footer row, so a row's
    # port is the port on the first footer at or below it
    footers = []
    for idx in range(n):
        c0 = mdf.iat[idx, 0]
        if pd.notna(c0):
            m = PAGE_PORT_RE.search(str(c0))
            if m:
                footers.append((idx, m.group(1)))

    def port_of_row(idx):
        for fidx, port in footers:
            if fidx >= idx:
                return port
        return None

    header_rows = []
    for idx in range(n):
        c0 = mdf.iat[idx, 0]
        if pd.isna(c0):
            continue
        c0s = str(c0).strip()
        if c0s.startswith(("S :", "C :", "N :", "(")):
            continue
        if any(PKG_RE.match(v) for c, v in row_cells(idx) if c != 0):
            header_rows.append(idx)

    bl_blocks = []
    for hi, start in enumerate(header_rows):
        end = header_rows[hi + 1] if hi + 1 < len(header_rows) else n
        bl_no = str(mdf.iat[start, 0]).strip()
        consignee = None
        status_raw = None
        shed_no = None
        containers = []
        temps = []
        desc_all = []
        is_reefer_type = False
        for r in range(start, end):
            c0 = mdf.iat[r, 0]
            if pd.notna(c0):
                c0s = str(c0).strip()
                if c0s.startswith("C :"):
                    consignee = c0s[3:].strip()
            for c, vs in row_cells(r):
                if c == 0:
                    continue
                m = re.match(r"^\d+\.\s*([A-Z]{4}\d{6,7})", vs)
                if m:
                    containers.append(m.group(1))
                    continue
                # STATUS อาจมีข้อความต่อท้าย เช่น "LCL AGENT'S LABOUR" หรือ
                # "LCL(AGENT)" ไม่ใช่แค่ "LCL" เฉยๆ จึงจับแบบขึ้นต้นด้วยคำนี้
                if status_raw is None and re.match(r"^(CY|LCL(/CFS)?|FCL)\b", vs, re.IGNORECASE):
                    status_raw = vs
                    continue
                # เก็บข้อความคำอธิบายสินค้าทั้งหมดไว้เป็น "blob" เดียว เพราะ
                # CLASS/UN ของ DG มักถูกแยกคนละเซลล์/คนละแถวกัน
                desc_all.append(vs)
                if REEFER_TYPE_RE.match(vs):
                    is_reefer_type = True  # 22R1/45R1... = ตู้เย็น
                vsu = vs.upper()
                if "TEMP" in vsu or "DEGREE" in vsu or "CELSIUS" in vsu:
                    temps.append(vs)
                m2 = re.search(r"SHED\s*NO\.?\s*([0-9]+)", vsu)
                if m2:
                    shed_no = m2.group(1)

        vent_m = VENT_RE.search(" ".join(desc_all))
        bl_blocks.append({
            "bl_no": bl_no,
            "consignee": consignee,
            "status_raw": status_raw,
            "shed_no": shed_no,
            "shed_no_norm": norm_shed(shed_no),
            "containers": containers,
            "temps": temps,
            "desc_all": desc_all,
            "is_reefer_type": is_reefer_type,
            "vent": vent_m.group(1) if vent_m else None,
            "port": port_of_row(start),
        })

    containers_map = {}
    for b in bl_blocks:
        if b["port"] and b["port"].endswith(tuple(excluded_ports)):
            continue
        for c in b["containers"]:
            containers_map[c] = b
    return containers_map


def truncate_consignee(s):
    """ตัดชื่อ CONSIGNEE ให้เหลือแค่ถึงคำว่า 'C/O' เช่น 'HEUNG A LINE CO.,LTD. C/O'
    (ใช้กฎเดียวกันไม่ว่าชื่อก่อน C/O จะเป็นอะไร)"""
    if not s:
        return s
    idx = s.find("C/O")
    if idx != -1:
        return s[: idx + 3].strip()
    return s.strip()


def find_cntrs_columns(cdf, scan_rows=10):
    """CNTRS.xls column positions drift between shipments - REMARK has been
    seen directly under its own header label in one file, and one column to
    the LEFT of its header label in another (same for STATUS/CONSIGNEE), so a
    fixed offset guess isn't reliable. Instead: find each header label's
    column, then look at real container data rows and see which of
    {header-1, header, header+1} is actually populated there - that's the
    true data column for this particular file."""
    header = {}
    for idx in range(min(scan_rows, len(cdf))):
        for c in range(cdf.shape[1]):
            v = cdf.iat[idx, c]
            if pd.notna(v) and str(v).strip().upper() in ("STATUS", "POL", "REMARK", "CONSIGNEE"):
                header[str(v).strip().upper()] = c
        if "REMARK" in header and "STATUS" in header and "CONSIGNEE" in header:
            break

    sample_rows = []
    for idx in range(len(cdf)):
        c0 = cdf.iat[idx, 0]
        if pd.notna(c0) and CONTAINER_RE.match(str(c0).strip()):
            sample_rows.append(idx)
            if len(sample_rows) >= 20:
                break

    def resolve(label, default):
        hc = header.get(label)
        if hc is None:
            return default
        best_c, best_count = hc, -1
        for cand in (hc - 1, hc, hc + 1):
            if cand < 0 or cand >= cdf.shape[1]:
                continue
            count = sum(1 for idx in sample_rows if pd.notna(cdf.iat[idx, cand]))
            if count > best_count:
                best_count, best_c = count, cand
        return best_c

    return {
        "status": resolve("STATUS", 18),
        "remark": resolve("REMARK", 22),
        "consignee": resolve("CONSIGNEE", 9),
    }


# ---------------- CNTRS PARSING ----------------

# ---------------- CNTRS PDF PARSING ----------------
PDF_CONTAINER_LINE_RE = re.compile(
    r"^(?P<cno>[A-Z]{4}\d{6,7})\s+(?P<item>\d+)\s+(?P<type>\S+)\s+(?P<size>\S+\s*\(\d+'\))\s+"
    r"(?P<cons>.*?)\s+(?P<wt>[\d,]+(?:\.\d+)?)\s+KGM\s+(?P<status>\d=\S+)\s+(?P<pol>[A-Z]{5})\s*(?P<remark>.*)$")
PDF_HEADER_RE = re.compile(
    r"CONTAINER LIST FOR FEEDER\s+(?P<feeder>.+?)\s+VOYAGE\s+(?P<voyage>\S+)\s+ARRIVAL DATE\s+(?P<arr>\S+)\s+"
    r"Port Of Discharge\s+(?P<pod>\w+)(?:\s+Port Of Delivery\s+(?P<pod2>\w+))?", re.IGNORECASE)
PDF_SKIP_PREFIXES = ("TOTAL", "PORTOFDISCHARGE", "HEUNG A LINE (THAILAND)", "TEL.", "CONTAINER LIST", "CONTAINER NO.",
                     "AS AGENTS", "SHED NUMBER", "CO.,LTD.", "/")


def pdf_lines(path):
    import pdfplumber
    lines = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            lines.extend((page.extract_text() or "").splitlines())
    return lines


def parse_cntrs_pdf(path):
    """CNTRS.pdf (พิมพ์จากระบบเดียวกับ CNTRS.xls): ข้อความเป็นบรรทัดสะอาด 1 ตู้ต่อ 1
    บรรทัด จึงอ่านด้วย regex ต่อบรรทัด. หัวแต่ละส่วนบอก Port Of Delivery และ
    'SHED NUMBER: xxxx' ใช้กับตู้ที่ตามมา; บรรทัดที่ต่อท้ายตู้และไม่ใช่หัว/ยอดรวม
    ถือเป็นส่วนต่อของ REMARK"""
    containers = {}
    shed = shed_desc = delivery = None
    last = None
    for raw in pdf_lines(path):
        line = raw.strip()
        if not line:
            continue
        h = PDF_HEADER_RE.search(line)
        if h:
            delivery = h.group("pod2")
            last = None
            continue
        m = re.match(r"^SHED NUMBER:\s*(\d+)\s*(.*)$", line)
        if m:
            shed, shed_desc = format_shed_code(m.group(1)), m.group(2).strip() or None
            last = None
            continue
        m = PDF_CONTAINER_LINE_RE.match(line)
        if m and shed is not None:
            containers[m.group("cno")] = {
                "item": int(m.group("item")),
                "shed_no": shed,
                "shed_no_norm": norm_shed(shed),
                "shed_desc": shed_desc,
                "delivery": delivery,
                "consignee": truncate_consignee(m.group("cons")),
                "status_raw": m.group("status"),
                "remark_raw": m.group("remark").strip() or None,
            }
            last = m.group("cno")
            continue
        if last and not line.upper().startswith(PDF_SKIP_PREFIXES):
            prev = containers[last]["remark_raw"]
            containers[last]["remark_raw"] = f"{prev} {line}" if prev else line
    return containers


def parse_cntrs_header_pdf(path):
    for line in pdf_lines(path):
        h = PDF_HEADER_RE.search(line)
        if h:
            return {"feeder": h.group("feeder"), "voyage": h.group("voyage"),
                    "arrival_date": h.group("arr"), "port_of_discharge": h.group("pod")}
    return {"feeder": "-", "voyage": "-", "arrival_date": "-", "port_of_discharge": "-"}


def parse_cntrs(path):
    if path.lower().endswith(".pdf"):
        return parse_cntrs_pdf(path)
    """CNTRS.xls: rows are grouped under 'SHED NUMBER:' headers. Each container
    row carries CONSIGNEE (may wrap onto the next row), STATUS (e.g. '8=FCL')
    and REMARK (free text such as 'REEFER -3C BY TRUCK' or 'HAZARDOUS CLASS 3
    UN 1263') - see find_cntrs_columns() for how their positions are located."""
    cdf = pd.ExcelFile(path).parse("Sheet1", header=None)
    n = len(cdf)
    cols = find_cntrs_columns(cdf)
    status_col, remark_col, consignee_col = cols["status"], cols["remark"], cols["consignee"]
    containers = {}
    cur_shed = None
    cur_shed_desc = None
    delivery = None
    for c in range(cdf.shape[1]):
        v = cdf.iat[3, c] if len(cdf) > 3 else None
        if pd.notna(v) and str(v).strip().upper() == "PORT OF DELIVERY":
            for c2 in range(c + 1, cdf.shape[1]):
                if pd.notna(cdf.iat[3, c2]):
                    delivery = str(cdf.iat[3, c2]).strip()
                    break
    for idx in range(n):
        c0 = cdf.iat[idx, 0]
        if pd.notna(c0) and str(c0).strip() == "SHED NUMBER:":
            shed_code = cdf.iat[idx, 2]
            shed_desc = cdf.iat[idx, 5]
            cur_shed = format_shed_code(shed_code) if pd.notna(shed_code) else None
            cur_shed_desc = str(shed_desc).strip() if pd.notna(shed_desc) else None
            continue
        if pd.notna(c0) and pd.notna(cdf.iat[idx, 2]):
            c0s = str(c0).strip()
            item = cdf.iat[idx, 2]
            if CONTAINER_RE.match(c0s) and cur_shed is not None:
                status = cdf.iat[idx, status_col]
                remark = cdf.iat[idx, remark_col]
                consignee_parts = []
                c9 = cdf.iat[idx, consignee_col]
                if pd.notna(c9):
                    consignee_parts.append(str(c9).strip())
                if idx + 1 < n and pd.isna(cdf.iat[idx + 1, 0]) and pd.isna(cdf.iat[idx + 1, 2]):
                    c9_next = cdf.iat[idx + 1, consignee_col]
                    if pd.notna(c9_next):
                        consignee_parts.append(str(c9_next).strip())
                consignee_full = " ".join(consignee_parts) if consignee_parts else None
                containers[c0s] = {
                    "item": int(item) if pd.notna(item) else None,
                    "shed_no": cur_shed,
                    "shed_no_norm": norm_shed(cur_shed),
                    "shed_desc": cur_shed_desc,
                    "delivery": delivery,
                    "consignee": truncate_consignee(consignee_full),
                    "status_raw": str(status).strip() if pd.notna(status) else None,
                    "remark_raw": str(remark).strip() if pd.notna(remark) else None,
                }
    return containers


def parse_cntrs_header(path):
    if path.lower().endswith(".pdf"):
        return parse_cntrs_header_pdf(path)
    """ข้อมูลหัวฟอร์ม CNTRS: FEEDER / VOYAGE / ARRIVAL DATE / PORT OF DISCHARGE (แถวที่ 4 ของไฟล์)."""
    cdf = pd.ExcelFile(path).parse("Sheet1", header=None, nrows=5)
    row = cdf.iloc[3]
    def get(c):
        v = row.get(c)
        return str(v).strip() if pd.notna(v) else "-"
    return {
        "feeder": get(7),
        "voyage": get(12),
        "arrival_date": get(17),
        "port_of_discharge": get(24),
    }


HICUBE_AND_RE = re.compile(r"\bAND\b|\bHI-CUBE\b", re.IGNORECASE)
REEFER_TEMP_RE = re.compile(r"REEFER\s*[+-]?\d+\s*C", re.IGNORECASE)
HAZARDOUS_WORD_RE = re.compile(r"\bHAZARDOUS\b", re.IGNORECASE)


def simplify_remark(remark_raw, temp_c, dg_c):
    """REMARK (CNTRS) มักซ้ำกับสิ่งที่คอลัมน์ TEMP (CNTRS)/DG (CNTRS) แสดงอยู่แล้ว
    เข้าคู่กันเสมอ 3 แบบ: HAZARDOUS คู่ DG, HI-CUBE คู่ SIZE CNTR 45',
    HI-CUBE/REEFER คู่อุณหภูมิ - เพราะ temp_c/dg_c ก็ดึงมาจาก remark_raw
    ตัวเดียวกันนี้เองอยู่แล้ว โชว์ซ้ำอีกทีในช่อง REMARK จึงไม่มีประโยชน์ไม่ว่าค่านั้น
    จะตรงกับ MANIFEST หรือไม่ก็ตาม (ถ้าไม่ตรง คอลัมน์ TEMP/DG ทั้งสองฝั่งก็โชว์ค่า
    จริงพร้อมไอคอน ⚠ และหมายเหตุอธิบายอยู่แล้ว ไม่ต้องพึ่ง REMARK) จึงตัดทิ้งเสมอ
    ทีละส่วน: HI-CUBE/AND ตัดทิ้งเสมอเพราะไม่ใช่ข้อมูลเทียบกับ MANIFEST, REEFER+
    อุณหภูมิตัดถ้าดึง temp_c ได้, HAZARDOUS/CLASS/UN ตัดถ้าดึง dg_c ได้ เหลือแค่
    ส่วนที่ยัง "ตรวจสอบไม่ได้"/มีข้อมูลเพิ่มจริงๆ เท่านั้น (เช่น BY TRUCK/BY TRAIN,
    TANK, OVER WEIGHT, TRANSIT...) ถ้าตัดจนไม่เหลืออะไรเลยให้เป็นค่าว่าง (คืนค่า
    None แล้ว blankdash() จะใส่ '-' ให้เอง)"""
    if not remark_raw:
        return None
    core = re.sub(r"\s+", " ", remark_raw.strip())
    core = HICUBE_AND_RE.sub(" ", core)
    if temp_c:
        core = REEFER_TEMP_RE.sub(" ", core)
    if dg_c:
        core = HAZARDOUS_WORD_RE.sub(" ", core)
        core = CLASS_RE.sub(" ", core)
        core = UN_RE.sub(" ", core)
        core = re.sub(r"\b(UN|CLASS)\b", " ", core, flags=re.IGNORECASE)  # เศษคำที่เหลือค้าง
    core = re.sub(r"[\s,]+", " ", core).strip()
    return core or None


def vent_ok(mb, temp_m, temp_c):
    """ตู้เย็น (ประเภทตู้ xxRx เช่น 22R1/45R1 หรือมี TEMP/REEFER) ต้องมี VENT ใน
    MANIFEST เสมอ (เช่น VENT:15CBM/H, VENT:CLOSED). คืน False = ต้องแจ้งเตือน ⚠
    ตู้ที่ไม่ใช่ตู้เย็นไม่ต้องมี VENT จึงถือว่าผ่านเสมอ"""
    is_reefer = mb["is_reefer_type"] or bool(temp_m) or bool(temp_c)
    return (not is_reefer) or bool(mb["vent"])


# SHED ที่ล็อกไว้ผูกกับปลายทางอยู่แล้ว: ถ้า CNTRS ไม่ระบุ Port Of Delivery (หรือเป็น THLCH)
# ให้ดูปลายทางจากเลข SHED แทน
SHED_DELIVERY = {"0110": "THBMT", "0302": "THSCT", "0332": "THLKR"}


def effective_delivery(cc):
    d = (cc.get("delivery") or "").upper()
    if d in ("", "THLCH"):
        return SHED_DELIVERY.get(cc["shed_no_norm"], d or None)
    return d


def expected_shed(delivery, has_temp, has_dg):
    """SHED ที่ล็อกไว้ตามกฎ: DG=2826 เสมอ; ปลายทาง THBMT (BMT/ส่งทางเรือ barge)=0110;
    ปลายทาง THLKR (ลาดกระบัง) สินค้าทั่วไป=0332 (ถ้ามี TEMP/DG ต้องไปทางรถบรรทุก
    ไม่ผ่านโกดังลาดกระบัง จึงไม่บังคับ). คืน None = ไม่มีกฎบังคับ. (SCT=0302 ยังไม่มี
    เกณฑ์แยกสินค้า จึงไม่บังคับ)"""
    if has_dg:
        return "2826"
    if has_temp:
        return None
    d = (delivery or "").upper()
    if d == "THBMT":
        return "0110"
    if d == "THLKR":
        return "0332"
    return None


def remark_check(delivery, temp_present, dg_present, remark_raw):
    """REMARK ต้องบอกวิธีขนส่งตาม Port Of Delivery:
    BMT (THBMT) = BY BARGE ; SCT = BY TRUCK ;
    LKR (THLKR) = BY TRUCK ถ้าเป็นตู้เย็น/DG นอกนั้น BY TRAIN.
    ปลายทางอื่น (เช่น THLCH) ไม่มีกฎ -> คืน None"""
    d = (delivery or "").upper()
    if d == "THBMT":
        expected = "BY BARGE"
    elif d.endswith("SCT"):
        expected = "BY TRUCK"
    elif d == "THLKR":
        expected = "BY TRUCK" if (temp_present or dg_present) else "BY TRAIN"
    else:
        return None, None, None
    r = (remark_raw or "").upper()
    if "BARGE" in r:
        actual = "BY BARGE"
    elif "TRUCK" in r:
        actual = "BY TRUCK"
    elif "TRAN" in r or "TRAIN" in r:
        actual = "BY TRAIN"
    else:
        actual = "(ไม่ระบุ)"
    return expected, actual, (expected == actual)


def classify_inputs():
    """แยกไฟล์ใน input/ ตามเนื้อหา ไม่ใช่ตามชื่อ (เคยเจอไฟล์ CNTRS ถูกตั้งชื่อว่า
    'MANIFEST (1).xls'): มีคำว่า CARGO MANIFEST = MANIFEST,
    มี CONTAINER LIST FOR FEEDER = CNTRS"""
    manifests, cntrs = [], []
    for name in sorted(os.listdir(INPUT_DIR)):
        if not name.lower().endswith((".xls", ".xlsx", ".pdf")):
            continue
        path = f"{INPUT_DIR}/{name}"
        if name.lower().endswith(".pdf"):
            text = " ".join(pdf_lines(path)[:8]).upper()
        else:
            head = pd.ExcelFile(path).parse("Sheet1", header=None, nrows=8)
            text = " ".join(str(v).upper() for v in head.values.ravel() if pd.notna(v))
        if "CARGO MANIFEST" in text:
            manifests.append(path)
        elif "CONTAINER LIST FOR FEEDER" in text:
            cntrs.append(path)
    return manifests, cntrs


def main():
    manifest_files, cntrs_files = classify_inputs()
    if not manifest_files or not cntrs_files:
        raise SystemExit(
            f"ไม่พบไฟล์ครบใน {INPUT_DIR}/ - MANIFEST (CARGO MANIFEST): {len(manifest_files)} ไฟล์, "
            f"CNTRS (CONTAINER LIST FOR FEEDER): {len(cntrs_files)} ไฟล์")
    print("MANIFEST:", [os.path.basename(f) for f in manifest_files])
    print("CNTRS:", [os.path.basename(f) for f in cntrs_files])

    cntrs_containers = {}
    for f in cntrs_files:
        for cno, rec in parse_cntrs(f).items():
            cntrs_containers.setdefault(cno, rec)
    cntrs_header = parse_cntrs_header(cntrs_files[0])
    manifest_containers = {}
    for f in manifest_files:
        manifest_containers.update(parse_manifest(f))

    all_container_nos = sorted(
        set(manifest_containers) | set(cntrs_containers),
        key=lambda c: (cntrs_containers.get(c, {}).get("item") is None,
                        cntrs_containers.get(c, {}).get("item", 0)),
    )

    rows = []
    mismatch_count = 0
    for cno in all_container_nos:
        mb = manifest_containers.get(cno)
        cc = cntrs_containers.get(cno)

        if mb is None:
            rows.append({
                "critical": True,
                "item": cc["item"], "container_no": cno, "bl_no": "-", "consignee": cc["consignee"] or "-",
                "shed_m": "-", "shed_c": cc["shed_no"], "shed_ok": False,
                "status_m": "-", "status_c": norm_status(cc["status_raw"]), "status_ok": False,
                "temp_m": "-", "temp_c": extract_cntrs_temp(cc["remark_raw"]), "temp_ok": False,
                "dg_m": "-", "dg_c": extract_dg(cc["remark_raw"]), "dg_ok": False,
                "vent_ok": True,
                "remark_raw": cc["remark_raw"], "remark_expected": "-", "remark_ok": False,
                "note": "⚠ CONTAINER นี้มีใน CNTRS แต่ไม่พบใน MANIFEST",
            })
            mismatch_count += 1
            continue
        if cc is None:
            rows.append({
                "critical": True,
                "item": "-", "container_no": cno, "bl_no": mb["bl_no"], "consignee": mb["consignee"],
                "shed_m": mb["shed_no"], "shed_c": "-", "shed_ok": False,
                "status_m": norm_status(mb["status_raw"]), "status_c": "-", "status_ok": False,
                "temp_m": extract_manifest_temp(mb["temps"]), "temp_c": "-", "temp_ok": False,
                "dg_m": extract_dg(" ".join(mb["desc_all"])), "dg_c": "-", "dg_ok": False,
                "vent_ok": vent_ok(mb, extract_manifest_temp(mb["temps"]), None),
                "remark_raw": "-", "remark_expected": "-", "remark_ok": False,
                "note": "⚠ CONTAINER นี้มีใน MANIFEST แต่ไม่พบใน CNTRS",
            })
            mismatch_count += 1
            continue

        shed_m, shed_c = mb["shed_no"], cc["shed_no"]
        shed_ok = mb["shed_no_norm"] == cc["shed_no_norm"]

        status_m, status_c = norm_status(mb["status_raw"]), norm_status(cc["status_raw"])
        status_ok = status_m == status_c

        temp_m = extract_manifest_temp(mb["temps"])
        temp_c = extract_cntrs_temp(cc["remark_raw"])
        temp_ok = temp_m == temp_c

        mb_desc_blob = " ".join(mb["desc_all"])
        dg_m = extract_dg(mb_desc_blob)
        dg_c = extract_dg(cc["remark_raw"])
        dg_ok = dg_matches(mb_desc_blob, cc["remark_raw"])

        vent_good = vent_ok(mb, temp_m, temp_c)

        exp_shed = expected_shed(effective_delivery(cc), bool(temp_m or temp_c), bool(dg_m or dg_c))
        shed_rule_bad = bool(exp_shed) and cc["shed_no_norm"] != exp_shed
        shed_equal = shed_ok
        if shed_rule_bad:
            shed_ok = False

        remark_expected, remark_actual, remark_ok = remark_check(
            effective_delivery(cc), bool(temp_m or temp_c), bool(dg_m or dg_c), cc["remark_raw"]
        )
        remark_display = simplify_remark(cc["remark_raw"], temp_c, dg_c)
        if remark_ok and remark_display:
            # วิธีขนส่งถูกต้องตามกฎแล้ว ไม่ต้องโชว์ซ้ำ ตัด BY BARGE/TRUCK/TRAIN ออก
            remark_display = re.sub(r"\bBY\s+(BARGE|TRUCK|TRAIN|TRAN)\b", " ", remark_display, flags=re.IGNORECASE)
            remark_display = re.sub(r"\s+", " ", remark_display).strip() or None

        notes = []
        if not shed_equal:
            hint = SHED_RULES.get(cc["shed_no_norm"])
            hint_txt = f" (CNTRS จัดเก็บที่ {hint})" if hint else ""
            notes.append(f"SHED ไม่ตรง: MANIFEST={shed_m or '-'} / CNTRS={shed_c or '-'}{hint_txt}")
        if shed_rule_bad:
            notes.append(f"SHED ผิดกฎที่ล็อกไว้: ควรเป็น {exp_shed} แต่ CNTRS อยู่ SHED {cc['shed_no']}"
                         f" (ปลายทาง {cc.get('delivery') or '-'})")
        if not status_ok:
            notes.append(f"STATUS ไม่ตรง: MANIFEST={status_m or '-'} / CNTRS={status_c or '-'}")
        if not temp_ok:
            notes.append(f"TEMP ไม่ตรง: MANIFEST={temp_m or '-'} / CNTRS={temp_c or '-'}")
        if not dg_ok:
            notes.append(f"DG ไม่ตรง: MANIFEST={dg_m or '-'} / CNTRS={dg_c or '-'}")
        if not vent_good:
            notes.append("VENT ไม่ระบุใน MANIFEST")
        if remark_ok is False:
            notes.append(f"REMARK ไม่ตรงกฎ: ควรเป็น '{remark_expected}' แต่พบ '{remark_actual}'")

        row_mismatch = not (shed_ok and status_ok and temp_ok and dg_ok and vent_good
                            and (remark_ok is not False))
        if row_mismatch:
            mismatch_count += 1

        rows.append({
            "critical": False,
            "item": cc["item"], "container_no": cno, "bl_no": mb["bl_no"],
            "consignee": cc["consignee"] or mb["consignee"],
            "shed_m": shed_m, "shed_c": shed_c, "shed_ok": shed_ok,
            "status_m": status_m, "status_c": status_c, "status_ok": status_ok,
            "temp_m": temp_m, "temp_c": temp_c, "temp_ok": temp_ok,
            "dg_m": dg_m, "dg_c": dg_c, "dg_ok": dg_ok,
            "vent_ok": vent_good,
            "remark_raw": remark_display, "remark_expected": remark_expected, "remark_ok": remark_ok,
            "note": " | ".join(notes) if notes else "",
        })

    build_excel(rows, mismatch_count, len(all_container_nos), cntrs_header)
    print(f"เสร็จสิ้น: พบข้อมูลไม่ตรงกัน {mismatch_count} / {len(all_container_nos)} container")
    print(f"บันทึกรายงานที่ {OUTPUT_PATH}")


def build_excel(rows, mismatch_count, total, cntrs_header):
    wb = Workbook()
    ws = wb.active
    ws.title = "CNTRS_EDI"

    GREEN = PatternFill("solid", fgColor="C6EFCE")
    RED = PatternFill("solid", fgColor="FFC7CE")
    GREY = PatternFill("solid", fgColor="F2F2F2")
    HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
    HEADER_FONT = Font(bold=True, color="FFFFFF")
    GREEN_FONT = Font(bold=True, color="006100")
    RED_FONT = Font(bold=True, color="9C0006")
    thin = Side(style="thin", color="B7B7B7")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

    LAST_COL = 21  # A..U
    last_letter = get_column_letter(LAST_COL)
    # คอลัมน์รหัส/ค่าสั้นๆ (ITEM, CONTAINER NO., SHED, STATUS, TEMP, VENT) จัดกึ่งกลาง
    # ส่วนคอลัมน์ข้อความ (NO., B/L NO., CONSIGNEE, DG, REMARK, NOTE) ชิดซ้าย
    CENTER_COLS = {2, 3, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15}

    ws.merge_cells(f"A1:{last_letter}1")
    ws["A1"] = "CNTRS EDI - MANIFEST vs CNTRS Discrepancy Report (SHED NUMBER / STATUS / TEMP / DG / REMARK)"
    ws["A1"].font = Font(bold=True, size=11, color="000000")

    ws.merge_cells(f"A2:{last_letter}2")
    ws["A2"] = (f"FEEDER: {cntrs_header['feeder']}   |   VOYAGE: {cntrs_header['voyage']}   |   "
                f"ARRIVAL DATE: {cntrs_header['arrival_date']}   |   "
                f"PORT OF DISCHARGE: {cntrs_header['port_of_discharge']}")
    ws["A2"].font = Font(bold=True, size=11, color="000000")

    ws.merge_cells(f"A3:{last_letter}3")
    ws["A3"] = ("กฎ SHED: DG=2826, SCT=0302, BMT=0110, ลาดกระบัง=0332  |  "
                "กฎ REMARK ตาม Port Of Delivery: BMT=BY BARGE, SCT=BY TRUCK, LKR=BY TRUCK (ตู้เย็น/DG) นอกนั้น BY TRAIN  |  "
                "✓ สีเขียว = ตรงกัน, ⚠ สีแดง = ไม่ตรงกัน")
    ws["A3"].font = Font(bold=True, size=8, color="FF0000")

    headers = [
        "NO.", "ITEM", "CONTAINER NO.", "B/L NO.", "CONSIGNEE",
        "SHED (MANIFEST)", "SHED (CNTRS)", "SHED ✓/⚠",
        "STATUS (MANIFEST)", "STATUS (CNTRS)", "STATUS ✓/⚠",
        "TEMP (MANIFEST)", "TEMP (CNTRS)", "TEMP ✓/⚠",
        "VENT (MANIFEST)",
        "DG CLASS/UN (MANIFEST)", "DG CLASS/UN (CNTRS)", "DG ✓/⚠",
        "REMARK (CNTRS)", "REMARK ✓/⚠",
        "หมายเหตุ / NOTE",
    ]
    header_row = 5
    for j, h in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=j, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    # ล็อกคอลัมน์ A-C (ถึง CONTAINER NO.) ไว้ไม่ให้เลื่อนหายตอนสกอลไปทางขวา
    ws.freeze_panes = f"D{header_row + 1}"

    def blankdash(v):
        return "-" if v is None or v == "" else v

    def set_icon(cell, ok, na=False):
        if na:
            cell.value = "-"
            cell.fill = GREY
        elif ok:
            cell.value = "✓"
            cell.font = GREEN_FONT
            cell.fill = GREEN
        else:
            cell.value = "⚠"
            cell.font = RED_FONT
            cell.fill = RED

    def write_pair(m_cell, c_cell, m_val, c_val, ok, hide_when_ok=False):
        """เขียนคู่ MANIFEST/CNTRS - ถ้า hide_when_ok และค่าตรงกัน (✓) ไม่ต้องโชว์
        รายละเอียดซ้ำ (ไอคอน ✓ บอกอยู่แล้ว) ใส่ '-' แทนทั้งคู่; ถ้าไม่ตรง (⚠) ยังคง
        โชว์ค่าจริงทั้งสองฝั่งเพื่อดูว่าต่างกันตรงไหน พร้อมทำพื้นแดงฝั่ง MANIFEST"""
        if hide_when_ok and ok:
            m_cell.value = "-"
            c_cell.value = "-"
            return
        m_cell.value = blankdash(m_val)
        c_cell.value = blankdash(c_val)
        if not ok:
            m_cell.fill = RED
            m_cell.font = RED_FONT

    r = header_row + 1
    for i, row in enumerate(rows, start=1):
        ws.cell(row=r, column=1, value=i)
        ws.cell(row=r, column=2, value=blankdash(row["item"]))
        ws.cell(row=r, column=3, value=row["container_no"])
        ws.cell(row=r, column=4, value=blankdash(row["bl_no"]))
        ws.cell(row=r, column=5, value=blankdash(row["consignee"]))
        write_pair(ws.cell(row=r, column=6), ws.cell(row=r, column=7),
                   row["shed_m"], row["shed_c"], row["shed_ok"], hide_when_ok=True)
        set_icon(ws.cell(row=r, column=8), row["shed_ok"])
        write_pair(ws.cell(row=r, column=9), ws.cell(row=r, column=10),
                   row["status_m"], row["status_c"], row["status_ok"], hide_when_ok=True)
        set_icon(ws.cell(row=r, column=11), row["status_ok"])
        write_pair(ws.cell(row=r, column=12), ws.cell(row=r, column=13),
                   row["temp_m"], row["temp_c"], row["temp_ok"], hide_when_ok=True)
        set_icon(ws.cell(row=r, column=14), row["temp_ok"])
        # VENT: ใส่ถูกต้อง (หรือไม่ใช่ตู้เย็น) = "-" ; ตู้เย็นแต่ไม่มี VENT = ⚠ แดง
        vent_cell = ws.cell(row=r, column=15)
        if row["vent_ok"]:
            vent_cell.value = "-"
        else:
            vent_cell.value = "❗"
            vent_cell.font = Font(bold=True, color="9C0006", size=8)
            vent_cell.fill = RED
        write_pair(ws.cell(row=r, column=16), ws.cell(row=r, column=17),
                   row["dg_m"], row["dg_c"], row["dg_ok"], hide_when_ok=True)
        set_icon(ws.cell(row=r, column=18), row["dg_ok"])
        remark_cell = ws.cell(row=r, column=19, value=blankdash(row["remark_raw"]))
        if row["remark_ok"] is False:  # ไม่ตรงกฎ = ใส่สีแดง (ตรง = "-")
            remark_cell.fill = RED
            remark_cell.font = RED_FONT
        set_icon(ws.cell(row=r, column=20), row["remark_ok"], na=(row["remark_ok"] is None))
        note_cell = ws.cell(row=r, column=21, value=row["note"])
        note_cell.font = Font(bold=True, color="FF0000")  # หมายเหตุ/NOTE ตัวอักษรแดง

        if row["critical"]:
            for c in range(1, LAST_COL + 1):
                ws.cell(row=r, column=c).fill = PatternFill("solid", fgColor="FFE699")

        for c in range(1, LAST_COL + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            horiz = "center" if c in CENTER_COLS else "left"
            cell.alignment = Alignment(horizontal=horiz, vertical="center", wrap_text=(c == LAST_COL))
        r += 1

    # Summary block (label merged across A:D so it never overlaps/clips the value in E)
    def summary_label(row_r, text, bold=True, color="000000"):
        ws.merge_cells(start_row=row_r, start_column=1, end_row=row_r, end_column=4)
        cell = ws.cell(row=row_r, column=1, value=text)
        cell.font = Font(bold=bold, size=11, color=color)
        cell.alignment = Alignment(horizontal="left", vertical="center")
        return cell

    r += 1
    summary_label(r, "สรุปผล / SUMMARY", bold=True, color="000000")
    r += 1
    summary_label(r, "จำนวน CONTAINER ทั้งหมด:", bold=True, color="000000")
    ws.cell(row=r, column=5, value=total).font = Font(bold=True, size=11)
    r += 1
    summary_label(r, "จำนวนที่ไม่ตรงกัน (⚠):", bold=True, color="FF0000")
    c = ws.cell(row=r, column=5, value=mismatch_count)
    c.font = Font(bold=True, size=11, color=(RED_FONT.color.rgb if mismatch_count else GREEN_FONT.color.rgb))
    r += 1
    summary_label(r, "จำนวนที่ตรงกัน (✓):", bold=True, color="000000")
    ws.cell(row=r, column=5, value=total - mismatch_count).font = Font(bold=True, size=11)

    widths = [5, 6, 15, 20, 30, 15, 13, 9, 15, 15, 9, 13, 13, 9, 11, 20, 20, 9, 22, 9, 45]
    for j, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(j)].width = w

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    global OUTPUT_PATH
    try:
        wb.save(OUTPUT_PATH)
    except PermissionError:
        # ไฟล์เดิมเปิดค้างอยู่ใน Excel: บันทึกเป็นชื่อใหม่แทน ไม่ให้ผลหาย
        base, ext = os.path.splitext(OUTPUT_PATH)
        n = 2
        while os.path.exists(f"{base}_{n}{ext}"):
            try:
                os.remove(f"{base}_{n}{ext}")
                break
            except PermissionError:
                n += 1
        OUTPUT_PATH = f"{base}_{n}{ext}"
        wb.save(OUTPUT_PATH)


if __name__ == "__main__":
    main()
