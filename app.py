"""
Quarterly Results Review — Presentation Generator
Flask backend: loads KPI data from all XLSX files, builds PPTX from the template.
"""
import os
import io
import re
import copy
import json
import base64
import shutil
import tempfile
import zipfile
from pathlib import Path
from datetime import datetime, date
from urllib import parse, request as urlrequest

import openpyxl
import pytesseract
from PIL import Image
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from lxml import etree

# ─── Paths ───────────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).parent
TEMPLATE_PATH = BASE_DIR / "Quarterly Results Review Template_8.18.pptx"
AIRTABLE_TOKEN = os.environ.get("AIRTABLE_TOKEN", "")
AIRTABLE_BASE_ID = os.environ.get("AIRTABLE_BASE_ID", "")
AIRTABLE_TABLE_ID = os.environ.get("AIRTABLE_TABLE_ID", "")
AIRTABLE_BRIEF_FIELD = "🔷EMB #"
AIRTABLE_EXPECTED_RESPONSES_FIELD = "🔹Expected Responses"
AIRTABLE_DATE_FIELDS = ["🔷Start Date"]


def get_xlsx_files():
    return sorted(
        path.name
        for path in BASE_DIR.glob("*.xlsx")
        if path.name != TEMPLATE_PATH.name and not path.name.startswith("~$")
    )

# ─── KPI Data Loader ──────────────────────────────────────────────────────────
def _find_header_row(rows):
    """Find the row index (0-based) that contains 'Event name' and 'Brief'."""
    for i, row in enumerate(rows):
        row_str = [str(c).strip().lower() if c else "" for c in row]
        if "event name" in row_str and "brief" in row_str:
            return i
    return None


def _col_index(headers, *candidates):
    """Return the first matching column index from a list of candidate names (exact or prefix match)."""
    for name in candidates:
        name_l = name.lower()
        # 1st pass: exact match (case-insensitive)
        for i, h in enumerate(headers):
            if h and str(h).strip().lower() == name_l:
                return i
        # 2nd pass: header starts with the candidate (handles "CQ (Q1 26)" etc.)
        for i, h in enumerate(headers):
            if h and str(h).strip().lower().startswith(name_l):
                return i
        # 3rd pass: candidate is contained in header (e.g. "Total SQL")
        for i, h in enumerate(headers):
            if h and name_l in str(h).strip().lower():
                return i
    return None


def _col_index_total(headers):
    """
    Resolve the 'Total' column carefully — must not pick up 'Total SQLR'
    (a separate metric in Data files) when a plain 'Total'/'Totals'/'Total SQL'
    column also exists.
    Preference order: exact 'totals', 'total', then 'total sql', then anything
    starting with 'total' that is NOT 'total sqlr'.
    """
    for exact in ("totals", "total", "total sql"):
        for i, h in enumerate(headers):
            if h and str(h).strip().lower() == exact:
                return i
    for i, h in enumerate(headers):
        h_l = str(h).strip().lower() if h else ""
        if h_l.startswith("total") and h_l != "total sqlr" and "sqlr" not in h_l:
            return i
    return None



def _parse_airtable_date(value):
    if not value:
        return None
    if isinstance(value, list):
        for item in value:
            parsed = _parse_airtable_date(item)
            if parsed:
                return parsed
        return None
    text = str(value).strip()
    if not text:
        return None
    text = text.replace("Z", "+00:00")
    for parser in (
        lambda v: datetime.fromisoformat(v).date(),
        lambda v: datetime.strptime(v, "%Y-%m-%d").date(),
        lambda v: datetime.strptime(v, "%m/%d/%Y").date(),
        lambda v: datetime.strptime(v, "%m/%d/%y").date(),
    ):
        try:
            return parser(text)
        except ValueError:
            pass
    return None



def _quarter_bucket_for_event(event_date):
    if not event_date:
        return None
    months = (date.today().year - event_date.year) * 12 + (date.today().month - event_date.month)
    quarter_diff = months // 3
    if quarter_diff <= 0:
        return "CQ"
    if quarter_diff == 1:
        return "NQ"
    if quarter_diff == 2:
        return "NQ+1"
    return "NQ+2"



def load_airtable_events():
    if not AIRTABLE_TOKEN:
        return {}

    base_id = parse.quote(AIRTABLE_BASE_ID, safe="")
    table_id = parse.quote(AIRTABLE_TABLE_ID, safe="")
    url = f"https://api.airtable.com/v0/{base_id}/{table_id}?pageSize=100"
    headers = {"Authorization": f"Bearer {AIRTABLE_TOKEN}"}
    events = {}

    while url:
        req = urlrequest.Request(url, headers=headers)
        with urlrequest.urlopen(req) as resp:
            payload = json.loads(resp.read().decode("utf-8"))

        for record in payload.get("records", []):
            fields = record.get("fields", {})
            brief = str(fields.get(AIRTABLE_BRIEF_FIELD, "")).strip()
            if not brief:
                continue
            event_date = None
            for field_name in AIRTABLE_DATE_FIELDS:
                event_date = _parse_airtable_date(fields.get(field_name))
                if event_date:
                    break
            events[brief.upper()] = {
                "airtable_record_id": record.get("id"),
                "event_date": event_date.isoformat() if event_date else "",
                "expected_individuals": fields.get(AIRTABLE_EXPECTED_RESPONSES_FIELD),
                "expected_sqor_bucket": _quarter_bucket_for_event(event_date),
            }

        offset = payload.get("offset")
        url = f"https://api.airtable.com/v0/{base_id}/{table_id}?pageSize=100&offset={parse.quote(offset)}" if offset else None
    return events



def _extract_metrics_from_image(image_bytes):
    if not shutil.which("tesseract"):
        return {
            "actual_individuals": None,
            "actual_sqor": None,
        }

    text = pytesseract.image_to_string(Image.open(io.BytesIO(image_bytes)))
    compact_text = text.replace(",", "")

    attendees = None
    attendees_match = re.search(r"Attendees\D{0,20}(\d+)", compact_text, re.IGNORECASE)
    if attendees_match:
        attendees = int(attendees_match.group(1))

    sqo_creation = None
    sqo_match = re.search(r"SQO\s*Creation\D{0,20}\$?([\d,]+(?:\.\d+)?)", text, re.IGNORECASE)
    if sqo_match:
        sqo_creation = float(sqo_match.group(1).replace(",", ""))

    return {
        "actual_individuals": attendees,
        "actual_sqor": sqo_creation,
    }



def _find_table_by_header(slide, header_text):
    for shape in slide.shapes:
        if not getattr(shape, "has_table", False):
            continue
        table = shape.table
        for row in table.rows:
            for cell in row.cells:
                if header_text.lower() in cell.text.lower():
                    return table
    return None



def _set_table_cell_text(cell, text):
    cell.text = text
    for paragraph in cell.text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.RIGHT



def _fill_summary_tables(prs, event_data):
    slide = prs.slides[1]
    expected_individuals = event_data.get("expected_individuals")
    actual_individuals = event_data.get("actual_individuals")
    expected_sqor_bucket = event_data.get("expected_sqor_bucket") or ""
    actual_sqor = event_data.get("actual_sqor")
    epm_pull_date = datetime.today().strftime("%-m/%-d/%y")

    for shape in slide.shapes:
        if hasattr(shape, "text") and "[Date of EPM pull]" in shape.text:
            shape.text = shape.text.replace("[Date of EPM pull]", f"Pulled {epm_pull_date}")

    attendees_table = _find_table_by_header(slide, "Expected Individuals")
    if attendees_table and len(attendees_table.rows) > 1:
        exp_text = str(int(expected_individuals)) if isinstance(expected_individuals, (int, float)) else (str(expected_individuals) if expected_individuals is not None else "")
        act_text = str(actual_individuals) if actual_individuals is not None else ""
        delta_val = None
        if isinstance(expected_individuals, (int, float)) and actual_individuals is not None:
            delta_val = actual_individuals - int(expected_individuals)
        delta_text = f"{delta_val:+,}" if delta_val is not None else ""
        _set_table_cell_text(attendees_table.cell(1, 0), exp_text)
        _set_table_cell_text(attendees_table.cell(1, 1), act_text)
        _set_table_cell_text(attendees_table.cell(1, 2), delta_text)

    sqor_table = _find_table_by_header(slide, "Expected SQOR by end of QX")
    if sqor_table and len(sqor_table.rows) > 1:
        sqor_table.cell(0, 0).text = f"Expected SQOR by end of {expected_sqor_bucket}" if expected_sqor_bucket else "Expected SQOR by end of QX"
        sqor_table.cell(0, 1).text = f"Actual SQOR as of {epm_pull_date}"
        value = None
        bucket_key = {
            "CQ": "cq",
            "NQ": "nq",
            "NQ+1": "nq1",
            "NQ+2": "nq2",
        }.get(expected_sqor_bucket)
        if bucket_key:
            value = sum((row.get(bucket_key) or 0) for row in event_data.get("kpi_rows", []))
        delta_value = actual_sqor - value if actual_sqor is not None and value is not None else None
        _set_table_cell_text(sqor_table.cell(1, 0), _fmt_currency(value) if value is not None else "")
        _set_table_cell_text(sqor_table.cell(1, 1), _fmt_currency(actual_sqor) if actual_sqor is not None else "")
        _set_table_cell_text(sqor_table.cell(1, 2), _fmt_currency(delta_value) if delta_value is not None else "")

    return prs



def load_all_events():
    """
    Returns a dict:  { brief_code: { event data } }
    and a list of unique event dicts for the search index.
    """
    events = {}  # keyed by brief
    airtable_events = load_airtable_events()

    for filename in get_xlsx_files():
        path = BASE_DIR / filename
        if not path.exists():
            continue
        try:
            wb = openpyxl.load_workbook(path, data_only=True)
        except Exception:
            continue

        for sheet_name in wb.sheetnames:
            if "option" in sheet_name.lower() or sheet_name.lower() == "sheet1":
                continue
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(values_only=True))
            header_row_idx = _find_header_row(rows)
            if header_row_idx is None:
                continue

            headers = [str(c).strip() if c else "" for c in rows[header_row_idx]]

            # Resolve CQ/NQ/NQ+1/NQ+2 precisely:
            # Headers like "CQ (Q1 26)", "NQ (Q2 26)", "NQ+1 (Q3 26)" must map correctly.
            cq_idx = nq_idx = nq1_idx = nq2_idx = None
            for i, h in enumerate(headers):
                h_u = str(h).strip().upper() if h else ""
                h_clean = h_u.split("(")[0].strip()  # strip any "(Q1 26)" suffix
                if h_clean == "CQ":
                    cq_idx = i
                elif h_clean == "NQ+2":
                    nq2_idx = i
                elif h_clean == "NQ+1":
                    nq1_idx = i
                elif h_clean == "NQ":
                    nq_idx = i

            icol = {
                "name": _col_index(headers, "event name"),
                "brief": _col_index(headers, "brief"),
                "ut15": _col_index(headers, "ut 15"),
                "ut17": _col_index(headers, "ut 17"),
                "cq": cq_idx,
                "nq": nq_idx,
                "nq1": nq1_idx,
                "nq2": nq2_idx,
                "total": _col_index_total(headers),
            }

            if icol["name"] is None or icol["brief"] is None:
                continue

            # Determine quarter label from sheet name
            quarter = "Q2"
            for q in ["Q2", "Q3", "Q4", "Q1"]:
                if q in sheet_name.upper():
                    quarter = q
                    break

            current_name = None
            current_brief = None

            for row in rows[header_row_idx + 1:]:
                name_val = row[icol["name"]] if icol["name"] is not None else None
                brief_val = row[icol["brief"]] if icol["brief"] is not None else None

                # Carry forward name/brief for continuation rows
                if name_val:
                    current_name = str(name_val).strip()
                if brief_val:
                    current_brief = str(brief_val).strip()

                if not current_brief:
                    continue

                def get_num(col_key):
                    if icol.get(col_key) is None:
                        return None
                    v = row[icol[col_key]]
                    try:
                        return float(v) if v is not None else None
                    except (TypeError, ValueError):
                        return None

                kpi_row = {
                    "event_name": current_name or "",
                    "brief": current_brief,
                    "ut15": str(row[icol["ut15"]]).strip() if icol.get("ut15") is not None and row[icol["ut15"]] else "",
                    "ut17": str(row[icol["ut17"]]).strip() if icol.get("ut17") is not None and row[icol["ut17"]] else "",
                    "cq": get_num("cq"),
                    "nq": get_num("nq"),
                    "nq1": get_num("nq1"),
                    "nq2": get_num("nq2"),
                    "total": get_num("total"),
                    "quarter": quarter,
                    "source_file": filename,
                }

                # Only keep rows that have at least one KPI number
                if all(kpi_row[k] is None for k in ["cq", "nq", "nq1", "nq2", "total"]):
                    continue

                key = current_brief.upper()
                if key not in events:
                    airtable_data = airtable_events.get(key, {})
                    events[key] = {
                        "event_name": current_name or "",
                        "brief": current_brief,
                        "kpi_rows": [],
                        **airtable_data,
                    }
                events[key]["kpi_rows"].append(kpi_row)
                # Update the name in case it was set on a later row
                if current_name:
                    events[key]["event_name"] = current_name
                if key in airtable_events:
                    events[key].update(airtable_events[key])

    return events


# ─── PPTX Generator ──────────────────────────────────────────────────────────
NSMAP = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

def _fmt_currency(val):
    if val is None:
        return "—"
    try:
        v = float(val)
        return f"${v:,.0f}"
    except (TypeError, ValueError):
        return str(val)


def _make_kpi_table_xml(kpi_rows):
    """
    Build the KPI table XML matching the style in slide 4 of the template.
    Returns an lxml element <a:tbl>.
    """
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"

    def tag(name):
        return f"{{{A}}}{name}"

    def text_run(text, sz=1000, bold=False, color=None):
        rpr = etree.SubElement(etree.Element("dummy"), tag("rPr"))
        rpr.set("lang", "en-US")
        rpr.set("sz", str(sz))
        if bold:
            rpr.set("b", "1")
        rpr.set("dirty", "0")
        if color:
            sf = etree.SubElement(rpr, tag("solidFill"))
            clr = etree.SubElement(sf, tag("srgbClr"))
            clr.set("val", color)
        else:
            sf = etree.SubElement(rpr, tag("solidFill"))
            sc = etree.SubElement(sf, tag("schemeClr"))
            sc.set("val", "tx1")
        r = etree.Element(tag("r"))
        r.append(rpr)
        t = etree.SubElement(r, tag("t"))
        t.text = text
        return r

    def make_cell(text, sz=1000, bold=False, color=None, bg=None, align="l"):
        tc = etree.Element(tag("tc"))
        txBody = etree.SubElement(tc, tag("txBody"))
        etree.SubElement(txBody, tag("bodyPr"))
        etree.SubElement(txBody, tag("lstStyle"))
        p = etree.SubElement(txBody, tag("p"))
        pPr = etree.SubElement(p, tag("pPr"))
        pPr.set("algn", align)
        buNone = etree.SubElement(pPr, tag("buNone"))
        p.append(text_run(text, sz=sz, bold=bold, color=color))
        tcPr = etree.SubElement(tc, tag("tcPr"))
        tcPr.set("marL", "45720")
        tcPr.set("marR", "45720")
        tcPr.set("marT", "45720")
        tcPr.set("marB", "45720")
        if bg:
            sf = etree.SubElement(tcPr, tag("solidFill"))
            clr = etree.SubElement(sf, tag("srgbClr"))
            clr.set("val", bg)
        return tc

    def make_row(*cells, height=228600):
        tr = etree.Element(tag("tr"))
        tr.set("h", str(height))
        for cell in cells:
            tr.append(cell)
        return tr

    # Column widths (EMU) — 9 columns matching the template
    col_widths = [1909454, 747387, 945083, 945083, 700776, 700776, 700776, 700776, 777926]

    tbl = etree.Element(tag("tbl"))

    tblPr = etree.SubElement(tbl, tag("tblPr"))
    styleId = etree.SubElement(tblPr, tag("tableStyleId"))
    styleId.text = "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}"

    tblGrid = etree.SubElement(tbl, tag("tblGrid"))
    for w in col_widths:
        gc = etree.SubElement(tblGrid, tag("gridCol"))
        gc.set("w", str(w))

    # Header row 0: "KPI Projections" spanning all cols (visual — first cell bold, rest blank)
    title_cells = [make_cell("KPI Projections", sz=1400, bold=True, align="l")]
    for _ in range(8):
        title_cells.append(make_cell(" ", sz=1400, align="l"))
    tbl.append(make_row(*title_cells, height=261388))

    # Header row 1: column labels
    labels = ["Event name", "Brief", "UT 15", "UT 17", "CQ", "NQ", "NQ+1", "NQ+2", "Total"]
    label_cells = [make_cell(lbl, sz=900, bold=True, color="FFFFFF", bg="161616", align="l") for lbl in labels]
    tbl.append(make_row(*label_cells, height=228600))

    # Data rows
    for kpi in kpi_rows:
        total_val = kpi.get("total")
        if total_val is None:
            cq = kpi.get("cq") or 0
            nq = kpi.get("nq") or 0
            nq1 = kpi.get("nq1") or 0
            nq2 = kpi.get("nq2") or 0
            total_val = cq + nq + nq1 + nq2

        cells = [
            make_cell(kpi.get("event_name", ""), sz=900, align="l"),
            make_cell(kpi.get("brief", ""), sz=900, align="l"),
            make_cell(kpi.get("ut15", ""), sz=900, align="l"),
            make_cell(kpi.get("ut17", ""), sz=900, align="l"),
            make_cell(_fmt_currency(kpi.get("cq")), sz=900, align="r"),
            make_cell(_fmt_currency(kpi.get("nq")), sz=900, align="r"),
            make_cell(_fmt_currency(kpi.get("nq1")), sz=900, align="r"),
            make_cell(_fmt_currency(kpi.get("nq2")), sz=900, align="r"),
            make_cell(_fmt_currency(total_val), sz=900, bold=True, align="r"),
        ]
        tbl.append(make_row(*cells, height=228600))

    return tbl


def _inject_kpi_table_into_slide(slide_xml_bytes, kpi_rows):
    """
    Replace the '[Insert table of KPI projects]' text box on the slide with
    a proper KPI table, using the same position/size as the placeholder box.
    Returns new slide XML bytes.
    """
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    P = "http://schemas.openxmlformats.org/presentationml/2006/main"

    root = etree.fromstring(slide_xml_bytes)
    spTree = root.find(".//{%s}spTree" % P)
    if spTree is None:
        spTree = root.find(".//{%s}cSld/{%s}spTree" % (P, P))

    # Find the text box that contains the placeholder
    placeholder_sp = None
    for sp in root.iter("{%s}sp" % P):
        texts = "".join(
            t.text or "" for t in sp.iter("{%s}t" % A)
        )
        if "[Insert table of" in texts or "Insert table of" in texts:
            placeholder_sp = sp
            break

    if placeholder_sp is None:
        return slide_xml_bytes  # nothing to replace

    # Get position/size from the placeholder shape
    xfrm = placeholder_sp.find(".//{%s}xfrm" % A)
    off = xfrm.find("{%s}off" % A) if xfrm is not None else None
    ext = xfrm.find("{%s}ext" % A) if xfrm is not None else None

    x = int(off.get("x", 0)) if off is not None else 3926798
    y = int(off.get("y", 0)) if off is not None else 5411449
    cx = int(ext.get("cx", 0)) if ext is not None else 8128036
    cy = int(ext.get("cy", 0)) if ext is not None else 653470

    # Build the table height based on number of rows
    num_data_rows = len(kpi_rows)
    row_height = 228600
    header_height = 261388 + 228600
    total_height = header_height + num_data_rows * row_height
    cy = max(cy, total_height)

    # Build graphic frame XML for the table
    gf_xml = f"""<p:graphicFrame
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
      <p:nvGraphicFramePr>
        <p:cNvPr id="99" name="KPI Table"/>
        <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
        <p:nvPr/>
      </p:nvGraphicFramePr>
      <p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></p:xfrm>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
        </a:graphicData>
      </a:graphic>
    </p:graphicFrame>"""

    gf_el = etree.fromstring(gf_xml)

    # Insert the table element into graphicData
    graphic_data = gf_el.find(".//{%s}graphicData" % A)
    tbl_el = _make_kpi_table_xml(kpi_rows)
    graphic_data.append(tbl_el)

    # Replace placeholder sp with the graphic frame
    parent = placeholder_sp.getparent()
    idx = list(parent).index(placeholder_sp)
    parent.remove(placeholder_sp)
    parent.insert(idx, gf_el)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _inject_epm_image_into_slide(pptx_bytes, slide_idx, image_bytes, image_ext="png"):
    """
    Insert the EPM screenshot image into the slide, replacing the
    '[Paste screenshot of EPM dashboard]' placeholder area.
    Returns new pptx bytes.
    """
    from pptx import Presentation
    from pptx.util import Emu
    from io import BytesIO

    prs = Presentation(BytesIO(pptx_bytes))
    slide = prs.slides[slide_idx]

    # Find the EPM placeholder text box position
    epm_placeholder = None
    for shape in slide.shapes:
        if hasattr(shape, "text") and "[Paste screenshot of EPM dashboard]" in shape.text:
            epm_placeholder = shape
            break

    if epm_placeholder:
        # Use the position of an existing image placeholder or the dashed rectangle
        # Find the dashed rectangle (the big area left of center)
        left = epm_placeholder.left
        top = epm_placeholder.top
        # Find the big image area — look for a picture or rectangle with large width
        for shape in slide.shapes:
            if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE
                left = shape.left
                top = shape.top
                width = shape.width
                height = shape.height
                # Replace that picture's image
                from pptx.oxml.ns import qn
                blipFill = shape._element.find(qn("p:blipFill"))
                if blipFill is not None:
                    blip = blipFill.find(qn("a:blip"))
                    if blip is not None:
                        rId = blip.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                        if rId:
                            from pptx.opc.constants import RELATIONSHIP_TYPE as RT
                            part = slide.part
                            img_part, new_rid = part.get_or_add_image_part(BytesIO(image_bytes))
                            blip.set("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed", new_rid)
                            out = BytesIO()
                            prs.save(out)
                            return out.getvalue()

        # If no picture found, add new image in the dashed rectangle area
        # Find the dashed-border rectangle
        for shape in slide.shapes:
            if hasattr(shape, "line") and shape.shape_type == 1:  # AUTO_SHAPE
                try:
                    if shape.line.dash_style is not None:
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height
                        from io import BytesIO as _BIO
                        pic = slide.shapes.add_picture(_BIO(image_bytes), left, top, width, height)
                        out = BytesIO()
                        prs.save(out)
                        return out.getvalue()
                except Exception:
                    pass

        # Fallback: add image at a reasonable position
        slide.shapes.add_picture(
            BytesIO(image_bytes),
            Emu(4041735), Emu(1366293),
            Emu(8042877), Emu(3188528)
        )

    out = BytesIO()
    prs.save(out)
    return out.getvalue()


def _replace_text_in_xml(xml_bytes, replacements):
    """
    Simple byte-level text replacement in slide XML.
    replacements: list of (old_str, new_str)
    """
    content = xml_bytes.decode("utf-8")
    for old, new in replacements:
        content = content.replace(old, new)
    return content.encode("utf-8")


def generate_presentation(event_data, epm_image_bytes=None):
    """
    Work on the template PPTX as a zip, modify slide 2 in-memory,
    and return the resulting PPTX bytes.
    """
    from io import BytesIO

    event_name = event_data.get("event_name", "[Event Name]")
    brief = event_data.get("brief", "[Brief ID]")
    kpi_rows = event_data.get("kpi_rows", [])

    # ── Read template as zip ──────────────────────────────────────────────────
    with open(TEMPLATE_PATH, "rb") as f:
        template_bytes = f.read()

    out_buf = BytesIO()
    with zipfile.ZipFile(BytesIO(template_bytes), "r") as zin, \
         zipfile.ZipFile(out_buf, "w", zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename == "ppt/slides/slide2.xml":
                # ── Replace title text ────────────────────────────────────────
                data = _replace_text_in_xml(data, [
                    ("[Event Name]", event_name),
                    ("[Brief ID]", brief),
                    (", [Event Date]", ""),
                    ("[Event Date]", ""),
                ])

                # ── Replace KPI placeholder with table ────────────────────────
                if kpi_rows:
                    data = _inject_kpi_table_into_slide(data, kpi_rows)

            zout.writestr(item, data)

    pptx_bytes = out_buf.getvalue()

    prs = Presentation(BytesIO(pptx_bytes))
    prs = _fill_summary_tables(prs, event_data)
    out = BytesIO()
    prs.save(out)
    pptx_bytes = out.getvalue()

    # ── Inject EPM image if provided ──────────────────────────────────────────
    if epm_image_bytes:
        pptx_bytes = _inject_epm_image_into_slide(pptx_bytes, 1, epm_image_bytes)

    return pptx_bytes


# ─── Flask App ───────────────────────────────────────────────────────────────
app = Flask(__name__, static_folder="static", template_folder=".")
CORS(app)

_events_cache = None

def get_events():
    global _events_cache
    if _events_cache is None:
        _events_cache = load_all_events()
    return _events_cache


@app.route("/")
def index():
    with open(BASE_DIR / "index.html", encoding="utf-8") as f:
        return f.read()


@app.route("/api/events")
def api_events():
    """Return a flat list of all events for the search dropdown."""
    events = get_events()
    result = []
    seen = set()
    for key, ev in events.items():
        uid = ev["brief"].upper()
        if uid in seen:
            continue
        seen.add(uid)
        result.append({
            "event_name": ev["event_name"],
            "brief": ev["brief"],
        })
    result.sort(key=lambda x: x["event_name"])
    return jsonify(result)


@app.route("/api/event/<brief>")
def api_event_detail(brief):
    """Return full KPI detail for a specific event."""
    events = get_events()
    key = brief.strip().upper()
    ev = events.get(key)
    if not ev:
        return jsonify({"error": "Event not found"}), 404
    return jsonify(ev)


@app.route("/api/generate", methods=["POST"])
def api_generate():
    """Generate the PPTX presentation for the given event."""
    brief = request.form.get("brief", "").strip()
    epm_file = request.files.get("epm_image")

    if not brief:
        return jsonify({"error": "brief is required"}), 400

    events = get_events()
    key = brief.upper()
    ev = events.get(key)
    if not ev:
        return jsonify({"error": f"Event '{brief}' not found in KPI data"}), 404

    epm_bytes = None
    if epm_file:
        epm_bytes = epm_file.read()

    event_payload = copy.deepcopy(ev)
    if epm_bytes:
        event_payload.update(_extract_metrics_from_image(epm_bytes))

    try:
        pptx_bytes = generate_presentation(event_payload, epm_image_bytes=epm_bytes)
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

    safe_name = re.sub(r"[^\w\-]", "_", ev["event_name"])[:60]
    filename = f"{safe_name}_{brief}_Results.pptx"

    return send_file(
        io.BytesIO(pptx_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name=filename,
    )


@app.route("/api/reload")
def api_reload():
    """Force reload of KPI data from disk."""
    global _events_cache
    _events_cache = None
    events = get_events()
    return jsonify({"status": "ok", "event_count": len(events)})


if __name__ == "__main__":
    xlsx_files = get_xlsx_files()
    print(f"Loading KPI data from {len(xlsx_files)} spreadsheets...")
    ev = get_events()
    print(f"  → {len(ev)} unique events loaded.")
    app.run(debug=True, port=5050)
