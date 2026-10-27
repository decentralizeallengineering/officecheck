"""Builders for every test fixture. Nothing is checked in; every file is generated here."""

from __future__ import annotations

import io
import re
import struct
import zipfile
import zlib
from pathlib import Path
from typing import Any

# ---- tiny helpers ----------------------------------------------------------------------------


def png_bytes(w: int = 4, h: int = 4, rgb: tuple[int, int, int] = (200, 30, 30)) -> bytes:
    """A valid, deterministic RGB PNG without any imaging library."""
    def chunk(tag: bytes, data: bytes) -> bytes:
        return struct.pack(">I", len(data)) + tag + data + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF)

    raw = b"".join(b"\x00" + bytes(rgb) * w for _ in range(h))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw)) + chunk(b"IEND", b""))


def inject_cached_values(path: Path, values: dict[tuple[int, str], Any]) -> None:
    """Write cached results into formula cells of an openpyxl-saved xlsx.

    ``values`` maps (sheet index starting at 1, "B2") -> value. Strings starting with '#'
    are written as error values (t="e").
    """
    with zipfile.ZipFile(path) as zf:
        items = {n: zf.read(n) for n in zf.namelist()}
    for (sheet_no, ref), val in values.items():
        name = f"xl/worksheets/sheet{sheet_no}.xml"
        xml = items[name].decode()
        if isinstance(val, bool):
            t, v = ' t="b"', "1" if val else "0"
        elif isinstance(val, str) and val.startswith("#"):
            t, v = ' t="e"', val
        elif isinstance(val, str):
            t, v = ' t="str"', val
        else:
            t, v = "", repr(val)
        pat = re.compile(rf'<c r="{ref}"([^>]*)><f>(.*?)</f><v\s*/?>(?:</v>)?</c>')
        xml, n = pat.subn(lambda m: f'<c r="{ref}"{m.group(1)}{t}><f>{m.group(2)}</f><v>{v}</v></c>', xml)
        assert n == 1, f"formula cell {ref} not found in {name}"
        items[name] = xml.encode()
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for n, data in items.items():
            zf.writestr(n, data)


# ---- docx --------------------------------------------------------------------------------------


def _fld(paragraph: Any, instr: str, result: str | None = None) -> None:
    """Append a complex field {instr} with an optional pre-rendered result to a paragraph."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    def run_with(child: Any) -> None:
        r = OxmlElement("w:r")
        r.append(child)
        paragraph._p.append(r)

    def fldchar(kind: str) -> Any:
        fc = OxmlElement("w:fldChar")
        fc.set(qn("w:fldCharType"), kind)
        return fc

    run_with(fldchar("begin"))
    it = OxmlElement("w:instrText")
    it.set(qn("xml:space"), "preserve")
    it.text = f" {instr} "
    run_with(it)
    run_with(fldchar("separate"))
    if result is not None:
        t = OxmlElement("w:t")
        t.text = result
        run_with(t)
    run_with(fldchar("end"))


def _fld_char(paragraph: Any, kind: str, instr: str | None = None) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    r = OxmlElement("w:r")
    fc = OxmlElement("w:fldChar")
    fc.set(qn("w:fldCharType"), kind)
    r.append(fc)
    paragraph._p.append(r)
    if instr is not None:
        r2 = OxmlElement("w:r")
        it = OxmlElement("w:instrText")
        it.set(qn("xml:space"), "preserve")
        it.text = f" {instr} "
        r2.append(it)
        paragraph._p.append(r2)


def _toc(doc: Any, entries: list[tuple[str, str]], pageref_last: bool = True) -> None:
    """A Word-style TOC field whose cached result lists ``entries`` (title, page)."""
    first = doc.add_paragraph(style=None)
    _fld_char(first, "begin", 'TOC \\o "1-3" \\h \\z \\u')
    _fld_char(first, "separate")
    paras = [first] + [doc.add_paragraph() for _ in entries[1:]]
    for i, (para, (title, page)) in enumerate(zip(paras, entries)):
        para.add_run(title)
        para.add_run().add_tab()
        if pageref_last and i == len(entries) - 1:
            _fld(para, f"PAGEREF _Toc{i} \\h", page)
        else:
            para.add_run(page)
    end = doc.add_paragraph()
    _fld_char(end, "end")


def make_docx_good(path: Path) -> Path:
    import docx

    doc = docx.Document()
    doc.add_heading("Quarterly Report", level=0)
    _toc(doc, [("Introduction", "1"), ("Background", "1"), ("Results", "2")])
    doc.add_heading("Introduction", level=1)
    doc.add_paragraph("The quarterly revenue grew by 12 percent.")
    doc.add_heading("Background", level=2)
    doc.add_paragraph("Point one", style="List Bullet")
    doc.add_heading("Results", level=1)
    t = doc.add_table(rows=3, cols=3)
    for c, h in zip(t.rows[0].cells, ["Name", "Qty", "Price"]):
        c.text = h
    for r in t.rows[1:]:
        for c in r.cells:
            c.text = "1"
    sec = doc.sections[0]
    sec.header.paragraphs[0].text = "ACME Confidential"
    fp = sec.footer.paragraphs[0]
    fp.add_run("Page ")
    _fld(fp, "PAGE", "1")
    fp.add_run(" of ")
    _fld(fp, "NUMPAGES", "2")
    doc.save(path)
    return path


def make_docx_bad(path: Path, *, toc: bool = True) -> Path:
    import docx
    from docx.enum.section import WD_ORIENT

    doc = docx.Document()
    if toc:  # stale TOC: entries no longer match the headings
        _toc(doc, [("Introduction", "1"), ("Results", "2")])
    doc.add_heading("Summary", level=1)
    doc.add_paragraph("Introduction")  # looks like a heading but is Normal text
    doc.add_paragraph("DRAFT - numbers pending")
    t = doc.add_table(rows=1, cols=2)
    t.rows[0].cells[0].text, t.rows[0].cells[1].text = "A", "B"
    sec = doc.sections[0]
    sec.orientation = WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    sec.header.paragraphs[0].text = "Internal"
    fp = sec.footer.paragraphs[0]
    fp.add_run("Total pages: ")
    _fld(fp, "NUMPAGES", "2")
    doc.save(path)
    return path


# ---- xlsx --------------------------------------------------------------------------------------

DATA_ROWS = [
    (1, "Apple", 3, 0.5),
    (2, "Banana", 5, 0.25),
    (3, "Cherry", 2, 4.0),
    (4, "Date", 7, 1.5),
    (5, "Elder", 1, 2.0),
]


def make_xlsx_good(path: Path) -> Path:
    import openpyxl
    from openpyxl.chart import BarChart, Reference
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.styles import PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["ID", "Name", "Qty", "Price", "Total"])
    for i, row in enumerate(DATA_ROWS, start=2):
        ws.append([*row, f"=C{i}*D{i}"])
    ws.freeze_panes = "A2"
    ws.conditional_formatting.add("C2:C6", CellIsRule(operator="greaterThan", formula=["3"],
                                                      fill=PatternFill("solid", start_color="FFFF00")))
    dv = DataValidation(type="whole", operator="between", formula1="0", formula2="100")
    dv.add("C2:C6")
    ws.add_data_validation(dv)
    chart = BarChart()
    chart.type = "col"
    chart.title = "Sales by fruit"
    chart.add_data(Reference(ws, min_col=3, max_col=4, min_row=1, max_row=6), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=2, min_row=2, max_row=6))
    ws.add_chart(chart, "H2")

    s = wb.create_sheet("Summary")
    s["A1"], s["B1"] = "Total qty", "=SUM(Data!C2:C6)"          # 18
    s["A2"], s["B2"] = "Average price", "=AVERAGE(Data!D2:D6)"  # 1.65
    s["A3"], s["B3"] = "Count", "=COUNT(Data!A2:A6)"            # 5
    s["A4"], s["B4"] = "Level", '=IF(B1>10,"high","low")'       # high
    s["A5"], s["B5"] = "Lookup", "=VLOOKUP(3,Data!A2:E6,2,FALSE)"  # Cherry
    s["A6"], s["B6"] = "Arithmetic", "=(B1-3)*2^2/4+-1"        # 14
    s["A7"], s["B7"] = "Cross", "=Lookup!B2*10"                  # 70
    s["A8"], s["B8"] = "Cached", "=SUM(Data!E2:E6)"             # cached below
    lk = wb.create_sheet("Lookup")
    lk.append(["Key", "Val"])
    lk.append(["x", 7])
    wb.save(path)
    inject_cached_values(path, {(2, "B8"): 30.75})
    return path


def make_xlsx_bad(path: Path) -> Path:
    import openpyxl

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(["ID", "Name", "Qty", "Price"])
    ws.append([2, "Banana", 5, 0.25])
    ws.append([1, "Apple", 3, "=1/0"])
    ws.append([2, "Banana", 5, 0.25])  # duplicate ID
    s = wb.create_sheet("Summary")
    s["A1"], s["B1"] = "Total qty", "=SUM(Data!C2:C3)"  # 8, not 18
    s["A3"], s["B3"] = "Count", 5                        # literal, no formula
    wb.save(path)
    inject_cached_values(path, {(1, "D3"): "#DIV/0!"})
    return path


# ---- pptx --------------------------------------------------------------------------------------


def make_pptx_good(path: Path) -> Path:
    import pptx
    from pptx.chart.data import CategoryChartData
    from pptx.enum.chart import XL_CHART_TYPE
    from pptx.util import Inches

    prs = pptx.Presentation()
    s1 = prs.slides.add_slide(prs.slide_layouts[0])  # Title Slide
    s1.shapes.title.text = "Quarterly Review"
    s1.placeholders[1].text = "Q3 2026"
    s1.notes_slide.notes_text_frame.text = "Welcome everyone."

    s2 = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only
    s2.shapes.title.text = "Numbers"
    tbl = s2.shapes.add_table(3, 3, Inches(1), Inches(2), Inches(6), Inches(2)).table
    tbl.cell(0, 0).text = "Region"
    s2.notes_slide.notes_text_frame.text = "Walk through the table."

    s3 = prs.slides.add_slide(prs.slide_layouts[5])
    s3.shapes.title.text = "Trend"
    cd = CategoryChartData()
    cd.categories = ["Q1", "Q2", "Q3"]
    cd.add_series("2025", (1, 2, 3))
    cd.add_series("2026", (2, 3, 4))
    gf = s3.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(2), Inches(5), Inches(4), cd)
    gf.chart.has_title = True
    gf.chart.chart_title.text_frame.text = "Revenue"
    s3.shapes.add_picture(io.BytesIO(png_bytes()), Inches(7), Inches(2), Inches(1), Inches(1))
    s3.notes_slide.notes_text_frame.text = "Growth continues."
    prs.save(path)
    return path


def make_pptx_bad(path: Path) -> Path:
    import pptx
    from pptx.util import Inches, Pt

    prs = pptx.Presentation()
    s1 = prs.slides.add_slide(prs.slide_layouts[5])  # Title Only, not Title Slide
    s1.shapes.title.text = "Draft"
    box = s1.shapes.add_textbox(Inches(1), Inches(2), Inches(2), Inches(0.5))
    box.text_frame.word_wrap = True
    p = box.text_frame.paragraphs[0]
    p.text = "This sentence is far too long to fit in a tiny text box at this size. " * 4
    p.runs[0].font.size = Pt(24)
    prs.save(path)
    return path


# ---- pdf ---------------------------------------------------------------------------------------


def _add_form_fields(pdf_bytes: bytes, fields: dict[str, Any]) -> bytes:
    """Add AcroForm fields (text for str, checkbox for bool) to page 1 with pypdf."""
    import pypdf
    from pypdf.generic import (ArrayObject, DictionaryObject, FloatObject, NameObject, NumberObject,
                               TextStringObject)

    w = pypdf.PdfWriter(clone_from=pypdf.PdfReader(io.BytesIO(pdf_bytes)))
    refs = []
    for i, (name, value) in enumerate(fields.items()):
        d = DictionaryObject({
            NameObject("/Type"): NameObject("/Annot"),
            NameObject("/Subtype"): NameObject("/Widget"),
            NameObject("/T"): TextStringObject(name),
            NameObject("/F"): NumberObject(4),
            NameObject("/Rect"): ArrayObject([FloatObject(50), FloatObject(40 + 30 * i),
                                              FloatObject(250), FloatObject(60 + 30 * i)]),
        })
        if isinstance(value, bool):
            d[NameObject("/FT")] = NameObject("/Btn")
            d[NameObject("/V")] = NameObject("/Yes" if value else "/Off")
        else:
            d[NameObject("/FT")] = NameObject("/Tx")
            d[NameObject("/V")] = TextStringObject(value)
        refs.append(w._add_object(d))
    page = w.pages[0]
    page[NameObject("/Annots")] = ArrayObject(refs)
    w._root_object[NameObject("/AcroForm")] = DictionaryObject({NameObject("/Fields"): ArrayObject(refs)})
    out = io.BytesIO()
    w.write(out)
    return out.getvalue()


def make_pdf_good(path: Path) -> Path:
    from fpdf import FPDF

    pdf = FPDF(format="Letter")
    pdf.set_creation_date(__import__("datetime").datetime(2026, 1, 1))
    pdf.set_font("Helvetica", size=12)
    pdf.add_page()
    pdf.start_section("Annual Report")
    pdf.cell(0, 10, "Annual Report 2026", new_x="LMARGIN", new_y="NEXT")
    pdf.start_section("Figures", level=1)
    with pdf.table() as table:
        for row in [("Region", "Sales"), ("North", "10"), ("South", "12"), ("West", "9")]:
            r = table.row()
            for c in row:
                r.cell(c)
    pdf.add_page()
    pdf.start_section("Appendix")
    pdf.cell(0, 10, "Appendix A: method notes", new_x="LMARGIN", new_y="NEXT")
    path.write_bytes(_add_form_fields(bytes(pdf.output()), {"applicant": "Jane Doe", "agree": True}))
    return path


def make_pdf_bad(path: Path) -> Path:
    from fpdf import FPDF

    pdf = FPDF(orientation="L", format="A4")
    pdf.set_creation_date(__import__("datetime").datetime(2026, 1, 1))
    pdf.add_page()
    pdf.rect(20, 20, 100, 60, style="F")  # graphics only: no extractable text
    path.write_bytes(_add_form_fields(bytes(pdf.output()), {"applicant": "John", "agree": False}))
    return path


# ---- masking fixtures (one secret hidden in each place the brief names) -------------------------

SECRET = "Jane Q. Public"


def make_masked_clean_docx(path: Path) -> Path:
    import docx

    doc = docx.Document()
    doc.add_paragraph("Customer: [REDACTED]")
    doc.sections[0].header.paragraphs[0].text = "Case file"
    doc.save(path)
    return path


def make_secret_docx_header(path: Path) -> Path:
    import docx

    doc = docx.Document()
    doc.add_paragraph("Customer: [REDACTED]")
    doc.sections[0].header.paragraphs[0].text = f"Prepared for {SECRET}"
    doc.save(path)
    return path


def make_secret_docx_footer_split_runs(path: Path) -> Path:
    import docx

    doc = docx.Document()
    doc.add_paragraph("Customer: [REDACTED]")
    fp = doc.sections[0].footer.paragraphs[0]
    fp.add_run("Jane Q.")  # secret split across runs with different formatting
    fp.add_run(" Pub").bold = True
    fp.add_run("lic")
    doc.save(path)
    return path


def make_secret_docx_comment(path: Path) -> Path:
    import docx

    doc = docx.Document()
    p = doc.add_paragraph("Customer: [REDACTED]")
    doc.add_comment(p.runs, text=f"Original name was {SECRET}", author="Reviewer")
    doc.save(path)
    return path


def make_secret_pptx_notes(path: Path) -> Path:
    import pptx

    prs = pptx.Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Case summary"
    s.notes_slide.notes_text_frame.text = f"Remember to call {SECRET.upper()}"  # case differs
    prs.save(path)
    return path


def make_secret_xlsx_comment(path: Path) -> Path:
    import openpyxl
    from openpyxl.comments import Comment

    wb = openpyxl.Workbook()
    wb.active["A1"] = "[REDACTED]"
    wb.active["A1"].comment = Comment(f"was {SECRET}", "auditor")
    wb.save(path)
    return path


def make_secret_pdf_form(path: Path) -> Path:
    from fpdf import FPDF

    pdf = FPDF()
    pdf.set_creation_date(__import__("datetime").datetime(2026, 1, 1))
    pdf.set_font("Helvetica", size=12)
    pdf.add_page()
    pdf.cell(0, 10, "Applicant: [REDACTED]")
    path.write_bytes(_add_form_fields(bytes(pdf.output()), {"applicant": SECRET}))
    return path


def make_secret_pdf_under_box(path: Path) -> Path:
    """A fake redaction: the text is still there, just covered by a black rectangle."""
    from fpdf import FPDF

    pdf = FPDF()
    pdf.set_creation_date(__import__("datetime").datetime(2026, 1, 1))
    pdf.set_font("Helvetica", size=12)
    pdf.add_page()
    pdf.text(20, 30, f"Applicant: {SECRET}")
    pdf.rect(40, 24, 60, 8, style="F")
    path.write_bytes(bytes(pdf.output()))
    return path
