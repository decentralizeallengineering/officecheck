"""Edge cases, robustness and the result contract."""

from __future__ import annotations

import json
from xml.sax.saxutils import escape

import jsonschema
import pytest

from officecheck import SpecError, grade, load_spec, result_schema

SECRET = "Jane Q. Public"


def one(path, check):
    return grade(path, {"checks": [check]})["checks"][0]


# ---- masked: every hiding place the brief names --------------------------------------------------

@pytest.mark.parametrize("fixture, location", [
    ("secret_docx_header", "word/header"),
    ("secret_docx_footer", "word/footer"),          # split across runs with different formatting
    ("secret_docx_comment", "word/comments.xml"),
    ("secret_pptx_notes", "ppt/notesSlides/"),        # different case, ignore_case defaults to true
    ("secret_xlsx_comment", "xl/comments"),
    ("secret_pdf_form", "form field applicant"),
    ("secret_pdf_box", "page 1"),                     # fake redaction: text under a black box
])
def test_masked_finds_secret(files, fixture, location):
    r = one(files[fixture], {"type": "any.masked", "strings": [SECRET, "never-present"]})
    assert not r["pass"]
    (hit,) = r["actual"]["found"]
    assert hit["string"] == SECRET
    assert any(loc.startswith(location) for loc in hit["locations"]), hit


def test_masked_case_sensitive_option(files):
    r = one(files["secret_pptx_notes"], {"type": "any.masked", "strings": [SECRET], "ignore_case": False})
    assert r["pass"]


# ---- docx ------------------------------------------------------------------------------------

def test_docx_without_toc_field(files):
    r = one(files["docx_no_toc"], {"type": "docx.has_toc"})
    assert not r["pass"] and r["actual"]["toc_field"] is False


def test_docx_toc_presence_only(files):
    assert one(files["docx_bad"], {"type": "docx.has_toc", "check_entries": False})["pass"]


def test_docx_heading_regex_and_title_level(files):
    assert one(files["docx_good"], {"type": "docx.has_heading", "regex": "^Res", "level": 1})["pass"]
    assert one(files["docx_good"], {"type": "docx.has_heading", "text": "Quarterly Report", "level": 0})["pass"]
    assert not one(files["docx_good"], {"type": "docx.has_heading", "text": "Results", "level": 2})["pass"]


def test_docx_text_scope_all_includes_headers(files):
    chk = {"type": "docx.contains_text", "text": "ACME Confidential"}
    assert not one(files["docx_good"], chk)["pass"]
    assert one(files["docx_good"], {**chk, "scope": "all"})["pass"]


# ---- xlsx: evaluator through cell_equals ----------------------------------------------------------

@pytest.mark.parametrize("cell, value", [
    ("Summary!B1", 18), ("Summary!B2", 1.65), ("Summary!B3", 5), ("Summary!B4", "high"),
    ("Summary!B5", "Cherry"), ("Summary!B6", 14), ("Summary!B7", 70), ("Data!E3", 1.25),
])
def test_cell_equals_evaluated(files, cell, value):
    r = one(files["xlsx_good"], {"type": "xlsx.cell_equals", "cell": cell, "value": value})
    assert r["pass"], r
    assert r["actual"]["source"] == "evaluated"


def test_cell_equals_cached_vs_evaluated(files):
    # B8 carries a (deliberately stale) cached value 30.75; the formula actually gives 23.25.
    base = {"type": "xlsx.cell_equals", "cell": "Summary!B8"}
    r = one(files["xlsx_good"], {**base, "value": 30.75})
    assert r["pass"] and r["actual"]["source"] == "cached"
    assert one(files["xlsx_good"], {**base, "value": 23.25, "source": "evaluate"})["pass"]
    assert not one(files["xlsx_good"], {**base, "value": 23.25})["pass"]


def test_cell_equals_literal_and_sheet_param(files):
    assert one(files["xlsx_good"], {"type": "xlsx.cell_equals", "sheet": "Data", "cell": "B2", "value": "Apple"})["pass"]
    assert not one(files["xlsx_good"], {"type": "xlsx.cell_equals", "sheet": "Data", "cell": "B2", "value": "apple"})["pass"]
    assert one(files["xlsx_good"], {"type": "xlsx.cell_equals", "sheet": "Data", "cell": "B2", "value": "apple",
                                    "ignore_case": True})["pass"]


def test_cell_equals_error_value(files):
    r = one(files["xlsx_bad"], {"type": "xlsx.cell_equals", "cell": "Data!D3", "value": "#DIV/0!"})
    assert r["pass"], r


def test_no_error_values_evaluate_uncached(tmp_path):
    import openpyxl

    wb = openpyxl.Workbook()
    wb.active["A1"] = 0
    wb.active["A2"] = "=1/A1"
    p = tmp_path / "e.xlsx"
    wb.save(p)
    assert one(p, {"type": "xlsx.no_error_values"})["pass"]  # nothing cached -> nothing to see
    r = one(p, {"type": "xlsx.no_error_values", "evaluate_uncached": True})
    assert not r["pass"] and r["actual"]["errors"] == ["Sheet!A2=#DIV/0! (evaluated)"]


def test_unsupported_formula_fails_cleanly(tmp_path):
    import openpyxl

    wb = openpyxl.Workbook()
    wb.active["A1"] = "=XLOOKUP(1,B1:B2,C1:C2)"
    p = tmp_path / "u.xlsx"
    wb.save(p)
    r = one(p, {"type": "xlsx.cell_equals", "cell": "A1", "value": 1})
    assert not r["pass"] and "not supported" in r["actual"]["error"]


def test_range_coverage_modes(files):
    chk = {"type": "xlsx.conditional_format_on", "sheet": "Data", "range": "C2:C10"}
    assert not one(files["xlsx_good"], chk)["pass"]
    assert one(files["xlsx_good"], {**chk, "mode": "intersects"})["pass"]
    assert one(files["xlsx_good"], {**chk, "range": "C3:C4", "rule_type": "cellIs"})["pass"]
    assert not one(files["xlsx_good"], {**chk, "range": "C3:C4", "rule_type": "colorScale"})["pass"]


def test_sorted_desc_and_freeze_cell(files):
    assert not one(files["xlsx_good"], {"type": "xlsx.sorted_by", "sheet": "Data", "column": "ID", "order": "desc"})["pass"]
    assert one(files["xlsx_good"], {"type": "xlsx.sorted_by", "sheet": "Data", "column_letter": "A"})["pass"]
    assert one(files["xlsx_good"], {"type": "xlsx.freeze_panes", "sheet": "Data", "cell": "A2"})["pass"]
    assert one(files["xlsx_bad"], {"type": "xlsx.freeze_panes", "sheet": "Data", "cell": None})["pass"]


def test_chart_constraints(files):
    base = {"type": "xlsx.chart_exists"}
    assert one(files["xlsx_good"], {**base, "chart_type": "column"})["pass"]
    assert not one(files["xlsx_good"], {**base, "chart_type": "line"})["pass"]
    assert not one(files["xlsx_good"], {**base, "series": 3})["pass"]
    assert one(files["xlsx_good"], {**base, "title_regex": "^Sales"})["pass"]


# ---- pptx / pdf ------------------------------------------------------------------------------

def test_pptx_slide_out_of_range_is_a_clean_failure(files):
    r = one(files["pptx_bad"], {"type": "pptx.title_on_slide", "slide": 5, "text": "x"})
    assert not r["pass"] and "no slide 5" in r["actual"]["error"]


def test_pdf_checkbox_and_dimensions(files):
    assert one(files["pdf_good"], {"type": "pdf.form_field_equals", "name": "agree", "value": True})["pass"]
    assert one(files["pdf_bad"], {"type": "pdf.form_field_equals", "name": "agree", "value": False})["pass"]
    assert one(files["pdf_good"], {"type": "pdf.page_size", "width": 612, "height": 792})["pass"]
    assert one(files["pdf_bad"], {"type": "pdf.page_size", "size": "A4"})["pass"]  # named size: either orientation
    assert not one(files["pdf_bad"], {"type": "pdf.page_size", "size": "A4", "orientation": "portrait"})["pass"]
    assert one(files["pdf_good"], {"type": "pdf.contains_text", "regex": r"Appendix [A-Z]"})["pass"]


# ---- wrong inputs never crash -----------------------------------------------------------------

def test_wrong_format_is_a_failed_check(files):
    r = one(files["pdf_good"], {"type": "docx.has_heading", "text": "x"})
    assert not r["pass"] and r["actual"] == {"error": "file is not a docx (detected: pdf)"}


def test_missing_file(tmp_path):
    res = grade(tmp_path / "nope.docx", {"checks": [{"type": "any.file_opens"}, {"type": "docx.has_toc"}]})
    assert res == {**res, "pass": False, "score": 0.0}
    assert all("error" in c["actual"] for c in res["checks"])


def test_corrupt_zip_reported(files):
    r = one(files["corrupt_docx"], {"type": "any.file_opens"})
    assert "corrupt" in r["actual"]["error"]


# ---- spec and result contract -------------------------------------------------------------------

@pytest.mark.parametrize("spec, message", [
    ({}, "'checks' is a required property"),
    ({"checks": []}, "should be non-empty"),
    ({"checks": [{"type": "docx.bogus"}]}, "unknown check type 'docx.bogus'"),
    ({"checks": [{"type": "docx.has_heading"}]}, "needs exactly one of: text | regex"),
    ({"checks": [{"type": "docx.has_heading", "text": "a", "lvl": 1}]}, "Additional properties"),
    ({"checks": [{"type": "xlsx.row_count"}]}, "needs at least one of"),
    ({"checks": [{"type": "any.file_opens", "id": "a"}, {"type": "any.size_under", "id": "a", "max_bytes": 1}]},
     "duplicate id"),
])
def test_invalid_specs(spec, message):
    with pytest.raises(SpecError, match=message.replace("|", r"\|").replace("(", r"\(")):
        load_spec(spec)


def test_result_shape_score_weights_threshold(files):
    spec = {
        "threshold": 0.7,
        "checks": [
            {"id": "toc", "type": "docx.has_toc", "weight": 3},
            {"type": "docx.heading_count", "min": 99},
        ],
    }
    res = grade(files["docx_good"], spec)
    jsonschema.validate(res, result_schema())
    assert [c["id"] for c in res["checks"]] == ["toc", "2:docx.heading_count"]
    assert res["score"] == 0.75 and res["pass"] is True
    res = grade(files["docx_good"], {**spec, "threshold": 1})
    assert res["pass"] is False


def test_deterministic_and_json_serialisable(files):
    spec = {"checks": [{"type": t} for t in ("any.file_opens", "docx.has_toc", "docx.footer_has_page_number")]}
    a = json.dumps(grade(files["docx_good"], spec), sort_keys=True)
    b = json.dumps(grade(files["docx_good"], spec), sort_keys=True)
    assert a == b


def test_spec_from_path_and_string(files, tmp_path):
    spec = {"checks": [{"type": "pdf.page_count", "equals": 2}]}
    p = tmp_path / "s.json"
    p.write_text(json.dumps(spec), encoding="utf-8")
    assert grade(files["pdf_good"], p)["pass"]
    assert grade(files["pdf_good"], str(p))["pass"]
    assert grade(files["pdf_good"], json.dumps(spec))["pass"]


# ---- a TOC exactly as Word 365 writes it ---------------------------------------------------------

_W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'


def _word_toc_entry(title: str, page: str, n: int, first: bool = False, last: bool = False) -> str:
    begin = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
             '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
             '<w:r><w:fldChar w:fldCharType="separate"/></w:r>') if first else ""
    end = '<w:r><w:fldChar w:fldCharType="end"/></w:r>' if last else ""
    return (
        f'<w:p {_W}><w:pPr><w:pStyle w:val="TOC1"/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9350"/>'
        f'</w:tabs></w:pPr>{begin}<w:hyperlink w:anchor="_Toc{n}" w:history="1">'
        f'<w:r><w:t>{escape(title)}</w:t></w:r><w:r><w:tab/></w:r>'
        f'<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r><w:instrText xml:space="preserve"> PAGEREF _Toc{n} \\h </w:instrText></w:r>'
        f'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{page}</w:t></w:r>'
        f'<w:r><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink>{end}</w:p>'
    )


def _make_word_toc_doc(path, toc_titles, headings):
    import docx
    from docx.oxml import parse_xml

    doc = docx.Document()
    body = doc.element.body
    sdt = parse_xml(
        f'<w:sdt {_W}><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>'
        f'<w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtContent>'
        f'<w:p><w:pPr><w:pStyle w:val="TOCHeading"/></w:pPr><w:r><w:t>Contents</w:t></w:r></w:p>'
        + "".join(_word_toc_entry(t, str(i + 1), i, i == 0, i == len(toc_titles) - 1)
                  for i, t in enumerate(toc_titles))
        + "</w:sdtContent></w:sdt>"
    )
    body.insert(0, sdt)
    for level, text in headings:
        doc.add_heading(text, level=level)
    doc.save(path)
    return path


def test_word_style_toc_in_sdt(tmp_path):
    heads = [(1, "Scope"), (2, "Method & data"), (3, "Sensors"), (4, "Too deep for TOC"), (1, "Findings")]
    good = _make_word_toc_doc(tmp_path / "w.docx", ["Scope", "Method & data", "Sensors", "Findings"], heads)
    r = one(good, {"type": "docx.has_toc"})
    assert r["pass"], r
    assert r["actual"]["entries"] == ["Scope", "Method & data", "Sensors", "Findings"]
    stale = _make_word_toc_doc(tmp_path / "s.docx", ["Scope", "Findings"], heads)
    r = one(stale, {"type": "docx.has_toc"})
    assert not r["pass"] and r["actual"]["entries"] == ["Scope", "Findings"]


def test_footer_page_number_as_simple_field(tmp_path):
    import docx
    from docx.oxml import parse_xml

    doc = docx.Document()
    fp = doc.sections[0].footer.paragraphs[0]
    fp._p.append(parse_xml(f'<w:fldSimple {_W} w:instr=" PAGE   \\* MERGEFORMAT "><w:r><w:t>1</w:t></w:r></w:fldSimple>'))
    doc.save(tmp_path / "f.docx")
    assert one(tmp_path / "f.docx", {"type": "docx.footer_has_page_number"})["pass"]


def test_text_fits_uses_inherited_placeholder_size(tmp_path):
    import pptx

    prs = pptx.Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[5])
    # ~100 chars at the master's 44pt title size needs 2+ lines; the title box holds ~1.5 lines
    s.shapes.title.text = "A deliberately long slide title that keeps going well past one line of text"
    prs.save(tmp_path / "t.pptx")
    r = one(tmp_path / "t.pptx", {"type": "pptx.text_fits"})
    assert not r["pass"] and r["actual"]["overflowing"][0]["shape"] == "Title 1"
    assert one(tmp_path / "t.pptx", {"type": "pptx.text_fits", "max_fill": 2.0})["pass"]
