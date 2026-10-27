"""A passing and a failing fixture for every check type."""

from __future__ import annotations

import pytest

from officecheck import REGISTRY, grade, spec_schema

SECRET = "Jane Q. Public"

# (check, passing fixture, failing fixture)
CASES = [
    # docx
    ({"type": "docx.has_heading", "text": "Introduction", "level": 1}, "docx_good", "docx_bad"),
    ({"type": "docx.heading_count", "min": 3}, "docx_good", "docx_bad"),
    ({"type": "docx.has_toc"}, "docx_good", "docx_bad"),
    ({"type": "docx.has_table", "min_rows": 3, "min_cols": 3, "header_contains": ["Name", "Price"]},
     "docx_good", "docx_bad"),
    ({"type": "docx.contains_text", "text": "Quarterly  REVENUE", "ignore_case": True}, "docx_good", "docx_bad"),
    ({"type": "docx.not_contains_text", "text": "DRAFT"}, "docx_good", "docx_bad"),
    ({"type": "docx.style_used", "style": "List Bullet"}, "docx_good", "docx_bad"),
    ({"type": "docx.page_orientation", "orientation": "portrait"}, "docx_good", "docx_bad"),
    ({"type": "docx.header_contains", "text": "Confidential"}, "docx_good", "docx_bad"),
    ({"type": "docx.footer_has_page_number"}, "docx_good", "docx_bad"),
    # xlsx
    ({"type": "xlsx.sheet_exists", "sheet": "Lookup"}, "xlsx_good", "xlsx_bad"),
    ({"type": "xlsx.cell_equals", "cell": "Summary!B1", "value": 18}, "xlsx_good", "xlsx_bad"),
    ({"type": "xlsx.cell_has_formula", "cell": "Summary!B3", "pattern": r"^=COUNT\("}, "xlsx_good", "xlsx_bad"),
    ({"type": "xlsx.column_exists", "sheet": "Data", "header": "Total"}, "xlsx_good", "xlsx_bad"),
    ({"type": "xlsx.row_count", "sheet": "Data", "min": 5}, "xlsx_good", "xlsx_bad"),
    ({"type": "xlsx.chart_exists", "chart_type": "bar", "series": 2, "title": "Sales by fruit"},
     "xlsx_good", "xlsx_bad"),
    ({"type": "xlsx.conditional_format_on", "sheet": "Data", "range": "C2:C6"}, "xlsx_good", "xlsx_bad"),
    ({"type": "xlsx.data_validation_on", "sheet": "Data", "range": "C3:C5", "validation_type": "whole"},
     "xlsx_good", "xlsx_bad"),
    ({"type": "xlsx.no_error_values"}, "xlsx_good", "xlsx_bad"),
    ({"type": "xlsx.sorted_by", "sheet": "Data", "column": "ID", "order": "asc"}, "xlsx_good", "xlsx_bad"),
    ({"type": "xlsx.no_duplicates", "sheet": "Data", "columns": ["ID"]}, "xlsx_good", "xlsx_bad"),
    ({"type": "xlsx.freeze_panes", "sheet": "Data", "rows": 1}, "xlsx_good", "xlsx_bad"),
    # pptx
    ({"type": "pptx.slide_count", "equals": 3}, "pptx_good", "pptx_bad"),
    ({"type": "pptx.title_on_slide", "slide": 1, "text": "Quarterly Review"}, "pptx_good", "pptx_bad"),
    ({"type": "pptx.layout_on_slide", "slide": 1, "layout": "Title Slide"}, "pptx_good", "pptx_bad"),
    ({"type": "pptx.has_notes"}, "pptx_good", "pptx_bad"),
    ({"type": "pptx.has_table", "min_rows": 3, "min_cols": 3}, "pptx_good", "pptx_bad"),
    ({"type": "pptx.has_chart", "chart_type": "column", "series": 2, "title": "Revenue"}, "pptx_good", "pptx_bad"),
    ({"type": "pptx.has_image"}, "pptx_good", "pptx_bad"),
    ({"type": "pptx.text_fits"}, "pptx_good", "pptx_bad"),
    # pdf
    ({"type": "pdf.page_count", "equals": 2}, "pdf_good", "pdf_bad"),
    ({"type": "pdf.contains_text", "text": "Annual Report 2026", "page": 1}, "pdf_good", "pdf_bad"),
    ({"type": "pdf.form_field_equals", "name": "applicant", "value": "Jane Doe"}, "pdf_good", "pdf_bad"),
    ({"type": "pdf.is_text_searchable"}, "pdf_good", "pdf_bad"),
    ({"type": "pdf.page_size", "size": "Letter", "orientation": "portrait"}, "pdf_good", "pdf_bad"),
    ({"type": "pdf.bookmarks_present", "titles": ["Annual Report", "Appendix"]}, "pdf_good", "pdf_bad"),
    ({"type": "pdf.table_on_page", "page": 1, "min_rows": 4, "min_cols": 2}, "pdf_good", "pdf_bad"),
    # any
    ({"type": "any.file_opens"}, "docx_good", "corrupt_docx"),
    ({"type": "any.size_under", "max_bytes": 1_000_000}, "docx_good", "big_blob"),
    ({"type": "any.masked", "strings": [SECRET]}, "masked_clean", "secret_docx_header"),
]


def _ids(case):
    return case[0]["type"]


def test_every_check_type_has_a_case():
    spec_schema()  # populate registry
    assert sorted({c[0]["type"] for c in CASES}) == sorted(REGISTRY)


@pytest.mark.parametrize("case", CASES, ids=_ids)
def test_passing_fixture(case, files):
    chk, good, _bad = case
    res = grade(files[good], {"checks": [chk]})
    assert res["pass"], res
    assert res["score"] == 1.0


@pytest.mark.parametrize("case", CASES, ids=_ids)
def test_failing_fixture(case, files):
    chk, _good, bad = case
    res = grade(files[bad], {"checks": [chk]})
    assert not res["pass"], res
    assert res["score"] == 0.0
    actual = res["checks"][0]["actual"]
    # it must fail on the merits, not because the check crashed (file_opens is the exception)
    if chk["type"] != "any.file_opens":
        assert not (isinstance(actual, dict) and "error" in actual), actual
