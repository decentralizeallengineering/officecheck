# Build brief: officecheck

Build a Python library "officecheck": deterministic checks that grade whether an office file meets a goal, so an automated test harness needs no model to judge. Common rules: Python 3.12, works on Windows and Linux. Allowed dependencies only: openpyxl, python-docx, docxcompose, python-pptx, pypdf, pdfplumber, fpdf2, pandas, numpy, jsonschema. No network, no telemetry. Pytest tests generate their own fixtures. Apache-2.0.

Input: a file path plus a spec (JSON, validated with jsonschema). Output: {"pass": bool, "score": 0..1, "checks": [{"id", "pass", "expected", "actual"}]}.

Spec check types:
- docx: has_heading(text|regex, level), heading_count>=, has_toc (TOC field present and pre-filled entries match headings), has_table(min_rows, min_cols, header contains), contains_text, not_contains_text, style_used, page_orientation, header_contains, footer_has_page_number.
- xlsx: sheet_exists, cell_equals (value, or formula result via a small evaluator for SUM/AVERAGE/COUNT/IF/VLOOKUP/arithmetic), cell_has_formula(pattern), column_exists(header), row_count, chart_exists(type, series count, title), conditional_format_on(range), data_validation_on(range), no_error_values (#REF! #DIV/0! #VALUE! in cached values), sorted_by(column, asc|desc), no_duplicates(key columns), freeze_panes.
- pptx: slide_count, title_on_slide(n, text), layout_on_slide, has_notes, has_table, has_chart, has_image, text_fits (a heuristic: characters per placeholder area).
- pdf: page_count, contains_text(page|any), form_field_equals, is_text_searchable, page_size, bookmarks_present, table_on_page(min_rows).
- any: file_opens, size_under, masked (none of a given list of strings appears anywhere, including headers, footers, comments, notes and form values).

A CLI: `officecheck <file> <spec.json>` prints a table and exits 0 or 1.
Tests: a passing and a failing fixture for every check type.
