# officecheck

Deterministic grading of office outputs. Give it a file (`.docx`, `.xlsx`, `.pptx`, `.pdf`)
and a JSON spec; it returns pass/fail per check, so an automated test harness never needs a
model to judge the output.

- Python 3.12+, Windows and Linux. Pure Python, no network, no telemetry.
- Runtime dependencies: `openpyxl`, `python-docx`, `python-pptx`, `pypdf`, `pdfplumber`,
  `jsonschema`. (`fpdf2` and `pytest` for tests only.)
- Apache-2.0.

```bash
pip install .            # or: pip install -e ".[test]" && pytest
officecheck report.docx examples/report.spec.json            # table, exit 0 / 1
officecheck report.docx examples/report.spec.json --json     # raw result
officecheck --print-schema                                   # spec JSON Schema
```

```python
from officecheck import grade
result = grade("report.docx", {"checks": [{"type": "docx.has_toc"}]})
# {"pass": True, "score": 1.0, "checks": [{"id": "1:docx.has_toc", "pass": True,
#   "expected": {...}, "actual": {...}}]}
```

## Contract

**Spec** (validated with JSON Schema draft 2020-12, see `schema/spec.schema.json`):

```json
{"version": 1, "threshold": 1.0, "checks": [{"id": "toc", "type": "docx.has_toc", "weight": 2}]}
```

Every check takes `type` (required), plus optional `id` (default `"<n>:<type>"`, must be unique),
`weight` (default 1) and `description`. Unknown keys are rejected.

**Result**: `{"pass": bool, "score": 0..1, "checks": [{"id", "pass", "expected", "actual"}]}`
(`result_schema()` gives the JSON Schema).

- `score` = weighted fraction of passing checks.
- `pass` = every check passed (default `threshold` 1.0), or `score >= threshold` when lower.
- A check that cannot run (wrong file type, missing sheet, slide out of range, unsupported
  formula, corrupt file) **fails** with `actual: {"error": "..."}`. Grading never raises because
  of the file; only an invalid spec raises `SpecError`.
- Output is deterministic: same file + spec gives byte-identical JSON.

**CLI exit codes**: `0` pass, `1` fail, `2` usage error or invalid spec (so a broken spec is
never mistaken for a bad document).

## Matching conventions

- `text` matches compare after collapsing whitespace (incl. NBSP/tabs/newlines);
  case-sensitive unless `ignore_case: true`. `regex` uses Python `re.search` — anchor it
  yourself for a full match. Every text check takes exactly one of `text` / `regex`.
- Count checks take any of `equals` / `min` / `max` (at least one).
- Slides and pages are 1-based. File type is detected from content, not the extension.

## Check reference

### docx

| type | parameters | passes when |
|---|---|---|
| `docx.has_heading` | `text`\|`regex`, `level` (0-9), `ignore_case` | a heading paragraph matches (whole text) at that level. Headings = style "Heading N", any style based on one, or an outline level; "Title" = level 0. Paragraphs in tables, content controls and text boxes are included. |
| `docx.heading_count` | `equals`/`min`/`max`, `level` | count of headings (levels 1-9, or exactly `level`) is in bounds |
| `docx.has_toc` | `check_entries` (true), `levels` [lo, hi], `ignore_case` | a `TOC` field exists **and** its pre-filled entries equal, in order, the headings in the field's `\o "lo-hi"` range (default 1-9 if no `\o`). Page numbers, PAGEREF results and leading numbering are stripped from entries. `check_entries: false` = presence only. |
| `docx.has_table` | `min_rows`, `min_cols`, `header_contains` (str or list), `ignore_case`, `min_count` | at least `min_count` (1) tables meet all bounds; each `header_contains` value must equal a first-row cell |
| `docx.contains_text` | `text`\|`regex`, `ignore_case`, `scope` | substring/regex found. `scope`: `body` (default; incl. tables) or `all` (+ headers, footers, footnotes, endnotes, comments) |
| `docx.not_contains_text` | same | not found |
| `docx.style_used` | `style` | a paragraph, run or table in body/headers/footers/notes uses it (name or style id, case-insensitive) |
| `docx.page_orientation` | `orientation`, `section` | every section (or the given one) is that orientation, judged from page width vs height |
| `docx.header_contains` | `text`\|`regex`, `ignore_case` | any header part (default, first-page, even; any section) contains it |
| `docx.footer_has_page_number` | — | a footer contains a `PAGE` field (complex or simple; `NUMPAGES` alone does not count) |

### xlsx

`sheet` defaults to the first sheet; a `cell` may also be written `Sheet!B2` / `'My Sheet'!B2`.
Sheet names resolve case-insensitively (as in Excel) except in `sheet_exists`.

| type | parameters | passes when |
|---|---|---|
| `xlsx.sheet_exists` | `sheet`, `ignore_case` | the sheet exists |
| `xlsx.cell_equals` | `cell`, `value` (str/number/bool/null), `sheet`, `tolerance` (1e-9), `ignore_case`, `source` | the value matches. For formula cells, `source: auto` (default) uses the cached value if the file has one, otherwise evaluates; `cached` uses only the cached value; `evaluate` ignores all cached values. Numbers match numbers only (within `tolerance`), strings match strings; a date matches its ISO string; an error matches its code (`"#DIV/0!"`). `actual.source` says which path was used. |
| `xlsx.cell_has_formula` | `cell`, `pattern`, `sheet`, `ignore_case` (**true**) | the cell holds a formula (normal or array) and `pattern` (regex, searched in the text incl. `=`) matches |
| `xlsx.column_exists` | `header`, `sheet`, `header_row` (1), `ignore_case` | a header cell equals it |
| `xlsx.row_count` | `equals`/`min`/`max`, `sheet`, `header_rows` (1) | non-empty rows below the header rows are in bounds |
| `xlsx.chart_exists` | `chart_type`, `series` (exact), `min_series`, `title`\|`title_regex`, `ignore_case`, `min_count`, `sheet` | at least `min_count` (1) charts meet every constraint. `chart_type` ∈ bar, column, line, pie, doughnut, area, scatter, bubble, radar, stock, surface; `bar` also accepts column charts. Series are counted across all plots of a combo chart. |
| `xlsx.conditional_format_on` | `range`, `sheet`, `rule_type` (e.g. `cellIs`, `colorScale`, `expression`), `mode` | `mode: covers` (default) = every cell of `range` is inside some conditional-format range; `intersects` = any overlap |
| `xlsx.data_validation_on` | `range`, `sheet`, `validation_type` (whole, decimal, list, date, time, textLength, custom), `mode` | same coverage rule for data validations |
| `xlsx.no_error_values` | `sheet` (default all), `errors` (default all 7 codes), `evaluate_uncached` (false) | no cell has a cached error value. Formula cells without a cached value are counted in `actual` but not judged unless `evaluate_uncached: true`. |
| `xlsx.sorted_by` | `column` (header text) \| `column_letter`, `order` (asc), `sheet`, `header_row` (1), `ignore_case` | values below the header are in non-strict order using Excel's ordering (numbers < text < booleans, text case-insensitive). Blank cells are skipped. |
| `xlsx.no_duplicates` | `columns` (headers) \| `column_letters`, `sheet`, `header_row`, `ignore_case` | no two data rows share the same key tuple (all-blank keys skipped) |
| `xlsx.freeze_panes` | `cell` (e.g. `"B2"`, or `null` for none) \| `rows`/`cols`, `sheet` | the frozen pane's top-left cell equals it (`rows: 1` = `"A2"`) |

**Formula evaluator** (`officecheck.formula`): numbers, strings, booleans, error literals,
refs/ranges (`$`, cross-sheet, whole columns), `+ - * / ^ % & = <> < > <= >=` with Excel
precedence (`=-2^2` is 4), and `SUM AVERAGE COUNT COUNTA MIN MAX IF IFERROR VLOOKUP ROUND ABS
AND OR NOT CONCATENATE CONCAT`. Errors propagate as values. Anything else — other functions,
defined names, circular references — fails the check with "cannot evaluate" rather than
guessing.

### pptx

| type | parameters | passes when |
|---|---|---|
| `pptx.slide_count` | `equals`/`min`/`max` | — |
| `pptx.title_on_slide` | `slide`, `text`\|`regex`, `ignore_case` | the slide's title placeholder text matches |
| `pptx.layout_on_slide` | `slide`, `layout`, `ignore_case` | the slide's layout name equals it |
| `pptx.has_notes` | `slide`, `mode` (`all` default \| `any`), `min_chars` (1) | the slide / all slides / any slide has speaker notes |
| `pptx.has_table` | `slide`, `min_rows`, `min_cols`, `min_count` | — (whole deck unless `slide`) |
| `pptx.has_chart` | `slide` + same constraints as `xlsx.chart_exists` | — |
| `pptx.has_image` | `slide`, `min_count` | picture shapes (incl. picture placeholders and image-filled shapes, inside groups) |
| `pptx.text_fits` | `slide`, `max_fill` (1.0), `char_width_em` (0.5), `line_height_em` (1.2), `default_font_pt` (18), `placeholders_only` | heuristic: for each text shape, lines = wrapped character width / box width (after insets), height = lines × line height × font size; `fill` = needed height / available height must be ≤ `max_fill`. Font size comes from the run, else the shape's list style, layout placeholder, master placeholder, master text styles. `normAutofit` font scale is applied; `wrap="none"` also checks width. |

### pdf

| type | parameters | passes when |
|---|---|---|
| `pdf.page_count` | `equals`/`min`/`max` | — |
| `pdf.contains_text` | `text`\|`regex`, `ignore_case`, `page` (default any) | extracted text (pdfplumber) contains it |
| `pdf.form_field_equals` | `name` (full or unique last segment), `value`, `ignore_case` | AcroForm value matches. `true`/`false` test a checkbox on/off; numbers compare numerically; lists compare multi-select values |
| `pdf.is_text_searchable` | `min_chars` (1), `min_ratio` (1.0) | at least `min_ratio` of pages have `min_chars` of extractable text (a scan without OCR fails) |
| `pdf.page_size` | `size` (A3, A4, A5, Letter, Legal, Tabloid) \| `width`+`height` (pt), `orientation`, `tolerance` (2pt), `page` | every page (or `page`) matches; visible box (CropBox) with `/Rotate` applied. Named sizes accept either orientation unless `orientation` is given. |
| `pdf.bookmarks_present` | `min_count` (1), `titles`, `ignore_case` | outline has ≥ `min_count` entries (nested included) and every listed title |
| `pdf.table_on_page` | `page` (default any), `min_rows` (2), `min_cols` (1), `strategy` (`lines` \| `text`) | pdfplumber finds a table that big. `lines` needs ruled tables; `text` infers from alignment. |

### any

| type | parameters | passes when |
|---|---|---|
| `any.file_opens` | `format` | the file parses with the library for its detected type (and is that `format`) |
| `any.size_under` | `max_bytes` | size < `max_bytes` |
| `any.masked` | `strings`, `ignore_case` (**true**) | none of the strings appears anywhere (whitespace-normalised): every XML part of an OOXML package — body, headers, footers, comments, notes, charts, document properties, attribute values (alt text, field codes), deleted tracked changes — recursing into embedded packages; for PDFs page text (two extractors), annotations, form values, metadata/XMP, bookmarks, attachments, and every inflated stream in the raw file (which also catches text left under a black box or in an earlier incremental revision). `actual` lists where each hit was found. Not covered: text drawn as images/outlines and legacy OLE embeddings. |

## Limits worth knowing

- `text_fits` is a heuristic. Calibrate `char_width_em` for your fonts (≈0.5 for Calibri/Arial
  mixed case, ≈0.6 for all caps) and treat `max_fill` near 1.0 as a soft line.
- Cached xlsx values are trusted under `source: auto`. Files saved by openpyxl carry none, so
  they are evaluated; files saved by Excel carry them. Use `source: evaluate` to re-derive
  results and catch stale or hand-written cached values.
- The TOC check compares against cached entries — exactly what a reader sees until the
  field is updated, which is the point.
- `any.masked` with very short strings (e.g. `"42"`) can hit binary stream content in PDFs;
  use distinctive strings.

## Development

```bash
python -m venv .venv && .venv/bin/pip install -e ".[test]"
.venv/bin/pytest            # every fixture is generated by tests/fixtures.py
officecheck --print-schema > schema/spec.schema.json   # after changing a check's parameters
```

New check types register with `@officecheck.check("fmt.name", properties, required=...)`; the
spec schema is assembled from the registry.
