"""Low-level OOXML helpers on top of zipfile + xml.etree (works on lxml elements too)."""

from __future__ import annotations

import posixpath
import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, Callable, Iterator

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
CHART_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
DRAWING_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"


def q(prefix_local: str) -> str:
    """'w:p' -> '{namespace}p' (Clark notation, as used by ElementTree and lxml)."""
    prefix, local = prefix_local.split(":")
    return f"{{{NS[prefix]}}}{local}"


def local(tag: Any) -> str:
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def read_xml(zf: zipfile.ZipFile, name: str) -> ET.Element:
    return ET.fromstring(zf.read(name))


def rels_path(part: str) -> str:
    d, f = posixpath.split(part)
    return posixpath.join(d, "_rels", f + ".rels")


def part_rels(zf: zipfile.ZipFile, part: str) -> dict[str, tuple[str, str]]:
    """{rId: (type, absolute target part)} for internal relationships of ``part``."""
    rp = rels_path(part)
    if rp not in zf.namelist():
        return {}
    out = {}
    base = posixpath.dirname(part)
    for rel in read_xml(zf, rp):
        if rel.get("TargetMode") == "External":
            continue
        target = rel.get("Target", "")
        tgt = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
        out[rel.get("Id", "")] = (rel.get("Type", ""), tgt)
    return out


def parts_matching(zf: zipfile.ZipFile, pattern: str) -> list[str]:
    rx = re.compile(pattern)
    return sorted(n for n in zf.namelist() if rx.fullmatch(n))


# ---- WordprocessingML text ------------------------------------------------------------------

_W_SKIP = {q("w:pPr"), q("w:rPr"), q("w:instrText"), q("w:delText"), q("w:tblPr"), q("w:trPr"),
           q("w:tcPr"), q("w:sectPr"), q("w:fldData"), q("w:tblGrid")}
_W_T, _W_TAB, _W_BR, _W_CR, _W_NBH = q("w:t"), q("w:tab"), q("w:br"), q("w:cr"), q("w:noBreakHyphen")
W_P = q("w:p")


def w_text(el: Any) -> str:
    """Visible text of a WordprocessingML element (no field codes, no deleted text)."""
    out: list[str] = []

    def walk(e: Any) -> None:
        for ch in e:
            t = ch.tag
            if t in _W_SKIP:
                continue
            if t == _W_T:
                out.append(ch.text or "")
            elif t == _W_TAB:
                out.append("\t")
            elif t in (_W_BR, _W_CR):
                out.append("\n")
            elif t == _W_NBH:
                out.append("-")
            else:
                walk(ch)

    walk(el)
    return "".join(out)


def w_paragraph_texts(root: Any) -> list[str]:
    return [w_text(p) for p in root.iter(W_P)]


def field_instructions(root: Any) -> list[str]:
    """Instruction text of every field (complex and simple) under ``root``, in order."""
    instrs: list[str] = []
    stack: list[int] = []
    for el in root.iter():
        t = el.tag
        if t == q("w:fldChar"):
            kind = el.get(q("w:fldCharType"))
            if kind == "begin":
                instrs.append("")
                stack.append(len(instrs) - 1)
            elif kind == "end" and stack:
                stack.pop()
        elif t == q("w:instrText") and stack:
            instrs[stack[-1]] += el.text or ""
        elif t == q("w:fldSimple"):
            instrs.append(el.get(q("w:instr"), ""))
    return [i.strip() for i in instrs]


# ---- DrawingML text --------------------------------------------------------------------------


def a_paragraph_texts(root: Any) -> list[str]:
    out = []
    for p in root.iter(q("a:p")):
        parts = []
        for el in p.iter():
            if el.tag == q("a:t"):
                parts.append(el.text or "")
            elif el.tag == q("a:br"):
                parts.append("\n")
        out.append("".join(parts))
    return out


# ---- Charts ---------------------------------------------------------------------------------

CHART_KINDS = {
    "barChart": "bar", "bar3DChart": "bar", "lineChart": "line", "line3DChart": "line",
    "pieChart": "pie", "pie3DChart": "pie", "ofPieChart": "pie", "doughnutChart": "doughnut",
    "areaChart": "area", "area3DChart": "area", "scatterChart": "scatter", "bubbleChart": "bubble",
    "radarChart": "radar", "stockChart": "stock", "surfaceChart": "surface",
    "surface3DChart": "surface",
}


def parse_chart(chart_space: Any) -> dict[str, Any]:
    """Summarise a c:chartSpace element: plot types, series count and title text.

    ``types`` uses normalised names; a bar chart with barDir=col is reported as "column"
    (a spec asking for "bar" accepts both).
    """
    chart = chart_space.find(q("c:chart"))
    if chart is None:
        return {"types": [], "series": 0, "title": None}
    types: list[str] = []
    series = 0
    plot = chart.find(q("c:plotArea"))
    if plot is not None:
        for el in plot:
            kind = CHART_KINDS.get(local(el.tag))
            if kind is None:
                continue
            if kind == "bar":
                bd = el.find(q("c:barDir"))
                if bd is not None and bd.get("val") == "col":
                    kind = "column"
            types.append(kind)
            series += len(el.findall(q("c:ser")))
    title = None
    t = chart.find(q("c:title"))
    if t is not None:
        rich = [x.text or "" for x in t.iter(q("a:t"))]
        if rich:
            title = "".join(rich)
        else:  # title from a cell reference: use the cached string
            cache = [x.text or "" for x in t.iter(q("c:v"))]
            title = "".join(cache) if cache else ""
    return {"types": types, "series": series, "title": title}


def chart_type_matches(wanted: str, types: list[str]) -> bool:
    wanted = wanted.lower().removesuffix("chart")
    if wanted == "bar":
        return any(t in ("bar", "column") for t in types)
    return wanted in types


def iter_charts_in_xlsx(zf: zipfile.ZipFile) -> Iterator[tuple[str, str]]:
    """Yield (sheet name, chart part name) for every chart in every worksheet/chartsheet."""
    wb_part = "xl/workbook.xml"
    wb = read_xml(zf, wb_part)
    rels = part_rels(zf, wb_part)
    sheets = wb.find(q("x:sheets"))
    for s in (sheets if sheets is not None else []):
        rid = s.get(q("r:id"))
        if rid not in rels:
            continue
        sheet_part = rels[rid][1]
        for _typ, drawing in part_rels(zf, sheet_part).values():
            if _typ != DRAWING_REL:
                continue
            for ctyp, chart_part in part_rels(zf, drawing).values():
                if ctyp == CHART_REL and chart_part in zf.namelist():
                    yield s.get("name", ""), chart_part


def safe(fn: Callable[[], Any], default: Any = None) -> Any:
    try:
        return fn()
    except Exception:  # noqa: BLE001
        return default
