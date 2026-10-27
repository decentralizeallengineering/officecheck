"""Collect every piece of text a file carries, for leak detection (any.masked).

OOXML: every XML part (body, headers, footers, comments, notes, charts, document properties,
custom XML...), text nodes *and* attribute values (alt text, field values), deleted
tracked-change text and field codes; embedded packages and PDFs are searched recursively.
PDF: page text (two extractors), annotations, form fields, metadata/XMP, outline, attachments,
plus every inflated stream in the raw file - which also catches content left behind by
incremental updates or fake redactions.

Not covered: text rendered as images or vector outlines, and text in legacy binary
(OLE) embeddings.
"""

from __future__ import annotations

import io
import re
import zipfile
import zlib
import xml.etree.ElementTree as ET
from typing import Any, Iterator

from .context import detect_format
from .ooxml import local

# Text inside these elements is joined without separators (runs of one paragraph/cell).
_BLOCKS = {"p", "si", "is", "c", "tc", "tr", "row", "comment", "pt", "v", "f", "cell", "text", "Relationship"}
_XMLISH = (".xml", ".rels", ".vml")
MAX_DEPTH = 4


def _xml_blobs(data: bytes) -> list[str]:
    try:
        root = ET.fromstring(data)
    except ET.ParseError:
        txt = data.decode("utf-8", "replace")
        return [re.sub(r"<[^>]+>", "\n", txt)]
    blocks: list[str] = []
    attrs: list[str] = []

    def walk(el: ET.Element, buf: list[str]) -> None:
        attrs.extend(v for v in el.attrib.values() if v)
        is_block = local(el.tag) in _BLOCKS
        target: list[str] = [] if is_block else buf
        if el.text:
            target.append(el.text)
        for ch in el:
            walk(ch, target)
            if ch.tail:
                target.append(ch.tail)
        if is_block:
            blocks.append("".join(target))

    top: list[str] = []
    walk(root, top)
    blocks.append("".join(top))
    return blocks + attrs


def _zip_blobs(data: bytes, prefix: str, depth: int) -> Iterator[tuple[str, str]]:
    with zipfile.ZipFile(io.BytesIO(data)) as zf:
        for info in sorted(zf.infolist(), key=lambda i: i.filename):
            if info.is_dir() or info.file_size > 200 * 1024 * 1024:
                continue
            name = info.filename
            raw = zf.read(name)
            loc = f"{prefix}{name}"
            if name.lower().endswith(_XMLISH):
                for b in _xml_blobs(raw):
                    yield loc, b
            elif depth < MAX_DEPTH and (raw[:4] == b"PK\x03\x04" or b"%PDF-" in raw[:1024]):
                yield from blobs(raw, f"{loc}::", depth + 1)
            elif name.lower().endswith((".txt", ".csv", ".json", ".htm", ".html", ".md")):
                yield loc, raw.decode("utf-8", "replace")


def _pdf_strings(obj: Any, seen: set[int], depth: int = 0) -> Iterator[str]:
    """All string values reachable from a PDF dict/array (shallow, cycle-safe)."""
    from pypdf.generic import ArrayObject, DictionaryObject, IndirectObject, NameObject, TextStringObject, ByteStringObject

    if depth > 6:
        return
    if isinstance(obj, IndirectObject):
        if obj.idnum in seen:
            return
        seen.add(obj.idnum)
        obj = obj.get_object()
    if isinstance(obj, DictionaryObject):
        for k, v in obj.items():
            if k in ("/Parent", "/P", "/Kids", "/Font", "/Resources", "/AP"):
                continue
            yield from _pdf_strings(v, seen, depth + 1)
    elif isinstance(obj, ArrayObject):
        for v in obj:
            yield from _pdf_strings(v, seen, depth + 1)
    elif isinstance(obj, (TextStringObject, str)) and not isinstance(obj, NameObject):
        yield str(obj)
    elif isinstance(obj, ByteStringObject):
        yield bytes(obj).decode("latin-1")
    elif isinstance(obj, NameObject):
        yield str(obj)[1:]


_STREAM = re.compile(rb"stream\r?\n")


def _raw_pdf_streams(data: bytes) -> Iterator[str]:
    yield data.decode("latin-1")
    for m in _STREAM.finditer(data):
        start = m.end()
        end = data.find(b"endstream", start)
        if end < 0:
            continue
        chunk = data[start:end]
        try:
            out = zlib.decompressobj().decompress(chunk)
        except zlib.error:
            continue
        if out:
            yield out.decode("latin-1")


def _pdf_blobs(data: bytes, prefix: str, depth: int) -> Iterator[tuple[str, str]]:
    import pdfplumber
    import pypdf

    reader = pypdf.PdfReader(io.BytesIO(data))
    if reader.is_encrypted:
        reader.decrypt("")
    for i, page in enumerate(reader.pages, 1):
        try:
            yield f"{prefix}page {i}", page.extract_text() or ""
        except Exception:  # noqa: BLE001
            pass
        annots = page.get("/Annots")
        if annots is not None:
            for s in _pdf_strings(annots, set()):
                yield f"{prefix}page {i} annotations", s
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        for i, page in enumerate(pdf.pages, 1):
            yield f"{prefix}page {i}", page.extract_text() or ""
    for name, field in (reader.get_fields() or {}).items():
        yield f"{prefix}form field {name}", name
        for key in ("/V", "/DV", "/Opt", "/TU"):
            if key in field:
                for s in _pdf_strings(field[key], set()):
                    yield f"{prefix}form field {name}", s
    if reader.metadata:
        for k, v in reader.metadata.items():
            yield f"{prefix}metadata {k}", str(v)
    xmp = reader.trailer["/Root"].get("/Metadata")
    if xmp is not None:
        yield f"{prefix}XMP metadata", xmp.get_object().get_data().decode("utf-8", "replace")

    def outline(items: Any) -> Iterator[str]:
        for it in items:
            if isinstance(it, list):
                yield from outline(it)
            else:
                yield str(getattr(it, "title", "") or "")

    for t in outline(reader.outline):
        yield f"{prefix}bookmarks", t
    for name, contents in (reader.attachments or {}).items():
        yield f"{prefix}attachment", name
        for content in contents:
            if depth < MAX_DEPTH:
                yield from blobs(content, f"{prefix}attachment {name}::", depth + 1)
    for s in _raw_pdf_streams(data):
        yield f"{prefix}raw pdf streams", s


def blobs(data: bytes, prefix: str = "", depth: int = 0) -> Iterator[tuple[str, str]]:
    """Yield (location, text) for every text-bearing piece of ``data``."""
    fmt = detect_format(data)
    try:
        if fmt == "pdf":
            yield from _pdf_blobs(data, prefix, depth)
            return
        if data[:4] == b"PK\x03\x04":
            yield from _zip_blobs(data, prefix, depth)
            return
    except Exception as e:  # noqa: BLE001 - fall back to the raw bytes below
        yield f"{prefix}(unparseable: {type(e).__name__})", data.decode("latin-1")
        return
    yield f"{prefix}raw", data.decode("utf-8", "replace")
