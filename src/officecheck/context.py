"""Lazy, cached loaders for the file under test."""

from __future__ import annotations

import contextlib
import io
import os
import zipfile
from pathlib import Path
from typing import Any

from .registry import CheckError

OOXML_MAIN = {
    "docx": "word/document.xml",
    "xlsx": "xl/workbook.xml",
    "pptx": "ppt/presentation.xml",
}


def detect_format(data: bytes) -> str | None:
    """Detect docx/xlsx/pptx/pdf from content (not the extension)."""
    if b"%PDF-" in data[:1024]:  # the spec tolerates leading junk before the header
        return "pdf"
    if data[:4] == b"PK\x03\x04":
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as zf:
                names = set(zf.namelist())
        except zipfile.BadZipFile:
            return None
        for fmt, main in OOXML_MAIN.items():
            if main in names:
                return fmt
    return None


class Context:
    """Holds the file bytes and lazily opens it with each library at most once."""

    def __init__(self, path: str | os.PathLike[str]):
        self.path = Path(path)
        self._cache: dict[str, Any] = {}
        self._errors: dict[str, str] = {}
        self._stack = contextlib.ExitStack()
        try:
            self.data: bytes | None = self.path.read_bytes()
            self.read_error: str | None = None
        except OSError as e:
            self.data = None
            self.read_error = f"{type(e).__name__}: {e}"
        self.format = detect_format(self.data) if self.data is not None else None

    def close(self) -> None:
        self._stack.close()

    def __enter__(self) -> "Context":
        return self

    def __exit__(self, *exc: object) -> None:
        self.close()

    # -- generic -----------------------------------------------------------------------------

    def require_bytes(self) -> bytes:
        if self.data is None:
            raise CheckError(f"cannot read file: {self.read_error}")
        return self.data

    def _load(self, key: str, fmt: str, loader) -> Any:
        if key in self._cache:
            return self._cache[key]
        if key in self._errors:
            raise CheckError(self._errors[key])
        data = self.require_bytes()
        if self.format != fmt:
            msg = f"file is not a {fmt} (detected: {self.format or 'unknown'})"
            self._errors[key] = msg
            raise CheckError(msg)
        try:
            obj = loader(data)
        except CheckError:
            raise
        except Exception as e:  # noqa: BLE001 - any parse failure means "does not open"
            msg = f"cannot open as {fmt}: {type(e).__name__}: {e}"
            self._errors[key] = msg
            raise CheckError(msg) from e
        self._cache[key] = obj
        return obj

    def zip(self) -> zipfile.ZipFile:
        def load(data: bytes) -> zipfile.ZipFile:
            return self._stack.enter_context(zipfile.ZipFile(io.BytesIO(data)))

        self.require_bytes()
        if self.format not in OOXML_MAIN:
            raise CheckError(f"file is not an OOXML package (detected: {self.format or 'unknown'})")
        return self._load("zip", self.format, load)

    # -- per-format --------------------------------------------------------------------------

    def docx(self):
        import docx

        return self._load("docx", "docx", lambda d: docx.Document(io.BytesIO(d)))

    def xlsx(self, data_only: bool = False):
        """Workbook with formulas (data_only=False) or cached values (data_only=True)."""
        import openpyxl

        key = "xlsx_values" if data_only else "xlsx_formulas"
        return self._load(
            key, "xlsx", lambda d: openpyxl.load_workbook(io.BytesIO(d), data_only=data_only)
        )

    def pptx(self):
        import pptx

        return self._load("pptx", "pptx", lambda d: pptx.Presentation(io.BytesIO(d)))

    def pypdf(self):
        import pypdf

        def load(d: bytes):
            r = pypdf.PdfReader(io.BytesIO(d))
            if r.is_encrypted:
                try:
                    if not r.decrypt(""):
                        raise CheckError("pdf is encrypted with a user password")
                except CheckError:
                    raise
                except Exception as e:
                    raise CheckError(f"pdf is encrypted and cannot be opened: {e}") from e
            _ = len(r.pages)
            return r

        return self._load("pypdf", "pdf", load)

    def plumber(self):
        import pdfplumber

        def load(d: bytes):
            return self._stack.enter_context(pdfplumber.open(io.BytesIO(d)))

        return self._load("plumber", "pdf", load)

    def pdf_page_texts(self) -> list[str]:
        """Text of each page (pdfplumber), cached."""
        if "pdf_texts" not in self._cache:
            pdf = self.plumber()
            self._cache["pdf_texts"] = [(p.extract_text() or "") for p in pdf.pages]
        return self._cache["pdf_texts"]

    def open_native(self) -> None:
        """Open the file with the library matching its detected format (raises CheckError)."""
        self.require_bytes()
        if self.format == "docx":
            self.docx()
        elif self.format == "xlsx":
            self.xlsx(False)
            self.xlsx(True)
        elif self.format == "pptx":
            self.pptx()
        elif self.format == "pdf":
            self.pypdf()
            self.pdf_page_texts()
        elif self.data[:4] == b"PK\x03\x04":
            try:
                zipfile.ZipFile(io.BytesIO(self.data)).close()
            except zipfile.BadZipFile as e:
                raise CheckError(f"corrupt ZIP/OOXML package: {e}") from e
            raise CheckError("ZIP file is not a docx, xlsx or pptx package")
        else:
            raise CheckError("unrecognised file format (expected docx, xlsx, pptx or pdf)")
