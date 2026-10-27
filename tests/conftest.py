from __future__ import annotations

import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

import fixtures as fx  # noqa: E402

BUILDERS = {
    "docx_good": ("good.docx", fx.make_docx_good),
    "docx_bad": ("bad.docx", fx.make_docx_bad),
    "docx_no_toc": ("no_toc.docx", lambda p: fx.make_docx_bad(p, toc=False)),
    "xlsx_good": ("good.xlsx", fx.make_xlsx_good),
    "xlsx_bad": ("bad.xlsx", fx.make_xlsx_bad),
    "pptx_good": ("good.pptx", fx.make_pptx_good),
    "pptx_bad": ("bad.pptx", fx.make_pptx_bad),
    "pdf_good": ("good.pdf", fx.make_pdf_good),
    "pdf_bad": ("bad.pdf", fx.make_pdf_bad),
    "corrupt_docx": ("corrupt.docx", lambda p: (p.write_bytes(fx_bytes_truncated()), p)[1]),
    "big_blob": ("big.bin", lambda p: (p.write_bytes(bytes(range(256)) * 8192), p)[1]),
    "masked_clean": ("clean.docx", fx.make_masked_clean_docx),
    "secret_docx_header": ("s_header.docx", fx.make_secret_docx_header),
    "secret_docx_footer": ("s_footer.docx", fx.make_secret_docx_footer_split_runs),
    "secret_docx_comment": ("s_comment.docx", fx.make_secret_docx_comment),
    "secret_pptx_notes": ("s_notes.pptx", fx.make_secret_pptx_notes),
    "secret_xlsx_comment": ("s_comment.xlsx", fx.make_secret_xlsx_comment),
    "secret_pdf_form": ("s_form.pdf", fx.make_secret_pdf_form),
    "secret_pdf_box": ("s_box.pdf", fx.make_secret_pdf_under_box),
}


def fx_bytes_truncated() -> bytes:
    import io

    import docx

    buf = io.BytesIO()
    docx.Document().save(buf)
    data = buf.getvalue()
    return data[: len(data) // 2]  # zip magic intact, central directory gone


@pytest.fixture(scope="session")
def files(tmp_path_factory: pytest.TempPathFactory) -> dict[str, Path]:
    root = tmp_path_factory.mktemp("fixtures")
    out = {}
    for key, (name, build) in BUILDERS.items():
        out[key] = build(root / name)
    return out
