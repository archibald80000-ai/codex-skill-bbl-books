from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PARENT_EXPORT = ROOT / "scripts" / "export_pdf.py"

spec = importlib.util.spec_from_file_location("bbl_books_parent_export_pdf", PARENT_EXPORT)
if spec is None or spec.loader is None:
    raise ImportError(f"Cannot load parent PDF exporter: {PARENT_EXPORT}")
_module = importlib.util.module_from_spec(spec)
sys.modules["bbl_books_parent_export_pdf"] = _module
spec.loader.exec_module(_module)

export_pdf_book = _module.export_pdf_book
export_pdf_from_docx_word = _module.export_pdf_from_docx_word
update_docx_fields_with_word = _module.update_docx_fields_with_word
