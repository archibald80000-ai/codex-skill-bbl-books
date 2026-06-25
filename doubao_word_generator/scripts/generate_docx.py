from __future__ import annotations

import argparse
import json
import re
import sys
import uuid
from pathlib import Path

SERVICE_ROOT = Path(__file__).resolve().parents[1]
ROOT = Path(__file__).resolve().parents[2]
PARENT_SCRIPTS = ROOT / "scripts"
if str(PARENT_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(PARENT_SCRIPTS))
if str(Path(__file__).resolve().parent) not in sys.path:
    sys.path.insert(0, str(Path(__file__).resolve().parent))

from build_print_docx import build_print_docx  # type: ignore
from clean_materials import clean_items
from extract_materials import extract_zip, load_materials
from export_pdf import export_pdf_book, export_pdf_from_docx_word, update_docx_fields_with_word
from layout_docx import DOC_TYPE_NAME, build_sections
from validate_print_output import validate_docx, validate_pdf  # type: ignore


SUPPORTED_DOC_TYPES = {"manual", "supplier", "meeting", "training", "prd", "contract", "illustrated"}
SUPPORTED_FORMATS = {"docx", "pdf"}


def safe_name(text: str) -> str:
    return re.sub(r'[<>:"/\\|?*\r\n]+', "", text).strip(" ._") or "BBL-books"


def parse_formats(raw: str) -> set[str]:
    values = {part.strip().lower() for part in raw.replace("+", ",").split(",") if part.strip()}
    unknown = values - SUPPORTED_FORMATS
    if unknown or not values:
        raise ValueError("output_formats must be docx, pdf, or docx,pdf")
    return values


def generate_from_paths(
    inputs: list[str],
    doc_type: str,
    output_formats: str,
    title: str | None,
    output_dir: str | Path | None = None,
    job_id: str | None = None,
) -> dict:
    if doc_type not in SUPPORTED_DOC_TYPES:
        raise ValueError(f"unsupported doc_type: {doc_type}")
    formats = parse_formats(output_formats)
    job_id = job_id or uuid.uuid4().hex[:12]
    out_dir = Path(output_dir) if output_dir else SERVICE_ROOT / "outputs" / job_id
    out_dir.mkdir(parents=True, exist_ok=True)
    items = load_materials(inputs)
    paragraphs, tables, removed = clean_items(items)
    images = [item.path for item in items if item.kind == "image"]
    final_title = title or ("贝贝灵熊猫 AI 伴侣原创 IP 形象设计规范" if any("贝贝灵" in "\n".join(i.paragraphs[:5]) for i in items) else DOC_TYPE_NAME.get(doc_type, "正式文档"))
    sections, figure_subject, figure_prefix = build_sections(doc_type, items, paragraphs, tables, images)
    basename = safe_name(final_title) + "_打印版"
    docx_path = out_dir / f"{basename}.docx"
    pdf_path = out_dir / f"{basename}.pdf"
    outputs: list[Path] = []
    if "docx" in formats or "pdf" in formats:
        build_print_docx(
            docx_path,
            final_title,
            sections,
            images,
            subtitle=DOC_TYPE_NAME.get(doc_type, "正式打印版"),
            company="北京贝贝灵科技有限公司" if "贝贝灵" in final_title else "",
            purpose="正式打印 / 审阅 / 交付",
            figure_subject=figure_subject,
            figure_prefix=figure_prefix,
        )
        update_docx_fields_with_word(docx_path)
        if "docx" in formats:
            outputs.append(docx_path)
    if "pdf" in formats:
        ok, _ = export_pdf_from_docx_word(docx_path, pdf_path)
        if not ok:
            export_pdf_book(pdf_path, final_title, sections, images)
        outputs.append(pdf_path)
    validations = {}
    if docx_path.exists():
        validations.update(validate_docx(docx_path, expected_images=len(images)))
    if pdf_path.exists():
        validations.update(validate_pdf(pdf_path))
    report = out_dir / "BBL-books_内部处理报告.md"
    report.write_text(
        "\n".join(
            [
                "# BBL-books 内部处理报告",
                "",
                f"- job_id: {job_id}",
                f"- doc_type: {doc_type}",
                f"- title: {final_title}",
                f"- removed_ai_chatter: {len(removed)}",
                "",
                "## outputs",
                *[f"- {path}" for path in outputs],
                "",
                "## validations",
                *[f"- {k}: {v}" for k, v in validations.items()],
            ]
        )
        + "\n",
        encoding="utf-8",
    )
    return {
        "job_id": job_id,
        "status": "done",
        "output_dir": str(out_dir),
        "files": [str(path) for path in outputs],
        "report": str(report),
        "validations": validations,
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True)
    parser.add_argument("--doc-type", required=True, choices=sorted(SUPPORTED_DOC_TYPES))
    parser.add_argument("--formats", required=True)
    parser.add_argument("--title")
    parser.add_argument("--output-dir", default=str(SERVICE_ROOT / "outputs"))
    args = parser.parse_args()
    result = generate_from_paths([args.input], args.doc_type, args.formats, args.title, args.output_dir)
    print(json.dumps(result, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()

