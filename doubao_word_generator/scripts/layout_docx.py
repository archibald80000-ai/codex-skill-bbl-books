from __future__ import annotations

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SERVICE_ROOT = Path(__file__).resolve().parents[1]
PARENT_SCRIPTS = ROOT / "scripts"
if str(PARENT_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(PARENT_SCRIPTS))

from bbl_books import bbl_sections, is_bbl_panda  # type: ignore
from build_print_docx import build_print_docx  # type: ignore


DOC_TYPE_TEMPLATE = {
    "manual": "manual_template.json",
    "supplier": "supplier_template.json",
    "meeting": "meeting_template.json",
    "training": "training_template.json",
    "prd": "prd_template.json",
    "contract": "contract_template.json",
    "illustrated": "illustrated_template.json",
}

DOC_TYPE_NAME = {
    "manual": "打印版手册",
    "supplier": "供应商说明书",
    "meeting": "会议纪要",
    "training": "培训讲义",
    "prd": "产品需求文档",
    "contract": "合同整理版",
    "illustrated": "图文说明书",
}


def load_template(doc_type: str) -> dict:
    template_name = DOC_TYPE_TEMPLATE.get(doc_type, "manual_template.json")
    return json.loads((SERVICE_ROOT / "templates" / template_name).read_text(encoding="utf-8"))


def chunk(paragraphs: list[str], count: int) -> list[list[str]]:
    if count <= 0:
        return []
    size = max(1, (len(paragraphs) + count - 1) // count)
    return [paragraphs[i : i + size] for i in range(0, len(paragraphs), size)]


def generic_sections(doc_type: str, paragraphs: list[str], tables: list[list[list[str]]], images: list[Path]) -> list[dict]:
    template = load_template(doc_type)
    section_titles = [title for title in template["sections"] if title not in {"封面", "目录"}]
    paragraph_groups = chunk(paragraphs, len(section_titles))
    sections: list[dict] = []
    table_index = 0
    for idx, title in enumerate(section_titles):
        section: dict = {
            "title": title,
            "paragraphs": paragraph_groups[idx] if idx < len(paragraph_groups) else [],
        }
        if "交付" in title or "报价" in title or "验收" in title or "主体" in title or "需求" in title:
            if table_index < len(tables):
                section["tables"] = [tables[table_index]]
                table_index += 1
        if ("图册" in title or "附图" in title or "附录" in title) and images:
            section["album"] = True
            section["paragraphs"] = section["paragraphs"] or ["本章节收录输入素材中的有效图片，图片等比缩放并使用正式图注。"]
        if title == "结束页" and not section["paragraphs"]:
            section["paragraphs"] = ["本文件为正式打印版，可用于审阅、沟通、交付和归档。"]
        sections.append(section)
    return sections


def build_sections(doc_type: str, items, paragraphs: list[str], tables: list[list[list[str]]], images: list[Path]) -> tuple[list[dict], str, str]:
    if doc_type in {"manual", "illustrated"} and is_bbl_panda(items):
        return bbl_sections(), "贝贝灵熊猫 AI 伴侣形象", "12"
    figure_subject = DOC_TYPE_NAME.get(doc_type, "项目资料")
    return generic_sections(doc_type, paragraphs, tables, images), figure_subject, "8"
