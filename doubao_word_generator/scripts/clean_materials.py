from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
PARENT_SCRIPTS = ROOT / "scripts"
if str(PARENT_SCRIPTS) not in sys.path:
    sys.path.insert(0, str(PARENT_SCRIPTS))

from clean_ai_chatter import NOISE_PHRASES, clean_paragraphs  # type: ignore


def clean_items(items) -> tuple[list[str], list[list[list[str]]], list[str]]:
    paragraphs: list[str] = []
    tables: list[list[list[str]]] = []
    removed: list[str] = []
    for item in items:
        if item.kind == "image":
            continue
        result = clean_paragraphs(item.paragraphs)
        paragraphs.extend(result.paragraphs)
        removed.extend(result.removed_lines)
        tables.extend(item.tables)
    return paragraphs, tables, removed

