from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Iterable


NOISE_PHRASES = [
    "好！直接给你",
    "好，直接给你",
    "立刻给你完整成品文档",
    "我现在帮你做",
    "我现在直接帮你下一步",
    "我帮你免费审核是否侵权",
    "我直接给你最精准、无风险、可落地的结论与操作步骤，你按这套做，完全不用担法律风险。",
    "我可以直接把这份规范生成可下载的 PDF 文件",
    "我可以直接把这份规范生成可下载的PDF文件",
    "你只需告诉我",
    "告诉我这 3 个信息",
    "告诉我这3个信息",
    "需要我现在",
    "要不要我帮你",
    "我接下来可以给你",
    "你现在直接把上面内容复制粘贴",
    "你现在直接把上面三份内容复制粘贴",
    "文档已完成",
    "豆包AI生成",
    "豆包 AI 生成",
    "注：文档部分内容可能由 AI 生成",
    "注: 文档部分内容可能由 AI 生成",
    "想要可爱/科技/治愈哪种风格",
    "硬件是桌面/挂件/大型",
]

QUESTION_NOISE_RE = re.compile(
    r"(要不要|需要我|是否需要我|你告诉我|告诉我：|我可以帮你|我可以直接帮你|我现在就给你|我现在可以)"
)
ROLE_LABEL_RE = re.compile(r"^\s*(用户|User|assistant|Assistant|system|System|tool|Tool|AI|豆包)\s*[:：]\s*$")
CHAT_TITLE_RE = re.compile(r"^\s*和豆包的对话[_\-\s]*\d*(?:\(\d+\))?\s*$")


@dataclass
class CleanResult:
    paragraphs: list[str]
    removed_lines: list[str] = field(default_factory=list)
    duplicate_lines: list[str] = field(default_factory=list)


def is_substantive_question(text: str) -> bool:
    return any(token in text for token in ["要不要自己画", "注册版权", "申请外观", "法律风险", "商标"])


def is_noise_line(text: str) -> bool:
    stripped = text.strip().strip("|").strip()
    if not stripped:
        return True
    if ROLE_LABEL_RE.match(stripped) or CHAT_TITLE_RE.match(stripped):
        return True
    for phrase in NOISE_PHRASES:
        if phrase in stripped:
            return True
    if QUESTION_NOISE_RE.search(stripped) and not is_substantive_question(stripped):
        return True
    if stripped in {"你告诉我：", "我可以帮你：", "我现在帮你做："}:
        return True
    return False


def normalize_for_dedupe(text: str) -> str:
    return re.sub(r"\s+", "", text.strip())


def clean_paragraphs(paragraphs: Iterable[str], dedupe_level: str = "conservative") -> CleanResult:
    kept: list[str] = []
    removed: list[str] = []
    duplicates: list[str] = []
    seen: set[str] = set()

    for raw in paragraphs:
        text = str(raw).strip()
        if is_noise_line(text):
            if text:
                removed.append(text)
            continue
        key = normalize_for_dedupe(text)
        if dedupe_level == "conservative" and key in seen and len(key) > 10:
            duplicates.append(text)
            continue
        seen.add(key)
        kept.append(text)

    return CleanResult(paragraphs=kept, removed_lines=removed, duplicate_lines=duplicates)


def clean_text(text: str, dedupe_level: str = "conservative") -> CleanResult:
    return clean_paragraphs(text.splitlines(), dedupe_level=dedupe_level)
