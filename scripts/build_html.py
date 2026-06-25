from __future__ import annotations

import html
import shutil
from pathlib import Path
from typing import Iterable


def build_html(
    output_path: str | Path,
    title: str,
    sections: list[dict],
    images: Iterable[Path],
    margin_cm: float = 2,
    font_zh: str = "SimSun",
    font_west: str = "Arial",
) -> Path:
    output = Path(output_path)
    asset_dir = output.with_name(output.stem + "_assets")
    asset_dir.mkdir(exist_ok=True)

    image_tags: list[str] = []
    for idx, image_path in enumerate(images, 1):
        dest = asset_dir / image_path.name
        if image_path.resolve() != dest.resolve():
            shutil.copy2(image_path, dest)
        image_tags.append(
            f'<figure><img src="{html.escape(asset_dir.name + "/" + dest.name)}" alt="图 {idx}">'
            f"<figcaption>图 11-{idx} 贝贝灵熊猫 AI 伴侣形象图（{idx}）</figcaption></figure>"
        )

    parts = [
        "<!doctype html>",
        '<html lang="zh-CN"><head><meta charset="utf-8">',
        f"<title>{html.escape(title)}</title>",
        "<style>",
        f"@page {{ size: A4; margin: {margin_cm}cm; }}",
        f"body {{ font-family: {font_west}, {font_zh}, serif; font-size: 12pt; line-height: 1.5; color: #111; }}",
        "h1 { page-break-before: always; border-bottom: 1px solid #999; padding-bottom: .25em; }",
        "h1:first-of-type { page-break-before: auto; }",
        "table { border-collapse: collapse; width: 100%; margin: 1em 0; }",
        "td, th { border: 1px solid #777; padding: .35em; vertical-align: top; }",
        "figure { page-break-inside: avoid; text-align: center; margin: 1.2em 0; }",
        "img { max-width: 100%; max-height: 23cm; object-fit: contain; }",
        "figcaption { font-size: 10pt; color: #555; margin-top: .4em; }",
        "</style></head><body>",
        f"<h1>{html.escape(title)}</h1>",
        "<h1>目录</h1><p>请使用浏览器或 PDF 工具生成目录页。</p>",
    ]

    for section in sections:
        parts.append(f"<h1>{html.escape(section['title'])}</h1>")
        for paragraph in section.get("paragraphs", []):
            parts.append(f"<p>{html.escape(paragraph)}</p>")
        for table in section.get("tables", []):
            parts.append("<table>")
            for row in table:
                parts.append("<tr>" + "".join(f"<td>{html.escape(str(cell))}</td>" for cell in row) + "</tr>")
            parts.append("</table>")
        if section.get("album"):
            parts.extend(image_tags)
        for subsection in section.get("subsections", []):
            parts.append(f"<h2>{html.escape(subsection['title'])}</h2>")
            for paragraph in subsection.get("paragraphs", []):
                parts.append(f"<p>{html.escape(paragraph)}</p>")
            for table in subsection.get("tables", []):
                parts.append("<table>")
                for row in table:
                    parts.append("<tr>" + "".join(f"<td>{html.escape(str(cell))}</td>" for cell in row) + "</tr>")
                parts.append("</table>")

    if image_tags and not any(section.get("album") for section in sections):
        parts.append("<h1>贝贝灵熊猫 AI 伴侣形象图册</h1>")
        parts.extend(image_tags)
    parts.append("</body></html>")
    output.write_text("\n".join(parts), encoding="utf-8")
    return output
