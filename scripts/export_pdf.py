from __future__ import annotations

import subprocess
from pathlib import Path


def _quote_ps_path(path: Path) -> str:
    return "'" + str(path).replace("'", "''") + "'"


def _run_word_powershell(docx_path: Path, pdf_path: Path | None = None) -> tuple[bool, str]:
    docx_q = _quote_ps_path(docx_path.resolve())
    export_line = ""
    if pdf_path is not None:
        export_line = f"$doc.ExportAsFixedFormat({_quote_ps_path(pdf_path.resolve())}, 17);"
    script = f"""
$ErrorActionPreference='Stop'
$word = New-Object -ComObject Word.Application
$word.Visible = $false
try {{
  $doc = $word.Documents.Open({docx_q})
  $doc.Fields.Update() | Out-Null
  foreach ($story in $doc.StoryRanges) {{ $story.Fields.Update() | Out-Null }}
  $doc.Save()
  {export_line}
  $doc.Close($false)
}} finally {{
  $word.Quit()
}}
"""
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        timeout=180,
    )
    if result.returncode != 0:
        return False, (result.stderr or result.stdout or "PowerShell Word COM failed").strip()
    return True, "Word COM via PowerShell"


def update_docx_fields_with_word(docx_path: str | Path) -> tuple[bool, str]:
    ps_ok, ps_msg = _run_word_powershell(Path(docx_path), None)
    if ps_ok:
        return True, ps_msg
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        return False, f"win32com unavailable: {type(exc).__name__}: {exc}"
    docx_path = Path(docx_path).resolve()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(docx_path))
        doc.Fields.Update()
        for story in doc.StoryRanges:
            story.Fields.Update()
        doc.Save()
        doc.Close(False)
        return True, "Word fields updated"
    except Exception as exc:
        return False, f"Word update failed: {type(exc).__name__}: {exc}"
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def export_pdf_from_docx_word(docx_path: str | Path, pdf_path: str | Path) -> tuple[bool, str]:
    ps_ok, ps_msg = _run_word_powershell(Path(docx_path), Path(pdf_path))
    if ps_ok:
        return True, str(Path(pdf_path))
    try:
        import win32com.client  # type: ignore
    except Exception as exc:
        return False, f"win32com unavailable: {type(exc).__name__}: {exc}"
    docx_path = Path(docx_path).resolve()
    pdf_path = Path(pdf_path).resolve()
    word = None
    try:
        word = win32com.client.DispatchEx("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(str(docx_path))
        doc.Fields.Update()
        for story in doc.StoryRanges:
            story.Fields.Update()
        doc.Save()
        doc.ExportAsFixedFormat(str(pdf_path), 17)
        doc.Close(False)
        return True, str(pdf_path)
    except Exception as exc:
        return False, f"Word PDF export failed: {type(exc).__name__}: {exc}"
    finally:
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass


def export_pdf_from_html(html_path: str | Path, pdf_path: str | Path) -> tuple[bool, str]:
    html_path = Path(html_path)
    pdf_path = Path(pdf_path)
    cmd = [
        "pandoc",
        str(html_path),
        "-o",
        str(pdf_path),
        "--pdf-engine=xelatex",
        "-V",
        "CJKmainfont=SimSun",
        "-V",
        "mainfont=Arial",
    ]
    try:
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=180)
    except Exception as exc:
        return False, f"{type(exc).__name__}: {exc}"
    if result.returncode != 0:
        return False, (result.stderr or result.stdout or "pandoc failed").strip()
    return True, str(pdf_path)


def export_pdf_book(
    pdf_path: str | Path,
    title: str,
    sections: list[dict],
    images: list[Path],
    margin_cm: float = 2,
    min_font_pt: float = 12,
) -> tuple[bool, str]:
    try:
        from PIL import Image as PILImage
        from reportlab.lib import colors
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
        from reportlab.lib.units import cm
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont
        from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
    except Exception as exc:
        return False, f"ReportLab unavailable: {type(exc).__name__}: {exc}"

    pdf_path = Path(pdf_path)
    try:
        pdfmetrics.registerFont(UnicodeCIDFont("STSong-Light"))
    except Exception:
        pass

    styles = getSampleStyleSheet()
    body = ParagraphStyle(
        "BBLBody",
        parent=styles["Normal"],
        fontName="STSong-Light",
        fontSize=min_font_pt,
        leading=max(min_font_pt * 1.5, 18),
        spaceAfter=6,
    )
    h1 = ParagraphStyle(
        "BBLH1",
        parent=styles["Heading1"],
        fontName="STSong-Light",
        fontSize=18,
        leading=24,
        spaceBefore=12,
        spaceAfter=12,
    )
    caption = ParagraphStyle(
        "BBLCaption",
        parent=body,
        fontSize=10,
        leading=14,
        alignment=1,
        textColor=colors.HexColor("#555555"),
    )

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=A4,
        leftMargin=margin_cm * cm,
        rightMargin=margin_cm * cm,
        topMargin=margin_cm * cm,
        bottomMargin=margin_cm * cm,
        title=title,
    )
    page_width, page_height = A4
    usable_width = page_width - 2 * margin_cm * cm
    usable_height = page_height - 2 * margin_cm * cm

    story = [Spacer(1, 6 * cm), Paragraph(title, h1), Paragraph("资料汇编成书版", body), PageBreak()]
    story.append(Paragraph("目录", h1))
    story.append(Paragraph("PDF 后备导出版保留章节内容；可使用 DOCX 获取可更新目录。", body))
    story.append(PageBreak())

    for section in sections:
        story.append(Paragraph(section["title"], h1))
        for paragraph in section.get("paragraphs", []):
            safe = str(paragraph).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            story.append(Paragraph(safe, body))
        for rows in section.get("tables", []):
            table = Table(rows, repeatRows=1)
            table.setStyle(
                TableStyle(
                    [
                        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                        ("FONTNAME", (0, 0), (-1, -1), "STSong-Light"),
                        ("FONTSIZE", (0, 0), (-1, -1), min_font_pt),
                        ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ]
                )
            )
            story.append(table)
            story.append(Spacer(1, 0.3 * cm))
        for subsection in section.get("subsections", []):
            story.append(Paragraph(subsection["title"], h1))
            for paragraph in subsection.get("paragraphs", []):
                safe = str(paragraph).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
                story.append(Paragraph(safe, body))
            for rows in subsection.get("tables", []):
                table = Table(rows, repeatRows=1)
                table.setStyle(
                    TableStyle(
                        [
                            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                            ("FONTNAME", (0, 0), (-1, -1), "STSong-Light"),
                            ("FONTSIZE", (0, 0), (-1, -1), min_font_pt),
                            ("VALIGN", (0, 0), (-1, -1), "TOP"),
                        ]
                    )
                )
                story.append(table)
                story.append(Spacer(1, 0.3 * cm))

    if images:
        story.append(PageBreak())
        story.append(Paragraph("贝贝灵熊猫 AI 伴侣形象图册", h1))
        for idx, image_path in enumerate(images, 1):
            with PILImage.open(image_path) as pil:
                img_w, img_h = pil.size
            scale = min(usable_width / img_w, (usable_height - 2 * cm) / img_h, 1)
            story.append(Paragraph(f"图 11-{idx} 贝贝灵熊猫 AI 伴侣形象图（{idx}）", caption))
            story.append(Image(str(image_path), width=img_w * scale, height=img_h * scale))
            if idx != len(images):
                story.append(PageBreak())

    def page_canvas(canvas, document):
        canvas.saveState()
        canvas.setFont("STSong-Light", 9)
        canvas.drawCentredString(page_width / 2, page_height - 1.2 * cm, title)
        canvas.drawCentredString(page_width / 2, 1 * cm, f"第 {document.page} 页")
        canvas.restoreState()

    try:
        doc.build(story, onFirstPage=page_canvas, onLaterPages=page_canvas)
    except Exception as exc:
        return False, f"ReportLab export failed: {type(exc).__name__}: {exc}"
    return True, str(pdf_path)
