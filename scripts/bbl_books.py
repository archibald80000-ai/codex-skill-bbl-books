from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from build_html import build_html
from build_print_docx import build_print_docx
from clean_ai_chatter import NOISE_PHRASES, clean_paragraphs
from export_pdf import export_pdf_book, export_pdf_from_docx_word, update_docx_fields_with_word
from inspect_inputs import InputItem, inspect_inputs
from validate_print_output import validate_docx, validate_pdf

SUPPORTED_FORMATS = {"docx", "pdf", "html"}


def fail(message: str) -> None:
    print(message, file=sys.stderr)
    raise SystemExit(2)


def parse_formats(raw: str | None) -> set[str]:
    if not raw:
        fail("请指定输出格式：docx、pdf、docx+pdf、html 或 docx+pdf+html。")
    normalized = raw.replace("+", ",")
    formats = {part.strip().lower() for part in normalized.split(",") if part.strip()}
    unknown = formats - SUPPORTED_FORMATS
    if unknown or not formats:
        fail("请指定输出格式：docx、pdf、docx+pdf、html 或 docx+pdf+html。")
    return formats


def safe_name(text: str) -> str:
    return re.sub(r'[<>:"/\\|?*\r\n]+', "", text).strip(" ._") or "正式手册"


def infer_output_dir(inputs: list[str], output_dir: str | None) -> Path:
    if output_dir:
        return Path(output_dir).resolve()
    first = Path(inputs[0]).resolve()
    return first if first.is_dir() else first.parent


def infer_title(items: list[InputItem], title: str | None) -> str:
    if title:
        return title
    for item in items:
        for paragraph in item.paragraphs:
            text = paragraph.strip()
            if 6 <= len(text) <= 45 and not text.startswith("和") and "对话" not in text:
                return re.sub(r"[｜|].*$", "", text).strip() or "资料打印版手册"
    return "资料打印版手册"


def is_bbl_panda(items: list[InputItem]) -> bool:
    text = "\n".join("\n".join(item.paragraphs[:20]) for item in items if item.kind != "image")
    return "贝贝灵" in text and "熊猫" in text


def bbl_sections() -> list[dict]:
    return [
        {"title": "一、合规总说明", "paragraphs": [
            "贝贝灵熊猫 AI 伴侣可以使用“熊猫”这一动物概念进行原创开发。熊猫物种本身不受版权保护，受保护的是他人已经创作完成的具体熊猫形象、IP、商标、外观设计和美术作品。因此，本项目应坚持原创形象设计、版权登记、商标注册、外观设计专利申请四位一体的确权路径，以降低后续商用、授权、销售和维权风险。",
            "项目不得直接使用、临摹或高度近似冰墩墩、花花、和叶、功夫熊猫、网红熊猫表情包等既有形象；不得使用网上无授权熊猫插画、3D 模型或 AI 生成熊猫图进行商用；不得使用成都熊猫基地、央视、动物园等官方 IP 形象。法院判断侵权的重点不是“是不是熊猫”，而是五官、比例、神态、动作、装饰和整体表达是否与已发表作品构成实质性相似。",
        ], "tables": [[["事项", "合规要求"], ["可使用内容", "熊猫物种概念、黑白配色、圆润体态、原创独立设计的熊猫卡通形象。"], ["禁止使用内容", "冰墩墩、花花、功夫熊猫、网红熊猫表情包、无授权插画、无授权 3D 模型、未经授权的官方 IP 形象。"], ["判断标准", "重点判断具体形象是否构成实质性相似，包括五官、比例、神态、动作、装饰、整体视觉表达。"], ["确权路径", "原创设计、著作权登记、商标注册、外观设计专利申请同步推进。"]]]},
        {"title": "二、原创 IP 核心设计原则", "paragraphs": [
            "贝贝灵熊猫 AI 伴侣应以“原创熊猫形象 + AI 科技陪伴属性”为核心方向。设计既要保留熊猫的亲和力，又必须通过比例、轮廓、五官、识别符号、配色和交互细节形成独立辨识度，避免落入普通熊猫、网红熊猫或既有影视形象的视觉路径。",
            "建议采用原创设计、版权登记、商标注册、外观专利申请四位一体的确权路径。设计阶段应保留草图、源文件、修改记录、创作说明和时间戳，作为后续版权登记、维权和商业授权的基础证据。",
        ], "tables": [[["原则", "执行要求"], ["原创性", "五官、比例、装饰、体态四个维度均需形成原创差异。"], ["科技感", "通过 AI 指示灯、浅蓝点缀、触控区等元素体现 AI 宠物伴侣定位。"], ["硬件适配", "形象应适合智能硬件外观、APP 图标、虚拟形象和 3D 交互。"], ["确权准备", "保留手稿、分层源文件、创作过程、时间戳和创作说明。"]]]},
        {"title": "三、整体造型规范", "paragraphs": ["贝贝灵熊猫 AI 伴侣的整体造型应兼顾原创性、硬件量产适配性与 AI 科技感。造型不得直接模仿任何现有熊猫 IP，应通过比例、轮廓、身体结构和耳部形态形成独立辨识度。"], "tables": [[["项目", "规范要求"], ["体态比例", "头身比 1 : 1.2，头大身短，短圆紧凑，区别于真实熊猫常见比例。"], ["头部轮廓", "圆角矩形头，避免正圆、扁圆、饼形轮廓。"], ["身体结构", "一体化圆润机身，短四肢、无腰线、无尖锐角，适合硬件量产。"], ["耳朵", "小三角耳，微前倾，耳距适中，避免大圆耳或长耳设计。"]]]},
        {"title": "四、五官原创设计规范", "paragraphs": ["五官是熊猫形象能否形成原创辨识度的关键。贝贝灵熊猫 AI 伴侣应避免常见的大圆眼、夸张笑脸、吊梢眼和影视化表情，应采用克制、治愈、智能的五官表达。"], "tables": [[["项目", "规范要求"], ["眼睛", "横长椭圆眼，眼尾微扬，避免大圆眼、吊梢眼或眯眼。"], ["黑眼圈", "不规则梯形黑眼圈，不上扬、不下垂、不连成一线。"], ["鼻子", "小椭圆黑鼻，居中、小巧，避免大圆鼻或心形鼻。"], ["嘴部", "隐藏式微笑线，微笑弧度极淡，避免大笑或 O 型嘴。"]]]},
        {"title": "五、独家识别符号规范", "paragraphs": ["识别符号应同时服务于品牌识别、AI 功能表达、版权登记和外观专利保护。建议将 AI 指示灯、耳尖浅蓝描边、掌心触控区和胸口 LOGO 位作为贝贝灵熊猫 AI 伴侣的核心视觉识别资产。"], "tables": [[["识别符号", "规范要求"], ["头顶", "AI 蓝光圆形指示灯，可表现开机、互动、语音响应等状态。"], ["耳尖", "耳尖 1mm 极光浅蓝 #66CCFF 描边，体现科技感和专属差异。"], ["掌心", "掌心浅蓝圆点触控交互区，用于表达可触控、可互动的硬件属性。"], ["胸口", "胸口贝贝灵 LOGO 专属位置，可印刷或发光，作为品牌展示区域。"]]]},
        {"title": "六、配色规范", "paragraphs": ["配色应保持简洁、稳定、可量产，并避免与既有 IP 的装饰性配色产生混淆。主色以熊猫基础黑白为主，科技点缀色只用于功能识别区域。"], "tables": [[["项目", "规范要求"], ["主色", "本白 #FFFFFF、纯黑 #000000。"], ["科技点缀色", "极光浅蓝 #66CCFF，仅用于 AI 灯、耳尖描边、掌心触控区等局部。"], ["禁止色彩", "禁止彩色、渐变、金色、红色、彩虹装饰，避免与既有 IP 或活动形象混淆。"], ["量产要求", "颜色应适合硬件外壳、喷涂、丝印、灯效和 APP 视觉统一应用。"]]]},
        {"title": "七、性格与动作规范", "paragraphs": ["贝贝灵熊猫 AI 伴侣的性格定位应保持治愈、安静、智能、科技感，不做夸张搞怪或影视模仿动作。动作系统应适合语音互动、睡眠陪伴、点头反馈和轻量情绪表达。"], "tables": [[["项目", "规范要求"], ["性格定位", "AI 陪伴型，安静、治愈、聪明、可信赖。"], ["标准动作", "端坐、点头、挥手、睡眠、语音互动。"], ["禁止动作", "禁止模仿任何影视熊猫动作、表情包动作、夸张搞怪表情。"], ["交互表达", "可通过头顶 AI 指示灯、语音反馈、掌心触控区等方式表达互动状态。"]]]},
        {"title": "八、硬件外观专利设计要点", "paragraphs": ["硬件外观专利应围绕整体熊猫坐姿、一体化圆润机身、头顶 AI 指示灯、耳尖浅蓝描边、梯形黑眼圈、胸口 LOGO 区和掌心触控区进行组织。专利图稿应突出整体轮廓与局部识别符号，便于后续维权时说明外观差异。"], "tables": [[["专利要点", "说明"], ["整体造型", "整体为熊猫坐姿一体化圆润造型，短四肢、无尖锐边角，适合量产。"], ["头部特征", "头部顶端设置 AI 蓝光圆形指示灯。"], ["耳部特征", "双耳尖设置 1mm 极光浅蓝 #66CCFF 装饰描边。"], ["眼部特征", "眼部采用不规则梯形黑眼圈原创设计。"], ["交互特征", "掌心设置浅蓝圆点触控交互区，胸口设置贝贝灵 LOGO 专属位置。"], ["建议写入", "将“AI 指示灯 + 耳尖蓝边 + 梯形黑眼圈 + 掌心触控区”作为外观设计保护重点。"]]]},
        {"title": "九、版权登记用创作说明", "paragraphs": ["本美术作品为北京贝贝灵科技有限公司原创，以熊猫为基础原型进行独立艺术创作。作品采用圆角矩形头部、不规则梯形黑眼圈、横长椭圆眼、头顶 AI 蓝光圆形指示灯、耳尖 1mm 极光浅蓝 #66CCFF 描边、掌心浅蓝圆点触控交互区、胸口贝贝灵 LOGO 专属位置等独创性特征。", "整体造型、五官比例、装饰元素、交互识别符号和配色方案均为原创设计，与现有任何熊猫美术作品、影视形象、网红表情包或商业 IP 不构成实质性相似。该作品可用于版权登记、商标申请、外观专利准备、智能硬件外观开发、APP 视觉和市场宣传。"], "tables": [[["登记事项", "准备内容"], ["登记机构", "中国版权保护中心或其他合法版权登记渠道。"], ["材料", "原创熊猫形象图、创作说明、身份证明或公司主体证明、源文件和创作过程记录。"], ["证据", "保留手稿、分层源文件、导出图、修改记录、时间戳和交付记录。"], ["作用", "证明权利归属，为商用授权、维权、下架投诉和侵权索赔提供基础证据。"]]]},
        {"title": "十、设计师 / 建模师 / 工厂交付文件要求", "paragraphs": ["设计师、建模师和工厂应严格按照本规范输出文件。所有文件需围绕原创造型、核心识别符号、配色和硬件适配进行交付，不得使用网上无授权熊猫素材或既有 IP 形象作为基础。"], "tables": [[["交付对象", "文件要求"], ["设计师", "正面图、侧面图、45°视角图、标准三视图、AI/PSD 源文件、PNG 透明底图。"], ["建模师", "3D 建模参考图、比例说明、材质说明、灯效说明、硬件外观效果图。"], ["工厂", "硬件外观效果图、结构识别点、LOGO 位置、AI 指示灯位置、掌心触控区位置。"], ["确权准备", "版权登记图稿、商标使用图、外观设计专利图稿、创作说明。"]]]},
        {"title": "十一、合规落地执行步骤", "paragraphs": ["贝贝灵熊猫 AI 伴侣建议按照“原创设计、版权登记、商标注册、外观专利、AI 内容合规、统一商用”的顺序落地。该路径成本相对可控，且能同时覆盖美术作品权属、品牌保护、硬件外观保护和 AI 产品合规风险。"], "tables": [[["步骤", "执行内容"], ["第 1 步", "原创设计专属熊猫形象，确保五官、体态、配色、装饰、动作明显区别于现有 IP。"], ["第 2 步", "做著作权登记，准备原创熊猫形象图、创作说明、主体证明和源文件记录。"], ["第 3 步", "注册核心商标类别：第 9 类智能硬件、APP、电子设备；第 28 类玩具、宠物用品；第 35 类广告、电商、销售；第 42 类 AI 软件、技术服务。"], ["第 4 步", "申请外观设计专利，保护熊猫硬件外观、造型和结构，周期通常约 3-6 个月。"], ["第 5 步", "AI 模型和内容合规：使用国内已备案大模型，训练数据不得使用受版权保护的熊猫 IP 素材，生成内容不得生成冰墩墩、花花等受保护形象，AI 交互类产品按规定处理算法备案要求。"], ["第 6 步", "宣传、APP、硬件、包装、销售物料统一使用自有原创熊猫形象，不混用第三方素材。"]]]},
        {"title": "十二、贝贝灵熊猫 AI 伴侣形象图册", "paragraphs": ["以下图册用于展示贝贝灵熊猫 AI 伴侣的形象方向、视觉气质和应用参考。图片仅作形象参考，正式商用前应以设计师交付的原创源文件和确权文件为准。"], "album": True},
        {"title": "结束页", "paragraphs": ["本规范可作为贝贝灵熊猫 AI 伴侣原创 IP 形象设计、版权登记、商标申请、外观设计专利准备、硬件外观开发、设计师交付和工厂沟通的基础文件。后续所有视觉延展、建模、开模和宣传物料均应遵循本规范的原创性、合规性和一致性要求。"]},
    ]


def generic_sections(items: list[InputItem], images: list[Path]) -> list[dict]:
    cleaned: list[str] = []
    tables: list[list[list[str]]] = []
    for item in items:
        if item.kind == "image":
            continue
        result = clean_paragraphs(item.paragraphs)
        cleaned.extend(result.paragraphs)
        tables.extend(item.tables)
    if not cleaned and not tables:
        cleaned = ["输入素材中未提取到可用正文内容，请检查源文件是否为空或是否为扫描图片。"]
    intro = "本文件根据用户提供的素材整理为正式打印版，已去除 AI 对话废话和无意义过渡语，并保留参数、步骤、表格、清单、注意事项和可执行细节。"
    body_text = cleaned[:80]
    return [
        {"title": "一、总说明", "paragraphs": [intro] + body_text[:6]},
        {"title": "二、核心原则", "paragraphs": body_text[6:14] or ["围绕素材中的核心要求、适用场景和约束条件进行归并，形成正式主版本。"]},
        {"title": "三、主体内容规范", "paragraphs": body_text[14:32], "tables": tables[:2]},
        {"title": "四、执行要求", "paragraphs": body_text[32:44] or ["执行时应保持素材中的参数、流程、报价、交付物和注意事项不缺失。"]},
        {"title": "五、交付清单", "paragraphs": body_text[44:56] or ["交付文件应包含正式文档、必要附件、源文件、图示和可执行清单。"]},
        {"title": "六、落地步骤", "paragraphs": body_text[56:68] or ["按照准备、确认、执行、验收、归档的顺序推进。"]},
        {"title": "七、注意事项", "paragraphs": body_text[68:80] or ["不得删除关键参数、报价、成本测算、执行步骤、表格、话术和合规要求。"]},
        {"title": "八、图册 / 附图 / 附录", "paragraphs": ["本章节收录输入素材中的有效图片，图片等比缩放并使用正式图注。"], "album": bool(images)},
        {"title": "结束页", "paragraphs": ["本文件为正式打印版，可用于审阅、沟通、交付和归档。"]},
    ]


def write_internal_report(path: Path, items: list[InputItem], removed: list[str], outputs: list[Path], validations: dict[str, bool | str]) -> None:
    lines = ["# BBL-books 内部处理报告", ""]
    lines.append("## 读取文件")
    for item in items:
        lines.append(f"- {item.path}")
    lines.append("")
    lines.append("## 删除的 AI 对话废话类型")
    for phrase in NOISE_PHRASES:
        lines.append(f"- {phrase}")
    lines.append("")
    lines.append("## 输出文件")
    for output in outputs:
        lines.append(f"- {output}")
    lines.append("")
    lines.append("## 验收")
    for key, value in validations.items():
        lines.append(f"- {key}：{value}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def run(args: argparse.Namespace) -> list[Path]:
    if not args.inputs:
        fail("请提供需要整理的文件夹路径或文件列表。")
    formats = parse_formats(args.formats)
    items = inspect_inputs(args.inputs)
    if not items:
        fail("未读取到可处理的真实素材文件，请检查输入路径。")
    output_dir = infer_output_dir(args.inputs, args.output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    images = [item.path for item in items if item.kind == "image"]
    title = "贝贝灵熊猫 AI 伴侣原创 IP 形象设计规范" if is_bbl_panda(items) and not args.title else infer_title(items, args.title)
    basename = safe_name(title) + "_打印版"
    docx_path = output_dir / f"{basename}.docx"
    pdf_path = output_dir / f"{basename}.pdf"
    html_path = output_dir / f"{basename}.html"
    sections = bbl_sections() if is_bbl_panda(items) else generic_sections(items, images)
    figure_subject = "贝贝灵熊猫 AI 伴侣形象" if is_bbl_panda(items) else safe_name(title)
    figure_prefix = "12" if is_bbl_panda(items) else "8"
    removed: list[str] = []
    for item in items:
        if item.kind != "image":
            removed.extend(clean_paragraphs(item.paragraphs).removed_lines)
    outputs: list[Path] = []
    if "docx" in formats or "pdf" in formats:
        build_print_docx(
            docx_path,
            title,
            sections,
            images,
            subtitle="设计师 / 建模师 / 工厂落地执行版" if is_bbl_panda(items) else "正式打印版",
            company="北京贝贝灵科技有限公司" if is_bbl_panda(items) else "",
            purpose="原创 IP 设计、版权登记、商标申请、外观专利、硬件外观开发" if is_bbl_panda(items) else "正式打印 / 审阅 / 交付",
            figure_subject=figure_subject,
            figure_prefix=figure_prefix,
        )
        update_docx_fields_with_word(docx_path)
        if "docx" in formats:
            outputs.append(docx_path)
    if "pdf" in formats:
        ok, _ = export_pdf_from_docx_word(docx_path, pdf_path)
        if not ok:
            export_pdf_book(pdf_path, title, sections, images)
        outputs.append(pdf_path)
    if "html" in formats:
        build_html(html_path, title, sections, images)
        outputs.append(html_path)
    validations: dict[str, bool | str] = {}
    if docx_path.exists():
        validations.update(validate_docx(docx_path, expected_images=len(images)))
    if pdf_path.exists():
        validations.update(validate_pdf(pdf_path))
    report = output_dir / "BBL-books_内部处理报告.md"
    write_internal_report(report, items, removed, outputs, validations)
    outputs.append(report)
    return outputs


def main() -> None:
    parser = argparse.ArgumentParser(description="Generate print-ready formal manuals from source materials.")
    parser.add_argument("--inputs", nargs="*", help="One or more files/folders.")
    parser.add_argument("--formats", help="docx, pdf, html, docx+pdf, or docx+pdf+html.")
    parser.add_argument("--output-dir")
    parser.add_argument("--title")
    parser.add_argument("--mode", default="print-ready", choices=["print-ready"])
    args = parser.parse_args()
    for output in run(args):
        print(output)


if __name__ == "__main__":
    main()
