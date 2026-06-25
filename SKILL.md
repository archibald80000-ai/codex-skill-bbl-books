---
name: bbl-books
description: 将用户提供的文件夹或多个文件整理为可直接打印的正式手册/书册/报告；适用于老板资料、供应商资料、AI 对话素材、产品方案、培训资料、合同附件、图片图册等素材的清洗、去重、归并和专业排版。必须要求用户提供输入路径和输出格式；正文不得包含清洗日志、素材清单、错误信息、文件路径或内部处理说明。
---

# BBL-books

## 定位

作为高级文档排版专家 / 打印版手册生成专家，把用户提供的零散文件素材整理成一份可以直接打印、直接发给老板/客户/设计师/供应商的正式成品文档。

输出不是草稿、处理报告、素材清单或开发日志。正式 DOCX/PDF/HTML 正文只包含成品内容；内部处理报告必须单独保存为 `BBL-books_内部处理报告.md`。

## 使用前置检查

执行前必须先检查用户是否提供：

1. 输入素材路径：一个文件夹、多个文件、或文件夹加补充文件。
2. 输出格式：`docx`、`pdf`、`docx+pdf`、`html`、`docx+pdf+html`。

如果没有输入路径，先询问：

```text
请提供需要整理的文件夹路径或文件列表。
```

如果没有输出格式，先询问：

```text
请指定输出格式：docx、pdf、docx+pdf、html 或 docx+pdf+html。
```

不要在用户未指定输出格式时擅自默认生成。

## 命令接口

```powershell
python scripts/bbl_books.py `
  --inputs "E:\path\folder_or_file" `
  --formats docx,pdf `
  --output-dir "E:\path\output" `
  --title "项目资料打印版手册" `
  --mode print-ready
```

多个输入：

```powershell
python scripts/bbl_books.py `
  --inputs "E:\path\a.docx" "E:\path\b.docx" "E:\path\images" `
  --formats docx,pdf `
  --output-dir "E:\path\output" `
  --title "项目资料打印版手册" `
  --mode print-ready
```

输出命名：

- `<文档标题>_打印版.docx`
- `<文档标题>_打印版.pdf`
- `<文档标题>_打印版.html`
- `BBL-books_内部处理报告.md`

## 工作流

1. 用 `scripts/inspect_inputs.py` 扫描真实素材，并忽略旧输出、临时文件、日志和内部报告。
2. 提取 DOCX、MD、TXT、HTML、PDF、XLSX/CSV 和图片内容。
3. 用 `scripts/clean_ai_chatter.py` 删除 AI 对话废话。
4. 保留参数、色值、型号、报价、成本测算、步骤、话术、表格、交付清单、注意事项、合规要求、版本信息和适用场景。
5. 合并多版本重复内容为正式主版本，不按文件名拼接章节。
6. 用 `scripts/build_print_docx.py` 生成打印版 DOCX。
7. 用 `scripts/export_pdf.py` 通过 Word COM 优先导出 PDF；不可用时才回退。
8. 用 `scripts/build_html.py` 生成 HTML（用户指定时）。
9. 用 `scripts/validate_print_output.py` 检查正式正文是否不含内部词、路径、错误信息、图片文件名和尺寸。

## 成品标准

默认结构按素材自动调整，优先参考：

```text
封面
目录
一、总说明
二、核心原则
三、主体内容规范
四、执行要求
五、交付清单
六、落地步骤
七、注意事项
八、图册 / 附图 / 附录
结束页
```

每章应有一段说明，并尽量配表格或条目清单。不要写成碎片堆叠，也不要把细节压缩成泛泛概述。

## 禁止进入正式正文

正式 DOCX/PDF/HTML 正文不得出现清洗日志、素材清单、错误信息、文件路径、图片尺寸、相似度分析、版本差异分析、脚本说明、manifest、旧输出文件内容或临时文件内容。完整禁词见 `references/forbidden_internal_terms.md`。

## 黄金样例

第一版行为对齐本技能内置的黄金样例标准：

```text
项目资料打印版手册_黄金样例.docx
```

详细标准见 `references/golden_sample_notes.md`。

## Resource Map

- `scripts/bbl_books.py` - 主命令入口。
- `scripts/inspect_inputs.py` - 输入扫描与过滤。
- `scripts/clean_ai_chatter.py` - AI 对话废话清理。
- `scripts/build_print_docx.py` - 打印版 DOCX 生成。
- `scripts/export_pdf.py` - Word/PDF 导出。
- `scripts/build_html.py` - HTML 生成。
- `scripts/validate_print_output.py` - 正式输出验收。
- `references/print_document_standard.md` - 打印版排版标准。
- `references/cleaning_rules.md` - 清理规则。
- `references/preservation_policy.md` - 细节保留原则。
- `references/image_album_rules.md` - 图册规则。
- `references/forbidden_internal_terms.md` - 正文禁词。
- `references/golden_sample_notes.md` - 黄金样例说明。
