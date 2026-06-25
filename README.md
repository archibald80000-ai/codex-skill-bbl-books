# codex-skill-bbl-books

`bbl-books` 是一个 Codex Skill，用于把文件夹或多个素材文件整理成可直接打印的正式手册、书册或报告。适合老板资料、供应商资料、AI 对话素材、产品方案、培训资料、合同附件、图片图册等整理与排版场景。

## 安装

### Windows

```powershell
git clone https://github.com/archibald80000-ai/codex-skill-bbl-books.git "$env:USERPROFILE\.codex\skills\bbl-books"
```

### macOS / Linux

```bash
git clone https://github.com/archibald80000-ai/codex-skill-bbl-books.git ~/.codex/skills/bbl-books
```

安装后确认：

```text
~/.codex/skills/bbl-books/SKILL.md
```

并且 `SKILL.md` 中包含：

```yaml
name: bbl-books
```

## 使用

在 Codex 中调用：

```text
Use $bbl-books to turn <输入素材路径> into a print-ready formal manual. Output docx+pdf.
```

中文示例：

```text
使用 $bbl-books，把 E:\project\materials 整理成可打印的正式手册，输出 docx+pdf。
```

## 说明

正式输出正文不会包含清洗日志、素材清单、错误信息、文件路径或内部处理说明；内部处理报告会单独保存。

