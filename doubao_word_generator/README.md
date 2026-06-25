# BBL-books Word Generator

## 1. 解决什么问题

这个最小服务把用户上传的零散资料整理成可下载的 Word 小册子，并可选导出 PDF。适用于打印版手册、供应商说明书、会议纪要、培训讲义、产品需求文档、合同整理版和图文说明书。

## 2. 为什么 Markdown 版不够

Markdown 版只能在聊天里输出结构化文本，用户还要手动复制到 Word / WPS / 飞书，再自己处理封面、目录、页眉页脚、页码、字体、图片和 PDF。老板或客户场景需要的是可以直接下载、打印、转发的正式 Word / PDF 文件。

## 3. 现在如何实现可下载 Word

豆包或扣子智能体负责收集资料、文档类型和输出格式，然后调用本服务。服务读取文件或 zip，清理 AI 废话，保留关键参数、表格、步骤、报价、交付清单，生成 A4 竖版 DOCX，并按需生成 PDF，最后返回下载链接。

## 4. 本地如何运行

```powershell
cd "path\to\bbl-books\doubao_word_generator"
pip install -r requirements.txt
uvicorn app:app --host 127.0.0.1 --port 8765
```

健康检查：

```powershell
curl http://127.0.0.1:8765/health
```

## 5. 如何测试

本地脚本测试：

```powershell
python scripts/generate_docx.py --input "E:\path\materials" --doc-type manual --formats docx,pdf --title "项目资料打印版手册"
```

接口测试可向 `/generate` 上传 `files` 或 `zip_file`，也可以本地调试时传 `source_path`。

## 6. 后续如何接入扣子 / 豆包

1. 将服务部署到云服务器。
2. 将 `openapi.yaml` 导入扣子自定义插件。
3. 豆包 / 扣子智能体按提示词收集资料、文档类型和输出格式。
4. 调用 `/generate`。
5. 将 `/download/{job_id}/{filename}` 链接返回给用户。

## 7. 老板使用方式

1. 上传资料。
2. 选择类型：打印版手册 / 供应商说明书 / 会议纪要 / 培训讲义 / 产品需求文档 / 合同整理版 / 图文说明书。
3. 选择格式：Word / PDF / Word+PDF。
4. 下载 Word / PDF。
