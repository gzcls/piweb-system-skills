---
name: any-to-markdown
description: "普通文档转 Markdown / 文本抽取。支持有文本内容的 Word/Excel/PPT/PDF/HTML/CSV/JSON 等文档转 Markdown，便于大模型阅读、检索、总结或再加工。用于：文档转Markdown、Office转md、PDF转文本、提取文档内容、document to markdown, text extraction。扫描件、图片OCR、无文本层PDF不要优先用本技能：扫描PDF转Markdown用 pdf-ocr，图片/截图OCR用 smart-ocr；创建或编辑原格式文档用 docx/xlsx/pptx。"
required_tools:
  - bash
recommended_tools:
  - read
  - read_document
  - write
---
# any-to-markdown

把一份文档（Word/Excel/PPT/PDF/HTML/CSV/JSON/图片…）转成 **Markdown 文本**，
方便后续让模型阅读、总结、检索或二次加工。引擎是 `markitdown`。

## 适用场景
- 用户说「提取这个文档的文字 / 把这个 Word/PPT/PDF 转成文本 / 读一下这个文件内容」
- 需要把文档内容喂给大模型做总结、问答、对比、检索
- 想把 Office/PDF 内容转成 Markdown 入库或拼进提示词

## 依赖
- Python 库 `markitdown[all]`。沙箱镜像已预装，支持 Word/Excel/PPT/PDF/HTML/CSV/JSON/图片等常见输入。
  如果当前运行中的沙箱镜像早于本依赖变更，需重建镜像后再使用完整格式支持。

## 使用方法
```bash
# 输出到 .md 文件
python3 /skills/any-to-markdown/scripts/to_markdown.py /workspace/in.docx -o /workspace/out.md
# 直接打印（适合内容不大、想直接读进上下文时）
python3 /skills/any-to-markdown/scripts/to_markdown.py /workspace/in.pdf
```

### 参数表
| 参数 | 说明 |
|------|------|
| `src`（位置参数） | 源文件，放 `/workspace` |
| `-o, --out` | 输出 `.md` 路径；省略则把 Markdown 打印到 stdout |

## 输出说明
- 生成 Markdown 文本：表格转成 md 表格、标题层级尽量保留。
- 不传 `-o` 时直接打印，便于模型立刻读取内容。

## 注意事项
- 输入/输出都放 `/workspace`；`/skills` 只读。
- 普通 Word/PDF/Excel 的在线阅读、分段分析优先用 `read_document`；本技能适合需要落地 Markdown 文件、批量转换或处理 PPT/HTML/CSV/JSON/图片等更宽格式时使用。
- **扫描件 PDF / 图片型 PDF** 用本技能可能提不到文字——那种应改用 `pdf-ocr`（OCR 识别）。
- 转换是「尽力还原文本」，复杂版式/公式/图表可能丢失或简化，适合做内容理解而非像素级还原。
- 超大文档直接打印会撑爆上下文，建议用 `-o` 落文件后按需读取片段。
