---
name: pdf-toolkit
description: "PDF 常用操作工具箱：合并、拆分、旋转、加水印、加密码、提取文字、提取表格为 CSV。
  用于：合并多个PDF、拆分PDF、给PDF加水印、PDF加密、旋转页面、提取PDF文字、把PDF表格导出成表格/CSV。
  操作：合并PDF、拆分PDF、PDF加水印、PDF加密、旋转PDF、提取PDF文字、提取PDF表格、PDF转CSV。
  关键词：PDF合并, 合并PDF, 拆分PDF, PDF水印, 加水印, PDF加密, 密码, 旋转, 提取文字, 提取表格, PDF转CSV, merge split watermark encrypt pdf, extract text table。"
required_tools:
  - bash
recommended_tools:
  - read
  - write
---
# pdf-toolkit

PDF 的日常操作工具箱（子命令式）。合并/拆分/旋转/水印/加密用 `pypdf`，
水印文字用 `reportlab`（支持中文），提表格用 `pdfplumber`。

## 适用场景
- 「把这几个 PDF 合并成一个」「把这个 PDF 按页拆开」
- 「给 PDF 加个『机密』水印」「给 PDF 设个打开密码」
- 「把这页转正/旋转 90 度」「提取 PDF 里的文字 / 表格」

## 依赖
- `pypdf`、`reportlab`、`pdfplumber`（沙箱镜像均已预装）。

## 使用方法
```bash
# 合并
python3 /skills/pdf-toolkit/scripts/pdf_tool.py merge /workspace/a.pdf /workspace/b.pdf -o /workspace/out.pdf
# 拆分（不给 --ranges 则每页一个文件）
python3 /skills/pdf-toolkit/scripts/pdf_tool.py split /workspace/in.pdf -o /workspace/parts --ranges 1-3,4-6
# 旋转（可选 --pages 只转部分页）
python3 /skills/pdf-toolkit/scripts/pdf_tool.py rotate /workspace/in.pdf --angle 90 -o /workspace/r.pdf
# 加文字水印（支持中文）
python3 /skills/pdf-toolkit/scripts/pdf_tool.py watermark /workspace/in.pdf --text 机密 -o /workspace/wm.pdf
# 加密码
python3 /skills/pdf-toolkit/scripts/pdf_tool.py encrypt /workspace/in.pdf --password 1234 -o /workspace/enc.pdf
# 提取文字
python3 /skills/pdf-toolkit/scripts/pdf_tool.py extract-text /workspace/in.pdf -o /workspace/text.txt
# 提取表格为 CSV
python3 /skills/pdf-toolkit/scripts/pdf_tool.py extract-tables /workspace/in.pdf -o /workspace/tables.csv
```

### 子命令与参数
| 子命令 | 关键参数 |
|--------|----------|
| `merge` | `inputs...`（多个 PDF） `-o` |
| `split` | `src` `-o`(目录) `--ranges 1-3,4-6`(可选) |
| `rotate` | `src` `-o` `--angle`(须为 90 的倍数：90/180/270) `--pages`(可选) |
| `watermark` | `src` `-o` `--text 水印文字` |
| `encrypt` | `src` `-o` `--password 口令` |
| `extract-text` | `src` `-o` `--pages`(可选) |
| `extract-tables` | `src` `-o`(CSV) `--pages`(可选) |

## 输出说明
- `split` 输出到目录，文件名 `<名>_part01.pdf`…；其它输出单文件。
- `extract-text` 按页分段；`extract-tables` 写 UTF-8(BOM) CSV，多表用 `# 第x页 表y` 分隔。

## 注意事项
- 输入/输出都放 `/workspace`；`/skills` 只读。
- `extract-text`/`extract-tables` 只对**文字版 PDF** 有效；**扫描件**请用 `pdf-ocr`。
- 页码均为 **1 基**（如 `--pages 1-3,5`）；超出总页数的页码会被**静默忽略**，若全部越界则输出为空。
- `watermark` 中文水印依赖 reportlab 内置 CID 字体 `STSong-Light`；`encrypt` 使用 pypdf 默认加密（4.x 为 AES-128）。
- 把 PDF 转成图片用 `pdf-to-image`；本工具箱不做渲染。
