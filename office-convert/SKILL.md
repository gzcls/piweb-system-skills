---
name: office-convert
description: "用 LibreOffice 转换文档格式并保留原排版（无需 MS Office）。把 Word/Excel/PPT 等 Office 文档转成 PDF，或在 doc↔docx、xls↔xlsx、ppt↔pptx 等格式间互转。
  用于：Word转PDF、Excel转PDF、PPT转PDF、Office批量转PDF、老格式doc/xls/ppt转新格式、文档格式转换。
  操作：Word转PDF、Excel转PDF、PPT转PDF、docx转pdf、文档转换、格式互转、Office转PDF。
  关键词：Word转PDF, Excel转PDF, PPT转PDF, office to pdf, docx to pdf, xlsx to pdf, pptx to pdf, 格式转换, 文档转换, libreoffice, convert。"
required_tools:
  - bash
recommended_tools:
  - read
  - write
---
# office-convert

用 **LibreOffice 无头模式**转换文档格式，**保留原排版**。最常用是「Office → PDF」，
也支持 Office 各格式互转。引擎是镜像预装的 LibreOffice，**不需要 MS Office**。

## 适用场景
- 用户说「把这个 Word/Excel/PPT 转成 PDF」「批量把 Office 转 PDF」
- 老格式转新格式：`.doc→.docx`、`.xls→.xlsx`、`.ppt→.pptx`
- 需要**保留 Office 原排版**的 PDF（区别于 `data-to-pdf` 的数据驱动排版）

## 依赖
- LibreOffice（沙箱镜像已预装 writer/calc/impress）。

## 使用方法
```bash
# Word/Excel/PPT 转 PDF
python3 /skills/office-convert/scripts/office_convert.py /workspace/in.docx --to pdf  -o /workspace
python3 /skills/office-convert/scripts/office_convert.py /workspace/in.xlsx --to pdf  -o /workspace
python3 /skills/office-convert/scripts/office_convert.py /workspace/in.pptx --to pdf  -o /workspace
# 格式互转
python3 /skills/office-convert/scripts/office_convert.py /workspace/old.doc  --to docx -o /workspace
```

### 参数表
| 参数 | 说明 |
|------|------|
| `src`（位置参数） | 源文件，放 `/workspace` |
| `--to` | 目标格式：`pdf`/`docx`/`xlsx`/`pptx`/`odt` 等。⚠️ `png` 对**多页文档只导出第一页**，要整篇转图请改走「先转 PDF → `pdf-to-image`」 |
| `-o, --out` | 输出目录（建议 `/workspace`；默认源同目录） |
| `--soffice` | 显式指定 soffice 可执行文件（一般不需要） |

## 输出说明
- 输出文件名 = 源文件名 + 新扩展名，落在 `--out` 目录。
- 转 PDF 会**尽量还原 Office 排版**（字体、表格、图表、版式）。

## 注意事项
- 输入/输出都放 `/workspace`；`/skills` 只读。脚本已用独立临时 UserInstallation 配置，规避只读 HOME/并发冲突。
- 中文字体依赖镜像的 `fonts-noto-cjk`；若源文档用了特殊字体，LibreOffice 会做近似替换。
- 大文件/复杂 PPT 转换较吃内存（512MB 上限），必要时建议调高 `AP_SANDBOX_MEMORY`。
- 想把文档转成图片：先用本技能转 PDF，再用 `pdf-to-image` 技能转图（不要用 `--to png`，多页只出首页）。
- 与 `data-to-pdf` 的区别：本技能**保留 Office 原排版**；`data-to-pdf` 是用数据从零排版。
- 与 `html-to-pdf` 的区别：**源是 HTML/网页 → 用 `html-to-pdf`**（Chromium 渲染更还原）；本技能用于 Office 二进制文档（doc/xls/ppt 系）。
