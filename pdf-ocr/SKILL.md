---
name: pdf-ocr
description: "PDF OCR 转 Markdown。将扫描件 PDF、无文本层 PDF 或用户明确要求 OCR/文字版/Markdown 文件的 PDF 转换为 Markdown；电子 PDF 会直接提取文本，扫描件会渲染后 OCR。操作：OCR 识别、提取扫描件文字、PDF 转 Markdown、生成文字版。关键词：PDF, OCR, scanned PDF, no text layer, 扫描件, 看不了, 文字版, Markdown, 中文识别, 制度文档。"
required_tools:
  - bash
recommended_tools:
  - read
  - read_document
  - write
---
# PDF OCR 转 Markdown

将 PDF 文件转换为 Markdown 格式的工具。自动检测文档类型并选择最佳处理策略：
- **电子版 PDF**（有文本层）：直接提取文字，速度快
- **扫描版 PDF**（纯图片）：渲染为图片 → OCR 识别 → 后处理格式化

## 适用场景

- `read_document` 提示 PDF 是扫描件、无文本层或解析结果为空
- 用户上传了 PDF 文件并明确要求转换为 Markdown / 文字版
- 用户要求提取扫描件的文字内容
- 用户要求分析 PDF 中的制度、规章、合同等中文文档
- 用户提到 PDF 看不了、打不开、需要文字版

普通电子版 PDF 的阅读、总结、抽取信息优先用 `read_document`；只有需要 OCR、需要落地 Markdown 文件，或 `read_document` 无法提取正文时再用本技能。

## 前置条件

沙箱镜像已预装 OCR 依赖（PyMuPDF + RapidOCR），通常无需运行时安装。脚本仍保留缺失依赖检测，便于非标准环境使用。

如果自动安装失败，手动安装：
```bash
pip install --target /tmp/pdf_ocr_deps PyMuPDF>=1.23.0 rapidocr_onnxruntime==1.2.13
```

## 使用方法

### 基本转换

```bash
python3 /skills/pdf-ocr/scripts/ocr_convert.py <输入PDF路径> <输出Markdown路径>
```

**示例：**
```bash
python3 /skills/pdf-ocr/scripts/ocr_convert.py /workspace/document.pdf /workspace/document.md
```

### 完整参数

```bash
python3 /skills/pdf-ocr/scripts/ocr_convert.py <input.pdf> <output.md> [选项]
```

| 参数 | 默认值 | 说明 |
|------|--------|------|
| `--dpi N` | 150 | 渲染分辨率。内存不足时降至 100 或 72 |
| `--no-format` | - | 跳过中文制度文档格式化（章/条/节自动标题） |
| `--no-correction` | - | 跳过 OCR 纠错词典 |
| `--corrections PATH` | - | 自定义纠错词典 JSON 路径 |
| `--raw-dir DIR` | - | 保存每页 OCR 原文到指定目录（便于人工校对） |
| `--confidence-threshold N` | 60 | 低置信度词阈值 (0-100)，低于此值的词会在报告中标出 |
| `--text-threshold N` | 50 | 每页平均字符数阈值，低于此判定为扫描件 |

### 常用场景命令

**1. 标准转换（推荐）**
```bash
python3 /skills/pdf-ocr/scripts/ocr_convert.py /workspace/input.pdf /workspace/output.md
```

**2. 保存 OCR 原文便于校对**
```bash
python3 /skills/pdf-ocr/scripts/ocr_convert.py /workspace/input.pdf /workspace/output.md --raw-dir /workspace/ocr_raw
```

**3. 内存受限环境（降低 DPI）**
```bash
python3 /skills/pdf-ocr/scripts/ocr_convert.py /workspace/input.pdf /workspace/output.md --dpi 100
```

**4. 纯文本提取（不格式化）**
```bash
python3 /skills/pdf-ocr/scripts/ocr_convert.py /workspace/input.pdf /workspace/output.md --no-format --no-correction
```

**5. 使用自定义纠错词典**
```bash
python3 /skills/pdf-ocr/scripts/ocr_convert.py /workspace/input.pdf /workspace/output.md --corrections /workspace/my_corrections.json
```

## 输出说明

### Markdown 输出
- 电子版 PDF：提取原文，保留段落结构
- 扫描件 PDF：OCR 识别后自动处理：
  - 移除重复的页眉、页脚、页码
  - 合并因扫描断裂的行
  - 第X章 → `## 二级标题`
  - 第X节 → `### 三级标题`
  - 第X条 → `**条款加粗**`
  - 应用纠错词典修正常见 OCR 错误

### 处理报告（stderr）
脚本会在 stderr 输出详细处理报告，包括：
- 文档类型（电子版/扫描件）
- 每页识别结果和置信度
- 纠错修正记录
- 低置信度词列表（需人工校对的部分）

## 纠错词典格式

自定义纠错词典为 JSON 文件：

```json
{
  "corrections": {
    "错误词": "正确词",
    "已有": "己有",
    "每年未": "每年末"
  },
  "domain_terms": [
    "档案管理员",
    "归档"
  ]
}
```

内置词典位于 `/skills/pdf-ocr/scripts/corrections.json`，管理员可编辑。

## DPI 选择建议

| 文档质量 | 推荐 DPI | 说明 |
|----------|----------|------|
| 清晰扫描件 | 150 | 默认值，识别率最好 |
| 普通扫描件 | 120-150 | 速度与质量平衡 |
| 内存受限 (<512MB) | 72-100 | 牺牲部分识别率换取稳定性 |
| 模糊扫描件 | 200 | 提升识别率，但内存占用更高 |

## 注意事项

1. **依赖**在本项目沙箱镜像中已预装；若运行在非标准环境，首次运行可能尝试安装依赖
2. **内存限制**：默认 150dpi 下每页约占 50-100MB 内存。如进程被杀 (Killed)，降低 `--dpi`
3. **中文识别**使用 PP-OCRv3 模型，对清晰扫描件识别率约 85-92%
4. **纠错词典**仅做简单字符串替换，复杂上下文修正需人工校对
5. 建议使用 `--raw-dir` 保存原始 OCR 结果，便于出错时追溯
