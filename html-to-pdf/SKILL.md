---
name: html-to-pdf
description: "HTML+CSS 转 PDF（用无头 Chromium/Chrome/Edge，浏览器级还原渐变、卡片、阴影、配色、Web 字体）。
  把设计精美的 HTML 报告/海报/发票/证书/页面转成 PDF，或把任意网页 HTML 导出为 PDF。
  用于：需要漂亮排版的 PDF、HTML 模板出 PDF、网页转 PDF、把 .html 文件渲染成 PDF。
  操作：HTML转PDF、网页转PDF、HTML导出PDF、html渲染PDF、html文件转pdf。
  关键词：HTML转PDF, 网页转PDF, html to pdf, web page to pdf, 渲染, 打印, 美观报告, 海报, 发票, 证书, css, chromium, headless。"
required_tools:
  - bash
recommended_tools:
  - read
  - write
---
# html-to-pdf

用系统**无头浏览器**（Chromium/Chrome/Edge）把 HTML+CSS 渲染成 PDF，
能完整还原渐变、卡片、阴影、Flex/Grid 布局、Web 字体等「网页级」效果。
适合做精美报告/海报/发票/证书。若只是数据表格类规整文档、且想免浏览器，用 `data-to-pdf`。

## 适用场景
- 用户说「把这个 HTML/网页转成 PDF」「用 HTML 模板生成漂亮的 PDF」
- 需要渐变背景、卡片、复杂 CSS、emoji、Web 字体的 PDF
- 已有 `.html` 文件或一段 HTML 源码，要导出 PDF

## 依赖
- 本项目沙箱镜像已预装 `chromium`，无需安装。平台内通常不要手动传 `--browser`；本机调试时脚本也会自动探测浏览器。
- 中文需系统中文字体（镜像已装 `fonts-noto-cjk`）。

## 使用方法
```bash
python3 /skills/html-to-pdf/scripts/html2pdf.py /workspace/report.html -o /workspace/report.pdf
```
显式指定浏览器（一般不需要，自动探测）：
```bash
python3 /skills/html-to-pdf/scripts/html2pdf.py /workspace/report.html -o /workspace/report.pdf --browser /usr/bin/chromium
```

### 参数表
| 参数 | 默认 | 说明 |
|------|------|------|
| `html`（位置参数） | — | HTML 文件路径，放 `/workspace` |
| `-o, --out` | — | 输出 PDF 路径，**指向 `/workspace/...`** |
| `--browser` | 自动探测 | 显式指定浏览器可执行文件的完整路径；平台沙箱中如需指定，使用 `/usr/bin/chromium` |
| `--wait` | 2000 | 渲染等待毫秒，给 Web 字体/布局留时间；仅 `--headless=new` 主路径生效，极旧版浏览器回退模式下被忽略 |

## 写 HTML 时的硬性约定
- **纸张/页边距用 CSS @page**：`@page { size: A4; margin: 12mm; }`
- **要打印背景色/底纹必须加**：
  ```css
  * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  ```
- HTML 尽量**自包含**（内联 CSS、本地图片/base64），别依赖联网资源——沙箱网络受限。

## 输出说明
- 生成单个 PDF（`-o` 指定）；页码由 CSS/内容控制，浏览器默认页眉页脚已关闭。
- 配合 `pdf-to-image` 技能可把结果再转成图片预览。

## 注意事项 ⚠️
- **内存**：无头 Chromium 较吃内存，容器默认 512MB **可能 OOM**。建议把 `AP_SANDBOX_MEMORY` 调到 `1g+`；
  脚本已带容器友好参数 `--no-sandbox --disable-dev-shm-usage --disable-gpu`。
- **emoji**：镜像已装 `fonts-noto-color-emoji`；若当前运行中的沙箱镜像较旧，emoji 可能显示成方块，需要重建镜像。
- 输入 HTML、输出 PDF 都放 `/workspace`；`/skills` 只读。
- 容器为非 root（`apuser`），脚本已用 `--no-sandbox`，无需 root/sudo。
