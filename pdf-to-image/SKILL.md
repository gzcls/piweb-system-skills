---
name: pdf-to-image
description: "PDF 转图片。把 PDF 的每一页渲染成 PNG/JPG 等图片，也可抽取 PDF 内嵌的原始图片对象。
  用于：把 PDF 变成可预览/可发送的图片、给 PDF 做缩略图、把扫描件/合同/报告/图纸截成图、提取 PDF 里的插图。
  操作：PDF转图片、PDF转PNG、PDF转JPG、PDF生成图片、PDF每页截图、PDF导出图片、提取PDF图片、PDF做缩略图。
  关键词：PDF, 转图片, 转PNG, 转JPG, 截图, 图片, 渲染, 缩略图, 扫描件, 提取图片, 每页, pdf to image, pdf to png/jpg, rasterize, render, thumbnail, extract images。"
required_tools:
  - bash
recommended_tools:
  - read
  - write
---
# pdf-to-image

把 PDF 转成图片。默认把**每一页整页渲染**为位图（所见即所得，含文字/表格/矢量）；
加 `--extract` 则改为**抽取页面里内嵌的原始图片对象**（只取嵌入原图，不含文字）。

## 适用场景
- 用户说「把这个 PDF 转成图片 / 转成 PNG / 每页存一张图 / 做缩略图」
- 用户想把扫描件、合同、报告、图纸**截成图片**预览或发送
- 用户想**提取 PDF 里的插图/照片**（用 `--extract`）

## 依赖
- Python 库 `PyMuPDF`（沙箱镜像已预装，无需安装）。纯 Python，不需要 Poppler/Ghostscript/ImageMagick。

## 使用方法
### 整页渲染（最常用）
```bash
python3 /skills/pdf-to-image/scripts/pdf2img.py /workspace/input.pdf -o /workspace/out --dpi 200
```
### 只转部分页 / 指定格式
```bash
python3 /skills/pdf-to-image/scripts/pdf2img.py /workspace/input.pdf -o /workspace/out --pages 1-3,5 --format jpg
```
### 抽取内嵌原始图片
```bash
python3 /skills/pdf-to-image/scripts/pdf2img.py /workspace/input.pdf -o /workspace/imgs --extract --min-w 100 --min-h 100
```

### 参数表
| 参数 | 默认 | 说明 |
|------|------|------|
| `pdf`（位置参数） | — | 源 PDF 路径，放 `/workspace` 下 |
| `-o, --out` | `<pdf名>_images` | 输出目录，**务必指向 `/workspace/...`** |
| `--dpi` | 144 | 分辨率；预览 150~200，打印 300 |
| `--zoom` | — | 缩放倍数（覆盖 `--dpi`），factor=zoom×72，如 2.0 = 144DPI |
| `--format` | png | png/jpg/webp/tif... |
| `--pages` | 全部 | 如 `1-3,5,8-10` |
| `--prefix` | `page_` | 文件名前缀 |
| `--alpha` | 关 | 保留透明通道；仅 png 生效，jpg/jpeg 会自动忽略 |
| `--extract` | 关 | 改为抽取内嵌原图 |
| `--min-w` / `--min-h` | 0 | 抽取模式下过滤小图（像素） |

## 输出说明
- 渲染：`out/page_001.png`、`page_002.png` …（按页序，零填充）。
- 抽取：`imgs/img_p001_01.png` …（`p页号_序号`）。
- 脚本会把生成的每个文件路径打印到 stdout，最后一行汇总数量。

## 注意事项
- 输入输出都放 `/workspace`；技能目录 `/skills` 是**只读**的，不能往里写。
- DPI 越高越清晰也越吃内存/磁盘。容器内存上限默认 512MB，**超大/超多页 PDF 用高 DPI 可能 OOM**——可降 `--dpi` 或用 `--pages` 分批处理。
- 「渲染」得到整页位图；「抽取」只得到嵌入的原图，二者用途不同，按需选择。
