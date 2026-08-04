---
name: image-toolkit
description: "图片处理工具箱：缩放、转格式、压缩、生成缩略图、加文字水印、查看信息。
  用于：压缩图片、改图片尺寸、转PNG/JPG/WebP、批量做缩略图、给图片加水印、看图片分辨率。
  操作：压缩图片、缩放图片、调整尺寸、转格式、PNG转JPG、做缩略图、图片加水印、查看图片信息。
  关键词：图片, 压缩, 缩放, 调整尺寸, 转格式, PNG, JPG, WebP, 缩略图, 水印, thumbnail, watermark, resize, convert, compress, sharp, image。"
required_tools:
  - bash
recommended_tools:
  - read
  - write
---
# image-toolkit

图片常用处理（子命令式），引擎是 Node 的 `sharp`。缩放 / 转格式 / 压缩 / 缩略图 / 文字水印 / 看信息。

## 适用场景
- 「把这张图压小一点 / 改成 800 宽 / 转成 WebP」
- 「批量生成缩略图」「给图片加个水印」「看下这图多大分辨率」

## 依赖
- Node 包 `sharp`（`docker/Dockerfile.sandbox` 已声明全局安装）。脚本会自动从 `npm root -g` 解析全局模块。
  如果脚本提示找不到 `sharp`，说明当前运行中的沙箱镜像较旧，需要重建镜像。
- 文字水印的中文依赖系统 CJK 字体（镜像 `fonts-noto-cjk`）。

## 使用方法
```bash
# 单文件
node /skills/image-toolkit/scripts/img_tool.js resize    /workspace/in.png -o /workspace/out.png --width 800
node /skills/image-toolkit/scripts/img_tool.js convert   /workspace/in.png -o /workspace/out.webp
node /skills/image-toolkit/scripts/img_tool.js compress  /workspace/in.jpg -o /workspace/out.jpg --quality 70
node /skills/image-toolkit/scripts/img_tool.js thumbnail /workspace/in.png -o /workspace/thumb.png --width 200
node /skills/image-toolkit/scripts/img_tool.js watermark /workspace/in.png -o /workspace/wm.png --text 机密
node /skills/image-toolkit/scripts/img_tool.js info      /workspace/in.png

# 批量（--outdir：对多个文件/通配符逐个处理，输出同名文件到目录）
node /skills/image-toolkit/scripts/img_tool.js resize  "/workspace/imgs/*.png" --outdir /workspace/out --width 800
node /skills/image-toolkit/scripts/img_tool.js convert "/workspace/imgs/*.png" --outdir /workspace/out --ext webp --quality 70
# 若提示找不到模块：NODE_PATH=$(npm root -g) node ...
```

> 批量：用 `--outdir` 代替 `-o`，输入可给多个文件或通配符（仅支持单层 `目录/*.ext`，不支持 `**` 递归；脚本兜底展开）；输出同名文件落到该目录。`convert` 批量改扩展名用 `--ext webp`。`resize`/`convert`/`compress`/`thumbnail`/`watermark`/`info` 都支持批量，例如批量加水印：
> ```bash
> node /skills/image-toolkit/scripts/img_tool.js watermark "/workspace/imgs/*.jpg" --outdir /workspace/wm --text 机密
> ```

### 子命令与参数
| 子命令 | 关键参数 | 说明 |
|--------|----------|------|
| `resize` | `--width` `--height` `--fit`(inside/cover/contain) | 默认 `inside` 等比缩放且不放大；`--fit cover` 会裁剪、`contain` 留边 |
| `convert` | （输出扩展名决定格式）`--quality` | png/jpg/webp/avif 互转；jpg/webp/avif 可加 `--quality`（默认 80） |
| `compress` | `--quality`(默认80) | 按扩展名压缩 |
| `thumbnail` | `--width`(默认200) `--height` | 居中裁剪缩略图 |
| `watermark` | `--text` | 右下角文字水印（支持中文） |
| `info` | — | 打印格式/宽高/通道 |

## 输出说明
- 输出格式由 `-o` 的扩展名决定；jpg/webp/avif 用 `--quality`，png 用最高压缩级别。
- `info` 把元数据以 JSON 打印到 stdout。

## 注意事项
- 输入/输出都放 `/workspace`；`/skills` 只读。
- 批量处理：直接用内建 `--outdir`（见上方「批量」），无需在 bash 里手写循环。
- 超大图（如上亿像素）吃内存（512MB 上限），必要时先 resize 再做其它操作。
