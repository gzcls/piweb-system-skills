---
name: video-toolkit
description: "视频处理工具箱（基于 ffmpeg）：查看信息、提取音频、裁剪片段、截取封面帧、批量抽帧、转格式、压缩、转 GIF。
  用于：剪视频片段、视频转音频、视频压缩、视频转格式、截取视频封面/缩略图、视频抽帧、视频转GIF、查看视频时长分辨率。
  操作：剪视频、裁剪片段、提取音频、视频转mp3、视频压缩、视频转码、转格式、截封面、抽帧、转GIF、查看视频信息。
  关键词：视频, ffmpeg, 剪辑, 裁剪, 提取音频, 视频转音频, 视频转mp3, 压缩视频, 视频转码, 转格式, 封面, 缩略图, 抽帧, GIF, video, audio, trim, compress, convert, thumbnail。"
required_tools:
  - bash
recommended_tools:
  - read
  - write
---
# video-toolkit

视频常用处理（子命令式），封装系统 **ffmpeg / ffprobe**。

> ⚠️ **资源敏感**：`docker/Dockerfile.sandbox` 已包含 ffmpeg/ffprobe；若当前运行中的沙箱镜像较旧，需重建镜像。
> 重编码类视频任务建议**调高沙箱内存**（见「内存与资源建议」）。

## 适用场景
- 「剪一段视频 / 把视频转成 mp3 / 压缩这个视频 / 转成 mp4」
- 「截个视频封面 / 每隔几秒抽一帧 / 做成 GIF」「看下这视频多长、多大分辨率」

## 依赖
沙箱镜像预装 `ffmpeg` / `ffprobe`。脚本本身只用 Python 标准库，无需 pip 依赖。
如果脚本提示找不到 `ffmpeg`，说明当前运行中的沙箱镜像早于本技能依赖变更，需要重建镜像。

## 使用方法
```bash
# 信息
python3 /skills/video-toolkit/scripts/video_tool.py info /workspace/in.mp4
# 提取音频
python3 /skills/video-toolkit/scripts/video_tool.py audio /workspace/in.mp4 -o /workspace/out.mp3
# 裁剪片段（默认无损 -c copy，秒级、低内存）
python3 /skills/video-toolkit/scripts/video_tool.py trim /workspace/in.mp4 --start 00:00:05 --duration 10 -o /workspace/clip.mp4
# 截封面 / 批量抽帧
python3 /skills/video-toolkit/scripts/video_tool.py thumbnail /workspace/in.mp4 --time 00:00:03 -o /workspace/cover.jpg
python3 /skills/video-toolkit/scripts/video_tool.py frames /workspace/in.mp4 --every 5 -o /workspace/frames
# 转格式 / 压缩 / 转 GIF
python3 /skills/video-toolkit/scripts/video_tool.py convert  /workspace/in.mov -o /workspace/out.mp4
python3 /skills/video-toolkit/scripts/video_tool.py compress /workspace/in.mp4 --height 720 --crf 28 -o /workspace/small.mp4
python3 /skills/video-toolkit/scripts/video_tool.py gif      /workspace/in.mp4 --start 00:00:02 --duration 3 --width 480 -o /workspace/out.gif
```

### 子命令与参数
| 子命令 | 关键参数 | 说明 |
|--------|----------|------|
| `info` | — | ffprobe 输出 JSON（时长/分辨率/编码/码率） |
| `audio` | `--bitrate`(192k) | 提取音频，格式由 `-o` 扩展名决定 |
| `trim` | `--start` `--end`/`--duration` `--reencode` | 默认无损切割（`-c copy` 不重编码）；切点对齐最近关键帧，**起点可能不精确/开头有黑帧**，需精确到帧加 `--reencode` |
| `thumbnail` | `--time` | 指定时刻截一帧 |
| `frames` | `--every N`/`--fps F` | 批量抽帧到目录 |
| `convert` | `--vcodec`/`--acodec`/`--copy` | 转格式/转码 |
| `compress` | `--height`(720) `--crf`(28) `--preset` | 缩分辨率 + 重编码压缩 |
| `gif` | `--start` `--duration` `--width` `--fps` | 片段转 GIF |
| 通用 | `--threads`(默认 1) | 限制线程以控内存/CPU |

## 内存与资源建议 ⚠️
ffmpeg 重编码（convert/compress/gif）比文档类任务吃 **CPU、内存、时间**，且与分辨率强相关：

- **512MB 够用**：`info`、`thumbnail`、`frames`、`audio`、以及 `trim` 默认的无损切割（`-c copy`，几乎不耗内存）。
- **重编码（compress/convert/gif）**：1080p 及以上建议把 `AP_SANDBOX_MEMORY` 调到 **2g**（至少 1g）；并：
  - 先 `--height 720`（或更低）降分辨率，能大幅降内存与耗时；
  - 保持默认 `--threads 1`（线程越多内存越高），机器宽裕再加；
  - `--preset veryfast`（默认）比 slow 省内存/CPU。
- **耗时**：长视频转码可能很慢。bash 工具超时上限 600000ms（10 分钟）；脚本内部也设了 timeout。长任务建议**先 trim 截短再处理**，或分段。
- **磁盘/tmp**：输出落 `/workspace`；大文件注意工作区与 `/tmp` 空间。

## 与其它 skill 的组合链路
视频本身不能直接出图片成品/文档，常见做法是**先用本技能抽出帧/封面，再交给图片或文档技能加工**：

### 1) 视频封面 → 美化/加水印（接 image-toolkit）
```bash
# 截封面
python3 /skills/video-toolkit/scripts/video_tool.py thumbnail /workspace/in.mp4 --time 00:00:03 -o /workspace/cover.jpg
# 压缩 + 加水印
node /skills/image-toolkit/scripts/img_tool.js compress  /workspace/cover.jpg -o /workspace/cover_s.jpg --quality 75
node /skills/image-toolkit/scripts/img_tool.js watermark /workspace/cover_s.jpg -o /workspace/cover_wm.jpg --text 内部资料
```

### 2) 批量抽帧 → 批量转缩略图（frames + image-toolkit 循环）
```bash
# 每 5 秒抽一帧
python3 /skills/video-toolkit/scripts/video_tool.py frames /workspace/in.mp4 --every 5 -o /workspace/frames
# 逐帧生成缩略图（bash 循环）
for f in /workspace/frames/frame_*.jpg; do
  node /skills/image-toolkit/scripts/img_tool.js thumbnail "$f" -o "/workspace/thumbs/$(basename "$f")" --width 240
done
```

### 3) 关键帧 → 做成 PPT（frames → data-to-pptx）
```bash
python3 /skills/video-toolkit/scripts/video_tool.py frames /workspace/in.mp4 --every 10 -o /workspace/frames
# 用抽出的帧拼一个 spec.json（每帧一页 image slide），再：
node /skills/data-to-pptx/scripts/pptx_build.js /workspace/spec.json /workspace/out.pptx
```
> spec 里每页用 `{"layout":"image","title":"00:10","image":"/workspace/frames/frame_0001.jpg"}`。

### 4) 视频 → 音频 → 转写（接已有 pdf-ocr 之外的语音流程）
`audio` 提取出的 mp3/wav 可交给你平台的语音转写能力（若有，做成执行型/独立 skill）。本技能只负责到「出音频文件」。

## 注意事项
- 输入/输出都放 `/workspace`；`/skills` 只读。时间格式支持 `00:00:05` 或秒数。
- 组合链路里的中间产物（封面、帧、缩略图）都落 `/workspace` 子目录，按需清理以省空间。
- 优先用无损操作：剪辑/换封装尽量用 `trim`(默认) / `convert --copy`，避免重编码。
  - 但无损 `trim` 切点会对齐最近关键帧，起点可能不精确或开头有黑帧；要精确切割就加 `--reencode`。
  - `convert --copy` 仅在源编码与目标容器兼容时可用；若失败，去掉 `--copy` 改为重编码。
- 想从视频生成可发的图：`thumbnail`/`frames` 出图后可接 `image-toolkit` 压缩/加水印。
- 本技能依赖较重、资源敏感，建议**仅在确有视频需求时启用**；纯图片用 `image-toolkit`。
