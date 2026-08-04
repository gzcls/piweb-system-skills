---
name: qrcode
description: "生成二维码图片（PNG）。把网址或文本生成二维码，可设纠错等级、尺寸、边框、前景/背景色。
  用于：生成二维码、把链接做成二维码、文本转二维码、活动/名片/付款码用二维码图片。
  操作：生成二维码、做二维码、链接转二维码、文本转二维码、二维码图片。
  关键词：二维码, QR码, 生成二维码, qrcode, QR code, 链接二维码, 文本二维码, 扫码。"
required_tools:
  - bash
recommended_tools:
  - read
  - write
---
# qrcode

把网址/文本生成**二维码 PNG**。引擎 Python `qrcode`（+Pillow）。

## 适用场景
- 「把这个链接生成二维码」「做个二维码图片」「文本转二维码」

## 依赖
本技能依赖 `qrcode[pil]`，沙箱镜像已预装。若当前运行中的沙箱镜像较旧，需重建镜像。

## 使用方法
```bash
python3 /skills/qrcode/scripts/qrcode_gen.py "https://example.com" -o /workspace/qr.png
python3 /skills/qrcode/scripts/qrcode_gen.py "文本内容" -o /workspace/qr.png --ecc H --size 12 --border 4
```

### 参数表
| 参数 | 默认 | 说明 |
|------|------|------|
| `data`（位置参数） | — | 二维码内容（网址或文本） |
| `-o, --out` | — | 输出 PNG（建议 `/workspace`） |
| `--ecc` | M | 纠错等级 L/M/Q/H（H 容错最高） |
| `--size` | 10 | 每个点像素(box_size) |
| `--border` | 4 | 边框格数 |
| `--fill` / `--bg` | black / white | 前景 / 背景色；支持颜色名或 `#RRGGBB` |

## 输出说明
- 单个 PNG 二维码。内容含中文也可（按文本编码）。

## 注意事项
- 输入输出在 `/workspace`；`/skills` 只读。
- 若脚本报缺少 `qrcode` 依赖，说明当前运行中的沙箱镜像较旧，重建镜像即可。
- 想加 Logo/美化二维码属于进一步需求，可后续扩展。
