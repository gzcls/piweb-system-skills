# -*- coding: utf-8 -*-
"""
video_tool —— 视频常用处理工具箱（独立脚本，子命令式），封装 ffmpeg / ffprobe。

子命令：
  info        查看视频信息（时长/分辨率/编码/码率）
  audio       提取音频（mp3/wav/aac…）
  trim        裁剪片段（默认无损 -c copy，秒级超快、低内存）
  thumbnail   在某时刻截一帧图片
  frames      按间隔/帧率批量抽帧
  convert     转封装/转码到指定格式
  compress    压缩（缩分辨率 + CRF，重编码）
  gif         片段转 GIF

依赖：系统 ffmpeg / ffprobe（沙箱镜像已预装；若当前运行镜像较旧，需重建）。
脚本本身只用 Python 标准库。

设计取向（控内存/CPU）：
  - trim/convert 在可行时用 -c copy 不重编码；
  - 重编码默认 -preset veryfast、-threads 1（可 --threads 调整），并建议先 --height 降分辨率。
"""
from __future__ import annotations

import json
import os
import shutil
import subprocess


def _bin(name: str) -> str:
    p = shutil.which(name)
    if p:
        return p
    for c in (f"/usr/bin/{name}", f"/usr/local/bin/{name}"):
        if os.path.isfile(c):
            return c
    raise FileNotFoundError(f"未找到 {name}。沙箱镜像应已预装 ffmpeg；若仍缺失，请重建镜像。")


def _run(args, timeout):
    proc = subprocess.run(args, capture_output=True, text=True, timeout=timeout)
    return proc


def _ensure_out(out):
    out = os.path.abspath(out)
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    return out


def info(src, timeout=120):
    pf = _bin("ffprobe")
    proc = _run([pf, "-v", "quiet", "-print_format", "json",
                 "-show_format", "-show_streams", src], timeout)
    data = json.loads(proc.stdout or "{}")
    fmt = data.get("format", {})
    out = {"filename": os.path.basename(src),
           "duration_s": round(float(fmt.get("duration", 0)), 2) if fmt.get("duration") else None,
           "size_bytes": int(fmt.get("size", 0)) if fmt.get("size") else None,
           "bit_rate": int(fmt.get("bit_rate", 0)) if fmt.get("bit_rate") else None,
           "streams": []}
    for s in data.get("streams", []):
        st = {"type": s.get("codec_type"), "codec": s.get("codec_name")}
        if s.get("codec_type") == "video":
            st.update(width=s.get("width"), height=s.get("height"),
                      fps=s.get("r_frame_rate"))
        elif s.get("codec_type") == "audio":
            st.update(sample_rate=s.get("sample_rate"), channels=s.get("channels"))
        out["streams"].append(st)
    print(json.dumps(out, ensure_ascii=False, indent=2))
    return out


def audio(src, out, fmt=None, bitrate="192k", threads=1, timeout=600):
    out = _ensure_out(out)
    args = [_bin("ffmpeg"), "-y", "-threads", str(threads), "-i", src, "-vn",
            "-b:a", bitrate, out]
    p = _run(args, timeout)
    _check(out, p)
    return out


def trim(src, out, start=None, end=None, duration=None, reencode=False, threads=1, timeout=600):
    out = _ensure_out(out)
    args = [_bin("ffmpeg"), "-y", "-threads", str(threads)]
    if start:
        args += ["-ss", str(start)]
    args += ["-i", src]
    if duration:
        args += ["-t", str(duration)]
    elif end:
        args += ["-to", str(end)]
    args += (["-c", "copy"] if not reencode else ["-preset", "veryfast"])
    args += [out]
    p = _run(args, timeout)
    _check(out, p)
    return out


def thumbnail(src, out, time="00:00:01", threads=1, timeout=120):
    out = _ensure_out(out)
    args = [_bin("ffmpeg"), "-y", "-threads", str(threads), "-ss", str(time),
            "-i", src, "-frames:v", "1", out]
    p = _run(args, timeout)
    _check(out, p)
    return out


def frames(src, out_dir, every=None, fps=None, threads=1, timeout=600):
    out_dir = os.path.abspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    if fps:
        vf = f"fps={fps}"
    elif every:
        vf = f"fps=1/{every}"
    else:
        vf = "fps=1"  # 默认每秒一帧
    pattern = os.path.join(out_dir, "frame_%04d.jpg")
    args = [_bin("ffmpeg"), "-y", "-threads", str(threads), "-i", src,
            "-vf", vf, pattern]
    p = _run(args, timeout)
    if p.returncode != 0:
        raise RuntimeError("抽帧失败:\n" + p.stderr[-800:])
    files = sorted(os.path.join(out_dir, f) for f in os.listdir(out_dir) if f.startswith("frame_"))
    return files


def convert(src, out, vcodec=None, acodec=None, copy=False, threads=1, timeout=1200):
    out = _ensure_out(out)
    args = [_bin("ffmpeg"), "-y", "-threads", str(threads), "-i", src]
    if copy:
        args += ["-c", "copy"]
    else:
        args += ["-c:v", vcodec or "libx264", "-preset", "veryfast",
                 "-c:a", acodec or "aac"]
    args += [out]
    p = _run(args, timeout)
    _check(out, p)
    return out


def compress(src, out, height=720, crf=28, preset="veryfast", threads=1, timeout=1800):
    out = _ensure_out(out)
    vf = f"scale=-2:{height}"
    args = [_bin("ffmpeg"), "-y", "-threads", str(threads), "-i", src,
            "-vf", vf, "-c:v", "libx264", "-crf", str(crf), "-preset", preset,
            "-c:a", "aac", "-b:a", "128k", out]
    p = _run(args, timeout)
    _check(out, p)
    return out


def gif(src, out, start=None, duration=3, width=480, fps=10, threads=1, timeout=600):
    out = _ensure_out(out)
    vf = f"fps={fps},scale={width}:-1:flags=lanczos"
    args = [_bin("ffmpeg"), "-y", "-threads", str(threads)]
    if start:
        args += ["-ss", str(start)]
    args += ["-t", str(duration), "-i", src, "-vf", vf, out]
    p = _run(args, timeout)
    _check(out, p)
    return out


def _check(out, proc):
    if not os.path.isfile(out) or os.path.getsize(out) == 0:
        raise RuntimeError("处理失败:\n" + (proc.stderr[-800:] if proc.stderr else "(无 stderr)"))


def _main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="video_tool", description="视频工具箱（ffmpeg）")
    sub = ap.add_subparsers(dest="cmd", required=True)

    p = sub.add_parser("info"); p.add_argument("src")
    p = sub.add_parser("audio"); p.add_argument("src"); p.add_argument("-o", "--out", required=True); p.add_argument("--bitrate", default="192k"); p.add_argument("--threads", type=int, default=1)
    p = sub.add_parser("trim"); p.add_argument("src"); p.add_argument("-o", "--out", required=True); p.add_argument("--start"); p.add_argument("--end"); p.add_argument("--duration"); p.add_argument("--reencode", action="store_true"); p.add_argument("--threads", type=int, default=1)
    p = sub.add_parser("thumbnail"); p.add_argument("src"); p.add_argument("-o", "--out", required=True); p.add_argument("--time", default="00:00:01"); p.add_argument("--threads", type=int, default=1)
    p = sub.add_parser("frames"); p.add_argument("src"); p.add_argument("-o", "--out", required=True); p.add_argument("--every", type=float); p.add_argument("--fps", type=float); p.add_argument("--threads", type=int, default=1)
    p = sub.add_parser("convert"); p.add_argument("src"); p.add_argument("-o", "--out", required=True); p.add_argument("--vcodec"); p.add_argument("--acodec"); p.add_argument("--copy", action="store_true"); p.add_argument("--threads", type=int, default=1)
    p = sub.add_parser("compress"); p.add_argument("src"); p.add_argument("-o", "--out", required=True); p.add_argument("--height", type=int, default=720); p.add_argument("--crf", type=int, default=28); p.add_argument("--preset", default="veryfast"); p.add_argument("--threads", type=int, default=1)
    p = sub.add_parser("gif"); p.add_argument("src"); p.add_argument("-o", "--out", required=True); p.add_argument("--start"); p.add_argument("--duration", default="3"); p.add_argument("--width", type=int, default=480); p.add_argument("--fps", type=int, default=10); p.add_argument("--threads", type=int, default=1)

    a = ap.parse_args(argv)
    res = None
    if a.cmd == "info":
        info(a.src)
    elif a.cmd == "audio":
        res = audio(a.src, a.out, bitrate=a.bitrate, threads=a.threads)
    elif a.cmd == "trim":
        res = trim(a.src, a.out, a.start, a.end, a.duration, a.reencode, a.threads)
    elif a.cmd == "thumbnail":
        res = thumbnail(a.src, a.out, a.time, a.threads)
    elif a.cmd == "frames":
        files = frames(a.src, a.out, a.every, a.fps, a.threads)
        for f in files:
            print(f)
        print(f"[OK] frames 完成，共 {len(files)} 帧", flush=True)
        return 0
    elif a.cmd == "convert":
        res = convert(a.src, a.out, a.vcodec, a.acodec, a.copy, a.threads)
    elif a.cmd == "compress":
        res = compress(a.src, a.out, a.height, a.crf, a.preset, a.threads)
    elif a.cmd == "gif":
        res = gif(a.src, a.out, a.start, a.duration, a.width, a.fps, a.threads)
    if res:
        print(f"[OK] {a.cmd} 完成 {res}", flush=True)
    return 0


def _cli():
    import sys
    try:
        return _main()
    except SystemExit:
        raise
    except Exception as e:
        print(f"[ERROR] {type(e).__name__}: {e}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(_cli())
