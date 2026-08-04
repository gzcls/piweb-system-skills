#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
PDF OCR 转 Markdown 工具

功能:
  1. 自动检测 PDF 是否有可提取文本层
  2. 有文本层: 直接提取文字
  3. 无文本层 (扫描件): 渲染为图片 → OCR 识别 → 后处理
  4. 中文制度文档格式化 (章/条/节 → Markdown 标题/加粗)
  5. OCR 纠错词典自动修正

用法:
  python3 ocr_convert.py <input.pdf> <output.md> [options]

选项:
  --dpi N                渲染分辨率 (默认 150, 内存不足可降至 100 或 72)
  --no-format            跳过制度文档格式化
  --no-correction        跳过纠错词典
  --corrections PATH     自定义纠错词典 JSON 文件路径
  --raw-dir DIR          保存每页 OCR 原文到指定目录
  --confidence-threshold N  低置信度词阈值 (默认 60, 0-100)
  --text-threshold N     每页平均字符数阈值 (默认 50), 低于此值判定为扫描件
"""

import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path

# ============================================================================
# 依赖管理
# ============================================================================

def ensure_dependencies():
    """检测并安装缺失的 Python 依赖"""
    missing = []

    try:
        import fitz  # noqa: F401
    except ImportError:
        missing.append("PyMuPDF>=1.23.0")

    try:
        from rapidocr_onnxruntime import RapidOCR  # noqa: F401
    except ImportError:
        missing.append("rapidocr_onnxruntime==1.2.13")

    if missing:
        print(f"[pdf-ocr] 正在安装缺失依赖: {', '.join(missing)}", file=sys.stderr)
        # 尝试安装到用户目录 (沙箱容器根只读, 但 /tmp 可写)
        cmd = [
            sys.executable, "-m", "pip", "install",
            "--target", "/tmp/pdf_ocr_deps",
            "--no-cache-dir",
            "--quiet",
        ] + missing
        try:
            subprocess.check_call(cmd, timeout=300)
            sys.path.insert(0, "/tmp/pdf_ocr_deps")
            print("[pdf-ocr] 依赖安装完成", file=sys.stderr)
        except subprocess.CalledProcessError as e:
            print(f"[pdf-ocr] 依赖安装失败: {e}", file=sys.stderr)
            sys.exit(1)
        except subprocess.TimeoutExpired:
            print("[pdf-ocr] 依赖安装超时 (5分钟)", file=sys.stderr)
            sys.exit(1)


# ============================================================================
# PDF 文本提取与扫描件检测
# ============================================================================

def extract_text_and_detect(pdf_path: str, text_threshold: int = 50):
    """
    提取 PDF 文本层并检测是否为扫描件。

    Returns:
        (is_scanned, page_texts, page_count)
        - is_scanned: True 表示扫描件 (无文本层)
        - page_texts: 每页提取的文本列表
        - page_count: 总页数
    """
    import fitz

    doc = fitz.open(pdf_path)
    page_count = len(doc)
    page_texts = []
    total_chars = 0

    for page in doc:
        text = page.get_text("text")
        page_texts.append(text)
        # 统计非空白字符数
        total_chars += len(text.replace(" ", "").replace("\n", "").replace("\t", ""))

    doc.close()

    if page_count == 0:
        return True, page_texts, page_count

    avg_chars = total_chars / page_count
    is_scanned = avg_chars <= text_threshold
    return is_scanned, page_texts, page_count


# ============================================================================
# PDF 页面渲染
# ============================================================================

def render_pdf_pages(pdf_path: str, dpi: int = 150):
    """
    将 PDF 各页渲染为 PNG 图片。

    Args:
        pdf_path: PDF 文件路径
        dpi: 渲染分辨率

    Returns:
        list of (page_num, png_bytes)
    """
    import fitz

    doc = fitz.open(pdf_path)
    pages = []
    zoom = dpi / 72.0  # PyMuPDF 默认 72 DPI
    matrix = fitz.Matrix(zoom, zoom)

    for i, page in enumerate(doc):
        try:
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            png_bytes = pix.tobytes("png")
            pages.append((i + 1, png_bytes))
            print(f"  页面 {i+1}/{len(doc)} 渲染完成 ({pix.width}x{pix.height})",
                  file=sys.stderr)
        except Exception as e:
            print(f"  页面 {i+1} 渲染失败: {e}", file=sys.stderr)

    doc.close()
    return pages


# ============================================================================
# OCR 识别
# ============================================================================

def ocr_pages(rendered_pages, confidence_threshold: int = 60):
    """
    对渲染的页面图片进行 OCR 识别。

    Args:
        rendered_pages: [(page_num, png_bytes), ...]
        confidence_threshold: 低置信度词阈值

    Returns:
        (page_texts, avg_confidence, low_confidence_words)
    """
    from rapidocr_onnxruntime import RapidOCR

    ocr = RapidOCR()
    page_texts = []
    all_confidences = []
    low_confidence_words = []

    for page_num, png_bytes in rendered_pages:
        try:
            result, _ = ocr(png_bytes)

            if result is None or len(result) == 0:
                page_texts.append("")
                print(f"  页面 {page_num}: 未识别到文字", file=sys.stderr)
                continue

            # result 格式: [[bbox, text, confidence], ...]
            lines = []
            page_confidences = []

            for item in result:
                bbox, text, confidence = item[0], item[1], item[2]
                lines.append(text)
                page_confidences.append(confidence)
                all_confidences.append(confidence)

                if confidence * 100 < confidence_threshold:
                    low_confidence_words.append({
                        "word": text,
                        "confidence": round(confidence * 100, 1),
                        "page": page_num,
                    })

            page_text = "\n".join(lines)
            page_texts.append(page_text)

            avg_page_conf = (sum(page_confidences) / len(page_confidences) * 100
                           if page_confidences else 0)
            print(f"  页面 {page_num}: {len(lines)} 行, "
                  f"平均置信度 {avg_page_conf:.1f}%", file=sys.stderr)

        except Exception as e:
            page_texts.append("")
            print(f"  页面 {page_num}: OCR 失败 - {e}", file=sys.stderr)

    avg_confidence = (sum(all_confidences) / len(all_confidences) * 100
                     if all_confidences else 0)

    return page_texts, avg_confidence, low_confidence_words


# ============================================================================
# 后处理: 页眉页脚清洗
# ============================================================================

def remove_header_footer(page_texts, similarity_threshold: float = 0.7):
    """
    检测并移除页眉页脚。

    策略: 对比各页首 2 行和末 2 行, 如果跨页重复出现 (相似度 > 阈值) 则删除。
    """
    if len(page_texts) < 3:
        return page_texts

    def get_head_tail(text, n=2):
        lines = text.strip().split("\n")
        head = lines[:n] if len(lines) >= n else lines
        tail = lines[-n:] if len(lines) >= n else []
        return head, tail

    def similarity(a, b):
        """简单的字符串相似度 (Jaccard)"""
        if not a or not b:
            return 0
        set_a = set(a)
        set_b = set(b)
        intersection = len(set_a & set_b)
        union = len(set_a | set_b)
        return intersection / union if union else 0

    # 收集各页首尾行
    heads = []
    tails = []
    for text in page_texts:
        h, t = get_head_tail(text)
        heads.append(h)
        tails.append(t)

    # 检测重复的页眉
    header_patterns = set()
    for i in range(len(heads)):
        for j in range(i + 1, len(heads)):
            for line_i in heads[i]:
                for line_j in heads[j]:
                    if similarity(line_i.strip(), line_j.strip()) > similarity_threshold:
                        header_patterns.add(line_i.strip())

    # 检测重复的页脚
    footer_patterns = set()
    for i in range(len(tails)):
        for j in range(i + 1, len(tails)):
            for line_i in tails[i]:
                for line_j in tails[j]:
                    if similarity(line_i.strip(), line_j.strip()) > similarity_threshold:
                        footer_patterns.add(line_i.strip())

    # 额外的页码模式
    page_num_pattern = re.compile(r"^[\s]*[-—]?\s*\d+\s*[-—]?\s*$")

    # 清理
    cleaned = []
    removed_count = 0
    for text in page_texts:
        lines = text.split("\n")
        filtered = []
        for line in lines:
            stripped = line.strip()
            # 跳过匹配页眉页脚模式的行
            if stripped in header_patterns or stripped in footer_patterns:
                removed_count += 1
                continue
            # 跳过纯页码行
            if page_num_pattern.match(stripped):
                removed_count += 1
                continue
            filtered.append(line)
        cleaned.append("\n".join(filtered))

    if removed_count > 0:
        print(f"  移除了 {removed_count} 行页眉/页脚/页码", file=sys.stderr)

    return cleaned


# ============================================================================
# 后处理: 段落合并
# ============================================================================

def merge_broken_lines(text: str) -> str:
    """
    合并因 OCR 断行产生的碎片行。

    规则:
    - 行末无句末标点 (。！？；：) 且下一行首字非章节/条款标记 → 合并
    - 保留空行分段
    """
    sentence_end = set("。！？；：.!?;:")
    # 章节条款开头模式
    section_start = re.compile(
        r"^(第[一二三四五六七八九十百千零]+[章节条款]|"
        r"[（(][一二三四五六七八九十]+[)）]|"
        r"附件\s*\d|"
        r"\d+[.、]|"
        r"[一二三四五六七八九十]+[、.])"
    )

    lines = text.split("\n")
    merged = []
    buffer = ""

    for line in lines:
        stripped = line.strip()

        # 空行: flush buffer, 保留空行
        if not stripped:
            if buffer:
                merged.append(buffer)
                buffer = ""
            merged.append("")
            continue

        # 如果 buffer 非空, 判断是否合并
        if buffer:
            # 下一行以章节标记开头 → 不合并
            if section_start.match(stripped):
                merged.append(buffer)
                buffer = stripped
            # 上一行末尾有句末标点 → 不合并
            elif buffer[-1] in sentence_end:
                merged.append(buffer)
                buffer = stripped
            else:
                # 合并
                buffer += stripped
        else:
            buffer = stripped

    if buffer:
        merged.append(buffer)

    return "\n".join(merged)


# ============================================================================
# 后处理: 中文制度文档格式化
# ============================================================================

def apply_document_formatting(text: str) -> str:
    """
    将中文制度文档格式化为 Markdown。

    规则:
    - 第X章 → ## 二级标题
    - 第X节 → ### 三级标题
    - 第X条 → **条款加粗**
    - （X）→ 保持段落列表
    - 附件X → 附件标题
    """
    lines = text.split("\n")
    formatted = []

    for line in lines:
        stripped = line.strip()
        if not stripped:
            formatted.append("")
            continue

        # 第X章 → ## 标题
        m = re.match(r"^(第[一二三四五六七八九十百千零]+章)\s*(.*)", stripped)
        if m:
            formatted.append(f"\n## {m.group(1)} {m.group(2).strip()}\n")
            continue

        # 第X节 → ### 标题
        m = re.match(r"^(第[一二三四五六七八九十百千零]+节)\s*(.*)", stripped)
        if m:
            formatted.append(f"\n### {m.group(1)} {m.group(2).strip()}\n")
            continue

        # 第X条 → 加粗
        m = re.match(r"^(第[一二三四五六七八九十百千零]+条)\s*(.*)", stripped)
        if m:
            formatted.append(f"\n**{m.group(1)}** {m.group(2).strip()}")
            continue

        # 附件 → 标题
        m = re.match(r"^(附件\s*\d+)\s*[:.：]?\s*(.*)", stripped)
        if m:
            formatted.append(f"\n### {m.group(1)} {m.group(2).strip()}\n")
            continue

        formatted.append(stripped)

    return "\n".join(formatted)


# ============================================================================
# 后处理: 纠错词典
# ============================================================================

def load_corrections(corrections_path: str = None) -> dict:
    """加载纠错词典"""
    corrections = {}

    # 1. 内置默认词典
    default_corrections = {
        "已有": "己有",
        "据为已有": "据为己有",
        "每年未": "每年末",
        "楚烧": "焚烧",
        "仃止": "停止",
        "洗理": "清理",
        "挡案": "档案",
        "仑库": "仓库",
        "人事部门": "人事部门",  # 保持正确词
    }
    corrections.update(default_corrections)

    # 2. 加载自定义词典
    if corrections_path and os.path.exists(corrections_path):
        try:
            with open(corrections_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "corrections" in data:
                    corrections.update(data["corrections"])
            print(f"  加载自定义纠错词典: {len(data.get('corrections', {}))} 条",
                  file=sys.stderr)
        except Exception as e:
            print(f"  加载纠错词典失败: {e}", file=sys.stderr)

    # 3. 尝试加载技能目录下的默认词典
    script_dir = Path(__file__).parent
    default_dict_path = script_dir / "corrections.json"
    if default_dict_path.exists() and str(default_dict_path) != corrections_path:
        try:
            with open(default_dict_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if "corrections" in data:
                    corrections.update(data["corrections"])
        except Exception:
            pass

    return corrections


def apply_dict_correction(text: str, corrections: dict) -> tuple:
    """
    应用纠错词典。

    Returns:
        (corrected_text, corrections_made)
    """
    corrections_made = []
    for wrong, right in corrections.items():
        if wrong == right:
            continue
        if wrong in text:
            count = text.count(wrong)
            text = text.replace(wrong, right)
            corrections_made.append({
                "from": wrong,
                "to": right,
                "count": count,
            })

    return text, corrections_made


# ============================================================================
# 主流程
# ============================================================================

def convert_pdf(
    input_path: str,
    output_path: str,
    dpi: int = 150,
    no_format: bool = False,
    no_correction: bool = False,
    corrections_path: str = None,
    raw_dir: str = None,
    confidence_threshold: int = 60,
    text_threshold: int = 50,
):
    """PDF 转 Markdown 主流程"""

    print(f"\n{'='*60}", file=sys.stderr)
    print(f"PDF OCR 转 Markdown", file=sys.stderr)
    print(f"输入: {input_path}", file=sys.stderr)
    print(f"输出: {output_path}", file=sys.stderr)
    print(f"{'='*60}\n", file=sys.stderr)

    # 1. 检测文本层
    print("[1/6] 检测文本层...", file=sys.stderr)
    is_scanned, page_texts, page_count = extract_text_and_detect(
        input_path, text_threshold
    )
    print(f"  页数: {page_count}", file=sys.stderr)
    print(f"  类型: {'扫描件 (无文本层)' if is_scanned else '电子版 (有文本层)'}",
          file=sys.stderr)

    avg_confidence = 100.0
    low_confidence_words = []
    corrections_made = []

    if not is_scanned:
        # 有文本层: 直接使用提取的文本
        print("\n[2/6] 直接使用提取的文本 (跳过 OCR)", file=sys.stderr)
        result_pages = page_texts
    else:
        # 扫描件: 渲染 + OCR
        print(f"\n[2/6] 渲染页面为图片 (DPI={dpi})...", file=sys.stderr)
        rendered = render_pdf_pages(input_path, dpi=dpi)

        if not rendered:
            print("错误: 无法渲染任何页面", file=sys.stderr)
            sys.exit(1)

        print(f"\n[3/6] OCR 识别中...", file=sys.stderr)
        result_pages, avg_confidence, low_confidence_words = ocr_pages(
            rendered, confidence_threshold
        )

    # 保存原始 OCR 文本
    if raw_dir:
        print(f"\n[4/6] 保存 OCR 原文到 {raw_dir}...", file=sys.stderr)
        os.makedirs(raw_dir, exist_ok=True)
        for i, text in enumerate(result_pages):
            raw_path = os.path.join(raw_dir, f"page-{i+1}.txt")
            with open(raw_path, "w", encoding="utf-8") as f:
                f.write(text)
    else:
        print("\n[4/6] 跳过保存 OCR 原文", file=sys.stderr)

    # 后处理
    print("\n[5/6] 后处理...", file=sys.stderr)

    # 页眉页脚清洗
    if is_scanned and len(result_pages) >= 3:
        print("  - 页眉页脚清洗", file=sys.stderr)
        result_pages = remove_header_footer(result_pages)

    # 合并为完整文本
    full_text = "\n\n".join(result_pages)

    # 段落合并
    if is_scanned:
        print("  - 段落合并", file=sys.stderr)
        full_text = merge_broken_lines(full_text)

    # 制度文档格式化
    if not no_format:
        print("  - 制度文档格式化", file=sys.stderr)
        full_text = apply_document_formatting(full_text)

    # 纠错词典
    if not no_correction:
        print("  - 纠错词典", file=sys.stderr)
        corrections = load_corrections(corrections_path)
        full_text, corrections_made = apply_dict_correction(full_text, corrections)
        if corrections_made:
            for c in corrections_made:
                print(f"    {c['from']} → {c['to']} ({c['count']}处)",
                      file=sys.stderr)

    # 清理多余空行
    full_text = re.sub(r"\n{3,}", "\n\n", full_text).strip()

    # 输出
    print(f"\n[6/6] 写入 {output_path}...", file=sys.stderr)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(full_text + "\n")

    # 处理报告
    print(f"\n{'='*60}", file=sys.stderr)
    print(f"处理完成", file=sys.stderr)
    print(f"  页数: {page_count}", file=sys.stderr)
    print(f"  类型: {'扫描件 (OCR)' if is_scanned else '电子版 (直接提取)'}",
          file=sys.stderr)
    if is_scanned:
        print(f"  平均置信度: {avg_confidence:.1f}%", file=sys.stderr)
    print(f"  输出字符数: {len(full_text)}", file=sys.stderr)
    if corrections_made:
        print(f"  纠错: {sum(c['count'] for c in corrections_made)} 处", file=sys.stderr)
    if low_confidence_words:
        print(f"  低置信度词 ({len(low_confidence_words)} 个):", file=sys.stderr)
        for w in low_confidence_words[:20]:
            print(f"    [{w['confidence']:.0f}%] p{w['page']}: {w['word']}",
                  file=sys.stderr)
        if len(low_confidence_words) > 20:
            print(f"    ... 还有 {len(low_confidence_words)-20} 个",
                  file=sys.stderr)
    print(f"{'='*60}\n", file=sys.stderr)


def main():
    parser = argparse.ArgumentParser(
        description="PDF OCR 转 Markdown",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("input", help="输入 PDF 文件路径")
    parser.add_argument("output", help="输出 Markdown 文件路径")
    parser.add_argument("--dpi", type=int, default=150,
                        help="渲染分辨率 (默认 150)")
    parser.add_argument("--no-format", action="store_true",
                        help="跳过制度文档格式化")
    parser.add_argument("--no-correction", action="store_true",
                        help="跳过纠错词典")
    parser.add_argument("--corrections", type=str, default=None,
                        help="自定义纠错词典 JSON 路径")
    parser.add_argument("--raw-dir", type=str, default=None,
                        help="保存每页 OCR 原文到指定目录")
    parser.add_argument("--confidence-threshold", type=int, default=60,
                        help="低置信度词阈值 (默认 60)")
    parser.add_argument("--text-threshold", type=int, default=50,
                        help="每页平均字符数阈值 (默认 50), 低于此判定为扫描件")

    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"错误: 输入文件不存在: {args.input}", file=sys.stderr)
        sys.exit(1)

    # 确保依赖就绪
    ensure_dependencies()

    convert_pdf(
        input_path=args.input,
        output_path=args.output,
        dpi=args.dpi,
        no_format=args.no_format,
        no_correction=args.no_correction,
        corrections_path=args.corrections,
        raw_dir=args.raw_dir,
        confidence_threshold=args.confidence_threshold,
        text_threshold=args.text_threshold,
    )


if __name__ == "__main__":
    main()
