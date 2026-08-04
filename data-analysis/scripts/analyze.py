# -*- coding: utf-8 -*-
"""
analyze —— CSV/Excel 数据分析与出图（独立脚本）。

能力：
  --summary           输出统计摘要(描述统计/字段类型/缺失值)为 Markdown
  --groupby/--agg     分组聚合，结果存 CSV
  --chart             用 matplotlib 出图(bar/line/pie/hist/scatter)为 PNG

依赖：pandas、matplotlib、openpyxl（读 xlsx）。沙箱镜像已预装。

示例：
  python3 analyze.py /workspace/data.csv  --summary -o /workspace/summary.md
  python3 analyze.py /workspace/data.xlsx --groupby 地区 --agg sum --y 销售额 -o /workspace/g.csv
  python3 analyze.py /workspace/data.csv  --chart bar --x 月份 --y 销售额,成本 -o /workspace/chart.png
"""
from __future__ import annotations

import os

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
import pandas as pd  # noqa: E402

# 中文字体：优先按「字体文件」注册，规避平台 matplotlib 按名找不到 Noto CJK 的常见问题
def _setup_cjk_font():
    import glob
    from matplotlib import font_manager as fm
    cands = [
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/noto/NotoSansCJKsc-Regular.otf",
        "/usr/share/fonts/truetype/noto/NotoSansCJK-Regular.ttc",
        "C:/Windows/Fonts/msyh.ttc", "C:/Windows/Fonts/simhei.ttf",
    ]
    cands += glob.glob("/usr/share/fonts/**/NotoSansCJK*.*", recursive=True)
    cands += glob.glob("/usr/share/fonts/**/*CJK*.*", recursive=True)
    for path in cands:
        if os.path.isfile(path):
            try:
                fm.fontManager.addfont(path)
                fam = fm.FontProperties(fname=path).get_name()
                plt.rcParams["font.sans-serif"] = [fam, "sans-serif"]
                return fam
            except Exception:
                continue
    plt.rcParams["font.sans-serif"] = ["Noto Sans CJK SC", "Microsoft YaHei",
                                       "SimHei", "WenQuanYi Zen Hei", "sans-serif"]
    return None


_setup_cjk_font()
plt.rcParams["axes.unicode_minus"] = False


def read_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(path)
    return pd.read_csv(path, encoding="utf-8-sig")


def summary_md(df: pd.DataFrame) -> str:
    lines = [f"# 数据摘要\n", f"- 行数：{len(df)}　列数：{df.shape[1]}\n", "## 字段与类型\n"]
    lines.append("| 字段 | 类型 | 非空 | 缺失 |\n|---|---|---|---|")
    for c in df.columns:
        lines.append(f"| {c} | {df[c].dtype} | {df[c].notna().sum()} | {df[c].isna().sum()} |")
    lines.append("\n## 描述统计（数值列）\n")
    desc = df.describe(include="number").round(3)
    if not desc.empty:
        lines.append("| 指标 | " + " | ".join(map(str, desc.columns)) + " |")
        lines.append("|---|" + "---|" * len(desc.columns))
        for idx, row in desc.iterrows():
            lines.append(f"| {idx} | " + " | ".join(str(v) for v in row.values) + " |")
    else:
        lines.append("（无数值列）")
    return "\n".join(lines) + "\n"


def make_chart(df, ctype, x, ycols, out, title=""):
    fig, ax = plt.subplots(figsize=(8, 4.5), dpi=150)
    ys = [c.strip() for c in ycols.split(",")] if ycols else []
    if ctype == "pie":
        col = ys[0] if ys else df.select_dtypes("number").columns[0]
        labels = df[x] if x else df.index
        ax.pie(df[col], labels=labels, autopct="%1.1f%%", startangle=90)
        ax.axis("equal")
    elif ctype == "hist":
        for c in ys:
            ax.hist(df[c].dropna(), bins=20, alpha=0.6, label=c)
        ax.legend()
    elif ctype == "scatter":
        ax.scatter(df[x], df[ys[0]])
        ax.set_xlabel(x); ax.set_ylabel(ys[0])
    else:  # bar / line
        xvals = df[x] if x else df.index
        for c in ys:
            if ctype == "line":
                ax.plot(xvals, df[c], marker="o", label=c)
            else:
                ax.bar([str(v) for v in xvals], df[c], label=c)
        ax.legend()
        if x:
            ax.set_xlabel(x)
        plt.xticks(rotation=30, ha="right")
    ax.set_title(title or ctype)
    fig.tight_layout()
    out = os.path.abspath(out)
    os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
    fig.savefig(out)
    plt.close(fig)
    return out


def _main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="analyze", description="CSV/Excel 分析与出图")
    ap.add_argument("src", help="CSV 或 Excel 文件（建议放 /workspace）")
    ap.add_argument("-o", "--out", required=True, help="输出路径（.md/.csv/.png）")
    ap.add_argument("--summary", action="store_true", help="输出统计摘要 Markdown")
    ap.add_argument("--groupby", default=None, help="分组列")
    ap.add_argument("--agg", default="sum", help="聚合方式 sum/mean/count/max/min")
    ap.add_argument("--chart", default=None, help="图表类型 bar/line/pie/hist/scatter")
    ap.add_argument("--x", default=None, help="X 轴/分类列")
    ap.add_argument("--y", default=None, help="Y 轴列，逗号分隔，如 销售额,成本")
    ap.add_argument("--title", default="", help="图表标题")
    args = ap.parse_args(argv)

    df = read_table(args.src)

    if args.summary:
        out = os.path.abspath(args.out)
        os.makedirs(os.path.dirname(out) or ".", exist_ok=True)
        with open(out, "w", encoding="utf-8") as f:
            f.write(summary_md(df))
        print(f"[OK] 摘要已生成 {out}", flush=True)
    elif args.groupby:
        ycols = [c.strip() for c in args.y.split(",")] if args.y else None
        g = df.groupby(args.groupby)
        g = g[ycols] if ycols else g
        res = getattr(g, args.agg)(numeric_only=True) if args.agg != "count" else g.count()
        out = os.path.abspath(args.out)
        res.to_csv(out, encoding="utf-8-sig")
        print(f"[OK] 分组聚合({args.agg}) 已生成 {out}", flush=True)
    elif args.chart:
        out = make_chart(df, args.chart, args.x, args.y, args.out, args.title)
        print(f"[OK] 图表已生成 {out}", flush=True)
    else:
        ap.error("请指定 --summary / --groupby / --chart 之一")
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
