# -*- coding: utf-8 -*-
"""
xlsx_build —— 生成 Excel（.xlsx）。两种输入：
  1) JSON spec（多 sheet、样式、图表）
  2) CSV 文件（--from-csv，快速转 xlsx）

openpyxl 实现，无需 MS Office。

用法：
    python3 xlsx_build.py /workspace/spec.json -o /workspace/out.xlsx
    python3 xlsx_build.py --from-csv /workspace/data.csv -o /workspace/out.xlsx [--sheet 数据]
"""
from __future__ import annotations

import csv
import json
import os

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HEAD_FILL = PatternFill("solid", fgColor="1F3B5C")
HEAD_FONT = Font(color="FFFFFF", bold=True, name="微软雅黑", size=10)
BODY_FONT = Font(name="微软雅黑", size=10)
THIN = Side(style="thin", color="BFC8D2")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
_CHART = {"bar": BarChart, "line": LineChart, "pie": PieChart}


def _write_sheet(ws, sheet: dict):
    headers = sheet.get("headers", [])
    rows = sheet.get("rows", [])
    r0 = 1
    if headers:
        ws.append(headers)
        for c in range(1, len(headers) + 1):
            cell = ws.cell(row=1, column=c)
            cell.fill = HEAD_FILL
            cell.font = HEAD_FONT
            cell.alignment = CENTER
            cell.border = BORDER
        r0 = 2
    for row in rows:
        ws.append(list(row))
    # 正文样式
    ncol = len(headers) if headers else (len(rows[0]) if rows else 0)
    for r in range(r0, ws.max_row + 1):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = BODY_FONT
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center")

    # 数字格式 {"列号(0基)": "#,##0.00"}
    for k, fmt in (sheet.get("number_formats") or {}).items():
        c = int(k) + 1
        for r in range(r0, ws.max_row + 1):
            ws.cell(row=r, column=c).number_format = fmt

    # 列宽
    widths = sheet.get("widths")
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(i)].width = w
    else:  # 自动估算
        for c in range(1, ncol + 1):
            maxlen = 0
            for r in range(1, ws.max_row + 1):
                v = ws.cell(row=r, column=c).value
                if v is not None:
                    maxlen = max(maxlen, len(str(v)))
            ws.column_dimensions[get_column_letter(c)].width = min(40, max(8, maxlen * 1.6))

    if headers and sheet.get("freeze_header", True):
        ws.freeze_panes = "A2"
    if headers and sheet.get("autofilter", True):
        ws.auto_filter.ref = f"A1:{get_column_letter(ncol)}{ws.max_row}"

    # 图表
    ch = sheet.get("chart")
    if ch and headers and rows:
        cls = _CHART.get(ch.get("type", "bar"), BarChart)
        chart = cls()
        chart.title = ch.get("title", "")
        cats_col = ch.get("cats_col", 0) + 1
        data_cols = ch.get("data_cols", [1])
        min_dc = min(data_cols) + 1
        max_dc = max(data_cols) + 1
        data = Reference(ws, min_col=min_dc, max_col=max_dc, min_row=1, max_row=ws.max_row)
        cats = Reference(ws, min_col=cats_col, max_col=cats_col, min_row=2, max_row=ws.max_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height, chart.width = 8, 16
        ws.add_chart(chart, ch.get("anchor", f"{get_column_letter(ncol + 2)}2"))


def build_xlsx(spec: dict, out_path: str) -> str:
    wb = Workbook()
    wb.remove(wb.active)
    sheets = spec.get("sheets") or [spec]  # 允许单 sheet 直接传
    for i, sheet in enumerate(sheets):
        ws = wb.create_sheet(title=str(sheet.get("name", f"Sheet{i + 1}"))[:31])
        _write_sheet(ws, sheet)
    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    wb.save(out_path)
    return out_path


def _smart_num(v):
    """纯数字转数值，但保留前导零(编号/手机号)、超长数字(身份证)等，避免精度/格式损坏。"""
    s = str(v).strip()
    if s == "":
        return v
    core = s.lstrip("-")
    if len(core) > 1 and core[0] == "0" and not core.startswith("0."):
        return v  # 前导零 → 按文本保留
    if core.isdigit() and len(core) > 15:
        return v  # 超长整数 → 避免变科学计数/丢精度
    try:
        return int(s)
    except ValueError:
        try:
            return float(s)
        except ValueError:
            return v


def csv_to_xlsx(csv_path: str, out_path: str, sheet_name="数据", delimiter=",", convert=True) -> str:
    with open(csv_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = list(csv.reader(f, delimiter=delimiter))
    headers = reader[0] if reader else []
    rows = reader[1:] if len(reader) > 1 else []
    if convert:
        rows = [[_smart_num(x) for x in row] for row in rows]
    return build_xlsx({"sheets": [{"name": sheet_name, "headers": headers, "rows": rows}]}, out_path)


def _main(argv=None) -> int:
    import argparse
    ap = argparse.ArgumentParser(prog="xlsx_build", description="生成 Excel（JSON spec 或 CSV）")
    ap.add_argument("spec", nargs="?", help="JSON spec 路径（与 --from-csv 二选一）")
    ap.add_argument("--from-csv", default=None, help="CSV 文件路径，快速转 xlsx")
    ap.add_argument("--sheet", default="数据", help="CSV 模式下的工作表名")
    ap.add_argument("--delimiter", default=",", help="CSV 分隔符，默认逗号")
    ap.add_argument("--no-convert", action="store_true", help="CSV 模式：不把数字字符串转成数值（全部按文本）")
    ap.add_argument("-o", "--out", required=True, help="输出 .xlsx 路径（建议 /workspace）")
    args = ap.parse_args(argv)

    if args.from_csv:
        out = csv_to_xlsx(args.from_csv, args.out, args.sheet, args.delimiter, convert=not args.no_convert)
    elif args.spec:
        with open(args.spec, "r", encoding="utf-8") as f:
            out = build_xlsx(json.load(f), args.out)
    else:
        ap.error("需要提供 JSON spec 或 --from-csv")
    print(f"[OK] 已生成 {out}", flush=True)
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
