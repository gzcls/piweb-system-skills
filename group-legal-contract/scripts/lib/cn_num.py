# -*- coding: utf-8 -*-
"""中文数字与阿拉伯数字互转（用于法条条号“第五百八十五条” ↔ 585）。

仅依赖标准库，兼容 Python 3.10+。
"""
from __future__ import annotations

import re

_DIGITS = {"零": 0, "〇": 0, "○": 0, "一": 1, "二": 2, "两": 2, "三": 3, "四": 4,
           "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}
_UNITS = {"十": 10, "百": 100, "千": 1000, "万": 10000}
_CN_CHARS = "".join(_DIGITS) + "".join(_UNITS)
CN_NUM_RE = f"[{_CN_CHARS}]+"


def cn_to_int(s: str) -> int | None:
    """将中文数字（如“一千二百六十”“十五”“二〇二三”）转为整数；无法识别返回 None。"""
    if s is None:
        return None
    s = s.strip()
    if not s:
        return None
    if s.isdigit():
        return int(s)
    # 纯数字串写法（二〇二三）
    if all(ch in _DIGITS for ch in s) and len(s) > 1:
        return int("".join(str(_DIGITS[ch]) for ch in s))
    total, section, number = 0, 0, 0
    for ch in s:
        if ch in _DIGITS:
            number = _DIGITS[ch]
        elif ch in _UNITS:
            unit = _UNITS[ch]
            if unit == 10000:
                section = (section + number) * unit
                total += section
                section = 0
            else:
                if number == 0 and unit == 10:
                    number = 1  # “十五”=15
                section += number * unit
            number = 0
        else:
            return None
    return total + section + number


def int_to_cn(n: int) -> str:
    """将 0–99999 的整数转为法条常用的中文写法（如 585 → 五百八十五，10 → 十，110 → 一百一十）。"""
    if n == 0:
        return "零"
    digits = "零一二三四五六七八九"
    units = [(10000, "万"), (1000, "千"), (100, "百"), (10, "十"), (1, "")]
    out = []
    zero_pending = False
    started = False
    for value, name in units:
        d = n // value
        n %= value
        if d:
            if zero_pending and started:
                out.append("零")
            zero_pending = False
            out.append(digits[d] + name)
            started = True
        elif started:
            zero_pending = True
    s = "".join(out)
    if s.startswith("一十"):
        s = s[1:]
    return s


def parse_num(s: str) -> int | None:
    """解析阿拉伯数字或中文数字。"""
    if s is None:
        return None
    s = s.strip()
    if re.fullmatch(r"\d+", s):
        return int(s)
    return cn_to_int(s)


if __name__ == "__main__":  # 简易自检
    for t in ["五百八十五", "十五", "一千二百六十", "一百一十", "二百零五", "二〇二三"]:
        print(t, cn_to_int(t))
    for v in [585, 15, 1260, 110, 205, 10, 1000, 1001]:
        print(v, int_to_cn(v))
