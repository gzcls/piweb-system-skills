# -*- coding: utf-8 -*-
"""法条引用抽取与核验模块。

支持识别：
- 《法名》（可选版本/文号括注）第X条（之Y）（第Z款）（第W项）
- 同一法名后连续引用多条：第143、522、523条 / 第20-21条 / 第620条至第624条 / 第5条、第23条
- 文号直接引用：法释〔2023〕13号第16条、国务院令第790号第22条、法释〔2024〕15号

仅依赖标准库，兼容 Python 3.10+。
"""
from __future__ import annotations

import re
from datetime import date, datetime
from dataclasses import dataclass, field, asdict

from .cn_num import CN_NUM_RE, parse_num, int_to_cn
from .kb import (
    resolve_law,
    find_articles,
    entry_by_id,
    split_paragraphs,
    norm_name,
)

# 任意条号数字（中文数字或阿拉伯数字）
_NUM_TOKEN = rf"(?:{CN_NUM_RE}|\d+)"

# 单个条号表达式，如“第585条”“第五百八十五条之二”“23条”
_SINGLE_ART_RE = re.compile(
    rf"第?(?P<num>{_NUM_TOKEN})条(?:之(?P<suffix>{_NUM_TOKEN}))?"
    rf"(?:第(?P<para>{_NUM_TOKEN})款)?"
    rf"(?:第[（(]?(?P<item>{_NUM_TOKEN})[）)]?项)?"
)

# 区间表达式，如“第620条-第624条”“第20-21条”“第620条至第624条”
_RANGE_ART_RE = re.compile(
    rf"第?(?P<start>{_NUM_TOKEN})条?\s*(?:[\-—~～]|至|到)\s*第?(?P<end>{_NUM_TOKEN})条"
)

# 顿号分隔共用“条”字，如“第143、522、523条”或“第23、24-25条”
_SHARED_TIAO_RE = re.compile(
    rf"第(?P<list>{_NUM_TOKEN}(?:\s*[、，,及和与\-—~～]\s*{_NUM_TOKEN})+)条"
)

# 法名后跟随的条文串（可含多条、款、项）
_TAIL_ART_CHAIN_RE = re.compile(
    rf"^\s*(?:[（(][^）)]*[）)]\s*)*"
    rf"(?P<chain>第{_NUM_TOKEN}(?:条|[\-—~～至到、，,]\s*第?{_NUM_TOKEN})"
    rf"(?:[\s、，,及和与\-—~～至到]*第?[（(]?{_NUM_TOKEN}[）)]?(?:条(?:之{_NUM_TOKEN})?|款|项))*)"
)

# 书名号法源
_BOOK_LAW_RE = re.compile(r"《(?P<law>(?:[^《》\n]|〈[^〉\n]*〉)+)》")

# 独立文号引用（如 法释〔2025〕4号第5条、国务院令第790号第22条）
_STANDALONE_DOCNO_RE = re.compile(
    r"(?P<docno>(?:法释|法发|法|国办发|国发|国粮仓规|国粮发)\s*[〔\[［(（【]\s*\d{4}\s*[〕\]］)）】]\s*\d+\s*号"
    r"|(?:国务院|住房和城乡建设部|住建部|商务部)令\s*第?\s*\d+\s*号)"
    rf"(?P<tail>(?:\s*第{_NUM_TOKEN}条(?:之{_NUM_TOKEN})?(?:第{_NUM_TOKEN}款)?(?:第[（(]?{_NUM_TOKEN}[）)]?项)?[\s、，,及和与]*)+)?"
)

# 非法规文书后缀（排除《保密协议》第8条、《房屋租赁合同》第3条等对合同自身的引用）
_NON_LAW_SUFFIX_RE = re.compile(
    r"(?:合同|协议|补充协议|章程|决议|确认书|申请书|答辩状|起诉状|担保书|委托书|告知书|承诺书|"
    r"律师函|通知书|催告函|函|手册|报价单|结算单|验收单|清单|线索表|报告|范本|模板|指引|说明|附件[一二三四五六七八九十\d]*)$"
)

# 法规特征词
_LAW_INDICATOR_RE = re.compile(
    r"(?:法|条例|规定|解释|办法|纪要|批复|决定|通则|细则|准则|规则|宪章|公约|令|号)"
)


def is_likely_law_name(name: str, index_path: str | None = None) -> bool:
    """判断《…》中的名称是否为法律法规/司法解释（而非合同或文书名称）。"""
    cleaned = name.strip()
    if not cleaned or "[" in cleaned or "●" in cleaned:
        return False
    # 先尝试知识库能否直接识别
    status, _ = resolve_law(cleaned, index_path or "")
    if status != "not_in_kb":
        return True
    normed = norm_name(cleaned)
    if _NON_LAW_SUFFIX_RE.search(normed) and not normed.endswith(("合同法", "劳动合同法")):
        return False
    return bool(_LAW_INDICATOR_RE.search(normed))


@dataclass
class ArticleRef:
    num: int
    suffix: int | None = None
    paragraph: int | None = None
    item: int | None = None
    raw: str = ""

    @property
    def label(self) -> str:
        s = f"第{int_to_cn(self.num)}条"
        if self.suffix:
            s += f"之{int_to_cn(self.suffix)}"
        if self.paragraph:
            s += f"第{int_to_cn(self.paragraph)}款"
        if self.item:
            s += f"第（{int_to_cn(self.item)}）项"
        return s


@dataclass
class CitationMatch:
    raw_text: str
    law_query: str
    articles: list[ArticleRef] = field(default_factory=list)
    line: int = 1
    context: str = ""
    acknowledged_out_of_kb: bool = False


@dataclass
class VerificationItem:
    raw_citation: str
    law_query: str
    article_label: str | None
    line: int
    status: str  # verified | article_not_found | out_of_kb | ambiguous_law | status_unverified | law_not_in_force
    status_cn: str
    law_id: str | None = None
    law_title: str | None = None
    law_status: str | None = None
    effective_date: str | None = None
    excerpt: bool = False
    acknowledged_out_of_kb: bool = False
    matched_text_preview: str | None = None
    note: str = ""

    def to_dict(self) -> dict:
        return asdict(self)


def parse_article_chain(chain: str) -> list[ArticleRef]:
    """解析条文串，返回展开后的 ArticleRef 列表。"""
    if not chain:
        return []
    refs: list[ArticleRef] = []
    seen: set[tuple[int, int | None, int | None, int | None]] = set()

    def add_ref(num: int | None, suffix: int | None = None, para: int | None = None, item: int | None = None, raw: str = ""):
        if num is None or num <= 0:
            return
        key = (num, suffix, para, item)
        if key not in seen:
            seen.add(key)
            refs.append(ArticleRef(num=num, suffix=suffix, paragraph=para, item=item, raw=raw))

    # 1. 先处理区间：“第620条-第624条” 或 “第20-21条”
    used_spans: list[tuple[int, int]] = []
    for rm in _RANGE_ART_RE.finditer(chain):
        s_num = parse_num(rm.group("start"))
        e_num = parse_num(rm.group("end"))
        if s_num and e_num and s_num <= e_num and (e_num - s_num) <= 50:
            used_spans.append(rm.span())
            for n in range(s_num, e_num + 1):
                add_ref(n, raw=rm.group(0))

    def in_used(pos: int) -> bool:
        return any(a <= pos < b for a, b in used_spans)

    # 2. 处理共用“条”字的顿号列表：“第143、522、523条”
    for sm in _SHARED_TIAO_RE.finditer(chain):
        if in_used(sm.start()):
            continue
        used_spans.append(sm.span())
        parts = re.split(r"[、，,及和与]", sm.group("list"))
        for part in parts:
            part = part.strip()
            if not part:
                continue
            rng = re.match(rf"^({_NUM_TOKEN})\s*[\-—~～]\s*({_NUM_TOKEN})$", part)
            if rng:
                s_num, e_num = parse_num(rng.group(1)), parse_num(rng.group(2))
                if s_num and e_num and s_num <= e_num and (e_num - s_num) <= 50:
                    for n in range(s_num, e_num + 1):
                        add_ref(n, raw=part)
            else:
                add_ref(parse_num(part), raw=part)

    # 3. 处理常规单条引用：“第585条”“第22条第3款”
    for am in _SINGLE_ART_RE.finditer(chain):
        if in_used(am.start()):
            continue
        num = parse_num(am.group("num"))
        suf = parse_num(am.group("suffix")) if am.group("suffix") else None
        para = parse_num(am.group("para")) if am.group("para") else None
        item = parse_num(am.group("item")) if am.group("item") else None
        add_ref(num, suf, para, item, raw=am.group(0))

    return refs


def extract_citations(text: str, index_path: str | None = None) -> list[CitationMatch]:
    """从文本中按行抽取全部法条引用。"""
    matches: list[CitationMatch] = []
    lines = text.splitlines()

    # “库外法源”说明只对相邻引用生效。此前只要文档任意位置出现一次
    # 说明，所有库外引用都会被误判为已标注，导致漏报。允许前后各一行
    # 是为了覆盖单独一行的脚注/括注，同时避免整篇文档的全局污染。
    def line_acknowledged(idx: int, start: int | None = None, end: int | None = None) -> bool:
        lo = max(0, idx - 1)
        hi = min(len(lines), idx + 2)
        # 独立一行的说明只覆盖相邻引用；同一行则按引用跨度附近的
        # 说明匹配，避免一行同时出现两部库外法源时互相污染。
        if any("库外法源" in lines[i] or "待人工核实" in lines[i] for i in range(lo, hi) if i != idx):
            return True
        line = lines[idx]
        if start is None or end is None:
            return "库外法源" in line or "待人工核实" in line
        for marker in re.finditer(r"库外法源|待人工核实", line):
            # 说明紧跟在引用后时归属于该引用；引用前的说明仅在行首或
            # 由标点/空白隔开时归属，避免前一引用的尾注污染后一引用。
            if end <= marker.start() <= end + 40:
                return True
            if marker.end() <= start and start - marker.end() <= 20:
                prefix = line[: marker.start()].rstrip()
                if not prefix or prefix[-1] in " \t:：;；,，。|｜":
                    return True
        return False

    for line_idx, line in enumerate(lines, start=1):
        occupied: list[tuple[int, int]] = []

        # 1. 书名号法源 + 后续条号
        for bm in _BOOK_LAW_RE.finditer(line):
            law_raw = bm.group("law").strip()
            if not is_likely_law_name(law_raw, index_path):
                continue
            after = line[bm.end():]
            tm = _TAIL_ART_CHAIN_RE.match(after)
            if tm:
                chain_str = tm.group("chain")
                art_refs = parse_article_chain(chain_str)
                span_end = bm.end() + tm.end()
                occupied.append((bm.start(), span_end))
                matches.append(
                    CitationMatch(
                        raw_text=line[bm.start():span_end].strip(),
                        law_query=law_raw,
                        articles=art_refs,
                        line=line_idx,
                        context=line.strip(),
                        acknowledged_out_of_kb=line_acknowledged(line_idx - 1, bm.start(), span_end),
                    )
                )
            else:
                # 仅在紧跟文号或明确属于法规引用场景时记录无条号法源，避免噪音
                pass

        # 2. 独立文号引用（未被书名号区间覆盖的）
        for dm in _STANDALONE_DOCNO_RE.finditer(line):
            if any(a <= dm.start() < b for a, b in occupied):
                continue
            docno = dm.group("docno")
            tail = dm.group("tail") or ""
            art_refs = parse_article_chain(tail) if tail else []
            if not art_refs:
                continue
            matches.append(
                CitationMatch(
                    raw_text=dm.group(0).strip(),
                    law_query=docno,
                    articles=art_refs,
                    line=line_idx,
                    context=line.strip(),
                    acknowledged_out_of_kb=line_acknowledged(line_idx - 1, dm.start(), dm.end()),
                )
            )

    return matches


def _item_exists(article_text: str, item: int) -> bool:
    """Check whether an article contains the requested list item.

    Item markers are commonly written as ``（一）``/``(1)`` or ``一、``/``1、``.
    Restrict the match to a line/list boundary so a number embedded in prose is
    not mistaken for a legal item.
    """
    cn = re.escape(int_to_cn(item))
    arabic = re.escape(str(item))
    return bool(re.search(
        rf"(?:^|[\n；;。])\s*(?:[（(](?:{cn}|{arabic})[）)]|(?:{cn}|{arabic})[、.．])",
        article_text,
    ))


def verify_citation_matches(
    matches: list[CitationMatch],
    index_path: str | None = None,
    as_of: str | None = None,
) -> list[VerificationItem]:
    """逐条核验抽取出的法条引用。"""
    results: list[VerificationItem] = []
    as_of_date: date | None = None
    if as_of:
        try:
            as_of_date = datetime.strptime(str(as_of)[:10], "%Y-%m-%d").date()
        except ValueError as exc:
            raise ValueError(f"as_of 必须是 YYYY-MM-DD: {as_of}") from exc

    def date_value(value: object) -> date | None:
        if not value:
            return None
        raw = str(value).strip()
        match = re.search(r"(\d{4})[-年/.](\d{1,2})[-月/.](\d{1,2})", raw)
        if not match:
            return None
        try:
            return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        except ValueError:
            return None

    for m in matches:
        status, targets = resolve_law(m.law_query, index_path or "")
        if status == "not_in_kb":
            for art in (m.articles or [None]):
                results.append(
                    VerificationItem(
                        raw_citation=m.raw_text,
                        law_query=m.law_query,
                        article_label=art.label if art else None,
                        line=m.line,
                        status="out_of_kb",
                        status_cn="库外法源",
                        acknowledged_out_of_kb=m.acknowledged_out_of_kb,
                        note="已标注【库外法源·待人工核实】" if m.acknowledged_out_of_kb else "本地库未收录该法源，须标注【库外法源·待人工核实】",
                    )
                )
            continue

        if status == "ambiguous":
            cand_titles = [entry_by_id(t.entry_id, index_path or "")["title"] for t in targets if entry_by_id(t.entry_id, index_path or "")]
            for art in (m.articles or [None]):
                results.append(
                    VerificationItem(
                        raw_citation=m.raw_text,
                        law_query=m.law_query,
                        article_label=art.label if art else None,
                        line=m.line,
                        status="ambiguous_law",
                        status_cn="法源有歧义",
                        note=f"匹配到多个法源：{', '.join(cand_titles)}",
                    )
                )
            continue

        # status == "ok"
        target = targets[0]
        entry = entry_by_id(target.entry_id, index_path or "")
        law_status = entry.get("status", "现行有效")
        status_needs_warn = any(k in (law_status or "") for k in ("待核实", "失效", "废止"))

        for art in m.articles:
            found_arts = []
            used_target = target
            for cand in targets:
                hits = find_articles(cand, art.num, art.suffix, index_path=index_path)
                if hits:
                    found_arts = hits
                    used_target = cand
                    break

            used_entry = entry_by_id(used_target.entry_id, index_path or "") or entry
            used_status = used_entry.get("status", "现行有效")
            law_not_in_force = any(k in (used_status or "") for k in ("失效", "废止"))
            used_warn = any(k in (used_status or "") for k in ("待核实", "部分失效"))
            effective = date_value(used_entry.get("effective_date"))
            not_yet_effective = bool(as_of_date and effective and as_of_date < effective)

            if not found_arts:
                if used_target.excerpt:
                    results.append(
                        VerificationItem(
                            raw_citation=m.raw_text,
                            law_query=m.law_query,
                            article_label=art.label,
                            line=m.line,
                            status="out_of_kb",
                            status_cn="库外法源（仅收录汇编节选）",
                            law_id=used_entry["id"],
                            law_title=used_entry["title"],
                            law_status=used_status,
                            effective_date=used_entry.get("effective_date"),
                            excerpt=True,
                            acknowledged_out_of_kb=m.acknowledged_out_of_kb,
                            note=f"《{used_entry['title']}》仅为节选汇编，未包含{art.label}，须按库外法源标注",
                        )
                    )
                else:
                    results.append(
                        VerificationItem(
                            raw_citation=m.raw_text,
                            law_query=m.law_query,
                            article_label=art.label,
                            line=m.line,
                            status="article_not_found",
                            status_cn="条号不存在",
                            law_id=used_entry["id"],
                            law_title=used_entry["title"],
                            law_status=used_status,
                            effective_date=used_entry.get("effective_date"),
                            excerpt=False,
                            note=f"在《{used_entry['title']}》（共{used_entry.get('article_count')}条）中未找到{art.label}",
                        )
                    )
                continue

            art_obj = found_arts[0]
            preview = art_obj.text.replace("\n", " ")[:120]
            note_parts = []
            if law_not_in_force:
                final_status = "law_not_in_force"
                final_cn = "法规已失效/废止"
            else:
                final_status = "status_unverified" if used_warn else "verified"
                final_cn = "状态待核实" if used_warn else "已核验"
            if not_yet_effective:
                final_status = "not_yet_effective"
                final_cn = "截至日尚未生效"
            # 检查款号是否存在。款号越界属于引用错误，应进入失败计数，
            # 而不是只在备注里提示后仍标记为“已核验”。
            if art.paragraph:
                paras = split_paragraphs(art_obj.text)
                if art.paragraph > len(paras):
                    note_parts.append(f"条文仅有 {len(paras)} 款，引用了第{art.paragraph}款")
                    final_status = "article_not_found"
                    final_cn = "款号不存在"
                else:
                    preview = paras[art.paragraph - 1].replace("\n", " ")[:120]

            # 检查项号是否真的存在于指定条文。即使款号有效，项号越界
            # 也不能生成“已核验”的结论。
            if art.item is not None and not _item_exists(art_obj.text, art.item):
                note_parts.append(f"条文中未找到第（{int_to_cn(art.item)}）项")
                final_status = "article_not_found"
                final_cn = "项号不存在"

            if used_target.excerpt:
                note_parts.append("来源为专题汇编节选")
            if used_warn or law_not_in_force:
                note_parts.append(f"法规状态为【{used_status}】，请留意时效性")
            if not_yet_effective:
                note_parts.append(f"截至 {as_of_date.isoformat()} 尚未达到生效日 {effective.isoformat()}")

            results.append(
                VerificationItem(
                    raw_citation=m.raw_text,
                    law_query=m.law_query,
                    article_label=art.label,
                    line=m.line,
                    status=final_status,
                    status_cn=final_cn,
                    law_id=used_entry["id"],
                    law_title=used_entry["title"],
                    law_status=used_status,
                    effective_date=used_entry.get("effective_date"),
                    excerpt=used_target.excerpt,
                    acknowledged_out_of_kb=m.acknowledged_out_of_kb,
                    matched_text_preview=preview,
                    note="；".join(note_parts),
                )
            )

    return results


__all__ = [
    "ArticleRef",
    "CitationMatch",
    "VerificationItem",
    "is_likely_law_name",
    "parse_article_chain",
    "extract_citations",
    "verify_citation_matches",
]
