# -*- coding: utf-8 -*-
"""知识库公共模块：索引加载、法规名称归一化、法源解析、条文切分。

被 kb_search.py、cite_law.py 与 verify_citations.py 共同使用。
仅依赖标准库，兼容 Python 3.10+。
"""
from __future__ import annotations

import json
import hashlib
import os
import re
from dataclasses import dataclass, field, asdict
from functools import lru_cache
from pathlib import Path

from .cn_num import CN_NUM_RE, cn_to_int, int_to_cn, parse_num

# ---------------------------------------------------------------------------
# 路径与索引
# ---------------------------------------------------------------------------

_SELF_ROOT = Path(__file__).resolve().parents[2]
_DIST_FALLBACK = _SELF_ROOT.parents[1] / "dist" / _SELF_ROOT.name


def _resolve_skill_root() -> Path:
    """选择包含知识库索引的 Skill 根目录。

    开发树通常只有手写源码，生成索引位于 ``dist``；发布包则索引与脚本
    位于同一目录。允许通过环境变量显式指定根目录，便于 CI、临时构建
    目录和多份 Skill 并存时避免静默读取错误版本。
    """
    override = os.environ.get("GROUP_LEGAL_CONTRACT_SKILL_ROOT", "").strip()
    candidates = [Path(override).expanduser()] if override else []
    candidates.extend([_SELF_ROOT, _DIST_FALLBACK])
    for candidate in candidates:
        root = candidate.resolve()
        if (root / "assets" / "kb_index.json").is_file():
            return root
    # 保留原始根目录，让 load_index() 产生清晰的缺失索引错误。
    return _SELF_ROOT


SKILL_ROOT = _resolve_skill_root()
INDEX_PATH = SKILL_ROOT / "assets" / "kb_index.json"


def _index_key(index_path: str | Path | None = None) -> str:
    """Return a stable cache key for an optional index path.

    The original helpers cached all secondary lookups against the module-level
    ``INDEX_PATH``.  That made ``--index`` unsafe: the index summary came from
    the requested file while aliases and article text still came from the
    default knowledge base.  Keeping the path in every cache key prevents that
    cross-index contamination while retaining the fast default path.
    """
    if index_path is None or str(index_path).strip() == "":
        return ""
    return str(Path(index_path).expanduser().resolve())


def _root_for_index(index_path: str | Path | None = None) -> Path:
    key = _index_key(index_path)
    if not key:
        return SKILL_ROOT
    # kb_index.json lives at <skill-root>/assets/kb_index.json.
    return Path(key).parent.parent


def kb_fingerprint(index_path: str | Path | None = None) -> str:
    """Return a cheap content fingerprint for cache/manifest binding.

    The index records each entry's content hash; hashing the index itself is
    enough to invalidate routing and article lookup caches when the built KB
    changes, without reading every law file during every preflight.
    """
    path = Path(_index_key(index_path) or INDEX_PATH)
    if not path.is_file():
        return f"missing:{path}"
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def kb_source_consistency(index_path: str | Path | None = None) -> dict:
    """Compare the published index and snapshot with packaged files and optional source KB.

    Verifies two layers of integrity:
    1. Every entry in ``kb_index.json`` exists on disk under the skill package
       and its raw file bytes match ``entry['sha256']``, and every entry recorded
       in ``kb_source_manifest.json`` matches ``kb_index.json``.
    2. When a mutable ``legal_kb`` source tree is available, every source file is
       normalized (UTF-8 without BOM, LF line endings, single trailing newline)
       and compared against ``kb_source_manifest.json``. Installed/dist-only
       packages return ``snapshot_only`` when layer 1 passes and no mutable
       source tree exists.
    """
    root = _root_for_index(index_path)
    manifest_path = root / "assets" / "kb_source_manifest.json"
    if not manifest_path.is_file():
        return {"status": "unavailable", "manifest": str(manifest_path), "mismatches": []}
    try:
        manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        idx_entries = load_index(_index_key(index_path) or None).get("entries", [])
        indexed = {e.get("source_rel_path"): e.get("sha256") for e in idx_entries if e.get("source_rel_path")}
        mismatches = []
        for item in manifest.get("files", []):
            if indexed.get(item.get("source_rel_path")) != item.get("normalized_sha256"):
                mismatches.append(item.get("source_rel_path"))

        package_missing: list[str] = []
        package_mismatches: list[str] = []
        for entry in idx_entries:
            rel_path = str(entry.get("path") or "")
            if not rel_path:
                continue
            entry_file = Path(rel_path) if Path(rel_path).is_absolute() else (root / rel_path)
            label = str(entry.get("source_rel_path") or entry.get("id") or rel_path)
            if not entry_file.is_file():
                package_missing.append(label)
                continue
            expected_sha = str(entry.get("sha256") or "").strip().lower()
            if expected_sha and hashlib.sha256(entry_file.read_bytes()).hexdigest() != expected_sha:
                package_mismatches.append(label)

        mismatches = sorted(set(mismatches) | set(package_missing) | set(package_mismatches))

        source_override = os.environ.get("GROUP_LEGAL_CONTRACT_SOURCE_KB", "").strip()
        candidates = [Path(source_override).expanduser()] if source_override else []
        # Repository layout: <repo>/skills/group-legal-contract and
        # <repo>/legal_kb. For a copied skill, also accept <copy>/legal_kb.
        explicit_index = Path(_index_key(index_path)).resolve() if _index_key(index_path) else INDEX_PATH.resolve()
        if not source_override and explicit_index == INDEX_PATH.resolve():
            candidates.append(_SELF_ROOT.parents[1] / "legal_kb")
        candidates.append(root.parent.parent / "legal_kb")
        source_root = next((p.resolve() for p in candidates if p.is_dir()), None)
        if source_root is None:
            return {
                "status": "snapshot_only" if not mismatches else "mismatch",
                "manifest": str(manifest_path),
                "source_root": None,
                "checked_count": len(idx_entries),
                "mismatches": mismatches,
                "package_missing": package_missing,
                "package_mismatches": package_mismatches,
                "missing": package_missing,
                "extra": [],
            }

        expected = {
            str(item.get("source_rel_path")): str(item.get("normalized_sha256") or "")
            for item in manifest.get("files", [])
            if item.get("source_rel_path")
        }
        current: dict[str, str] = {}
        for source_path in source_root.rglob("*.md"):
            rel = source_path.relative_to(source_root).as_posix()
            # The builder consumes only the four KB source buckets. README and
            # audit drafts are repository material, not indexed KB entries.
            if not rel.startswith(("01_", "02_", "03_", "04_")):
                continue
            text = source_path.read_text(encoding="utf-8")
            if text.startswith("\ufeff"):
                text = text[1:]
            text = text.replace("\r\n", "\n").replace("\r", "\n")
            if text and not text.endswith("\n"):
                text += "\n"
            raw = text.encode("utf-8")
            current[rel] = hashlib.sha256(raw).hexdigest()
        changed = sorted(rel for rel, expected_hash in expected.items() if current.get(rel) != expected_hash and rel in current)
        missing = sorted(set(rel for rel in expected if rel not in current) | set(package_missing))
        extra = sorted(rel for rel in current if rel not in expected)
        all_mismatches = sorted(set(mismatches) | set(changed) | set(missing) | set(extra))
        return {
            "status": "ok" if not all_mismatches else "mismatch",
            "manifest": str(manifest_path),
            "source_root": str(source_root),
            "checked_count": len(expected),
            "mismatches": all_mismatches,
            "package_missing": package_missing,
            "package_mismatches": package_mismatches,
            "changed": changed,
            "missing": missing,
            "extra": extra,
        }
    except Exception as exc:
        return {"status": "error", "manifest": str(manifest_path), "mismatches": [], "error": str(exc)}


_INDEX_CACHE: dict[tuple[str, int, int], dict] = {}


def load_index(path: str | None = None) -> dict:
    p = (Path(path) if path else INDEX_PATH).resolve()
    if not p.exists():
        raise FileNotFoundError(
            f"未找到知识库索引 {p}。请使用构建产物（dist/group-legal-contract），"
            "或使用带有 assets/kb_index.json 的发布包。")
    stat = p.stat()
    key = (str(p), stat.st_mtime_ns, stat.st_size)
    cached = _INDEX_CACHE.get(key)
    if cached is not None:
        return cached
    # 同一路径出现新版本时清掉依赖旧索引的二级缓存，避免一个长驻进程
    # 在构建后继续返回旧别名或旧条文。
    stale = [old for old in _INDEX_CACHE if old[0] == str(p) and old != key]
    for old in stale:
        _INDEX_CACHE.pop(old, None)
    for fn_name in ("entry_by_id", "_alias_tables", "resolve_law", "_parsed", "_find_articles_cached"):
        fn = globals().get(fn_name)
        if fn is not None and hasattr(fn, "cache_clear"):
            fn.cache_clear()
    data = json.loads(p.read_text(encoding="utf-8"))
    _INDEX_CACHE[key] = data
    return data


def entries(kind: str | None = None, index_path: str | Path | None = None) -> list[dict]:
    items = load_index(_index_key(index_path) or None)["entries"]
    return [e for e in items if kind is None or e["kind"] == kind]


@lru_cache(maxsize=1024)
def entry_by_id(entry_id: str, index_path: str = "") -> dict | None:
    for e in load_index(index_path or None)["entries"]:
        if e["id"] == entry_id:
            return e
    return None


def read_entry_text(entry: dict, index_path: str | Path | None = None) -> str:
    path = Path(entry["path"])
    if not path.is_absolute():
        path = _root_for_index(index_path) / path
    raw = path.read_bytes()
    expected = str(entry.get("sha256") or "").strip().lower()
    if expected:
        actual = hashlib.sha256(raw).hexdigest()
        if actual != expected:
            raise ValueError(
                f"知识库条目校验失败: {entry.get('id') or entry.get('title') or path} "
                f"(expected={expected}, actual={actual})"
            )
    return raw.decode("utf-8")


# ---------------------------------------------------------------------------
# 名称 / 文号归一化
# ---------------------------------------------------------------------------

_VERSION_PAREN = re.compile(
    r"[（(][^（）()]*(?:修订|修正|施行|\d{4}|令|网信办|国家粮储局|版|试行)[^（）()]*[）)]")
_DOCNO_RE = re.compile(
    r"(?:法释|法发|法|国办发|国发|国粮仓规|国粮发)\s*[〔\[［(（【]\s*\d{4}\s*[〕\]］)）】]\s*\d+\s*号"
    r"|(?:国务院|住房和城乡建设部|住建部|商务部|国家互联网信息办公室)令\s*第?\s*\d+\s*号")


def norm_docno(s: str) -> str:
    s = re.sub(r"\s+", "", s)
    s = re.sub(r"[〔\[［(（【](\d{4})[〕\]］)）】]", r"〔\1〕", s)
    s = re.sub(r"令第?(\d+)号", r"令第\1号", s)
    s = s.replace("住房和城乡建设部令", "住建部令")
    return s


def find_docnos(text: str) -> list[str]:
    seen, out = set(), []
    for m in _DOCNO_RE.finditer(text):
        d = norm_docno(m.group(0))
        if d not in seen:
            seen.add(d)
            out.append(d)
    return out


def norm_name(s: str) -> str:
    """法规名称归一化：去书名号、空白、“中华人民共和国”、版本括注。保留（一）（二）等序号括注。"""
    s = s.replace("〈", "").replace("〉", "").replace("《", "").replace("》", "")
    s = re.sub(r"\s+", "", s)
    s = s.replace("(", "（").replace(")", "）")
    s = s.replace("中华人民共和国", "")
    prev = None
    while prev != s:
        prev = s
        s = _VERSION_PAREN.sub("", s)
    return s.strip("：:，,。.")


def title_aliases(title: str) -> list[str]:
    """由文件标题自动派生别名。"""
    cands = {title}
    if "_" in title:
        cands.add(title.split("_", 1)[0])
    out = set()
    for c in cands:
        n = norm_name(c)
        if n:
            out.add(n)
    return sorted(out)


# ---------------------------------------------------------------------------
# 条文切分
# ---------------------------------------------------------------------------

ARTICLE_RE = re.compile(
    r"^\s*(?:[\*\-]\s+)?\*\*\s*(?:《(?P<law>[^》]+)》\s*)?"
    rf"第(?P<num>{CN_NUM_RE})条(?:之(?P<suffix>{CN_NUM_RE}))?"
    r"(?:\s*【(?P<title>[^】]*)】)?\s*\*\*")
HEADING_RE = re.compile(r"^(?P<hashes>#{1,6})\s+(?P<text>.+?)\s*$")
BOOK_RE = re.compile(r"《((?:[^《》]|〈[^〉]*〉)+)》")
TITLE_LIKE_RE = re.compile(r"^(?:最高人民法院|最高人民检察院|国务院|全国人民代表大会)关于")
HEADING_PREFIX_RE = re.compile(
    r"^(?:第[一二三四五六七八九十]+部分|附|[一二三四五六七八九十]+、|\d+[\.、])\s*[：:]?\s*")


@dataclass
class Component:
    key: str
    heading: str
    names: list[str] = field(default_factory=list)   # 已归一化
    doc_nos: list[str] = field(default_factory=list)


@dataclass
class Article:
    num: int
    suffix: int | None
    label: str
    title: str | None
    line: int            # 1-based 行号
    text: str
    component: str | None = None
    explicit_law: str | None = None   # 条号前显式标注的《法名》（已归一化）

    def to_dict(self) -> dict:
        return asdict(self)


def heading_sources(text: str) -> tuple[list[str], list[str]]:
    """从标题中提取法源名称与文号。"""
    raw = HEADING_PREFIX_RE.sub("", text.strip())
    doc_nos = find_docnos(raw)
    names: list[str] = []
    if TITLE_LIKE_RE.match(raw):
        names.append(norm_name(re.split(r"[（(]法释|——", raw)[0]))
    else:
        for m in BOOK_RE.finditer(raw):
            n = norm_name(m.group(1))
            if n and n not in names:
                names.append(n)
    return names, doc_nos


def parse_law_text(text: str, doc_kind: str = "single") -> tuple[list[Article], list[Component]]:
    """把法规 Markdown 切分为条文；合编/汇编文件同时识别组件（按 ##/### 标题）。"""
    lines = text.splitlines()
    multi = doc_kind in ("combined", "compilation")
    components: list[Component] = []
    h2_comp: Component | None = None
    h3_comp: Component | None = None
    h2_idx = h3_idx = 0
    articles: list[Article] = []
    cur: Article | None = None
    buf: list[str] = []

    def flush():
        nonlocal cur, buf
        if cur is not None:
            cur.text = "\n".join(x for x in buf).strip()
            articles.append(cur)
        cur, buf = None, []

    for i, line in enumerate(lines, start=1):
        hm = HEADING_RE.match(line)
        if hm:
            flush()
            level = len(hm.group("hashes"))
            if multi and level in (2, 3):
                names, docnos = heading_sources(hm.group("text"))
                if level == 2:
                    h2_idx += 1
                    h3_idx = 0
                    key = f"s{h2_idx}"
                    h2_comp = Component(key, hm.group("text").strip(), names, docnos)
                    h3_comp = None
                    components.append(h2_comp)
                else:
                    h3_idx += 1
                    key = f"s{h2_idx}.{h3_idx}"
                    h3_comp = Component(key, hm.group("text").strip(), names, docnos)
                    components.append(h3_comp)
            continue
        if line.strip() == "---":
            flush()
            continue
        am = ARTICLE_RE.match(line)
        if am:
            flush()
            num = cn_to_int(am.group("num"))
            suffix = cn_to_int(am.group("suffix")) if am.group("suffix") else None
            label = f"第{am.group('num')}条" + (f"之{am.group('suffix')}" if suffix else "")
            comp_key = None
            if multi:
                for c in (h3_comp, h2_comp):
                    if c is not None and (c.names or c.doc_nos):
                        comp_key = c.key
                        break
            explicit = norm_name(am.group("law")) if am.group("law") else None
            cur = Article(num or 0, suffix, label, am.group("title"), i, "", comp_key, explicit)
            buf = [line]
            continue
        if cur is not None:
            buf.append(line)
    flush()
    return articles, components


@lru_cache(maxsize=128)
def _parsed(entry_id: str, index_path: str = ""):
    e = entry_by_id(entry_id, index_path)
    if e is None:
        raise KeyError(entry_id)
    return parse_law_text(read_entry_text(e, index_path or None), e.get("doc_kind", "single"))


def law_articles(entry_id: str, index_path: str | Path | None = None) -> tuple[list[Article], list[Component]]:
    return _parsed(entry_id, _index_key(index_path))


# ---------------------------------------------------------------------------
# 法源解析
# ---------------------------------------------------------------------------

@dataclass
class LawTarget:
    entry_id: str
    component: str | None      # 组件 key（仅合编/汇编）
    via: str                   # 命中方式：alias / docno / component
    excerpt: bool              # 是否为汇编节选


@lru_cache(maxsize=32)
def _alias_tables(index_path: str = ""):
    alias: dict[str, list[LawTarget]] = {}
    docno: dict[str, list[LawTarget]] = {}

    def add(table, key, tgt):
        if not key:
            return
        lst = table.setdefault(key, [])
        if not any(t.entry_id == tgt.entry_id and t.component == tgt.component for t in lst):
            lst.append(tgt)

    for e in entries("law", index_path or None):
        excerpt = e.get("doc_kind") == "compilation"
        for a in e.get("aliases", []):
            add(alias, norm_name(a), LawTarget(e["id"], None, "alias", excerpt))
        if e.get("doc_kind") != "compilation":
            for d in e.get("doc_nos", []):
                add(docno, norm_docno(d), LawTarget(e["id"], None, "docno", False))
        for c in e.get("components", []):
            for n in c.get("names", []):
                add(alias, norm_name(n), LawTarget(e["id"], c["key"], "component", excerpt))
            for d in c.get("doc_nos", []):
                add(docno, norm_docno(d), LawTarget(e["id"], c["key"], "component", excerpt))
    return alias, docno


@lru_cache(maxsize=2048)
def resolve_law(query: str, index_path: str = "") -> tuple[str, list[LawTarget]]:
    """解析法源名称或文号。

    返回 (status, targets)：status ∈ {"ok", "ambiguous", "not_in_kb"}。
    优先级：独立收录文件（single/combined） > 合编组件 > 汇编节选。
    """
    q = query.strip()
    alias, docno = _alias_tables(index_path)
    cands: list[LawTarget] = []
    dn = find_docnos(q)
    if dn:
        for d in dn:
            cands += docno.get(d, [])
    if not cands:
        cands = list(alias.get(norm_name(q), []))
    if not cands:
        return "not_in_kb", []
    full = [t for t in cands if not t.excerpt]
    if full:
        # 同一文件的多个命中合并
        ids = {t.entry_id for t in full}
        if len(ids) == 1:
            # 若同一文件既有整体命中又有组件命中，组件命中更精确
            comp = [t for t in full if t.component]
            return "ok", comp[:1] if len(comp) == 1 else full[:1]
        return "ambiguous", full
    return "ok", cands


@lru_cache(maxsize=4096)
def _find_articles_cached(
    entry_id: str,
    component: str | None,
    num: int,
    suffix: int | None,
    index_path: str = "",
) -> tuple[Article, ...]:
    """缓存按法源/组件/条号的精确查找。

    引用核验通常会重复引用同一条法条（意见书正文、修订批注、附录），
    之前每次都要重新遍历该法规全部条文。返回 tuple 避免调用方修改缓存
    中的列表容器。
    """
    arts, _comps = law_articles(entry_id, index_path or None)
    hits = [a for a in arts if a.num == num and a.suffix == suffix]
    if component:
        scoped = [a for a in hits if a.component and (
            a.component == component or a.component.startswith(component + ".")
        )]
        if scoped:
            return tuple(scoped)
        e = entry_by_id(entry_id, index_path)
        comp = next((c for c in (e or {}).get("components", []) if c["key"] == component), None)
        if comp:
            names = set(comp.get("names", []))
            return tuple(a for a in hits if a.explicit_law and a.explicit_law in names)
        return ()
    if len(hits) > 1:
        primary = [a for a in hits if a.explicit_law is None]
        return tuple(primary or hits)
    return tuple(hits)


def find_articles(
    target: LawTarget,
    num: int,
    suffix: int | None = None,
    index_path: str | Path | None = None,
) -> list[Article]:
    return list(_find_articles_cached(
        target.entry_id, target.component, num, suffix, _index_key(index_path)
    ))


def split_paragraphs(article_text: str) -> list[str]:
    """按“款”切分：以非空行为单位，列表项（（一）、1.）并入上一款。"""
    paras: list[str] = []
    for raw in article_text.splitlines():
        s = raw.strip()
        if not s:
            continue
        if paras and re.match(r"^(?:[（(][一二三四五六七八九十]+[）)]|\d+[\.、]|[\*\-]\s)", s):
            paras[-1] += "\n" + s
        else:
            paras.append(s)
    return paras


__all__ = [
    "SKILL_ROOT", "INDEX_PATH", "load_index", "kb_fingerprint", "kb_source_consistency", "entries", "entry_by_id", "read_entry_text",
    "norm_name", "norm_docno", "find_docnos", "title_aliases", "parse_law_text",
    "heading_sources", "Article", "Component", "LawTarget", "resolve_law", "find_articles",
    "law_articles", "split_paragraphs", "parse_num", "int_to_cn", "cn_to_int",
]
