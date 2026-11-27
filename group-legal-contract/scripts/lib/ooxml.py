# -*- coding: utf-8 -*-
"""OOXML (.docx) 底层读写、自动编号解析、字符级 Run 切分与修订/批注写入核心模块。

仅依赖标准库与 lxml / python-docx，兼容 Python 3.10+。
"""
from __future__ import annotations

import copy
import difflib
import io
import os
import posixpath
import re
import tempfile
import zipfile
from urllib.parse import unquote
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from lxml import etree

from .cn_num import int_to_cn
from .textio import atomic_write_bytes, sanitize_xml_attribute, sanitize_xml_text
from .textnorm import extract_clause_ref

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
WPS_NS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"

NSMAP = {
    "w": W_NS,
    "w14": W14_NS,
    "r": R_NS,
}

# ECMA-376 (ISO/IEC 29500-1 wml.xsd) 严格 <xsd:sequence> 子元素顺序定义
PPR_ORDER: list[str] = [
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr",
    "widowControl", "numPr", "suppressLineNumbers", "pBdr", "shd",
    "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct",
    "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
    "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents",
    "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
]

RPR_ORDER: list[str] = [
    "ins", "del", "moveFrom", "moveTo",
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps",
    "strike", "dstrike", "outline", "shadow", "emboss", "imprint",
    "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
    "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
    "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
    "eastAsianLayout", "specVanish", "oMath", "rPrChange",
]

TBLPR_ORDER: list[str] = [
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
    "tblStyleRowBandSize", "tblStyleColBandSize", "tblW", "jc",
    "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout",
    "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange",
]

TCPR_ORDER: list[str] = [
    "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders",
    "shd", "noWrap", "tcMar", "textDirection", "tcFitText", "vAlign",
    "hideMark", "headers", "cellIns", "cellDel", "cellMerge", "tcPrChange",
]

SECTPR_ORDER: list[str] = [
    "headerReference", "footerReference", "footnotePr", "endnotePr",
    "type", "pgSz", "pgMar", "paperSrc", "pgBorders", "lnNumType",
    "pgNumType", "cols", "formProt", "vAlign", "noEndnote", "titlePg",
    "textDirection", "bidi", "rtlGutter", "docGrid", "printerSettings",
    "sectPrChange",
]

SETTINGS_ORDER: list[str] = [
    "writeProtection", "view", "zoom", "removePersonalInformation",
    "removeDateAndTime", "doNotDisplayPageBoundaries",
    "displayBackgroundShape", "printPostScriptOverText",
    "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts",
    "embedSystemFonts", "saveSubsetFonts", "saveFormsData", "mirrorMargins",
    "alignBordersAndEdges", "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
    "hideGrammaticalErrors", "activeWritingStyle", "proofState",
    "formsDesign", "attachedTemplate", "linkStyles",
    "stylePaneFormatFilter", "stylePaneSortMethod", "documentType",
    "mailMerge", "revisionView", "trackRevisions", "doNotTrackMoves",
    "doNotTrackFormatting", "documentProtection", "autoFormatOverride",
    "styleLockTheme", "styleLockQFSet", "defaultTabStop", "autoHyphenation",
    "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps",
    "showEnvelope", "summaryLength", "clickAndTypeStyle",
    "defaultTableStyle", "evenAndOddHeaders", "bookFoldRevPrinting",
    "bookFoldPrinting", "bookFoldPrintingSheets",
    "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing",
    "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery",
    "doNotUseMarginsForDrawingGridOrigin", "drawingGridHorizontalOrigin",
    "drawingGridVerticalOrigin", "doNotShadeFormData", "noPunctuationKerning",
    "characterSpacingControl", "printTwoOnOne", "strictFirstAndLastChars",
    "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture",
    "doNotValidateAgainstSchema", "saveInvalidXml", "ignoreMixedContent",
    "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml",
    "saveXmlDataOnly", "useXSLTWhenSaving", "saveThroughXSLT",
    "showXMLTags", "alwaysMergeEmptyNamespace", "updateFields",
    "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars",
    "rsids", "mathPr", "attachedSchema", "themeFontLang",
    "clrSchemeMapping", "doNotIncludeSubdocsInStats",
    "doNotAutoCompressPictures", "forceUpgrade", "captions",
    "readModeInkLockDown", "smartTagType", "shapeDefaults",
    "doNotEmbedSmartTags", "decimalSymbol", "listSeparator",
]

STYLE_ORDER: list[str] = [
    "name", "aliases", "basedOn", "next", "link", "autoRedefine",
    "hidden", "uiPriority", "semiHidden", "unhideWhenUsed", "qFormat",
    "locked", "personal", "personalCompose", "personalReply", "rsid",
    "pPr", "rPr", "tblPr", "trPr", "tcPr", "tblStylePr",
]

COMMENTS_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
COMMENTS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)


def w_tag(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def w14_tag(local: str) -> str:
    return f"{{{W14_NS}}}{local}"


def xml_tag(local: str) -> str:
    return f"{{{XML_NS}}}{local}"


def r_tag(local: str) -> str:
    return f"{{{R_NS}}}{local}"


def set_xml_space_preserve(elem: etree._Element) -> None:
    """设置 xml:space='preserve'，防止 Word 吞掉首尾空格。"""
    elem.set(xml_tag("space"), "preserve")


class IdAllocator:
    """全文唯一 w:id 分配器，确保新增修订与批注 ID 不与文档已有 ID 冲突。"""

    def __init__(self, used_ids: set[int] | None = None) -> None:
        self.used_ids: set[int] = set(used_ids or ())
        self._next: int = max(self.used_ids, default=0) + 1
        if self._next < 1:
            self._next = 1

    @classmethod
    def from_trees(cls, trees: list[etree._Element]) -> "IdAllocator":
        used: set[int] = set()
        id_attr = w_tag("id")
        for root in trees:
            for elem in root.iter():
                val = elem.get(id_attr)
                if val is not None:
                    try:
                        used.add(int(val))
                    except ValueError:
                        pass
        return cls(used)

    def next_id(self) -> str:
        while self._next in self.used_ids:
            self._next += 1
        val = self._next
        self.used_ids.add(val)
        self._next += 1
        return str(val)


class DocxPackage:
    """只读加载 .docx ZIP 包并在内存中维护各 XML 部件，支持另存为新 .docx。"""

    def __init__(self, source_path: str | Path) -> None:
        self.source_path = Path(source_path)
        if not self.source_path.is_file():
            raise FileNotFoundError(f"未找到输入 docx 文件: {self.source_path}")

        self.raw_parts: dict[str, bytes] = {}
        self.xml_parts: dict[str, etree._Element] = {}
        self._modified_xml: set[str] = set()

        with zipfile.ZipFile(self.source_path, "r") as zf:
            for name in zf.namelist():
                self.raw_parts[name] = zf.read(name)

        if "word/document.xml" not in self.raw_parts:
            raise ValueError(f"非法 docx 文件（缺少 word/document.xml）: {self.source_path}")

    def get_xml(self, part_name: str) -> etree._Element | None:
        if part_name in self.xml_parts:
            return self.xml_parts[part_name]
        raw = self.raw_parts.get(part_name)
        if raw is None:
            return None
        parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
        root = etree.fromstring(raw, parser=parser)
        self.xml_parts[part_name] = root
        return root

    def set_xml(self, part_name: str, root: etree._Element) -> None:
        self.xml_parts[part_name] = root
        self._modified_xml.add(part_name)

    def mark_modified(self, part_name: str) -> None:
        if part_name in self.xml_parts:
            self._modified_xml.add(part_name)

    def remove_part(self, part_name: str) -> None:
        self.raw_parts.pop(part_name, None)
        self.xml_parts.pop(part_name, None)
        self._modified_xml.discard(part_name)

    def all_word_xml_trees(self) -> list[etree._Element]:
        trees: list[etree._Element] = []
        for name in list(self.raw_parts.keys()) + list(self.xml_parts.keys()):
            if name.endswith(".xml") and (name.startswith("word/") or name == "[Content_Types].xml"):
                tree = self.get_xml(name)
                if tree is not None and tree not in trees:
                    trees.append(tree)
        return trees

    def create_id_allocator(self) -> IdAllocator:
        return IdAllocator.from_trees(self.all_word_xml_trees())

    def save(self, out_path: str | Path, unwrap_wps_ac: bool = False) -> None:
        out_path = Path(out_path)
        if out_path.resolve() == self.source_path.resolve():
            raise ValueError("禁止覆盖原始输入文件，--out 必须指向不同的文件路径")
        if out_path.suffix.lower() != ".docx":
            raise ValueError(f"Word 输出文件必须使用 .docx 扩展名: {out_path}")
        out_path.parent.mkdir(parents=True, exist_ok=True)

        # 保存前统一执行 ECMA-376 OOXML Schema 规范化，防止桌面版 Word 报“文件已损坏”
        normalize_ooxml_package(self, unwrap_wps_ac=unwrap_wps_ac)

        all_names = list(self.raw_parts.keys())
        for name in self.xml_parts:
            if name not in all_names:
                all_names.append(name)

        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
            for name in all_names:
                if name in self._modified_xml and name in self.xml_parts:
                    xml_bytes = etree.tostring(
                        self.xml_parts[name],
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone="yes",
                    )
                    zf.writestr(name, xml_bytes)
                elif name in self.raw_parts:
                    zf.writestr(name, self.raw_parts[name])
                elif name in self.xml_parts:
                    xml_bytes = etree.tostring(
                        self.xml_parts[name],
                        xml_declaration=True,
                        encoding="UTF-8",
                        standalone="yes",
                    )
                    zf.writestr(name, xml_bytes)

        # 先写同目录临时文件并原子替换，进程中断或磁盘写入失败时保留旧文件。
        atomic_write_bytes(out_path, buf.getvalue())


def _zip_part_name_from_rels(rels_name: str) -> str | None:
    """将 OPC relationships 部件名映射到其 source part。"""
    if rels_name == "_rels/.rels":
        return None
    if not rels_name.endswith(".rels") or "/_rels/" not in rels_name:
        return None
    base, rel_file = rels_name.rsplit("/_rels/", 1)
    if not rel_file.endswith(".rels"):
        return None
    return f"{base}/{rel_file[:-5]}"


def _resolve_opc_target(source_part: str | None, target: str) -> str:
    """按 OPC 规则解析 relationships 中的内部 Target。"""
    decoded = unquote(target).replace("\\", "/")
    if decoded.startswith("/"):
        resolved = decoded.lstrip("/")
    else:
        source_dir = posixpath.dirname(source_part or "")
        resolved = posixpath.join(source_dir, decoded)
    return posixpath.normpath(resolved).lstrip("/")


def _strict_package_errors(zf: zipfile.ZipFile) -> list[str]:
    """执行不依赖外部 schema 包的 OPC/OOXML 严格包级检查。"""
    errors: list[str] = []
    names = zf.namelist()
    name_set = set(names)

    unsafe = [n for n in names if not n or n.startswith("/") or "\\" in n or ".." in n.split("/")]
    if unsafe:
        errors.append(f"ZIP 包含不安全部件路径: {', '.join(sorted(unsafe)[:5])}")

    ct_name = "[Content_Types].xml"
    ct_root = None
    try:
        ct_root = etree.fromstring(zf.read(ct_name))
    except Exception as exc:
        errors.append(f"[Content_Types].xml 无法解析: {exc}")
    if ct_root is not None:
        defaults: dict[str, str] = {}
        overrides: dict[str, str] = {}
        for child in ct_root:
            local = etree.QName(child).localname
            if local == "Default":
                ext = child.get("Extension")
                content_type = child.get("ContentType")
                if not ext or not content_type:
                    errors.append("[Content_Types].xml 存在不完整 Default")
                elif ext.lower() in defaults and defaults[ext.lower()] != content_type:
                    errors.append(f"[Content_Types].xml Default 重复且类型冲突: {ext}")
                else:
                    defaults[ext.lower()] = content_type
            elif local == "Override":
                part_name = child.get("PartName")
                content_type = child.get("ContentType")
                if not part_name or not content_type:
                    errors.append("[Content_Types].xml 存在不完整 Override")
                else:
                    key = part_name.lstrip("/")
                    if key in overrides and overrides[key] != content_type:
                        errors.append(f"[Content_Types].xml Override 重复且类型冲突: {part_name}")
                    overrides[key] = content_type

        for part in sorted(name_set - {ct_name, "_rels/.rels"}):
            if part not in overrides and posixpath.splitext(part)[1].lstrip(".").lower() not in defaults:
                errors.append(f"部件未被 [Content_Types].xml 覆盖: {part}")
        if "_rels/.rels" in name_set and "rels" not in defaults and "_rels/.rels" not in overrides:
            errors.append("部件未被 [Content_Types].xml 覆盖: _rels/.rels")
        for part in sorted(overrides):
            if part not in name_set:
                errors.append(f"[Content_Types].xml Override 指向不存在部件: {part}")

    # 每个 relationships 部件中的内部目标必须解析到 ZIP 内现有部件；关系 ID 需唯一。
    for rels_name in sorted(
        n for n in names if n == "_rels/.rels" or ("/_rels/" in n and n.endswith(".rels"))
    ):
        source_part = _zip_part_name_from_rels(rels_name)
        if source_part is not None and source_part not in name_set:
            errors.append(f"relationships 部件的 source 不存在 ({rels_name}): {source_part}")
        try:
            rel_root = etree.fromstring(zf.read(rels_name))
        except Exception as exc:
            errors.append(f"relationships 解析失败 ({rels_name}): {exc}")
            continue
        seen_ids: set[str] = set()
        for rel in rel_root:
            if etree.QName(rel).localname != "Relationship":
                continue
            rid = rel.get("Id")
            target = rel.get("Target")
            if not rid or rid in seen_ids:
                errors.append(f"relationships 存在重复或缺失 Id ({rels_name}): {rid or '<empty>'}")
            if rid:
                seen_ids.add(rid)
            if not target:
                errors.append(f"relationships 缺少 Target ({rels_name})")
                continue
            if rel.get("TargetMode", "").lower() == "external":
                continue
            resolved = _resolve_opc_target(source_part, target)
            if resolved not in name_set:
                errors.append(f"relationships 内部 Target 不存在 ({rels_name}): {target} -> {resolved}")

    # 修订与批注标记的 w:id 在整个包内保持唯一，避免 Word 无法配对。
    revision_tags = {w_tag("ins"), w_tag("del"), w_tag("moveFrom"), w_tag("moveTo")}
    comment_marker_tags = {w_tag("commentRangeStart"), w_tag("commentRangeEnd"), w_tag("commentReference")}
    revision_ids: dict[str, str] = {}
    comment_ids_by_tag: dict[str, dict[str, str]] = {
        tag: {} for tag in comment_marker_tags
    }
    for part in names:
        if not part.endswith(".xml") or not part.startswith("word/"):
            continue
        try:
            root = etree.fromstring(zf.read(part))
        except Exception:
            continue
        for elem in root.iter():
            tag = elem.tag
            if tag in revision_tags:
                value = elem.get(w_tag("id"))
                if value is not None:
                    previous = revision_ids.get(value)
                    if previous:
                        errors.append(f"修订 w:id 重复: {value} ({previous}, {part})")
                    revision_ids[value] = part
            elif tag in comment_marker_tags:
                value = elem.get(w_tag("id"))
                if value is not None:
                    previous = comment_ids_by_tag[tag].get(value)
                    if previous:
                        errors.append(f"批注标记 w:id 重复: {value} ({previous}, {part})")
                    comment_ids_by_tag[tag][value] = part
    return errors


def validate_docx_file(
    docx_path: str | Path,
    reopen: bool = True,
    strict: bool = False,
) -> dict[str, Any]:
    """验证 `.docx` 是完整 ZIP、所有 XML 可解析且可被 python-docx 重开。

    这项检查不替代 `docx_verify.py` 的修订语义校验；它专门拦截输出阶段最
    常见的“文件已损坏/无法打开”问题。返回结构稳定，适合写入运行日志。
    """
    path = Path(docx_path)
    errors: list[str] = []
    zip_ok = False
    xml_ok = False
    reopen_ok = True
    required_parts = {"[Content_Types].xml", "_rels/.rels", "word/document.xml"}

    try:
        with zipfile.ZipFile(path, "r") as zf:
            names = zf.namelist()
            if len(names) != len(set(names)):
                errors.append("ZIP 包包含重复部件名称")
            bad_name = zf.testzip()
            if bad_name:
                errors.append(f"ZIP CRC 校验失败: {bad_name}")
            missing = sorted(required_parts - set(names))
            if missing:
                errors.append(f"缺少必需 OOXML 部件: {', '.join(missing)}")
            zip_ok = not errors

            parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
            for name in names:
                if not (name.endswith(".xml") or name.endswith(".rels")):
                    continue
                try:
                    etree.fromstring(zf.read(name), parser=parser)
                except Exception as exc:
                    errors.append(f"XML 解析失败 ({name}): {exc}")
            xml_ok = not any("XML 解析失败" in e for e in errors)
            if strict and not errors:
                errors.extend(_strict_package_errors(zf))
    except Exception as exc:
        errors.append(f"读取 DOCX ZIP 失败: {exc}")

    if reopen and not errors:
        try:
            import docx

            document = docx.Document(str(path))
            _ = len(document.paragraphs)
            _ = len(document.tables)
        except Exception as exc:
            reopen_ok = False
            errors.append(f"python-docx 重新打开失败: {exc}")

    return {
        "valid": not errors,
        "zip_ok": zip_ok,
        "xml_well_formed": xml_ok,
        "python_docx_reopen": reopen_ok,
        "strict_package_checked": strict,
        "errors": errors,
    }


# ============================================================================
# 样式与自动编号解析（支持多级列表、中文数字、带圈数字、甲乙丙丁、字母等）
# ============================================================================


@dataclass
class StyleDef:
    style_id: str
    name: str
    based_on: str | None = None
    outline_lvl: int | None = None
    num_id: str | None = None
    ilvl: int | None = None


class StyleResolver:
    """解析 word/styles.xml，支持样式名称、大纲级别、继承编号属性查询。"""

    def __init__(self, styles_root: etree._Element | None) -> None:
        self.styles: dict[str, StyleDef] = {}
        self.default_p_style_id: str | None = None
        if styles_root is None:
            return

        for st in styles_root.findall(w_tag("style")):
            st_type = st.get(w_tag("type"))
            st_id = st.get(w_tag("styleId"))
            if not st_id:
                continue
            if st_type == "paragraph" and st.get(w_tag("default")) in ("1", "true"):
                self.default_p_style_id = st_id
            name_el = st.find(w_tag("name"))
            name = name_el.get(w_tag("val")) if name_el is not None else st_id
            based_el = st.find(w_tag("basedOn"))
            based_on = based_el.get(w_tag("val")) if based_el is not None else None

            outline_lvl: int | None = None
            num_id: str | None = None
            ilvl: int | None = None

            pPr = st.find(w_tag("pPr"))
            if pPr is not None:
                out_el = pPr.find(w_tag("outlineLvl"))
                if out_el is not None and out_el.get(w_tag("val")) is not None:
                    try:
                        val = int(out_el.get(w_tag("val")))
                        if 0 <= val <= 8:
                            outline_lvl = val + 1
                    except ValueError:
                        pass
                numPr = pPr.find(w_tag("numPr"))
                if numPr is not None:
                    nid_el = numPr.find(w_tag("numId"))
                    ilvl_el = numPr.find(w_tag("ilvl"))
                    if nid_el is not None:
                        num_id = nid_el.get(w_tag("val"))
                    if ilvl_el is not None and ilvl_el.get(w_tag("val")) is not None:
                        try:
                            ilvl = int(ilvl_el.get(w_tag("val")))
                        except ValueError:
                            pass

            # Heading 1..9 兜底推断 outline_level
            if outline_lvl is None and name:
                m = re.match(r"^(?:heading|标题)\s*([1-9])$", name.strip(), re.IGNORECASE)
                if m:
                    outline_lvl = int(m.group(1))

            self.styles[st_id] = StyleDef(
                style_id=st_id,
                name=name or st_id,
                based_on=based_on,
                outline_lvl=outline_lvl,
                num_id=num_id,
                ilvl=ilvl,
            )

    def get_style_name(self, style_id: str | None) -> str:
        if not style_id:
            if self.default_p_style_id and self.default_p_style_id in self.styles:
                return self.styles[self.default_p_style_id].name
            return "Normal"
        st = self.styles.get(style_id)
        return st.name if st else style_id

    def get_outline_level(self, p_elem: etree._Element) -> int | None:
        pPr = p_elem.find(w_tag("pPr"))
        style_id: str | None = None
        if pPr is not None:
            out_el = pPr.find(w_tag("outlineLvl"))
            if out_el is not None and out_el.get(w_tag("val")) is not None:
                try:
                    val = int(out_el.get(w_tag("val")))
                    return val + 1 if 0 <= val <= 8 else None
                except ValueError:
                    pass
            pStyle = pPr.find(w_tag("pStyle"))
            if pStyle is not None:
                style_id = pStyle.get(w_tag("val"))

        visited: set[str] = set()
        curr = style_id
        while curr and curr not in visited:
            visited.add(curr)
            st = self.styles.get(curr)
            if not st:
                break
            if st.outline_lvl is not None:
                return st.outline_lvl
            curr = st.based_on
        return None

    def get_style_num_pr(self, style_id: str | None) -> tuple[str | None, int | None]:
        visited: set[str] = set()
        curr = style_id
        num_id: str | None = None
        ilvl: int | None = None
        while curr and curr not in visited:
            visited.add(curr)
            st = self.styles.get(curr)
            if not st:
                break
            if num_id is None and st.num_id is not None:
                num_id = st.num_id
            if ilvl is None and st.ilvl is not None:
                ilvl = st.ilvl
            curr = st.based_on
        return num_id, ilvl


@dataclass
class LevelDef:
    ilvl: int
    start: int = 1
    num_fmt: str = "decimal"
    lvl_text: str = "%1."
    lvl_restart: int | None = None


_HEAVENLY_STEMS = ["甲", "乙", "丙", "丁", "戊", "己", "庚", "辛", "壬", "癸"]
_EARTHLY_BRANCHES = ["子", "丑", "寅", "卯", "辰", "巳", "午", "未", "申", "酉", "戌", "亥"]
_LEGAL_CN_DIGITS = "零壹贰叁肆伍陆柒捌玖"
_ENCLOSED_CIRCLES = [
    "①", "②", "③", "④", "⑤", "⑥", "⑦", "⑧", "⑨", "⑩",
    "⑪", "⑫", "⑬", "⑭", "⑮", "⑯", "⑰", "⑱", "⑲", "⑳",
]


def _int_to_roman(n: int, upper: bool = True) -> str:
    if n <= 0:
        return str(n)
    val_map = [
        (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
        (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
        (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I"),
    ]
    res: list[str] = []
    rem = n
    for v, sym in val_map:
        while rem >= v:
            res.append(sym)
            rem -= v
    s = "".join(res)
    return s if upper else s.lower()


def _int_to_letter(n: int, upper: bool = False) -> str:
    if n <= 0:
        return str(n)
    # Word lowerLetter/upperLetter: 1->a, 26->z, 27->aa, 28->ab (或重复字母)
    chars: list[str] = []
    rem = n
    while rem > 0:
        rem -= 1
        chars.append(chr((ord("A") if upper else ord("a")) + (rem % 26)))
        rem //= 26
    return "".join(reversed(chars))


def _int_to_legal_cn(n: int) -> str:
    s = int_to_cn(n)
    trans = str.maketrans("零一二三四五六七八九十百千万", "零壹贰叁肆伍陆柒捌玖拾佰仟万")
    return s.translate(trans)


def format_num_value(n: int, num_fmt: str) -> str:
    """将计数器整数按 OOXML w:numFmt 格式化为对应字符串。"""
    fmt = (num_fmt or "decimal").strip()
    if fmt == "decimal":
        return str(n)
    if fmt == "decimalZero":
        return f"{n:02d}"
    if fmt in (
        "chineseCounting",
        "chineseCountingThousand",
        "japaneseCounting",
        "taiwaneseCounting",
        "ideographDigital",
    ):
        return int_to_cn(n)
    if fmt in ("chineseLegalSimplified", "ideographLegalTraditional"):
        return _int_to_legal_cn(n)
    if fmt == "ideographTraditional":
        if 1 <= n <= len(_HEAVENLY_STEMS):
            return _HEAVENLY_STEMS[n - 1]
        return int_to_cn(n)
    if fmt == "ideographZodiac":
        if 1 <= n <= len(_EARTHLY_BRANCHES):
            return _EARTHLY_BRANCHES[n - 1]
        return int_to_cn(n)
    if fmt == "lowerLetter":
        return _int_to_letter(n, upper=False)
    if fmt == "upperLetter":
        return _int_to_letter(n, upper=True)
    if fmt == "lowerRoman":
        return _int_to_roman(n, upper=False)
    if fmt == "upperRoman":
        return _int_to_roman(n, upper=True)
    if fmt in ("decimalEnclosedCircle", "decimalEnclosedCircleChinese"):
        if 1 <= n <= len(_ENCLOSED_CIRCLES):
            return _ENCLOSED_CIRCLES[n - 1]
        return f"({n})"
    if fmt == "decimalEnclosedParen":
        return f"({n})"
    if fmt == "decimalEnclosedFullstop":
        return f"{n}."
    if fmt in ("bullet", "none"):
        return ""
    return str(n)


class NumberingResolver:
    """解析 word/numbering.xml 并按文档顺序计算每个段落的自动编号文本 (list_label)。"""

    def __init__(
        self,
        numbering_root: etree._Element | None,
        style_resolver: StyleResolver,
    ) -> None:
        self.style_resolver = style_resolver
        self.abstract_nums: dict[str, dict[int, LevelDef]] = {}
        self.num_to_abstract: dict[str, str] = {}
        self.lvl_overrides: dict[tuple[str, int], LevelDef] = {}
        self.start_overrides: dict[tuple[str, int], int] = {}
        # 状态：每个 numId（或 abstractNumId）在各 ilvl 的当前计数值
        self._counters: dict[str, dict[int, int]] = {}

        if numbering_root is None:
            return

        for abs_el in numbering_root.findall(w_tag("abstractNum")):
            abs_id = abs_el.get(w_tag("abstractNumId"))
            if not abs_id:
                continue
            levels: dict[int, LevelDef] = {}
            for lvl_el in abs_el.findall(w_tag("lvl")):
                lvl_def = self._parse_lvl_element(lvl_el)
                if lvl_def is not None:
                    levels[lvl_def.ilvl] = lvl_def
            self.abstract_nums[abs_id] = levels

        for num_el in numbering_root.findall(w_tag("num")):
            num_id = num_el.get(w_tag("numId"))
            if not num_id:
                continue
            abs_id_el = num_el.find(w_tag("abstractNumId"))
            if abs_id_el is None or not abs_id_el.get(w_tag("val")):
                continue
            abs_id = abs_id_el.get(w_tag("val"))
            self.num_to_abstract[num_id] = abs_id

            for ovr_el in num_el.findall(w_tag("lvlOverride")):
                ilvl_str = ovr_el.get(w_tag("ilvl"))
                if ilvl_str is None:
                    continue
                try:
                    ilvl = int(ilvl_str)
                except ValueError:
                    continue
                st_ovr = ovr_el.find(w_tag("startOverride"))
                if st_ovr is not None and st_ovr.get(w_tag("val")) is not None:
                    try:
                        self.start_overrides[(num_id, ilvl)] = int(st_ovr.get(w_tag("val")))
                    except ValueError:
                        pass
                sub_lvl = ovr_el.find(w_tag("lvl"))
                if sub_lvl is not None:
                    parsed = self._parse_lvl_element(sub_lvl)
                    if parsed is not None:
                        self.lvl_overrides[(num_id, ilvl)] = parsed

    @staticmethod
    def _parse_lvl_element(lvl_el: etree._Element) -> LevelDef | None:
        ilvl_str = lvl_el.get(w_tag("ilvl"))
        if ilvl_str is None:
            return None
        try:
            ilvl = int(ilvl_str)
        except ValueError:
            return None
        start = 1
        start_el = lvl_el.find(w_tag("start"))
        if start_el is not None and start_el.get(w_tag("val")) is not None:
            try:
                start = int(start_el.get(w_tag("val")))
            except ValueError:
                start = 1
        fmt_el = lvl_el.find(w_tag("numFmt"))
        num_fmt = fmt_el.get(w_tag("val")) if fmt_el is not None and fmt_el.get(w_tag("val")) else "decimal"
        txt_el = lvl_el.find(w_tag("lvlText"))
        lvl_text = txt_el.get(w_tag("val")) if txt_el is not None and txt_el.get(w_tag("val")) is not None else f"%{ilvl + 1}."
        restart: int | None = None
        rst_el = lvl_el.find(w_tag("lvlRestart"))
        if rst_el is not None and rst_el.get(w_tag("val")) is not None:
            try:
                restart = int(rst_el.get(w_tag("val")))
            except ValueError:
                restart = None
        return LevelDef(ilvl=ilvl, start=start, num_fmt=num_fmt, lvl_text=lvl_text, lvl_restart=restart)

    def _get_level_def(self, num_id: str, ilvl: int) -> LevelDef | None:
        if (num_id, ilvl) in self.lvl_overrides:
            base = copy.copy(self.lvl_overrides[(num_id, ilvl)])
        else:
            abs_id = self.num_to_abstract.get(num_id)
            if not abs_id or abs_id not in self.abstract_nums:
                return None
            lvl_map = self.abstract_nums[abs_id]
            if ilvl not in lvl_map:
                return None
            base = copy.copy(lvl_map[ilvl])
        if (num_id, ilvl) in self.start_overrides:
            base.start = self.start_overrides[(num_id, ilvl)]
        return base

    def compute_paragraph_list_label(self, p_elem: etree._Element) -> str:
        """如果段落有自动编号，推进对应级别计数器并返回格式化后的编号文本（如'第八条'、'1.2'）。"""
        pPr = p_elem.find(w_tag("pPr"))
        num_id: str | None = None
        ilvl: int | None = None
        style_id: str | None = None

        if pPr is not None:
            pStyle = pPr.find(w_tag("pStyle"))
            if pStyle is not None:
                style_id = pStyle.get(w_tag("val"))
            numPr = pPr.find(w_tag("numPr"))
            if numPr is not None:
                nid_el = numPr.find(w_tag("numId"))
                ilvl_el = numPr.find(w_tag("ilvl"))
                if nid_el is not None:
                    num_id = nid_el.get(w_tag("val"))
                if ilvl_el is not None and ilvl_el.get(w_tag("val")) is not None:
                    try:
                        ilvl = int(ilvl_el.get(w_tag("val")))
                    except ValueError:
                        pass

        if num_id is None and style_id:
            st_num_id, st_ilvl = self.style_resolver.get_style_num_pr(style_id)
            num_id = st_num_id
            if ilvl is None:
                ilvl = st_ilvl

        if not num_id or num_id == "0":
            return ""
        if ilvl is None:
            ilvl = 0

        lvl_def = self._get_level_def(num_id, ilvl)
        if lvl_def is None:
            return ""
        if lvl_def.num_fmt in ("bullet", "none"):
            return ""

        # 使用 abstractNumId 作为共享计数键（除非 numId 有 startOverride 独立计数）
        abs_id = self.num_to_abstract.get(num_id, num_id)
        counter_key = num_id if any(k[0] == num_id for k in self.start_overrides) else f"abs:{abs_id}"
        state = self._counters.setdefault(counter_key, {})

        # 确保上层级别已初始化
        for parent_lvl in range(ilvl):
            if parent_lvl not in state:
                p_def = self._get_level_def(num_id, parent_lvl)
                state[parent_lvl] = p_def.start if p_def else 1

        # 当前级别递增
        if ilvl not in state:
            state[ilvl] = lvl_def.start
        else:
            state[ilvl] += 1

        # 重置更深级别
        for deeper_lvl in list(state.keys()):
            if deeper_lvl > ilvl:
                d_def = self._get_level_def(num_id, deeper_lvl)
                if d_def is None or d_def.lvl_restart is None or d_def.lvl_restart > ilvl:
                    del state[deeper_lvl]

        # 替换 lvlText 中的 %1..%9
        label = lvl_def.lvl_text
        is_dotted_decimal = lvl_def.num_fmt == "decimal" and "." in label
        for k in range(1, 10):
            placeholder = f"%{k}"
            if placeholder in label:
                k_idx = k - 1
                k_def = self._get_level_def(num_id, k_idx)
                k_val = state.get(k_idx, k_def.start if k_def else 1)
                if k_idx < ilvl and is_dotted_decimal:
                    k_fmt = "decimal"
                else:
                    k_fmt = k_def.num_fmt if k_def else "decimal"
                label = label.replace(placeholder, format_num_value(k_val, k_fmt))

        return label.strip()


# ============================================================================
# 段落纯文本提取、修订/批注检测与块遍历
# ============================================================================

_REVISION_TAGS = frozenset(
    {
        w_tag("ins"),
        w_tag("del"),
        w_tag("moveFrom"),
        w_tag("moveTo"),
        w_tag("rPrChange"),
        w_tag("pPrChange"),
    }
)
_TEXT_REVISION_TAGS = frozenset(
    {
        w_tag("ins"),
        w_tag("del"),
        w_tag("moveFrom"),
        w_tag("moveTo"),
    }
)
_COMMENT_TAGS = frozenset(
    {
        w_tag("commentRangeStart"),
        w_tag("commentRangeEnd"),
        w_tag("commentReference"),
    }
)


def paragraph_has_existing_revisions(p_elem: etree._Element) -> bool:
    """检查段落是否包含任何已有修订标记（w:ins / w:del 等）。"""
    for el in p_elem.iter():
        if el.tag in _REVISION_TAGS:
            return True
    return False


def paragraph_has_text_revisions(p_elem: etree._Element) -> bool:
    """检查段落是否包含 w:ins / w:del / w:moveFrom / w:moveTo。"""
    for el in p_elem.iter():
        if el.tag in _TEXT_REVISION_TAGS:
            return True
    return False


def paragraph_has_comments(p_elem: etree._Element) -> bool:
    """检查段落是否包含批注标记。"""
    for el in p_elem.iter():
        if el.tag in _COMMENT_TAGS:
            return True
    return False


def extract_paragraph_accepted_text(p_elem: etree._Element) -> str:
    """按“当前接受视图”提取段落纯文本（展开超链接、域结果、内容控件、w:ins；忽略 w:del 与域指令）。"""
    chunks: list[str] = []

    def _walk(node: etree._Element) -> None:
        tag = node.tag
        # 删除修订或移出内容在“当前接受视图”下不取文本
        if tag in (w_tag("del"), w_tag("moveFrom"), w_tag("pPr"), w_tag("rPr")):
            return
        if tag == w_tag("r"):
            for child in node:
                ctag = child.tag
                if ctag == w_tag("t") and child.text:
                    chunks.append(child.text)
                elif ctag in (w_tag("tab"), w_tag("ptab")):
                    chunks.append("\t")
                elif ctag in (w_tag("br"), w_tag("cr")):
                    br_type = child.get(w_tag("type"))
                    if br_type in (None, "textWrapping"):
                        chunks.append("\n")
                elif ctag == w_tag("noBreakHyphen"):
                    chunks.append("-")
            return
        for child in node:
            _walk(child)

    _walk(p_elem)
    return "".join(chunks)


@dataclass
class BlockRecord:
    block_id: str
    seq_id: str
    location: str
    list_label: str
    clause_ref: str
    text: str
    style: str
    outline_level: int | None
    has_existing_revisions: bool
    has_comments: bool
    p_elem: etree._Element = field(repr=False)

    def to_dict(self) -> dict[str, Any]:
        return {
            "block_id": self.block_id,
            "location": self.location,
            "list_label": self.list_label,
            "clause_ref": self.clause_ref,
            "text": self.text,
            "style": self.style,
            "outline_level": self.outline_level,
            "has_existing_revisions": self.has_existing_revisions,
            "has_comments": self.has_comments,
        }


def iter_document_blocks(
    pkg: DocxPackage,
    include_headers_footers: bool = True,
) -> list[BlockRecord]:
    """遍历文档中的所有段落块（body、table、header、footer、footnote），返回 BlockRecord 列表。"""
    doc_root = pkg.get_xml("word/document.xml")
    if doc_root is None:
        return []

    style_resolver = StyleResolver(pkg.get_xml("word/styles.xml"))
    num_resolver = NumberingResolver(pkg.get_xml("word/numbering.xml"), style_resolver)

    blocks: list[BlockRecord] = []
    used_block_ids: set[str] = set()
    body_p_counter = 0
    table_counter = 0

    def _make_block(p_elem: etree._Element, seq_id: str, location: str) -> BlockRecord:
        para_id = p_elem.get(w14_tag("paraId"))
        if para_id and para_id not in used_block_ids:
            block_id = para_id
        else:
            block_id = seq_id
        used_block_ids.add(block_id)

        pPr = p_elem.find(w_tag("pPr"))
        style_id: str | None = None
        if pPr is not None:
            pStyle = pPr.find(w_tag("pStyle"))
            if pStyle is not None:
                style_id = pStyle.get(w_tag("val"))

        style_name = style_resolver.get_style_name(style_id)
        outline_lvl = style_resolver.get_outline_level(p_elem)
        list_label = num_resolver.compute_paragraph_list_label(p_elem) if location in ("body", "table") else ""
        text = extract_paragraph_accepted_text(p_elem)
        clause_ref = extract_clause_ref(list_label, text)
        has_rev = paragraph_has_existing_revisions(p_elem)
        has_com = paragraph_has_comments(p_elem)

        return BlockRecord(
            block_id=block_id,
            seq_id=seq_id,
            location=location,
            list_label=list_label,
            clause_ref=clause_ref,
            text=text,
            style=style_name,
            outline_level=outline_lvl,
            has_existing_revisions=has_rev,
            has_comments=has_com,
            p_elem=p_elem,
        )

    def _walk_table(tbl_elem: etree._Element, t_idx: int) -> None:
        nonlocal table_counter
        for r_idx, tr in enumerate(tbl_elem.findall(w_tag("tr"))):
            for c_idx, tc in enumerate(tr.findall(w_tag("tc"))):
                p_idx = 0
                for child in _iter_container_children(tc):
                    if child.tag == w_tag("p"):
                        seq_id = f"t{t_idx}-r{r_idx}-c{c_idx}-p{p_idx}"
                        blocks.append(_make_block(child, seq_id, "table"))
                        p_idx += 1
                    elif child.tag == w_tag("tbl"):
                        table_counter += 1
                        _walk_table(child, table_counter)

    def _iter_container_children(container: etree._Element) -> list[etree._Element]:
        """遍历容器直接子节点；若遇到块级 w:sdt 则展开 w:sdtContent。"""
        out: list[etree._Element] = []
        for ch in container:
            if ch.tag == w_tag("sdt"):
                sdt_content = ch.find(w_tag("sdtContent"))
                if sdt_content is not None:
                    out.extend(_iter_container_children(sdt_content))
            else:
                out.append(ch)
        return out

    body = doc_root.find(w_tag("body"))
    if body is not None:
        for child in _iter_container_children(body):
            if child.tag == w_tag("p"):
                body_p_counter += 1
                seq_id = f"p-{body_p_counter:04d}"
                blocks.append(_make_block(child, seq_id, "body"))
            elif child.tag == w_tag("tbl"):
                table_counter += 1
                _walk_table(child, table_counter)

    if include_headers_footers:
        # 页眉 header*.xml
        hdr_names = sorted(
            n for n in pkg.raw_parts if re.match(r"^word/header\d*\.xml$", n)
        )
        for h_idx, h_name in enumerate(hdr_names, start=1):
            h_root = pkg.get_xml(h_name)
            if h_root is None:
                continue
            for p_idx, p_el in enumerate(h_root.findall(f".//{w_tag('p')}"), start=1):
                txt = extract_paragraph_accepted_text(p_el)
                if txt.strip() or paragraph_has_existing_revisions(p_el) or paragraph_has_comments(p_el):
                    blocks.append(_make_block(p_el, f"hdr{h_idx}-p-{p_idx:04d}", "header"))

        # 页脚 footer*.xml
        ftr_names = sorted(
            n for n in pkg.raw_parts if re.match(r"^word/footer\d*\.xml$", n)
        )
        for f_idx, f_name in enumerate(ftr_names, start=1):
            f_root = pkg.get_xml(f_name)
            if f_root is None:
                continue
            for p_idx, p_el in enumerate(f_root.findall(f".//{w_tag('p')}"), start=1):
                txt = extract_paragraph_accepted_text(p_el)
                if txt.strip() or paragraph_has_existing_revisions(p_el) or paragraph_has_comments(p_el):
                    blocks.append(_make_block(p_el, f"ftr{f_idx}-p-{p_idx:04d}", "footer"))

        # 脚注 footnotes.xml
        fn_root = pkg.get_xml("word/footnotes.xml")
        if fn_root is not None:
            for fn_el in fn_root.findall(w_tag("footnote")):
                fn_type = fn_el.get(w_tag("type"))
                fn_id = fn_el.get(w_tag("id"), "0")
                if fn_type in ("separator", "continuationSeparator"):
                    continue
                try:
                    if int(fn_id) <= 0:
                        continue
                except ValueError:
                    pass
                for p_idx, p_el in enumerate(fn_el.findall(f".//{w_tag('p')}"), start=1):
                    txt = extract_paragraph_accepted_text(p_el)
                    if txt.strip():
                        blocks.append(_make_block(p_el, f"fn{fn_id}-p-{p_idx:04d}", "footnote"))

    return blocks


# ============================================================================
# 字符级 Run 映射、切分、最小差异计算与修订/批注写入
# ============================================================================


@dataclass
class CharAtom:
    """段落当前接受视图中单个字符对应的 XML 节点信息。"""

    char: str
    run_elem: etree._Element | None
    t_elem: etree._Element | None
    offset_in_t: int
    is_protected: bool = False
    protect_reason: str | None = None


def normalize_paragraph_runs(p_elem: etree._Element) -> None:
    """将段落直接子级中含多个内容子节点的 w:r 拆分为单一内容子节点的独立 w:r。

    这样每个普通文本 w:r 都只包含至多一个 w:rPr 和一个 w:t，后续按字符偏移切分不会误伤同 run 的 w:tab 或 w:br。
    """
    for child in list(p_elem):
        if child.tag != w_tag("r"):
            continue
        rPr = child.find(w_tag("rPr"))
        content_children = [c for c in child if c.tag != w_tag("rPr")]
        if len(content_children) <= 1:
            continue
        idx = list(p_elem).index(child)
        for c in content_children:
            new_r = etree.Element(w_tag("r"))
            if rPr is not None:
                new_r.append(copy.deepcopy(rPr))
            new_r.append(c)
            p_elem.insert(idx, new_r)
            idx += 1
        p_elem.remove(child)


def build_char_atoms(p_elem: etree._Element) -> tuple[str, list[CharAtom]]:
    """构建段落当前接受视图的 (raw_text, atoms) 映射。

    同时识别不可安全拆分/修改的保护结构：w:tab、w:br、复杂域/简单域、超链接、SDT 内容控件、已有修订。
    """
    normalize_paragraph_runs(p_elem)
    atoms: list[CharAtom] = []
    in_complex_field_depth = 0

    def _collect_from_run(
        r_el: etree._Element,
        direct_in_p: bool,
        inherited_protect: str | None,
    ) -> None:
        nonlocal in_complex_field_depth
        has_fld_char = False
        for c in r_el:
            if c.tag == w_tag("fldChar"):
                has_fld_char = True
                fld_type = c.get(w_tag("fldCharType"))
                if fld_type == "begin":
                    in_complex_field_depth += 1
                elif fld_type == "end" and in_complex_field_depth > 0:
                    in_complex_field_depth -= 1

        protect_reason = inherited_protect
        if protect_reason is None and (in_complex_field_depth > 0 or has_fld_char):
            protect_reason = "跨越域代码(w:fldChar)"
        if protect_reason is None and not direct_in_p:
            protect_reason = "位于非直接段落子节点内"

        content_children = [c for c in r_el if c.tag != w_tag("rPr")]
        for c in content_children:
            ctag = c.tag
            if ctag == w_tag("t") and c.text:
                is_prot = protect_reason is not None
                for off, ch in enumerate(c.text):
                    atoms.append(
                        CharAtom(
                            char=ch,
                            run_elem=r_el if direct_in_p else None,
                            t_elem=c if direct_in_p else None,
                            offset_in_t=off,
                            is_protected=is_prot,
                            protect_reason=protect_reason,
                        )
                    )
            elif ctag in (w_tag("tab"), w_tag("ptab")):
                atoms.append(
                    CharAtom(
                        char="\t",
                        run_elem=r_el if direct_in_p else None,
                        t_elem=None,
                        offset_in_t=0,
                        is_protected=True,
                        protect_reason=protect_reason or "跨越制表符(w:tab)",
                    )
                )
            elif ctag in (w_tag("br"), w_tag("cr")):
                br_type = c.get(w_tag("type"))
                if br_type in (None, "textWrapping"):
                    atoms.append(
                        CharAtom(
                            char="\n",
                            run_elem=r_el if direct_in_p else None,
                            t_elem=None,
                            offset_in_t=0,
                            is_protected=True,
                            protect_reason=protect_reason or "跨越换行符(w:br)",
                        )
                    )
            elif ctag == w_tag("noBreakHyphen"):
                atoms.append(
                    CharAtom(
                        char="-",
                        run_elem=r_el if direct_in_p else None,
                        t_elem=None,
                        offset_in_t=0,
                        is_protected=True,
                        protect_reason=protect_reason or "跨越不换行连字符(w:noBreakHyphen)",
                    )
                )

    def _walk_child(node: etree._Element, direct_in_p: bool, protect: str | None) -> None:
        tag = node.tag
        if tag in (w_tag("del"), w_tag("moveFrom"), w_tag("pPr"), w_tag("rPr")):
            return
        if tag == w_tag("r"):
            _collect_from_run(node, direct_in_p=direct_in_p, inherited_protect=protect)
            return
        if tag == w_tag("hyperlink"):
            sub_prot = protect or "跨越超链接(w:hyperlink)"
        elif tag == w_tag("fldSimple"):
            sub_prot = protect or "跨越简单域(w:fldSimple)"
        elif tag == w_tag("sdt"):
            sub_prot = protect or "跨越内容控件(w:sdt)"
        elif tag in (w_tag("ins"), w_tag("moveTo")):
            sub_prot = protect or "跨越已有修订(w:ins)"
        else:
            sub_prot = protect or f"跨越复杂结构({etree.QName(tag).localname})"

        for ch in node:
            _walk_child(ch, direct_in_p=False, protect=sub_prot)

    for top_child in p_elem:
        _walk_child(top_child, direct_in_p=True, protect=None)

    raw_text = "".join(a.char for a in atoms)
    return raw_text, atoms


def split_run_at_char_index(
    p_elem: etree._Element,
    atoms: list[CharAtom],
    char_idx: int,
) -> list[CharAtom]:
    """确保在段落字符索引 char_idx 处存在 w:r 边界（若位于某 w:t 内部则将该 w:r 一分为二）。"""
    if char_idx <= 0 or char_idx >= len(atoms):
        return atoms

    target = atoms[char_idx]
    if target.offset_in_t == 0:
        return atoms
    if target.run_elem is None or target.t_elem is None:
        return atoms

    r_elem = target.run_elem
    t_elem = target.t_elem
    full_text = t_elem.text or ""
    split_pos = target.offset_in_t
    left_text = full_text[:split_pos]
    right_text = full_text[split_pos:]

    t_elem.text = left_text
    set_xml_space_preserve(t_elem)

    # 构造右侧新 w:r
    new_r = etree.Element(w_tag("r"))
    rPr = r_elem.find(w_tag("rPr"))
    if rPr is not None:
        new_r.append(copy.deepcopy(rPr))
    new_t = etree.SubElement(new_r, w_tag("t"))
    new_t.text = right_text
    set_xml_space_preserve(new_t)

    parent = r_elem.getparent()
    if parent is not None:
        r_pos = list(parent).index(r_elem)
        parent.insert(r_pos + 1, new_r)

    # 原地更新右侧同一 t_elem 的 CharAtom，无需全量重建以保持已有修订后的索引稳定
    for i in range(char_idx, len(atoms)):
        if atoms[i].t_elem is t_elem:
            atoms[i].run_elem = new_r
            atoms[i].t_elem = new_t
            atoms[i].offset_in_t -= split_pos
        else:
            break

    return atoms


def compute_minimal_diff_opcodes(
    orig_text: str,
    new_text: str,
    min_equal_len: int = 2,
) -> list[tuple[str, int, int, int, int]]:
    """计算字符级最小差异 opcodes，并将夹在修改之间且长度 < min_equal_len 的碎片 equal 段并入修改。

    首尾的 equal 片段无论长度均保留，避免误将首尾未变字符标为删除再插入。
    """
    matcher = difflib.SequenceMatcher(None, orig_text, new_text, autojunk=False)
    raw_ops = list(matcher.get_opcodes())
    if len(raw_ops) <= 1:
        return raw_ops

    # 将内部（非首、非尾）长度 < min_equal_len 的 equal 转为 replace，再合并相邻非 equal 块
    converted: list[tuple[str, int, int, int, int]] = []
    last_idx = len(raw_ops) - 1
    for idx, (tag, i1, i2, j1, j2) in enumerate(raw_ops):
        if tag == "equal" and 0 < idx < last_idx and (i2 - i1) < min_equal_len:
            converted.append(("replace", i1, i2, j1, j2))
        else:
            converted.append((tag, i1, i2, j1, j2))

    merged: list[tuple[str, int, int, int, int]] = []
    for tag, i1, i2, j1, j2 in converted:
        if not merged:
            merged.append((tag, i1, i2, j1, j2))
            continue
        prev_tag, pi1, pi2, pj1, pj2 = merged[-1]
        if prev_tag != "equal" and tag != "equal":
            ni1, ni2 = pi1, i2
            nj1, nj2 = pj1, j2
            if ni2 > ni1 and nj2 > nj1:
                ntag = "replace"
            elif ni2 > ni1:
                ntag = "delete"
            else:
                ntag = "insert"
            merged[-1] = (ntag, ni1, ni2, nj1, nj2)
        elif prev_tag == "equal" and tag == "equal":
            merged[-1] = ("equal", pi1, i2, pj1, j2)
        else:
            merged.append((tag, i1, i2, j1, j2))

    return merged


def clone_rPr(r_elem: etree._Element | None) -> etree._Element | None:
    """复制指定 w:r 的 w:rPr（若存在），剔除旧 rPrChange。"""
    if r_elem is None:
        return None
    rPr = r_elem.find(w_tag("rPr"))
    if rPr is None:
        return None
    cloned = copy.deepcopy(rPr)
    for chg in cloned.findall(w_tag("rPrChange")):
        cloned.remove(chg)
    return cloned


def make_text_run(
    text: str,
    rPr_template: etree._Element | None = None,
    is_del: bool = False,
) -> etree._Element:
    """创建包含 w:t 或 w:delText 的 w:r 元素。"""
    r_el = etree.Element(w_tag("r"))
    if rPr_template is not None:
        r_el.append(copy.deepcopy(rPr_template))
    t_tag = w_tag("delText") if is_del else w_tag("t")
    t_el = etree.SubElement(r_el, t_tag)
    clean_text, _ = sanitize_xml_text(str(text))
    t_el.text = clean_text
    set_xml_space_preserve(t_el)
    return r_el


def build_runs_preserving_rPr(
    atoms_slice: list[CharAtom],
    is_del: bool = False,
) -> list[etree._Element]:
    """根据原字符切片中的 run_elem 分组，生成保留各自原有 w:rPr 的 w:r 列表。"""
    if not atoms_slice:
        return []
    runs: list[etree._Element] = []
    curr_run = atoms_slice[0].run_elem
    curr_chars: list[str] = [atoms_slice[0].char]

    for atom in atoms_slice[1:]:
        if atom.run_elem is curr_run:
            curr_chars.append(atom.char)
        else:
            runs.append(make_text_run("".join(curr_chars), clone_rPr(curr_run), is_del=is_del))
            curr_run = atom.run_elem
            curr_chars = [atom.char]

    runs.append(make_text_run("".join(curr_chars), clone_rPr(curr_run), is_del=is_del))
    return runs


def make_del_element(
    rev_id: str,
    author: str,
    date_iso: str,
    runs: list[etree._Element],
) -> etree._Element:
    del_el = etree.Element(
        w_tag("del"),
        {
            w_tag("id"): sanitize_xml_attribute(str(rev_id))[0],
            w_tag("author"): sanitize_xml_attribute(str(author))[0],
            w_tag("date"): sanitize_xml_attribute(str(date_iso))[0],
        },
    )
    for r in runs:
        del_el.append(r)
    return del_el


def make_ins_element(
    rev_id: str,
    author: str,
    date_iso: str,
    runs: list[etree._Element],
) -> etree._Element:
    ins_el = etree.Element(
        w_tag("ins"),
        {
            w_tag("id"): sanitize_xml_attribute(str(rev_id))[0],
            w_tag("author"): sanitize_xml_attribute(str(author))[0],
            w_tag("date"): sanitize_xml_attribute(str(date_iso))[0],
        },
    )
    for r in runs:
        ins_el.append(r)
    return ins_el


# ============================================================================
# 批注（word/comments.xml、关系与 Content_Types）与修订跟踪开关
# ============================================================================


def format_comment_text(
    rev_id: str,
    comment: dict[str, Any] | None,
    fallback_suggestion: str | None = None,
) -> str:
    """按规范组装批注文本：
    “【R-07｜高风险】问题：…／依据：…／底线：…／建议：…”（字段为空则省略）。
    """
    comment = comment or {}
    risk_raw = (comment.get("risk") or "").strip()
    if risk_raw in ("高", "中", "低"):
        risk_label = f"{risk_raw}风险"
    elif risk_raw:
        risk_label = risk_raw
    else:
        risk_label = "提示"

    header = f"【{rev_id}｜{risk_label}】"
    parts: list[str] = []

    issue = (comment.get("issue") or "").strip()
    if issue:
        parts.append(f"问题：{issue}")

    basis = comment.get("basis") or []
    if isinstance(basis, list):
        basis_items = [str(b).strip() for b in basis if b and str(b).strip()]
        if basis_items:
            parts.append(f"依据：{'；'.join(basis_items)}")
    elif isinstance(basis, str) and basis.strip():
        parts.append(f"依据：{basis.strip()}")

    fallback = (comment.get("fallback") or "").strip()
    if fallback:
        parts.append(f"底线：{fallback}")

    suggestion = (comment.get("suggestion") or fallback_suggestion or "").strip()
    if suggestion:
        parts.append(f"建议：{suggestion}")

    if not parts and fallback_suggestion:
        parts.append(f"建议：{fallback_suggestion}")

    return header + "／".join(parts)


def ensure_comments_part(pkg: DocxPackage) -> etree._Element:
    """获取或创建 word/comments.xml，并补全 word/_rels/document.xml.rels 与 [Content_Types].xml。"""
    comments_root = pkg.get_xml("word/comments.xml")
    if comments_root is None:
        comments_root = etree.Element(w_tag("comments"), nsmap=NSMAP)
        pkg.set_xml("word/comments.xml", comments_root)
    else:
        pkg.mark_modified("word/comments.xml")

    # 1. 补全 word/_rels/document.xml.rels
    rels_path = "word/_rels/document.xml.rels"
    rels_root = pkg.get_xml(rels_path)
    if rels_root is None:
        rels_root = etree.Element(f"{{{PKG_REL_NS}}}Relationships", nsmap={None: PKG_REL_NS})
        pkg.set_xml(rels_path, rels_root)

    has_comments_rel = False
    existing_rids: set[int] = set()
    for rel in rels_root:
        rid = rel.get("Id", "")
        m = re.match(r"^rId(\d+)$", rid)
        if m:
            existing_rids.add(int(m.group(1)))
        if rel.get("Type") == COMMENTS_REL_TYPE:
            has_comments_rel = True

    if not has_comments_rel:
        next_rid_num = max(existing_rids, default=0) + 1
        rel_el = etree.SubElement(rels_root, f"{{{PKG_REL_NS}}}Relationship")
        rel_el.set("Id", f"rId{next_rid_num}")
        rel_el.set("Type", COMMENTS_REL_TYPE)
        rel_el.set("Target", "comments.xml")
        pkg.mark_modified(rels_path)

    # 2. 补全 [Content_Types].xml
    ct_path = "[Content_Types].xml"
    ct_root = pkg.get_xml(ct_path)
    if ct_root is not None:
        has_ct = False
        for child in ct_root:
            if child.get("PartName") in ("/word/comments.xml", "word/comments.xml"):
                has_ct = True
                break
        if not has_ct:
            ovr = etree.SubElement(ct_root, f"{{{CT_NS}}}Override")
            ovr.set("PartName", "/word/comments.xml")
            ovr.set("ContentType", COMMENTS_CONTENT_TYPE)
            pkg.mark_modified(ct_path)

    # 3. 确保 word/styles.xml 中存在 CommentText 与 CommentReference 样式
    ensure_comment_styles_in_styles_xml(pkg)

    return comments_root


def add_comment_entry(
    pkg: DocxPackage,
    comment_id: str,
    author: str,
    date_iso: str,
    text: str,
) -> etree._Element:
    """在 word/comments.xml 中追加一条 <w:comment> 元素，支持通过换行符分多段落。"""
    comments_root = ensure_comments_part(pkg)
    clean_author, _ = sanitize_xml_attribute(str(author))
    clean_date, _ = sanitize_xml_attribute(str(date_iso))
    clean_comment_id, _ = sanitize_xml_attribute(str(comment_id))
    comment_el = etree.SubElement(
        comments_root,
        w_tag("comment"),
        {
            w_tag("id"): clean_comment_id,
            w_tag("author"): clean_author,
            w_tag("date"): clean_date,
            w_tag("initials"): clean_author[:2] if clean_author else "法务",
        },
    )
    lines = text.splitlines() if text else [""]
    if not lines:
        lines = [""]

    for idx, line in enumerate(lines):
        p_el = etree.SubElement(comment_el, w_tag("p"))
        pPr = etree.SubElement(p_el, w_tag("pPr"))
        pStyle = etree.SubElement(pPr, w_tag("pStyle"))
        pStyle.set(w_tag("val"), "CommentText")

        if idx == 0:
            ref_r = etree.SubElement(p_el, w_tag("r"))
            ref_rPr = etree.SubElement(ref_r, w_tag("rPr"))
            ref_rStyle = etree.SubElement(ref_rPr, w_tag("rStyle"))
            ref_rStyle.set(w_tag("val"), "CommentReference")
            etree.SubElement(ref_r, w_tag("annotationRef"))

        txt_r = etree.SubElement(p_el, w_tag("r"))
        t_el = etree.SubElement(txt_r, w_tag("t"))
        clean_line, _ = sanitize_xml_text(str(line))
        t_el.text = clean_line
        set_xml_space_preserve(t_el)

    pkg.mark_modified("word/comments.xml")
    return comment_el


def make_comment_Markers(
    comment_id: str,
) -> tuple[etree._Element, etree._Element, etree._Element]:
    """创建正文中配对的 (commentRangeStart, commentRangeEnd, commentReference_run) 三元组。"""
    c_start = etree.Element(w_tag("commentRangeStart"), {w_tag("id"): comment_id})
    c_end = etree.Element(w_tag("commentRangeEnd"), {w_tag("id"): comment_id})
    c_ref_run = etree.Element(w_tag("r"))
    rPr = etree.SubElement(c_ref_run, w_tag("rPr"))
    rStyle = etree.SubElement(rPr, w_tag("rStyle"))
    rStyle.set(w_tag("val"), "CommentReference")
    etree.SubElement(c_ref_run, w_tag("commentReference"), {w_tag("id"): comment_id})
    return c_start, c_end, c_ref_run


def attach_comment_around_elements(
    parent_p: etree._Element,
    first_el: etree._Element,
    last_el: etree._Element,
    comment_id: str,
) -> None:
    """在段落内 first_el 之前插入 commentRangeStart，last_el 之后插入 commentRangeEnd 与 commentReference。"""
    c_start, c_end, c_ref_run = make_comment_Markers(comment_id)
    children = list(parent_p)
    start_idx = children.index(first_el)
    parent_p.insert(start_idx, c_start)
    children = list(parent_p)
    end_idx = children.index(last_el)
    parent_p.insert(end_idx + 1, c_end)
    parent_p.insert(end_idx + 2, c_ref_run)


def attach_comment_to_whole_paragraph(
    p_elem: etree._Element,
    comment_id: str,
) -> None:
    """将批注挂在整段上（在 pPr 之后插入 commentRangeStart，段末插入 commentRangeEnd 与 commentReference）。"""
    c_start, c_end, c_ref_run = make_comment_Markers(comment_id)
    pPr = p_elem.find(w_tag("pPr"))
    if pPr is not None:
        insert_pos = list(p_elem).index(pPr) + 1
    else:
        insert_pos = 0
    p_elem.insert(insert_pos, c_start)
    p_elem.append(c_end)
    p_elem.append(c_ref_run)


def enable_track_revisions(pkg: DocxPackage) -> None:
    """在 word/settings.xml 中开启 <w:trackRevisions/>（严格遵循 ECMA-376 CT_Settings 子元素顺序）。"""
    settings_path = "word/settings.xml"
    settings_root = pkg.get_xml(settings_path)
    if settings_root is None:
        settings_root = etree.Element(w_tag("settings"), nsmap=NSMAP)
        pkg.set_xml(settings_path, settings_root)
    if settings_root.find(w_tag("trackRevisions")) is None:
        etree.SubElement(settings_root, w_tag("trackRevisions"))
        _sort_children_ecma376(settings_root, SETTINGS_ORDER)
        pkg.mark_modified(settings_path)


# ============================================================================
# ECMA-376 OOXML Schema 规范化与 LibreOffice 转换污染修复
# ============================================================================


def _sort_children_ecma376(parent: etree._Element, order_list: list[str]) -> bool:
    """将 parent 的直接子元素按 ECMA-376 <xsd:sequence> 顺序稳定重排。返回是否发生了顺序调整。"""
    if len(parent) <= 1:
        return False
    order_map = {w_tag(name): idx for idx, name in enumerate(order_list)}
    # m:mathPr is part of CT_Settings but belongs to the math namespace rather
    # than the w: namespace. Treat it as the corresponding sequence item;
    # otherwise the generic unknown-element fallback moves it after listSeparator.
    if parent.tag == w_tag("settings") and "mathPr" in order_list:
        order_map[f"{{{M_NS}}}mathPr"] = order_list.index("mathPr")
    children = list(parent)
    indexed = [
        (order_map.get(ch.tag, 10000 + pos), pos, ch)
        for pos, ch in enumerate(children)
    ]
    sorted_indexed = sorted(indexed, key=lambda item: (item[0], item[1]))
    sorted_children = [item[2] for item in sorted_indexed]
    if sorted_children == children:
        return False
    parent[:] = sorted_children
    return True


def _dedup_single_children(parent: etree._Element, allowed_multi_localnames: set[str] | None = None) -> bool:
    """在属性容器（如 w:rPr、w:pPr、w:tblPr、w:tcPr）中合并并去重重复出现的 maxOccurs=1 子元素。"""
    if len(parent) <= 1:
        return False
    multi_tags = {w_tag(name) for name in (allowed_multi_localnames or set())}
    seen: dict[str, etree._Element] = {}
    changed = False
    for ch in list(parent):
        if not isinstance(ch.tag, str) or not ch.tag.startswith(f"{{{W_NS}}}"):
            continue
        if ch.tag in multi_tags:
            continue
        if ch.tag not in seen:
            seen[ch.tag] = ch
        else:
            prev = seen[ch.tag]
            for k, v in ch.attrib.items():
                if k not in prev.attrib:
                    prev.set(k, v)
            for sub in list(ch):
                if prev.find(sub.tag) is None:
                    prev.append(sub)
            parent.remove(ch)
            changed = True
    return changed


def _clean_single_font_val(val: str, is_east_asia: bool = False) -> str:
    """清理 LibreOffice 导出在 w:rFonts 中的分号拼接字体名（如 '宋体;SimSun'）与乱码字符（如 'ˎ̥;Times New Roman'）。"""
    if not val:
        return val
    # 检测分号或明显乱码控制/修饰符
    has_garbled = any(ord(c) in (0x02CE, 0x0325, 0xFFFD) for c in val)
    if ";" not in val and not has_garbled:
        return val
    parts = [p.strip() for p in val.split(";") if p.strip()]
    clean_parts: list[str] = []
    for p in parts:
        cleaned = "".join(c for c in p if ord(c) not in (0x02CE, 0x0325, 0xFFFD)).strip()
        if cleaned:
            clean_parts.append(cleaned)
    if not clean_parts:
        return "宋体" if is_east_asia else "Times New Roman"
    if is_east_asia:
        for preferred in ("宋体", "黑体", "仿宋", "楷体", "微软雅黑", "SimSun", "SimHei"):
            for cp in clean_parts:
                if preferred in cp:
                    return "宋体" if preferred == "SimSun" else ("黑体" if preferred == "SimHei" else preferred)
        return clean_parts[0]
    else:
        for preferred in ("Times New Roman", "Arial", "Calibri", "宋体", "黑体"):
            for cp in clean_parts:
                if preferred.lower() in cp.lower():
                    return preferred
        return clean_parts[0]


def _sanitize_rfonts_in_tree(root: etree._Element) -> bool:
    """清理树中所有 <w:rFonts> 的分号拼接字体名或乱码字符。"""
    changed = False
    for rf in root.iter(w_tag("rFonts")):
        for attr_local, is_ea in (
            ("ascii", False),
            ("hAnsi", False),
            ("eastAsia", True),
            ("cs", False),
        ):
            attr_q = w_tag(attr_local)
            val = rf.get(attr_q)
            if val is not None:
                new_val = _clean_single_font_val(val, is_east_asia=is_ea)
                if new_val != val:
                    rf.set(attr_q, new_val)
                    changed = True
    return changed


def _fix_or_unwrap_wps_alternate_content(root: etree._Element, unwrap: bool = False) -> bool:
    """修复或解包 LibreOffice .doc -> .docx 转换生成的损坏 <mc:AlternateContent>。
    - 当 unwrap=True 时：若 <mc:Fallback> 下含 <w:pict>，直接用 <mc:Fallback> 的子元素替换 <mc:AlternateContent>，
      彻底消除损坏的 <mc:Choice Requires="wps"> 及重复的封面文本框段落。
    - 当 unwrap=False 时：保持树结构不变，原地修补 <mc:Choice Requires="wps"> 中缺失的 mc:Ignorable、
      <wp:cNvGraphicFramePr/>、wp:docPr/@name、wp:extent/@cy 以及 <wps:style> 颜色子元素。
    """
    changed = False
    ac_tag = f"{{{MC_NS}}}AlternateContent"
    choice_tag = f"{{{MC_NS}}}Choice"
    fallback_tag = f"{{{MC_NS}}}Fallback"

    ac_list = list(root.iter(ac_tag))
    if not ac_list:
        return False

    if unwrap:
        for ac in ac_list:
            parent = ac.getparent()
            if parent is None:
                continue
            fb = ac.find(fallback_tag)
            if fb is not None and fb.find(w_tag("pict")) is not None:
                idx = list(parent).index(ac)
                for ch in list(fb):
                    parent.insert(idx, ch)
                    idx += 1
                parent.remove(ac)
                changed = True

        # 检查解包后树中是否不再使用 wps / wpg，确保 mc:Ignorable 仅包含已声明的前缀
        ig_attr = f"{{{MC_NS}}}Ignorable"
        ig_val = root.get(ig_attr)
        if ig_val:
            declared = set(root.nsmap.keys())
            valid_prefixes = [p for p in ig_val.split() if p in declared]
            new_ig = " ".join(valid_prefixes)
            if new_ig != ig_val:
                root.set(ig_attr, new_ig)
                changed = True
        return changed

    # unwrap=False：原地修补 DrawingML / WPS Schema 错误
    ig_attr = f"{{{MC_NS}}}Ignorable"
    ig_val = root.get(ig_attr) or ""
    ig_tokens = ig_val.split()
    declared = set(root.nsmap.keys())
    for req_p in ("wps", "wpg"):
        if req_p in declared and req_p not in ig_tokens:
            ig_tokens.append(req_p)
    new_ig = " ".join(p for p in ig_tokens if p in declared)
    if new_ig and new_ig != ig_val:
        root.set(ig_attr, new_ig)
        changed = True

    shape_counter = 100
    for ac in ac_list:
        ch_el = ac.find(choice_tag)
        if ch_el is None:
            continue
        # 1. 修补 wp:anchor / wp:inline 缺少 <wp:cNvGraphicFramePr/>、name=""、cy="0"
        for wp_container_local in ("anchor", "inline"):
            for wp_el in ch_el.iter(f"{{{WP_NS}}}{wp_container_local}"):
                ext_el = wp_el.find(f"{{{WP_NS}}}extent")
                if ext_el is not None and ext_el.get("cy") == "0":
                    ext_el.set("cy", "12700")
                    changed = True
                doc_pr = wp_el.find(f"{{{WP_NS}}}docPr")
                if doc_pr is not None and not (doc_pr.get("name") or "").strip():
                    shape_counter += 1
                    doc_pr.set("name", f"Shape_{shape_counter}")
                    changed = True
                cnv_el = wp_el.find(f"{{{WP_NS}}}cNvGraphicFramePr")
                graphic_el = wp_el.find(f"{{{A_NS}}}graphic")
                if cnv_el is None and graphic_el is not None:
                    new_cnv = etree.Element(f"{{{WP_NS}}}cNvGraphicFramePr")
                    g_idx = list(wp_el).index(graphic_el)
                    wp_el.insert(g_idx, new_cnv)
                    changed = True
        # 2. 修补 <wps:style> 下空的 <a:lnRef>/<a:fillRef>/<a:effectRef>/<a:fontRef>
        for style_el in ch_el.iter(f"{{{WPS_NS}}}style"):
            for ref_local in ("lnRef", "fillRef", "effectRef", "fontRef"):
                ref_el = style_el.find(f"{{{A_NS}}}{ref_local}")
                if ref_el is not None and len(ref_el) == 0:
                    clr = etree.SubElement(ref_el, f"{{{A_NS}}}srgbClr")
                    clr.set("val", "000000")
                    changed = True

    return changed


def normalize_ooxml_tree(root: etree._Element, unwrap_wps_ac: bool = False) -> bool:
    """对单个 OOXML 树执行 ECMA-376 规范化（去重、严格序列重排、字体清理、AlternateContent 修复）。"""
    changed = False

    if _fix_or_unwrap_wps_alternate_content(root, unwrap=unwrap_wps_ac):
        changed = True

    if _sanitize_rfonts_in_tree(root):
        changed = True

    if root.tag == w_tag("settings"):
        if _dedup_single_children(root, allowed_multi_localnames={"activeWritingStyle", "attachedSchema", "smartTagType"}):
            changed = True
        # ECMA-376 CT_Zoom requires w:percent even when w:val="bestFit" is used.
        # Older Word/python-docx templates commonly omit it; adding the neutral
        # 100% value keeps the setting visually unchanged while making the part
        # acceptable to strict OOXML validators.
        zoom_el = root.find(w_tag("zoom"))
        if zoom_el is not None and zoom_el.get(w_tag("percent")) is None:
            zoom_el.set(w_tag("percent"), "100")
            changed = True
        if _sort_children_ecma376(root, SETTINGS_ORDER):
            changed = True

    if root.tag == w_tag("styles"):
        for st in root.findall(w_tag("style")):
            if _dedup_single_children(st, allowed_multi_localnames={"rsid", "tblStylePr"}):
                changed = True
            if _sort_children_ecma376(st, STYLE_ORDER):
                changed = True

    for pPr in root.iter(w_tag("pPr")):
        if _dedup_single_children(pPr):
            changed = True
        if _sort_children_ecma376(pPr, PPR_ORDER):
            changed = True

    for rPr in root.iter(w_tag("rPr")):
        if _dedup_single_children(rPr, allowed_multi_localnames={"ins", "del", "moveFrom", "moveTo"}):
            changed = True
        if _sort_children_ecma376(rPr, RPR_ORDER):
            changed = True

    for tblPr in root.iter(w_tag("tblPr")):
        if _dedup_single_children(tblPr):
            changed = True
        if _sort_children_ecma376(tblPr, TBLPR_ORDER):
            changed = True

    for tcPr in root.iter(w_tag("tcPr")):
        if _dedup_single_children(tcPr, allowed_multi_localnames={"headers"}):
            changed = True
        if _sort_children_ecma376(tcPr, TCPR_ORDER):
            changed = True

    for sectPr in root.iter(w_tag("sectPr")):
        if _dedup_single_children(sectPr, allowed_multi_localnames={"headerReference", "footerReference"}):
            changed = True
        if _sort_children_ecma376(sectPr, SECTPR_ORDER):
            changed = True

    return changed


def ensure_comment_styles_in_styles_xml(pkg: DocxPackage) -> bool:
    """确保 word/styles.xml 中包含 CommentText 与 CommentReference 样式定义。"""
    styles_path = "word/styles.xml"
    styles_root = pkg.get_xml(styles_path)
    if styles_root is None:
        return False

    existing_ids = {
        st.get(w_tag("styleId")) for st in styles_root.findall(w_tag("style"))
    }
    changed = False

    if "CommentText" not in existing_ids:
        st_ct = etree.SubElement(
            styles_root,
            w_tag("style"),
            {w_tag("type"): "paragraph", w_tag("styleId"): "CommentText"},
        )
        nm = etree.SubElement(st_ct, w_tag("name"))
        nm.set(w_tag("val"), "annotation text")
        rPr = etree.SubElement(st_ct, w_tag("rPr"))
        rFonts = etree.SubElement(rPr, w_tag("rFonts"))
        rFonts.set(w_tag("ascii"), "Times New Roman")
        rFonts.set(w_tag("hAnsi"), "Times New Roman")
        rFonts.set(w_tag("eastAsia"), "宋体")
        sz = etree.SubElement(rPr, w_tag("sz"))
        sz.set(w_tag("val"), "18")
        szCs = etree.SubElement(rPr, w_tag("szCs"))
        szCs.set(w_tag("val"), "18")
        changed = True

    if "CommentReference" not in existing_ids:
        st_cr = etree.SubElement(
            styles_root,
            w_tag("style"),
            {w_tag("type"): "character", w_tag("styleId"): "CommentReference"},
        )
        nm = etree.SubElement(st_cr, w_tag("name"))
        nm.set(w_tag("val"), "annotation reference")
        rPr = etree.SubElement(st_cr, w_tag("rPr"))
        sz = etree.SubElement(rPr, w_tag("sz"))
        sz.set(w_tag("val"), "16")
        szCs = etree.SubElement(rPr, w_tag("szCs"))
        szCs.set(w_tag("val"), "16")
        changed = True

    if changed:
        pkg.mark_modified(styles_path)
    return changed


def reassign_colliding_bookmark_ids(pkg: DocxPackage) -> bool:
    """检查文档包中 w:bookmarkStart/w:bookmarkEnd 的 w:id 是否与修订/批注 w:id 冲突，若有冲突则重编号书签 ID。"""
    id_attr = w_tag("id")
    rev_and_com_ids: set[str] = set()
    all_int_ids: set[int] = set()
    bookmark_starts: list[tuple[str, etree._Element]] = []
    bookmark_ends: list[tuple[str, etree._Element]] = []

    part_names = [
        name
        for name in list(pkg.raw_parts.keys()) + list(pkg.xml_parts.keys())
        if name.startswith("word/") and name.endswith(".xml")
    ]
    seen_parts: set[str] = set()
    for name in part_names:
        if name in seen_parts:
            continue
        seen_parts.add(name)
        root = pkg.get_xml(name)
        if root is None:
            continue
        for el in root.iter():
            val = el.get(id_attr)
            if val is None:
                continue
            try:
                all_int_ids.add(int(val))
            except ValueError:
                pass
            if el.tag == w_tag("bookmarkStart"):
                bookmark_starts.append((name, el))
            elif el.tag == w_tag("bookmarkEnd"):
                bookmark_ends.append((name, el))
            elif el.tag in (
                w_tag("ins"),
                w_tag("del"),
                w_tag("moveFrom"),
                w_tag("moveTo"),
                w_tag("comment"),
                w_tag("commentRangeStart"),
                w_tag("commentRangeEnd"),
                w_tag("commentReference"),
            ):
                rev_and_com_ids.add(val)

    bm_ids = {el.get(id_attr, "") for _, el in bookmark_starts}
    if not (bm_ids & rev_and_com_ids):
        return False

    next_num = max(all_int_ids, default=0) + 100
    remap: dict[str, str] = {}
    modified_parts: set[str] = set()

    for part_name, el in bookmark_starts:
        old_id = el.get(id_attr, "")
        if old_id in rev_and_com_ids:
            if old_id not in remap:
                remap[old_id] = str(next_num)
                next_num += 1
            el.set(id_attr, remap[old_id])
            modified_parts.add(part_name)

    for part_name, el in bookmark_ends:
        old_id = el.get(id_attr, "")
        if old_id in remap:
            el.set(id_attr, remap[old_id])
            modified_parts.add(part_name)

    for p_name in modified_parts:
        pkg.mark_modified(p_name)
    return bool(modified_parts)


def normalize_ooxml_package(pkg: DocxPackage, unwrap_wps_ac: bool = False) -> None:
    """对 DocxPackage 中的所有 word/*.xml 部件执行 ECMA-376 规范化。"""
    # 清理主部件不存在的孤儿 relationships 部件（如 WPS 残留的空 word/_rels/comments.xml.rels）
    present_parts = set(pkg.raw_parts.keys()) | set(pkg.xml_parts.keys())
    for part_name in list(present_parts):
        source_part = _zip_part_name_from_rels(part_name)
        if source_part is not None and source_part not in present_parts:
            pkg.remove_part(part_name)
            present_parts.discard(part_name)

    # 清理 [Content_Types].xml 中指向不存在部件的悬空 Override
    ct_root = pkg.get_xml("[Content_Types].xml")
    if ct_root is not None:
        ct_changed = False
        for child in list(ct_root):
            if etree.QName(child).localname == "Override":
                part_attr = (child.get("PartName") or "").lstrip("/")
                if part_attr and part_attr not in present_parts:
                    ct_root.remove(child)
                    ct_changed = True
        if ct_changed:
            pkg.mark_modified("[Content_Types].xml")

    if "word/comments.xml" in pkg.raw_parts or "word/comments.xml" in pkg.xml_parts:
        ensure_comment_styles_in_styles_xml(pkg)

    reassign_colliding_bookmark_ids(pkg)

    part_names = [
        name
        for name in list(pkg.raw_parts.keys()) + list(pkg.xml_parts.keys())
        if name.startswith("word/") and name.endswith(".xml")
    ]
    seen: set[str] = set()
    for name in part_names:
        if name in seen:
            continue
        seen.add(name)
        root = pkg.get_xml(name)
        if root is None:
            continue
        if normalize_ooxml_tree(root, unwrap_wps_ac=unwrap_wps_ac):
            pkg.mark_modified(name)


def sanitize_docx_file(docx_path: str | Path, unwrap_wps_ac: bool = True) -> None:
    """原地规范化并修复一个 .docx 文件（用于清理 LibreOffice .doc -> .docx 转换产物或 python-docx 输出）。"""
    p = Path(docx_path)
    if p.suffix.lower() != ".docx":
        raise ValueError(f"待规范化文件必须使用 .docx 扩展名: {p}")
    fd, tmp_name = tempfile.mkstemp(
        prefix=f".{p.name}.", suffix=".sanitizing.docx", dir=str(p.parent)
    )
    os.close(fd)
    tmp_path = Path(tmp_name)
    try:
        pkg = DocxPackage(p)
        pkg.save(tmp_path, unwrap_wps_ac=unwrap_wps_ac)
        os.replace(tmp_path, p)
    finally:
        tmp_path.unlink(missing_ok=True)
