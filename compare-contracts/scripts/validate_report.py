#!/usr/bin/env python3
"""Validate a generated static contract comparison report."""

from __future__ import annotations

import argparse
import html
import json
import re
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import unquote, urlparse


DEFAULT_SECTIONS = {"summary", "entities", "matrix", "commercial", "risk", "portfolio", "method"}


class ReportParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.ids = []
        self.hrefs = []
        self.stylesheets = []
        self.source_citations = 0
        self.title = ""
        self._in_title = False

    def handle_starttag(self, tag, attrs):
        attributes = dict(attrs)
        if attributes.get("id"):
            self.ids.append(attributes["id"])
        if tag == "a" and attributes.get("href"):
            self.hrefs.append(attributes["href"])
        if tag == "link" and attributes.get("rel") == "stylesheet" and attributes.get("href"):
            self.stylesheets.append(attributes["href"])
        classes = (attributes.get("class") or "").split()
        if "source" in classes:
            self.source_citations += 1
        if tag == "title":
            self._in_title = True

    def handle_endtag(self, tag):
        if tag == "title":
            self._in_title = False

    def handle_data(self, data):
        if self._in_title:
            self.title += data


def local_target(report: Path, href: str) -> Path | None:
    parsed = urlparse(html.unescape(href))
    if parsed.scheme or href.startswith("#"):
        return None
    return (report.parent / unquote(parsed.path)).resolve()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("report", type=Path)
    parser.add_argument("--required-section", action="append", default=[])
    parser.add_argument("--min-citations", type=int, default=3)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()

    report = args.report.resolve()
    text = report.read_text(encoding="utf-8")
    parsed = ReportParser()
    parsed.feed(text)
    errors = []
    warnings = []

    if not parsed.title.strip():
        errors.append("missing document title")
    duplicate_ids = sorted({value for value in parsed.ids if parsed.ids.count(value) > 1})
    if duplicate_ids:
        errors.append(f"duplicate ids: {', '.join(duplicate_ids)}")
    required = set(args.required_section) or DEFAULT_SECTIONS
    missing_sections = sorted(required - set(parsed.ids))
    if missing_sections:
        errors.append(f"missing required sections: {', '.join(missing_sections)}")
    unresolved = sorted(set(re.findall(r"\{\{[^{}]+\}\}", text)))
    if unresolved:
        errors.append(f"unresolved template placeholders: {', '.join(unresolved[:12])}")
    if not parsed.stylesheets:
        errors.append("no stylesheet linked")

    local_links = []
    for href in [*parsed.stylesheets, *parsed.hrefs]:
        target = local_target(report, href)
        if target is None:
            continue
        local_links.append(str(target))
        if not target.exists():
            errors.append(f"missing local target: {href}")
    if parsed.source_citations < args.min_citations:
        warnings.append(
            f"only {parsed.source_citations} source citation elements; expected at least {args.min_citations}"
        )
    if "legal" not in text.lower() and "法律" not in text:
        warnings.append("no visible legal-review limitation detected")

    result = {
        "report": str(report),
        "title": parsed.title.strip(),
        "required_sections": sorted(required),
        "source_citations": parsed.source_citations,
        "local_link_count": len(local_links),
        "errors": errors,
        "warnings": warnings,
    }
    if args.output:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(result, ensure_ascii=False, indent=2))
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
