#!/usr/bin/env python3
"""Build and QA a structured seven-day weather presentation in one command."""

from __future__ import annotations

import argparse
import re
import subprocess
import sys
import zipfile
from pathlib import Path


SKILL_ROOT = Path(__file__).resolve().parent.parent
RENDERER = SKILL_ROOT / "scripts" / "render_weather_report.js"
THUMBNAIL = SKILL_ROOT / "scripts" / "thumbnail.py"
PLACEHOLDER_PATTERN = re.compile(
    r"(?:xxxx|lorem|ipsum|this\s+(?:page|slide)\s+layout)",
    re.IGNORECASE,
)


def run(command: list[str], label: str) -> str:
    result = subprocess.run(command, capture_output=True, text=True)
    output = "\n".join(part.strip() for part in (result.stdout, result.stderr) if part.strip())
    if result.returncode != 0:
        raise RuntimeError(f"{label} failed ({result.returncode})\n{output}")
    return output


def validate_package(output: Path) -> int:
    if not output.is_file() or output.stat().st_size == 0:
        raise RuntimeError(f"PPTX was not created: {output}")
    with zipfile.ZipFile(output, "r") as archive:
        broken = archive.testzip()
        if broken:
            raise RuntimeError(f"PPTX ZIP integrity failed at {broken}")
        names = set(archive.namelist())
        required = {"[Content_Types].xml", "ppt/presentation.xml"}
        missing = required - names
        if missing:
            raise RuntimeError(f"PPTX is missing required parts: {sorted(missing)}")
        slides = [
            name for name in names
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
        ]
    if len(slides) != 7:
        raise RuntimeError(f"Expected 7 slides, found {len(slides)}")
    return len(slides)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Validate JSON, render a Huaxin weather PPTX, and create text/visual QA artifacts.",
    )
    parser.add_argument("input", help="Structured weather JSON")
    parser.add_argument("output", help="Output .pptx path")
    parser.add_argument("--qa-dir", default="pptx-qa", help="Directory for content.txt and overview.jpg")
    parser.add_argument("--no-visual", action="store_true", help="Skip LibreOffice thumbnail rendering")
    args = parser.parse_args()

    input_path = Path(args.input).resolve()
    output_path = Path(args.output).resolve()
    qa_dir = Path(args.qa_dir).resolve()
    if not input_path.is_file():
        raise RuntimeError(f"Input JSON does not exist: {input_path}")
    if output_path.suffix.lower() != ".pptx":
        raise RuntimeError("Output path must end with .pptx")

    output_path.parent.mkdir(parents=True, exist_ok=True)
    qa_dir.mkdir(parents=True, exist_ok=True)

    run(["node", str(RENDERER), str(input_path), str(output_path)], "weather renderer")
    slide_count = validate_package(output_path)

    content = run([sys.executable, "-m", "markitdown", str(output_path)], "content QA")
    if PLACEHOLDER_PATTERN.search(content):
        raise RuntimeError("Content QA found placeholder text")
    if len(content.strip()) < 300:
        raise RuntimeError("Content QA extracted unexpectedly little text")
    content_path = qa_dir / "content.txt"
    content_path.write_text(content + "\n", encoding="utf-8")

    overview_path = qa_dir / "overview.jpg"
    if not args.no_visual:
        overview_prefix = overview_path.with_suffix("")
        run(
            [sys.executable, str(THUMBNAIL), str(output_path), str(overview_prefix), "--cols", "3"],
            "visual QA render",
        )
        if not overview_path.is_file() or overview_path.stat().st_size == 0:
            raise RuntimeError(f"Visual QA overview was not created: {overview_path}")

    print(f"PPTX={output_path}")
    print(f"SLIDES={slide_count}")
    print(f"QA_TEXT={content_path}")
    if not args.no_visual:
        print(f"QA_OVERVIEW={overview_path}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(str(error), file=sys.stderr)
        raise SystemExit(1)
