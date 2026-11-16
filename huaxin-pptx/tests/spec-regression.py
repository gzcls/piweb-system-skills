#!/usr/bin/env python3
"""Input preflight regression: use synthetic specs and never touch an existing deck."""
import copy
import importlib.util
import json
from pathlib import Path
import subprocess
import sys
import tempfile

sys.dont_write_bytecode = True
ROOT = Path(__file__).resolve().parents[1]
module_spec = importlib.util.spec_from_file_location("hx_spec_test", ROOT / "scripts/hx_pptx.py")
hx = importlib.util.module_from_spec(module_spec)
sys.modules[module_spec.name] = hx
module_spec.loader.exec_module(hx)


def content(layout):
    return {"complete": False, "slides": [{"title": "测试", "layout": layout}]}


def main():
    example = copy.deepcopy(hx.EXAMPLE_SPEC)
    assert hx.validate_spec(example)["status"] == "VALID"
    assert example == hx.EXAMPLE_SPEC
    chart = {"type": "chart", "kind": "doughnut", "categories": ["A", "B"],
             "series": [{"name": "示例", "values": [60, 40]}]}
    assert hx.validate_spec(content(chart))["status"] == "VALID"
    assert hx.validate_spec(content([chart]))["status"] == "VALID"
    assert hx.validate_spec({"slides": [{"title": "补画页"}]})["status"] == "VALID"
    invalid = [
        ([], "spec"), ({"slides": []}, "slides"), ({"slides": {}}, "slides"),
        ({"slides": [None]}, "slides[0]"), (content("chart"), "layout"),
        (content([None]), "layout[0]"), (content({"recipe": "kpis"}), "recipe"),
        (content({"type": "chart", "recipe": "chart"}), "recipe"),
        (content({"type": "doughnut"}), "kind"), (content({"type": []}), "layout.type"),
        (content({"type": "chart", "kind": []}), "kind"),
        (content({**chart, "series": [{"values": [1]}]}), "values"),
    ]
    for spec, fragment in invalid:
        for operation in (hx.validate_spec, hx.build_deck):
            try:
                operation(spec)
            except hx.SpecError as error:
                assert fragment in str(error), str(error)
            else:
                raise AssertionError(f"Accepted malformed spec: {spec!r}")
    with tempfile.TemporaryDirectory(prefix="hx-spec-test-") as directory:
        root = Path(directory)
        spec_path = root / "spec.json"
        deck = root / "existing.pptx"
        deck.write_bytes(b"existing deck must survive")
        script = ROOT / "scripts/hx_pptx.py"
        for value, expected in [(example, 0), (content({"recipe": "chart"}), 2)]:
            spec_path.write_text(json.dumps(value), encoding="utf-8")
            before = {p.name: p.read_bytes() for p in root.iterdir()}
            run = subprocess.run([sys.executable, "-B", str(script), "validate-spec", "--spec", str(spec_path)],
                                 capture_output=True, text=True, encoding="utf-8")
            assert run.returncode == expected, run.stderr
            assert json.loads(run.stdout)["status"] == ("VALID" if expected == 0 else "SPEC_ERROR")
            assert before == {p.name: p.read_bytes() for p in root.iterdir()}
        for bad_json in ('{"slides":', '{"slides": [NaN]}'):
            spec_path.write_text(bad_json, encoding="utf-8")
            for args in (["validate-spec"], ["build", "--output", str(deck), "--force"]):
                run = subprocess.run([sys.executable, "-B", str(script), *args, "--spec", str(spec_path)],
                                     capture_output=True, text=True, encoding="utf-8")
                assert run.returncode == 2 and json.loads(run.stdout)["status"] == "SPEC_ERROR", run.stderr
                assert deck.read_bytes() == b"existing deck must survive"
    print("Spec preflight regression checks PASSED")


if __name__ == "__main__":
    main()
