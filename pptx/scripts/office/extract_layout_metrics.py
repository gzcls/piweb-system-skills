"""Extract alignment metrics and styles from a PPTX file and save them to a theme config.

Usage: python extract_layout_metrics.py <pptx_path> <output_json_path>
"""

import json
import sys
import zipfile
from xml.etree import ElementTree


def extract_metrics(pptx_path: str, output_path: str) -> None:
    namespaces = {
        "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
        "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    }

    config = {
        "theme": "Extracted_Template_Theme",
        "fonts": {"header": "Microsoft YaHei", "body": "Microsoft YaHei"},
        "slides": {},
    }

    with zipfile.ZipFile(pptx_path, "r") as z:
        # 1. Parse Slide 1 (Title Slide) to find Title block
        try:
            slide1_xml = z.read("ppt/slides/slide1.xml")
            root1 = ElementTree.fromstring(slide1_xml)

            # Find the largest text box or placeholder
            best_title_box = None
            best_area = 0

            for sp in root1.findall(".//p:sp", namespaces):
                texts = [t.text for t in sp.findall(".//a:t", namespaces) if t.text]
                if not texts:
                    continue

                xfrm = sp.find(".//a:xfrm", namespaces)
                if xfrm is not None:
                    off = xfrm.find("a:off", namespaces)
                    ext = xfrm.find("a:ext", namespaces)
                    if off is not None and ext is not None:
                        x = int(off.get("x")) / 914400
                        y = int(off.get("y")) / 914400
                        w = int(ext.get("cx")) / 914400
                        h = int(ext.get("cy")) / 914400
                        area = w * h
                        # Assume the title is the box with text and largest area
                        if area > best_area:
                            best_area = area
                            best_title_box = {
                                "x": round(x, 2),
                                "y": round(y, 2),
                                "w": round(w, 2),
                                "h": round(h, 2),
                                "fontSize": 36,
                                "bold": True,
                                "color": "1E293B",
                                "align": "left",
                            }

            if best_title_box:
                config["slides"]["title_slide"] = {
                    "bg_image": "assets/bg_title.jpeg",
                    "title": best_title_box,
                    "subtitle": {
                        "x": best_title_box["x"],
                        "y": round(best_title_box["y"] + best_title_box["h"] + 0.3, 2),
                        "w": best_title_box["w"],
                        "h": 0.8,
                        "fontSize": 16,
                        "color": "64748B",
                        "align": "left",
                    },
                }
            else:
                # Default fallback
                config["slides"]["title_slide"] = {
                    "bg_image": "assets/bg_title.jpeg",
                    "title": {
                        "x": 1.0,
                        "y": 1.8,
                        "w": 8.0,
                        "h": 1.5,
                        "fontSize": 36,
                        "bold": True,
                        "color": "1E293B",
                        "align": "left",
                    },
                    "subtitle": {
                        "x": 1.0,
                        "y": 3.5,
                        "w": 8.0,
                        "h": 0.8,
                        "fontSize": 16,
                        "color": "64748B",
                        "align": "left",
                    },
                }
        except Exception as e:
            print(f"Warning parsing slide1: {e}", file=sys.stderr)

        # 2. Parse SlideLayout 2 (Content Slide Layout) for placeholders
        try:
            layout2_xml = z.read("ppt/slideLayouts/slideLayout2.xml")
            root2 = ElementTree.fromstring(layout2_xml)

            title_box = None
            body_box = None

            for sp in root2.findall(".//p:sp", namespaces):
                ph = sp.find(".//p:ph", namespaces)
                ph_type = ph.get("type") if ph is not None else None

                xfrm = sp.find(".//a:xfrm", namespaces)
                if xfrm is not None:
                    off = xfrm.find("a:off", namespaces)
                    ext = xfrm.find("a:ext", namespaces)
                    if off is not None and ext is not None:
                        x = int(off.get("x")) / 914400
                        y = int(off.get("y")) / 914400
                        w = int(ext.get("cx")) / 914400
                        h = int(ext.get("cy")) / 914400

                        if ph_type == "title":
                            title_box = {
                                "x": round(x, 2),
                                "y": round(y, 2),
                                "w": round(w, 2),
                                "h": round(h, 2),
                                "fontSize": 26,
                                "bold": True,
                                "color": "0F172A",
                                "align": "left",
                            }
                        elif ph_type is None or ph_type == "body":
                            # Standard layout body box
                            if (
                                w > 4.0 and h > 2.0
                            ):  # Make sure it's a reasonably large content box
                                body_box = {
                                    "x": round(x, 2),
                                    "y": round(y, 2),
                                    "w": round(w, 2),
                                    "h": round(h, 2),
                                }

            # Set configurations based on extraction, or use fallbacks
            config["slides"]["content_slide"] = {
                "bg_image": "assets/bg_content.jpeg",
                "title": title_box
                or {
                    "x": 0.5,
                    "y": 0.3,
                    "w": 9.0,
                    "h": 0.8,
                    "fontSize": 26,
                    "bold": True,
                    "color": "0F172A",
                    "align": "left",
                },
                "content_area": body_box
                or {"x": 0.5, "y": 1.3, "w": 9.0, "h": 3.7},
            }
        except Exception as e:
            print(f"Warning parsing slideLayout2: {e}", file=sys.stderr)

        # 3. Setup Slide 3 (End Slide) based on slide3.xml or default
        try:
            slide3_xml = z.read("ppt/slides/slide3.xml")
            root3 = ElementTree.fromstring(slide3_xml)

            end_box = None
            for sp in root3.findall(".//p:sp", namespaces):
                texts = [t.text for t in sp.findall(".//a:t", namespaces) if t.text]
                if texts:
                    xfrm = sp.find(".//a:xfrm", namespaces)
                    if xfrm is not None:
                        off = xfrm.find("a:off", namespaces)
                        ext = xfrm.find("a:ext", namespaces)
                        if off is not None and ext is not None:
                            x = int(off.get("x")) / 914400
                            y = int(off.get("y")) / 914400
                            w = int(ext.get("cx")) / 914400
                            h = int(ext.get("cy")) / 914400
                            end_box = {
                                "x": round(x, 2),
                                "y": round(y, 2),
                                "w": round(w, 2),
                                "h": round(h, 2),
                                "fontSize": 36,
                                "bold": True,
                                "color": "1E293B",
                                "align": "center",
                            }
                            break

            config["slides"]["end_slide"] = {
                "bg_image": "assets/bg_title.jpeg",
                "title": end_box
                or {
                    "x": 1.0,
                    "y": 2.2,
                    "w": 8.0,
                    "h": 1.5,
                    "fontSize": 36,
                    "bold": True,
                    "color": "1E293B",
                    "align": "center",
                },
            }
        except Exception as e:
            print(f"Warning parsing slide3: {e}", file=sys.stderr)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

    print(f"Successfully extracted layout metrics to {output_path}")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(
            "Usage: python extract_layout_metrics.py <pptx_path> <output_json_path>",
            file=sys.stderr,
        )
        sys.exit(1)

    extract_metrics(sys.argv[1], sys.argv[2])
