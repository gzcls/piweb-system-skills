"""Extract background images from a PPTX file.

Usage: python extract_bg.py <pptx_path> <output_dir>

This script extracts images from ppt/media/ inside the PPTX file
and saves them to the output directory.
"""

import sys
import zipfile
from pathlib import Path


def extract_backgrounds(pptx_path: Path, output_dir: Path) -> None:
    if not pptx_path.exists():
        print(f"Error: {pptx_path} does not exist", file=sys.stderr)
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(pptx_path, "r") as z:
        media_files = [f for f in z.namelist() if f.startswith("ppt/media/")]
        if not media_files:
            print("No media files found in the PPTX", file=sys.stderr)
            return

        print(f"Found {len(media_files)} media files in the PPTX.")

        # Extract all media files
        for f in media_files:
            name = Path(f).name
            dest = output_dir / name
            dest.write_bytes(z.read(f))
            print(f"Extracted {f} to {dest}")

        # Specific alias mapping for the known 3-slide template:
        # slide1 / slide3 typically use the first image (often title slide background)
        # slide2 typically uses the second image (content slide background)
        image_list = sorted(
            [
                p
                for p in output_dir.glob("image*")
                if p.suffix.lower() in [".jpg", ".jpeg", ".png"]
            ],
            key=lambda x: x.name,
        )

        if len(image_list) >= 2:
            # Map first image as title background
            title_bg = output_dir / "bg_title.jpeg"
            content_bg = output_dir / "bg_content.jpeg"

            import shutil

            shutil.copy2(image_list[0], title_bg)
            shutil.copy2(image_list[1], content_bg)

            print(f"Mapped {image_list[0].name} -> bg_title.jpeg")
            print(f"Mapped {image_list[1].name} -> bg_content.jpeg")


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python extract_bg.py <pptx_path> <output_dir>", file=sys.stderr)
        sys.exit(1)

    pptx_path = Path(sys.argv[1])
    output_dir = Path(sys.argv[2])

    extract_backgrounds(pptx_path, output_dir)
