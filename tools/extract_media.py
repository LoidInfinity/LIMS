#!/usr/bin/env python3
"""Extract the pictures used by selected slides so a page can be reproduced.

Composed content pages reference image files directly, so a contributor's
figures must be pulled out of the source deck first. Files are named
slide<NN>-shape<ID>.<ext> and a JSON manifest is printed for the compose spec.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path


def extract(presentation, slides, output):
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    source = Path(presentation).expanduser().resolve()
    if not source.is_file():
        raise ValueError(f"Missing presentation: {source}")
    deck = Presentation(str(source))
    target = Path(output).expanduser().resolve()
    target.mkdir(parents=True, exist_ok=True)
    records = []
    for number in sorted(set(slides)):
        if not 1 <= number <= len(deck.slides):
            raise ValueError(f"Slide {number} is out of range (1..{len(deck.slides)})")
        slide = deck.slides[number - 1]
        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            image = shape.image
            name = f"slide{number:02d}-shape{shape.shape_id}.{image.ext}"
            path = target / name
            data = image.blob
            path.write_bytes(data)
            records.append({"slide": number, "shape_id": shape.shape_id, "shape_name": shape.name,
                            "file": name, "width_px": image.size[0], "height_px": image.size[1],
                            "sha256": hashlib.sha256(data).hexdigest()})
    if not records:
        raise ValueError(f"No pictures found on slide(s) {sorted(set(slides))}")
    return target, records


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("presentation", type=Path, help="Source .pptx")
    parser.add_argument("--slide", type=int, action="append", required=True,
                        help="1-based slide number to extract from (repeatable)")
    parser.add_argument("--output", type=Path, required=True, help="Directory for extracted pictures")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()
    target, records = extract(args.presentation, args.slide, args.output)
    if args.json:
        print(json.dumps({"output": str(target), "pictures": records}, ensure_ascii=False, indent=2))
        return
    print(f"{target}  ({len(records)} picture(s))")
    for record in records:
        print(f"  slide {record['slide']} shape {record['shape_id']:>3}  {record['file']}  "
              f"{record['width_px']}x{record['height_px']}")


if __name__ == "__main__":
    main()
