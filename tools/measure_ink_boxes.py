"""Measure every text element on a white background and correct collision boxes.

PowerPoint is the only reliable source for line breaking and glyph metrics, so
the deck is rendered twice: once normally (for the visual pages) and once as an
isolated deck where each text shape sits alone on a white slide.  The measured
ink boxes replace the model-computed collision boxes in the manifest, and the
corrected manifest is re-checked for intersections.
"""
import argparse
import copy
import json
import subprocess
import sys
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.oxml.ns import qn

SLIDE_MM_FALLBACK = (338.667, 190.5)


def build_isolated(deck_path, output_path, mapping_path):
    deck = Presentation(str(deck_path))
    isolated = Presentation()
    isolated.slide_width, isolated.slide_height = deck.slide_width, deck.slide_height
    blank = isolated.slide_layouts[6]
    mapping = []
    for page_index, slide in enumerate(deck.slides, 1):
        for shape in slide.shapes:
            if not shape.has_text_frame or not shape.text_frame.text.strip():
                continue
            target = isolated.slides.add_slide(blank)
            for placeholder in list(target.shapes):
                placeholder._element.getparent().remove(placeholder._element)
            element = copy.deepcopy(shape._element)
            spPr = element.find(qn("p:spPr"))
            if spPr is not None:
                for tag in ("a:solidFill", "a:gradFill", "a:pattFill", "a:noFill",
                            "a:ln", "a:effectLst"):
                    for node in spPr.findall(qn(tag)):
                        spPr.remove(node)
                etree.SubElement(spPr, qn("a:noFill"))
                line = etree.SubElement(spPr, qn("a:ln"))
                etree.SubElement(line, qn("a:noFill"))
            target.shapes._spTree.append(element)
            mapping.append({"slide": len(mapping) + 1, "page": page_index,
                            "shape_id": int(shape.shape_id)})
    isolated.save(str(output_path))
    mapping_path.write_text(json.dumps(mapping, ensure_ascii=False, indent=1), encoding="utf-8")
    return mapping


def slide_path(directory, number):
    for candidate in (f"{number:02d}.png", f"{number:03d}.png", f"{number}.png"):
        path = directory / candidate
        if path.is_file():
            return path
    return None


def measure(directory, mapping, slide_mm):
    measured = {}
    for entry in mapping:
        path = slide_path(directory, entry["slide"])
        if path is None:
            continue
        image = Image.open(path).convert("L")
        scale = image.width / slide_mm[0]
        pixels = image.load()
        xs, ys = [], []
        for y in range(image.height):
            for x in range(image.width):
                if pixels[x, y] < 200:
                    xs.append(x)
                    ys.append(y)
        if xs:
            measured[(entry["page"], entry["shape_id"])] = [
                round(min(xs) / scale, 2), round(min(ys) / scale, 2),
                round((max(xs) + 1 - min(xs)) / scale, 2),
                round((max(ys) + 1 - min(ys)) / scale, 2)]
    return measured


def intersects(first, second, epsilon=0.05):
    ax, ay, aw, ah = first
    bx, by, bw, bh = second
    return (ax < bx + bw - epsilon and bx < ax + aw - epsilon
            and ay < by + bh - epsilon and by < ay + ah - epsilon)


def correct(manifest, measured):
    corrected = 0
    for page in manifest["pages"]:
        for box in page.get("collision_boxes", []):
            if box["kind"] in ("image", "badge"):
                continue
            key = (page["page"], box.get("shape_id"))
            if key in measured:
                box["rect_mm"] = measured[key]
                corrected += 1
    violations = []
    for page in manifest["pages"]:
        boxes = page.get("collision_boxes", [])
        for index, first in enumerate(boxes):
            if first["exempt"] or first["kind"] == "badge":
                continue
            for second in boxes[index + 1:]:
                if second["exempt"] or second["kind"] == "badge":
                    continue
                if intersects(first["rect_mm"], second["rect_mm"]):
                    violations.append((page["page"], first["label"][:20], second["label"][:20]))
    return corrected, violations


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--deck", required=True, type=Path)
    parser.add_argument("--manifest", required=True, type=Path)
    parser.add_argument("--renderer", type=Path, default=Path("tools/render_pptx.ps1"))
    parser.add_argument("--work-dir", type=Path, default=Path("test-workspace/render"))
    parser.add_argument("--slide-width-mm", type=float, default=SLIDE_MM_FALLBACK[0])
    parser.add_argument("--slide-height-mm", type=float, default=SLIDE_MM_FALLBACK[1])
    args = parser.parse_args()

    deck = Presentation(str(args.deck))
    slide_mm = (deck.slide_width / 36000, deck.slide_height / 36000)
    args.work_dir.mkdir(parents=True, exist_ok=True)
    isolated_deck = args.work_dir / "isolated.pptx"
    mapping_path = args.work_dir / "isolated-mapping.json"
    mapping = build_isolated(args.deck, isolated_deck, mapping_path)

    render_dir = args.work_dir / "isolated"
    if render_dir.exists():
        for stale in render_dir.glob("*.png"):
            stale.unlink()
    command = ["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass",
               "-File", str(args.renderer), "-Presentation", str(isolated_deck),
               "-OutputDirectory", str(render_dir)]
    subprocess.run(command, check=True)

    measured = measure(render_dir, mapping, slide_mm)
    manifest = json.loads(args.manifest.read_text(encoding="utf-8"))
    corrected, violations = correct(manifest, measured)
    args.manifest.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"measured {len(measured)} text shapes, corrected {corrected} collision boxes")
    if violations:
        print(f"WARNING: {len(violations)} intersection(s) remain on the measured boxes:")
        for page, first, second in violations[:20]:
            print(f"  page {page}: {first!r} x {second!r}")
    else:
        print("all measured collision boxes are disjoint")


if __name__ == "__main__":
    main()
