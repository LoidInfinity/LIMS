"""Ingest a PPTX into a slim, content-free template package plus descriptor.

The source deck is never copied into the library. Only the master/layout shells,
the theme and the media those parts reference survive; slide text, speaker notes
and slide-only pictures are dropped. Layout geometry is recorded so a builder can
place content without guessing.
"""
from __future__ import annotations

import argparse
import hashlib
import io
import math
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from . import __version__
from .core import SUITE_ID, atomic_bytes, write_json
from .template import DESCRIPTOR_NAME, now_iso, rebuild_catalog, slugify, validate_template

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"a": A, "p": P, "r": R}

PACKAGE_NAME = "template.pptx"
PREVIEW_NAME = "preview.png"

DROP_PREFIXES = ("ppt/slides/", "ppt/notesSlides/", "ppt/notesMasters/", "ppt/handoutMasters/",
                 "ppt/comments/", "ppt/commentAuthors.xml")
PRUNE_PREFIXES = ("ppt/media/", "ppt/embeddings/", "ppt/charts/", "ppt/diagrams/", "ppt/ink/")

LAYOUT_KINDS = {
    "title": "title", "titleOnly": "title-only", "secHead": "section", "blank": "blank",
    "obj": "content", "objOnly": "content", "tx": "content", "txAndObj": "content", "objAndTx": "content",
    "twoColTx": "two-content", "twoObj": "two-content", "twoObjAndTx": "two-content",
    "twoObjOverTx": "two-content", "twoTxTwoObj": "two-content", "objAndTwoObj": "two-content",
    "twoObjAndObj": "two-content", "txAndTwoObj": "two-content", "vertTitleAndTx": "two-content",
    "fourObj": "two-content", "objOverTx": "comparison", "txOverObj": "comparison",
    "txAndClipArt": "picture", "clipArtAndTx": "picture", "picTx": "picture",
    "txAndMedia": "picture", "mediaAndTx": "picture",
    "tbl": "table", "chart": "chart", "txAndChart": "chart", "chartAndTx": "chart",
}

ASPECTS = {"16:9": 16 / 9, "16:10": 16 / 10, "4:3": 4 / 3, "3:2": 3 / 2, "21:9": 21 / 9, "5:4": 5 / 4}


def read_parts(source):
    with zipfile.ZipFile(source) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _xml(parts, name):
    if name not in parts:
        raise ValueError(f"Missing package part: {name}")
    return ET.fromstring(parts[name])


def _rels(parts, part):
    directory, _, filename = part.rpartition("/")
    name = f"{directory}/_rels/{filename}.rels" if directory else f"_rels/{filename}.rels"
    if name not in parts:
        return {}
    result = {}
    for rel in ET.fromstring(parts[name]):
        if rel.get("TargetMode") == "External":
            continue
        result[rel.get("Id")] = (rel.get("Type", ""), _resolve(directory, rel.get("Target", "")))
    return result


def _resolve(base, target):
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(base, target)) if base else posixpath.normpath(target)


def _mm(emu):
    return round(int(emu) / 36000, 3)


def aspect_of(cx, cy):
    if not cx or not cy:
        raise ValueError("Slide size must be positive")
    ratio = cx / cy
    best = min(ASPECTS, key=lambda key: abs(ASPECTS[key] - ratio))
    if abs(ASPECTS[best] - ratio) <= 0.01:
        return best
    divisor = math.gcd(cx, cy)
    width, height = cx // divisor, cy // divisor
    if max(width, height) <= 50:
        return f"{width}:{height}"
    return f"{ratio:.3f}"


def _theme(parts, part):
    root = _xml(parts, part)
    colors, fonts = {}, {}
    scheme = root.find(".//a:clrScheme", NS)
    if scheme is not None:
        for node in scheme:
            slot = node.tag.rsplit("}", 1)[-1]
            color = node.find("a:srgbClr", NS)
            value = color.get("val") if color is not None else None
            if value is None:
                system = node.find("a:sysClr", NS)
                value = system.get("lastClr") if system is not None else None
            if value and re.fullmatch(r"[0-9A-Fa-f]{6}", value):
                colors[slot] = "#" + value.upper()
    for field, path in (("major_latin", ".//a:fontScheme/a:majorFont/a:latin"),
                        ("major_cjk", ".//a:fontScheme/a:majorFont/a:ea"),
                        ("minor_latin", ".//a:fontScheme/a:minorFont/a:latin"),
                        ("minor_cjk", ".//a:fontScheme/a:minorFont/a:ea")):
        node = root.find(path, NS)
        if node is not None and node.get("typeface"):
            fonts[field] = node.get("typeface")
    return colors, fonts


def _layout(parts, part, index):
    root = _xml(parts, part)
    container = root.find("p:cSld", NS)
    name = container.get("name") if container is not None else ""
    raw_type = root.get("type") or "cust"
    placeholders = []
    for shape in root.findall(".//p:sp", NS):
        properties = shape.find("p:nvSpPr/p:nvPr", NS)
        placeholder = properties.find("p:ph", NS) if properties is not None else None
        if placeholder is None:
            continue
        entry = {"type": placeholder.get("type") or "obj"}
        if placeholder.get("idx") is not None:
            entry["idx"] = int(placeholder.get("idx"))
        transform = shape.find("p:spPr/a:xfrm", NS)
        if transform is not None:
            offset, extent = transform.find("a:off", NS), transform.find("a:ext", NS)
            if offset is not None and extent is not None:
                entry.update(x_mm=_mm(offset.get("x")), y_mm=_mm(offset.get("y")),
                             w_mm=_mm(extent.get("cx")), h_mm=_mm(extent.get("cy")))
        placeholders.append(entry)
    return {"index": index, "part": part, "name": name or f"Layout {index + 1}",
            "type": LAYOUT_KINDS.get(raw_type, "other"), "raw_type": raw_type,
            "placeholders": placeholders}


def describe(parts, source_name, digest):
    presentation = _xml(parts, "ppt/presentation.xml")
    size = presentation.find("p:sldSz", NS)
    if size is None:
        raise ValueError("presentation.xml has no slide size")
    cx, cy = int(size.get("cx")), int(size.get("cy"))
    masters, layouts = [], []
    colors, fonts = {}, {}
    for _, (kind, target) in _rels(parts, "ppt/presentation.xml").items():
        if not kind.endswith("/slideMaster"):
            continue
        master = _xml(parts, target)
        container = master.find("p:cSld", NS)
        name = (container.get("name") if container is not None else "") or f"Master {len(masters) + 1}"
        record = {"name": name, "part": target}
        background = master.find("p:cSld/p:bg/p:bgPr/a:blipFill/a:blip", NS)
        if background is not None and background.get("{%s}embed" % R):
            media = _rels(parts, target).get(background.get("{%s}embed" % R))
            if media:
                record["background"] = media[1]
        masters.append(record)
        for _, (related, layout_part) in _rels(parts, target).items():
            if related.endswith("/slideLayout"):
                layouts.append(_layout(parts, layout_part, len(layouts)))
            elif related.endswith("/theme") and not colors:
                theme_colors, theme_fonts = _theme(parts, layout_part)
                colors.update(theme_colors)
                fonts.update(theme_fonts)
    if not layouts:
        raise ValueError("The deck declares no slide layouts")
    warnings = ["Slide content and speaker notes are dropped; only masters, layouts and referenced media are kept."]
    if any(not layout["placeholders"] for layout in layouts):
        warnings.append("Some layouts declare no explicit placeholder geometry; positions inherited from the master are not resolved.")
    return {
        "schema_version": 1,
        "id": "",
        "name": "",
        "school": {"name": "", "slug": ""},
        "tags": [],
        "aspect": aspect_of(cx, cy),
        "size": {"width_mm": _mm(cx), "height_mm": _mm(cy)},
        "theme": {"colors": colors or {"light1": "#FFFFFF"}, "fonts": fonts or {"minor_latin": "Calibri"}},
        "masters": masters,
        "layouts": layouts,
        "files": {"package": PACKAGE_NAME, "preview": PREVIEW_NAME, "preview_kind": "geometry"},
        "provenance": {
            "source_filename": source_name,
            "source_sha256": digest,
            "contributed_by": "unattributed",
            "license": "unspecified",
            "ingested_at": now_iso(),
            "tool": f"{SUITE_ID} {__version__}",
        },
        "warnings": warnings,
    }


def _drop_element(data, tag):
    pattern = re.compile(rb"<" + tag + rb"\b[^>]*/>|<" + tag + rb"\b[^>]*>.*?</" + tag + rb"\s*>", re.S)
    return pattern.sub(b"", data)


def _prune_relationships(data, drop_suffixes):
    root = ET.fromstring(data)
    for rel in list(root):
        if rel.get("Type", "").endswith(drop_suffixes):
            root.remove(rel)
    return _serialize(root, REL, "")


def _prune_content_types(data):
    root = ET.fromstring(data)
    for override in list(root):
        part = override.get("PartName", "")
        if part.startswith(("/ppt/slides/", "/ppt/notesSlides/", "/ppt/notesMasters/",
                            "/ppt/handoutMasters/", "/ppt/comments/", "/ppt/commentAuthors.xml",
                            "/docProps/thumbnail.")) or part in ("/ppt/commentAuthors.xml",):
            root.remove(override)
    return _serialize(root, CT, "")


def _serialize(root, namespace, prefix):
    ET.register_namespace(prefix, namespace)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _referenced(parts):
    targets = set()
    for name in parts:
        if not name.endswith(".rels"):
            continue
        directory = posixpath.dirname(posixpath.dirname(name))
        for rel in ET.fromstring(parts[name]):
            if rel.get("TargetMode") == "External":
                continue
            targets.add(_resolve(directory, rel.get("Target", "")))
    return targets


def slim(parts):
    kept = {name: data for name, data in parts.items()
            if not name.startswith(DROP_PREFIXES) and not name.startswith("docProps/thumbnail.")}
    presentation = kept["ppt/presentation.xml"]
    for tag in (b"p:sldIdLst", b"p:notesMasterIdLst", b"p:handoutMasterIdLst",
                b"p:custShowLst", b"p14:sectionLst"):
        presentation = _drop_element(presentation, tag)
    kept["ppt/presentation.xml"] = presentation
    kept["ppt/_rels/presentation.xml.rels"] = _prune_relationships(
        kept["ppt/_rels/presentation.xml.rels"],
        ("/slide", "/notesSlide", "/notesMaster", "/handoutMaster", "/commentAuthors", "/slideComments"))
    kept["_rels/.rels"] = _prune_relationships(kept["_rels/.rels"], ("/thumbnail",))
    kept["[Content_Types].xml"] = _prune_content_types(kept["[Content_Types].xml"])
    referenced = _referenced(kept)
    for name in [name for name in kept if name.startswith(PRUNE_PREFIXES) and name not in referenced]:
        del kept[name]
    return kept


def _zip_bytes(parts):
    buffer = io.BytesIO()
    order = ["[Content_Types].xml", *[name for name in parts if name != "[Content_Types].xml"]]
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name in order:
            archive.writestr(name, parts[name])
    return buffer.getvalue()


def validate_slim(data, expected_layouts):
    try:
        from pptx import Presentation
    except ImportError as exc:
        raise ValueError("Missing python-pptx; install requirements.txt in the active Python environment") from exc
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        if any(name.startswith("ppt/slides/") for name in archive.namelist()):
            raise ValueError("Slim package still contains slide parts")
    presentation = Presentation(io.BytesIO(data))
    layouts = [layout for master in presentation.slide_masters for layout in master.slide_layouts]
    if len(presentation.slides):
        raise ValueError("Slim package still opens with slides")
    if len(layouts) != expected_layouts:
        raise ValueError(f"Slim package has {len(layouts)} layouts, expected {expected_layouts}")


_FONT_CACHE = {}
_PLACEHOLDER_COLORS = {"title": (0x37, 0x4E, 0x7A), "body": (0xB6, 0xC4, 0xD8),
                       "picture": (0xD8, 0xC6, 0xA3), "table": (0xA9, 0xC4, 0xB2),
                       "other": (0xE2, 0xE5, 0xE9)}


def _font(size):
    if size in _FONT_CACHE:
        return _FONT_CACHE[size]
    from PIL import ImageFont
    font = None
    for candidate in ("C:/Windows/Fonts/msyh.ttc", "C:/Windows/Fonts/segoeui.ttf",
                      "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
                      "/System/Library/Fonts/PingFang.ttc"):
        try:
            font = ImageFont.truetype(candidate, size)
            break
        except OSError:
            continue
    _FONT_CACHE[size] = font or ImageFont.load_default()
    return _FONT_CACHE[size]


def _text(draw, xy, value, font, fill):
    try:
        draw.text(xy, value, font=font, fill=fill)
    except (UnicodeEncodeError, OSError):
        draw.text(xy, value.encode("ascii", "replace").decode("ascii"), font=font, fill=fill)


def _group(kind):
    if kind in ("title", "ctrTitle", "subTitle"):
        return "title"
    if kind in ("pic", "clipArt", "media", "dgm"):
        return "picture"
    if kind == "tbl":
        return "table"
    if kind in ("body", "obj", "sldNum", "ftr", "dt", "hdr"):
        return "other"
    return "other"


def _draw_layout(draw, layout, x, y, cell_w, cell_h, slide_w, slide_h):
    draw.rectangle([x, y, x + cell_w, y + cell_h], fill="#F3F4F6", outline="#D1D5DB")
    scale = min((cell_w - 22) / slide_w, (cell_h - 22) / slide_h)
    width, height = slide_w * scale, slide_h * scale
    left, top = x + (cell_w - width) / 2, y + (cell_h - height) / 2
    draw.rectangle([left, top, left + width, top + height], fill="#FFFFFF", outline="#9CA3AF")
    for placeholder in layout.get("placeholders", []):
        if "x_mm" not in placeholder:
            continue
        box = [left + placeholder["x_mm"] * scale, top + placeholder["y_mm"] * scale,
               left + (placeholder["x_mm"] + placeholder["w_mm"]) * scale,
               top + (placeholder["y_mm"] + placeholder["h_mm"]) * scale]
        draw.rectangle(box, fill=_PLACEHOLDER_COLORS[_group(placeholder["type"])], outline="#6B7280")
    label = f"{layout['index']}: {layout['name']}"
    _text(draw, (x + 4, y + cell_h + 6), label[:38], _font(12), (17, 24, 39))


def preview_png(descriptor, columns=4, cell_w=320, cell_h=200):
    from PIL import Image, ImageDraw
    layouts = descriptor["layouts"]
    columns = max(1, min(columns, len(layouts)))
    rows = math.ceil(len(layouts) / columns)
    pad, label_h, header_h = 18, 34, 60
    width = pad * 2 + columns * cell_w + (columns - 1) * pad
    height = header_h + pad + rows * (cell_h + label_h) + (rows - 1) * pad + pad
    image = Image.new("RGB", (width, height), "#FFFFFF")
    draw = ImageDraw.Draw(image)
    draw.rectangle([0, 0, width, header_h], fill="#1F2937")
    _text(draw, (pad, 12), f"{descriptor['school']['name']} · {descriptor['name']}", _font(19), (255, 255, 255))
    subtitle = (f"{descriptor['aspect']}   {descriptor['size']['width_mm']}×{descriptor['size']['height_mm']} mm"
                f"   {len(layouts)} layouts   theme preview")
    _text(draw, (pad, 38), subtitle, _font(13), (156, 163, 175))
    slide_w, slide_h = descriptor["size"]["width_mm"], descriptor["size"]["height_mm"]
    for position, layout in enumerate(layouts):
        column, row = position % columns, position // columns
        x = pad + column * (cell_w + pad)
        y = header_h + pad + row * (cell_h + label_h + pad)
        _draw_layout(draw, layout, x, y, cell_w, cell_h, slide_w, slide_h)
    buffer = io.BytesIO()
    image.save(buffer, format="PNG", optimize=True)
    return buffer.getvalue()


def _tags(values):
    result = []
    for value in values:
        text = value.strip()
        if text and text not in result:
            result.append(text[:32])
    if len(result) > 12:
        raise ValueError("At most 12 tags are allowed")
    return result


def ingest(source, root, school, school_slug=None, template_id=None, name=None, tags=(),
           contributor="unattributed", license_name="unspecified", force=False):
    source = Path(source)
    if source.suffix.lower() != ".pptx":
        raise ValueError("Only .pptx files can be ingested")
    if not source.is_file():
        raise ValueError(f"Missing source file: {source}")
    school = (school or "").strip()
    if not school:
        raise ValueError("--school is required so the library can group templates by institution")
    slug = (school_slug or slugify(school)).strip().lower()
    if not slug:
        raise ValueError("Cannot derive an ASCII slug from the school name; pass --school-slug (pinyin), e.g. --school-slug tsinghua")
    data = source.read_bytes()
    descriptor = describe(read_parts(source), source.name, hashlib.sha256(data).hexdigest())
    descriptor["school"] = {"name": school, "slug": slug}
    descriptor["id"] = (template_id or f"{slug}-{descriptor['aspect'].replace(':', 'x')}").strip().lower()
    descriptor["name"] = (name or f"{school} {descriptor['aspect']} 模板").strip()
    descriptor["tags"] = _tags(tags)
    descriptor["provenance"].update(contributed_by=contributor or "unattributed",
                                    license=license_name or "unspecified",
                                    source_mb=round(len(data) / 1048576, 2))
    package = _zip_bytes(slim(read_parts(source)))
    validate_slim(package, len(descriptor["layouts"]))
    validate_template(descriptor)
    target = Path(root) / slug / descriptor["id"]
    if target.exists() and not force:
        raise ValueError(f"Template already exists: {target}. Use --force to replace it")
    target.mkdir(parents=True, exist_ok=True)
    atomic_bytes(target / PACKAGE_NAME, package)
    atomic_bytes(target / PREVIEW_NAME, preview_png(descriptor))
    write_json(target / DESCRIPTOR_NAME, descriptor)
    rebuild_catalog(root)
    return descriptor, target, len(data), len(package)


def main(skill_dir):
    parser = argparse.ArgumentParser(prog="ingest_template.py",
                                     description="Add a PPTX to the template library as a content-free package.")
    parser.add_argument("source", type=Path, help="Source .pptx")
    parser.add_argument("--root", type=Path, help="Template library root (default: templates/ beside the skills)")
    parser.add_argument("--school", required=True, help="Institution display name, e.g. 清华大学")
    parser.add_argument("--school-slug", help="ASCII directory slug, e.g. tsinghua (required for non-ASCII names)")
    parser.add_argument("--id", dest="template_id", help="Template id, e.g. tsinghua-thesis-16x9")
    parser.add_argument("--name", help="Template display name")
    parser.add_argument("--tags", default="", help="Comma separated tags, e.g. 学术汇报,答辩")
    parser.add_argument("--contributor", default="unattributed", help="Who supplied the deck")
    parser.add_argument("--license", dest="license_name", default="unspecified",
                        help="Redistribution permission, e.g. 'authorized by the author'")
    parser.add_argument("--force", action="store_true", help="Replace an existing template of the same id")
    args = parser.parse_args()
    from .template import template_root
    root = template_root(skill_dir, args.root)
    descriptor, target, source_size, package_size = ingest(
        args.source, root, args.school, school_slug=args.school_slug, template_id=args.template_id,
        name=args.name, tags=[item for item in args.tags.split(",") if item.strip()],
        contributor=args.contributor, license_name=args.license_name, force=args.force)
    print(f"{target}")
    print(f"  id        {descriptor['id']}")
    print(f"  layouts   {len(descriptor['layouts'])}  aspect {descriptor['aspect']}")
    print(f"  package   {package_size / 1024:.1f} KB from {source_size / 1024:.1f} KB source")
    for warning in descriptor.get("warnings", []):
        print(f"Warning: {warning}", file=sys.stderr)
