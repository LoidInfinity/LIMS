#!/usr/bin/env python3
"""Turn source slides into compose page specs by reading their real geometry.

Frames (dashed/solid outlines), pictures and text are extracted with their
actual rectangles, font properties and z-order, then emitted as a compose spec
the renderer can rebuild. Nothing is guessed from shape names: the preset
geometry, corner adjustment and line style come from the XML.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import zipfile
from xml.etree import ElementTree as ET
import colorsys

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS = {"p": P, "a": A}
EMU_MM = 36000
DASH = {"ROUND_DOT": "round_dot", "SQUARE_DOT": "square_dot", "DASH": "dash",
        "DASH_DOT": "dash_dot", "LONG_DASH": "long_dash", "SOLID": "solid"}
THEME = {"DARK_1": "dk1", "LIGHT_1": "lt1", "DARK_2": "dk2", "LIGHT_2": "lt2",
         "TEXT_1": "dk1", "BACKGROUND_1": "lt1", "TEXT_2": "dk2", "BACKGROUND_2": "lt2",
         "ACCENT_1": "accent1", "ACCENT_2": "accent2", "ACCENT_3": "accent3",
         "ACCENT_4": "accent4", "ACCENT_5": "accent5", "ACCENT_6": "accent6",
         "HYPERLINK": "hlink", "FOLLOWED_HYPERLINK": "folHlink"}
CONTENT_AREA = (6.0, 24.0, 326.667, 162.0)


def theme_colors(path):
    colors = {}
    with zipfile.ZipFile(path) as archive:
        names = [name for name in archive.namelist() if re.fullmatch(r"ppt/theme/theme\d+\.xml", name)]
        if not names:
            return colors
        root = ET.fromstring(archive.read(sorted(names)[0]))
    scheme = root.find(".//a:clrScheme", NS)
    if scheme is None:
        return colors
    for node in scheme:
        slot = node.tag.rsplit("}", 1)[-1]
        value = None
        srgb = node.find("a:srgbClr", NS)
        if srgb is not None:
            value = srgb.get("val")
        if value is None:
            system = node.find("a:sysClr", NS)
            if system is not None:
                value = system.get("lastClr")
        if value:
            colors[slot] = "#" + value.upper()
    return colors


def mm(value):
    return round(int(value) / EMU_MM, 3)


def run_color(run, colors):
    from pptx.oxml.ns import qn
    if run is None:
        return None
    properties = run._r.find(qn("a:rPr"))
    if properties is None:
        return None
    solid = properties.find(qn("a:solidFill"))
    if solid is None:
        return None
    for child in solid:
        color, _ = resolve_color(child, colors)
        if color:
            return color
    return None


def shape_rect(shape):
    if shape.left is None or shape.top is None or shape.width is None or shape.height is None:
        return None
    return [mm(shape.left), mm(shape.top), mm(shape.width), mm(shape.height)]


def group_transform(shape):
    group = shape._element.find("p:grpSpPr", NS)
    if group is None:
        return None
    xfrm = group.find("a:xfrm", NS)
    if xfrm is None:
        return None
    off, ext = xfrm.find("a:off", NS), xfrm.find("a:ext", NS)
    ch_off, ch_ext = xfrm.find("a:chOff", NS), xfrm.find("a:chExt", NS)
    if None in (off, ext, ch_off, ch_ext):
        return None
    ox, oy = int(off.get("x")), int(off.get("y"))
    ex, ey = int(ext.get("cx")), int(ext.get("cy"))
    cx, cy = int(ch_off.get("x")), int(ch_off.get("y"))
    cw, ch = int(ch_ext.get("cx")), int(ch_ext.get("cy"))
    if not cw or not ch:
        return None
    sx, sy = ex / cw, ey / ch

    def transform(x, y, w, h):
        return [ox / EMU_MM + (x - cx / EMU_MM) * sx, oy / EMU_MM + (y - cy / EMU_MM) * sy, w * sx, h * sy]
    return transform


def flatten(shapes, transform=None):
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            inner = group_transform(shape)
            if inner is None:
                continue
            combined = inner if transform is None else (
                lambda x, y, w, h, outer=transform, nested=inner: outer(*nested(x, y, w, h)))
            yield from flatten(shape.shapes, combined)
        else:
            yield shape, transform





def _hex_rgb(value):
    value = value.lstrip("#")
    return tuple(int(value[index:index + 2], 16) for index in (0, 2, 4))


def resolve_color(node, colors):
    """Resolve an a:srgbClr / a:schemeClr element with lumMod/lumOff/alpha modifiers."""
    from pptx.oxml.ns import qn
    import colorsys
    if node is None:
        return None, None
    base = None
    if node.tag == qn("a:srgbClr"):
        base = "#" + (node.get("val") or "").upper()
    elif node.tag == qn("a:schemeClr"):
        base = colors.get(THEME.get(node.get("val", "").upper(), "")) or colors.get(node.get("val", ""))
    if not base or len(base) != 7:
        return None, None
    red, green, blue = [channel / 255 for channel in _hex_rgb(base)]
    lum_mod = lum_off = None
    alpha = None
    for child in node:
        tag = child.tag.rsplit("}", 1)[-1]
        if tag in ("lumMod", "lumOff", "alpha"):
            try:
                value = int(child.get("val")) / 100000
            except (TypeError, ValueError):
                continue
            if tag == "lumMod":
                lum_mod = value
            elif tag == "lumOff":
                lum_off = value
            else:
                alpha = value
    if lum_mod is not None or lum_off is not None:
        hue, lightness, saturation = colorsys.rgb_to_hls(red, green, blue)
        lightness = lightness * (lum_mod if lum_mod is not None else 1.0) + (lum_off or 0.0)
        red, green, blue = colorsys.hls_to_rgb(hue, max(0.0, min(1.0, lightness)), saturation)
    result = "#%02X%02X%02X" % tuple(round(channel * 255) for channel in (red, green, blue))
    return result, (round(alpha * 100, 1) if alpha is not None else None)


def fill_info(shape, colors):
    """Solid fill color plus alpha of an auto shape, read from the XML modifiers."""
    from pptx.oxml.ns import qn
    solid = shape._element.spPr.find(qn("a:solidFill"))
    if solid is None:
        return None
    for child in solid:
        color, alpha = resolve_color(child, colors)
        if color:
            return {"color": color, "alpha_pct": alpha} if alpha is not None else color
    return None


def container_of(shape, colors):
    """Border and/or fill of a text-free auto shape that can host content."""
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    if shape.shape_type != MSO_SHAPE_TYPE.AUTO_SHAPE:
        return None
    if shape.has_text_frame and shape.text_frame.text.strip():
        return None
    border = None
    try:
        width = shape.line.width
        if width is not None:
            dash = shape.line.dash_style
            line_color = shape.line.color
            border = {"width_pt": round(width.pt, 2),
                      "dash": "solid" if dash is None else DASH.get(dash.name, "dash"),
                      "color": ("#" + str(line_color.rgb)) if line_color.type == 1 else "#B8B8B7"}
    except (AttributeError, ValueError):
        border = None
    fill = fill_info(shape, colors)
    if border is None and fill is None:
        return None
    record = {"border": border, "radius_mm": corner_radius(shape)}
    if fill:
        record["fill"] = fill
    return record


def corner_radius(shape):
    geometry = shape._element.find(".//a:prstGeom", NS)
    if geometry is None or geometry.get("prst") != "roundRect":
        return 0.0
    guide = geometry.find("a:avLst/a:gd", NS)
    if guide is None:
        return 0.0
    match = re.fullmatch(r"val\s+(-?\d+)", guide.get("fmla", ""))
    if not match:
        return 0.0
    rect = shape_rect(shape)
    if rect is None:
        return 0.0
    return round(max(0.0, int(match.group(1)) / 100000) * min(rect[2], rect[3]), 2)


FONT_ALIASES = {"+mn-lt": "minor_latin", "+mj-lt": "major_latin",
                "+mn-ea": "minor_ea", "+mj-ea": "major_ea",
                "+mn-cs": "minor_cs", "+mj-cs": "major_cs"}


def theme_fonts(path):
    fonts = {}
    with zipfile.ZipFile(path) as archive:
        names = [name for name in archive.namelist() if re.fullmatch(r"ppt/theme/theme\d+\.xml", name)]
        if not names:
            return fonts
        root = ET.fromstring(archive.read(sorted(names)[0]))
    for kind in ("major", "minor"):
        for tag in ("latin", "ea", "cs"):
            node = root.find(".//a:fontScheme/a:%sFont/a:%s" % (kind, tag), NS)
            if node is not None and node.get("typeface"):
                fonts["%s_%s" % (kind, tag)] = node.get("typeface")
    return fonts


def resolve_typeface(value, fonts):
    if not value:
        return None
    if value.startswith("+"):
        return fonts.get(FONT_ALIASES.get(value, ""))
    return value


def bullet_of(paragraph, colors):
    from pptx.oxml.ns import qn
    properties = paragraph._p.find(qn("a:pPr"))
    if properties is None:
        return None
    char = properties.find(qn("a:buChar"))
    auto = properties.find(qn("a:buAutoNum"))
    if char is None and auto is None:
        return None
    record = {}
    if char is not None:
        record["char"] = char.get("char")
    else:
        record["auto_num"] = auto.get("type")
    font = properties.find(qn("a:buFont"))
    if font is not None and font.get("typeface"):
        record["font"] = font.get("typeface")
    size = properties.find(qn("a:buSzPct"))
    if size is not None:
        try:
            record["size_pct"] = round(int(size.get("val")) / 1000, 1)
        except (TypeError, ValueError):
            pass
    color = properties.find(qn("a:buClr"))
    if color is not None:
        for child in color:
            value, _ = resolve_color(child, colors)
            if value:
                record["color"] = value
            break
    for attribute, key in (("marL", "mar_l_mm"), ("indent", "indent_mm")):
        value = properties.get(attribute)
        if value is not None:
            try:
                record[key] = round(int(value) / EMU_MM, 3)
            except ValueError:
                pass
    return record


def paragraph_spacing(paragraph):
    from pptx.oxml.ns import qn
    properties = paragraph._p.find(qn("a:pPr"))
    if properties is None:
        return {}
    record = {}
    line = properties.find(qn("a:lnSpc"))
    if line is not None:
        percent = line.find(qn("a:spcPct"))
        if percent is not None:
            record["line_spacing_pct"] = round(int(percent.get("val")) / 1000, 1)
    after = properties.find(qn("a:spcAft"))
    if after is not None:
        points = after.find(qn("a:spcPts"))
        if points is not None:
            record["space_after_pt"] = round(int(points.get("val")) / 100, 1)
    return record


def text_payload(shape, colors, fonts, default_pt=18):
    from pptx.oxml.ns import qn
    paragraphs = shape.text_frame.paragraphs
    lines = [paragraph.text for paragraph in paragraphs]
    text = "\n".join(lines).rstrip("\n")
    if not text.strip():
        return None
    first_run = next((run for paragraph in paragraphs for run in paragraph.runs if run.text.strip()), None)
    size = first_run.font.size.pt if first_run is not None and first_run.font.size else default_pt
    payload = {"text": text, "font_pt": round(size, 1)}
    if first_run is not None:
        run_properties = first_run._r.find(qn("a:rPr"))
        if run_properties is not None:
            for tag, key in (("latin", "font_latin"), ("ea", "font_cjk"), ("cs", "font_cs")):
                node = run_properties.find(qn("a:" + tag))
                value = resolve_typeface(node.get("typeface") if node is not None else None, fonts)
                if value:
                    payload[key] = value
        if first_run.font.bold:
            payload["bold"] = True
        if first_run.font.italic:
            payload["italic"] = True
        color = run_color(first_run, colors)
        if color:
            payload["color"] = color
    alignment = next((p.alignment for p in paragraphs if p.alignment is not None), None)
    if alignment is not None:
        payload["align"] = {"LEFT": "left", "CENTER": "center", "RIGHT": "right",
                            "JUSTIFY": "justify"}.get(alignment.name, "left")
    anchor = shape.text_frame.vertical_anchor
    if anchor is not None:
        payload["anchor"] = {"TOP": "top", "MIDDLE": "middle", "BOTTOM": "bottom"}.get(anchor.name, "top")
    for paragraph in paragraphs:
        bullet = bullet_of(paragraph, colors)
        if bullet:
            payload["bullet"] = bullet
            break
    payload.update(paragraph_spacing(paragraphs[0]))
    return payload


def text_em(text):
    return sum(1.0 if ord(char) > 0x2E80 else 0.5 for char in text)


def text_height_mm(text, width_mm, font_pt, line_factor=1.35):
    per_line = max(1.0, width_mm / (font_pt * 25.4 / 72))
    lines = sum(max(1, int(-(-text_em(line) // per_line))) for line in text.split("\n"))
    return lines * font_pt * line_factor * 25.4 / 72


def inside(rect, outer, tolerance=0.3):
    return (rect[0] >= outer[0] - tolerance and rect[1] >= outer[1] - tolerance
            and rect[0] + rect[2] <= outer[0] + outer[2] + tolerance
            and rect[1] + rect[3] <= outer[1] + outer[3] + tolerance)


def overlap(first, second, tolerance=0.2):
    return (first[0] < second[0] + second[2] - tolerance and second[0] < first[0] + first[2] - tolerance
            and first[1] < second[1] + second[3] - tolerance and second[1] < first[1] + first[3] - tolerance)


def clip(item_rect, area):
    left = max(item_rect[0], area[0])
    top = max(item_rect[1], area[1])
    right = min(item_rect[0] + item_rect[2], area[0] + area[2])
    bottom = min(item_rect[1] + item_rect[3], area[1] + area[3])
    if right <= left or bottom <= top:
        return None
    original = item_rect[2] * item_rect[3]
    clipped = [left, top, right - left, bottom - top]
    return clipped if (clipped[2] * clipped[3]) / original >= 0.5 else None


def fit_text(payload, rect, inner):
    """Shrink the font until the text fits its original box, then expand the box if needed."""
    width = max(1.0, rect[2] - 1.2)
    size = payload.get("font_pt", 14)
    while size > 8 and text_height_mm(payload["text"], width, size) > rect[3]:
        size -= 1
    payload["font_pt"] = size
    needed = text_height_mm(payload["text"], width, size)
    if needed > rect[3]:
        grow = needed - rect[3]
        room_below = inner[1] + inner[3] - (rect[1] + rect[3])
        take = min(grow, max(0.0, room_below))
        rect[3] += take
        grow -= take
        if grow > 0:
            room_above = rect[1] - inner[1]
            take = min(grow, max(0.0, room_above))
            rect[1] -= take
            rect[3] += take
    return rect


def analyse_slide(slide, number, colors, assets, asset_prefix, size_mm, padding=1.0, fonts=None):
    from lxml import etree
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    width_mm, height_mm = size_mm
    title, title_xml, frames, items = "", None, [], []
    layout_part = str(slide.slide_layout.part.partname).lstrip("/")
    for shape, transform in flatten(slide.shapes):
        rect = shape_rect(shape)
        if rect is None:
            continue
        if transform is not None:
            rect = transform(*rect)
        if rect[0] + rect[2] <= 0 or rect[1] + rect[3] <= 0 or rect[0] >= width_mm or rect[1] >= height_mm:
            continue
        if shape.is_placeholder:
            if shape.placeholder_format.type in (1, 3):  # TITLE, CENTER_TITLE
                title = shape.text_frame.text
                title_xml = etree.tostring(shape._element, encoding="unicode")
            continue
        container = container_of(shape, colors)
        if container is not None and not (rect[2] > width_mm * 0.92 and rect[3] > height_mm * 0.92):
            frames.append({"rect": rect, **container})
            continue
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            image = shape.image
            name = f"slide{number:02d}-shape{shape.shape_id}.{image.ext}"
            target = Path(assets) / name
            target.parent.mkdir(parents=True, exist_ok=True)
            blob = image.blob
            if not target.is_file() or hashlib.sha256(target.read_bytes()).hexdigest() != hashlib.sha256(blob).hexdigest():
                target.write_bytes(blob)
            items.append({"rect": rect, "image": f"{asset_prefix}/{target.name}"})
            continue
        if shape.has_text_frame:
            payload = text_payload(shape, colors, fonts or {})
            if payload is not None:
                items.append({"rect": rect, **payload})
    modules = [{"rect": frame["rect"], "padding_mm": padding, "border": frame["border"],
                **({"radius_mm": frame["radius_mm"]} if frame.get("radius_mm") else {}),
                **({"fill": frame["fill"]} if frame.get("fill") else {}),
                "items": [], "_inner": [frame["rect"][0] + padding, frame["rect"][1] + padding,
                                        frame["rect"][2] - 2 * padding, frame["rect"][3] - 2 * padding]}
               for frame in frames]
    page_module = {"rect": list(CONTENT_AREA), "padding_mm": 0.0, "border": None,
                   "items": [], "_inner": list(CONTENT_AREA)}
    for item in items:
        center = [item["rect"][0] + item["rect"][2] / 2, item["rect"][1] + item["rect"][3] / 2]
        target = page_module
        for module in modules:
            if (inside(item["rect"], module["_inner"]) and
                    module["_inner"][0] <= center[0] <= module["_inner"][0] + module["_inner"][2] and
                    module["_inner"][1] <= center[1] <= module["_inner"][1] + module["_inner"][3]):
                target = module
                break
        target["items"].append(item)
    result = []
    for module in [*modules, page_module]:
        inner = module.pop("_inner")
        boxes = []
        rendered = []
        for item in list(module["items"]):
            rect = clip(item["rect"], inner)
            if rect is None:
                continue
            if "text" in item:
                rect = fit_text(item, list(rect), inner)
            box = [(rect[0] - inner[0]) / inner[2], (rect[1] - inner[1]) / inner[3],
                   rect[2] / inner[2], rect[3] / inner[3]]
            box = [round(max(0.0, min(1.0, value)), 5) for value in box]
            if box[2] <= 0.002 or box[3] <= 0.002:
                continue
            payload = {"box": box}
            if "image" in item:
                payload["image"] = item["image"]
            else:
                for key in ("text", "font_pt", "bold", "italic", "color", "align", "anchor",
                            "font_latin", "font_cjk", "font_cs", "bullet",
                            "line_spacing_pct", "space_after_pt"):
                    if key in item:
                        payload[key] = item[key]
                payload.setdefault("font_pt", 18)
            for _, existing_box in boxes:
                if overlap(box, existing_box, tolerance=0.0005):
                    payload["overlay"] = True
                    break
            boxes.append((payload, box))
            rendered.append(payload)
        module["items"] = rendered
        if module["items"] or module.get("fill"):
            result.append(module)
    return {"number": number, "title": title, "title_xml": title_xml,
            "layout_part": layout_part, "modules": result}


def build_spec(deck_path, slides, assets, asset_prefix, layout, template=None, report=None):
    from pptx import Presentation
    presentation = Presentation(str(deck_path))
    size_mm = (presentation.slide_width / EMU_MM, presentation.slide_height / EMU_MM)
    colors = theme_colors(deck_path)
    fonts = theme_fonts(deck_path)
    pages, raw = [], []
    for number in slides:
        if not 1 <= number <= len(presentation.slides):
            raise ValueError(f"Slide {number} is out of range")
        analysis = analyse_slide(presentation.slides[number - 1], number, colors, assets, asset_prefix,
                                 size_mm, fonts=fonts)
        page = {"layout": layout, "layout_part": analysis["layout_part"], "modules": analysis["modules"]}
        if analysis["title"].strip():
            page["title"] = analysis["title"]
            if analysis.get("title_xml"):
                page["title_xml"] = analysis["title_xml"]
        pages.append(page)
        raw.append(analysis)
    spec = {"pages": pages, "base_dir": "."}
    if template:
        spec["template"] = template
    if report:
        Path(report).write_text(json.dumps(raw, ensure_ascii=False, indent=2), encoding="utf-8")
    return spec


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("presentation", type=Path)
    parser.add_argument("--slide", type=int, action="append", required=True, help="1-based slide number (repeatable)")
    parser.add_argument("--assets", type=Path, required=True, help="Directory for extracted pictures")
    parser.add_argument("--spec", type=Path, required=True, help="Compose spec to write")
    parser.add_argument("--layout", type=int, required=True, help="Template layout index for these pages")
    parser.add_argument("--template", help="Template id to record in the spec")
    parser.add_argument("--report", type=Path, help="Optional raw analysis JSON")
    args = parser.parse_args()
    assets = args.assets.expanduser().resolve()
    spec_parent = args.spec.expanduser().resolve().parent
    try:
        prefix = assets.relative_to(spec_parent).as_posix()
    except ValueError:
        prefix = str(assets)
    spec = build_spec(args.presentation.expanduser().resolve(), sorted(set(args.slide)),
                      assets, prefix, args.layout, args.template, args.report)
    args.spec.write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"{args.spec}  ({len(spec['pages'])} page(s))")
    for page in spec["pages"]:
        modules = page.get("modules", [])
        items = sum(len(module["items"]) for module in modules)
        print(f"  title={page.get('title', '')!r:28} modules={len(modules)} items={items}")


if __name__ == "__main__":
    main()
