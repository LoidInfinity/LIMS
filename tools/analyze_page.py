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
import sys
import unicodedata
import zipfile
from xml.etree import ElementTree as ET

SHARED = Path(__file__).resolve().parents[1] / "shared"
if str(SHARED) not in sys.path:
    sys.path.insert(0, str(SHARED))
from research_style.annotation import (classify, color_role_for, load_catalog,
                                       merge_text_colors, parse_citation, run_annotation)
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


CITATION_RE = re.compile(r"(19|20)\d{2}")
ROLE_ALIGN = {"LEFT": "left", "CENTER": "center", "RIGHT": "right", "JUSTIFY": "justify"}


def paragraph_plain_text(paragraph_spec):
    return "".join(str(run.get("text", "")) for run in paragraph_spec.get("runs", []))


def payload_plain_text(payload):
    return "\n".join(paragraph_plain_text(paragraph) for paragraph in payload.get("paragraphs", []))


def run_payload(run, colors, fonts):
    from pptx.oxml.ns import qn
    record = {"text": run.text}
    if run.font.size:
        record["font_pt"] = round(run.font.size.pt, 1)
    if run.font.bold:
        record["bold"] = True
    if run.font.italic:
        record["italic"] = True
    color = run_color(run, colors)
    if color:
        record["color"] = color
    properties = run._r.find(qn("a:rPr"))
    if properties is not None:
        baseline = properties.get("baseline")
        if baseline is not None:
            try:
                record["baseline"] = int(baseline)
            except ValueError:
                pass
        for tag, key in (("latin", "font_latin"), ("ea", "font_cjk"), ("cs", "font_cs")):
            node = properties.find(qn("a:" + tag))
            value = resolve_typeface(node.get("typeface") if node is not None else None, fonts)
            if value:
                record[key] = value
    return record


def infer_role(paragraphs, size, bold):
    """Decide what a text block is used for, so formatting can follow its role."""
    text = " ".join("".join(str(run.get("text", "")) for run in paragraph.get("runs", []))
                    for paragraph in paragraphs).strip()
    lowered = text.lower()
    if (size and size <= 13 and len(text) <= 120 and CITATION_RE.search(text)
            and re.match(r"^[A-Z][A-Za-z.&'\- ]{2,}[ ,]", text)
            and not lowered.startswith(("http", "doi"))):
        return "citation"
    if bold and size and size >= 15:
        return "heading"
    if bold and size and size >= 12.5:
        return "subheading"
    if size and size <= 11:
        return "note"
    return "body"


def text_payload(shape, colors, fonts, default_pt=18, catalog=None, palette=None):
    from pptx.oxml.ns import qn
    frame = shape.text_frame
    paragraphs = list(frame.paragraphs)
    if not any(paragraph.text.strip() for paragraph in paragraphs):
        return None
    insets = [round(frame.margin_left / EMU_MM, 3), round(frame.margin_top / EMU_MM, 3),
              round(frame.margin_right / EMU_MM, 3), round(frame.margin_bottom / EMU_MM, 3)]
    alignment = next((p.alignment for p in paragraphs if p.alignment is not None), None)
    align = ROLE_ALIGN.get(alignment.name) if alignment is not None else None
    wrap_setting = frame.word_wrap
    anchor_name = frame.vertical_anchor
    anchor = {"TOP": "top", "MIDDLE": "middle", "BOTTOM": "bottom"}.get(anchor_name.name) if anchor_name else None
    if catalog:
        citation = parse_citation(frame.text)
        if citation:
            payload = {"annotation": "citation", "citation": citation, "insets_mm": insets,
                       "font_pt": float(catalog.get("citation", {}).get("font_pt", 12))}
            payload["allow_overflow"] = True
            if wrap_setting is False:
                payload["word_wrap"] = False
            if anchor:
                payload["anchor"] = anchor
            return payload
        first_run = next((run for paragraph in paragraphs
                          for run in paragraph.runs if run.text.strip()), None)
        size = first_run.font.size.pt if first_run is not None and first_run.font.size else default_pt
        bold = bool(first_run.font.bold) if first_run is not None else False
        bulleted = any(paragraph._p.find(qn("a:pPr")) is not None
                       and paragraph._p.find(qn("a:pPr")).find(qn("a:buChar")) is not None
                       for paragraph in paragraphs)
        annotation = classify(frame.text, size, bold, bulleted,
                               bool(first_run.font.italic) if first_run is not None else False)
        style = catalog.get(annotation, {})
        base_bold = bool(style.get("bold"))
        paragraph_colors = [run_color(run, colors) for paragraph in paragraphs
                            for run in paragraph.runs if run.text.strip()]
        item_color = next((value for value in paragraph_colors if value), None)
        spec_paragraphs = []
        for paragraph in paragraphs:
            runs = []
            for run in paragraph.runs:
                record = {"text": run.text}
                properties = run._r.find(qn("a:rPr"))
                baseline = properties.get("baseline") if properties is not None else None
                color = run_color(run, colors)
                if color and color != item_color:
                    role = color_role_for(color, palette)
                    if role:
                        record["color_role"] = role
                    else:
                        record["color"] = color
                run_name = run_annotation(color, item_color, baseline,
                                          bool(run.font.bold), base_bold)
                if run_name:
                    record["annotation"] = run_name
                runs.append(record)
            spec = {"runs": runs}
            bullet = bullet_of(paragraph, colors)
            if bullet:
                spec["bullet"] = bullet
            spec.update(paragraph_spacing(paragraph))
            spec_paragraphs.append(spec)
        payload = {"annotation": annotation, "paragraphs": spec_paragraphs,
                   "font_pt": float(style.get("font_pt", size)), "insets_mm": insets,
                   "allow_overflow": True}
        if wrap_setting is False:
            payload["word_wrap"] = False
        if item_color:
            item_role = color_role_for(item_color, palette)
            if item_role:
                payload["color_role"] = item_role
            else:
                payload["color"] = item_color
        if anchor:
            payload["anchor"] = anchor
        for key in ("bullet", "line_spacing_pct", "space_after_pt"):
            if key in spec_paragraphs[0] and len({str(p.get(key)) for p in spec_paragraphs}) == 1:
                payload[key] = spec_paragraphs[0][key]
                for paragraph in spec_paragraphs:
                    paragraph.pop(key, None)
        return payload
    raw = []
    for paragraph in paragraphs:
        entry = {"runs": [run_payload(run, colors, fonts) for run in paragraph.runs]}
        bullet = bullet_of(paragraph, colors)
        if bullet:
            entry["bullet"] = bullet
        if paragraph.alignment is not None:
            entry["align"] = ROLE_ALIGN.get(paragraph.alignment.name, "left")
        entry.update(paragraph_spacing(paragraph))
        raw.append(entry)
    first_run = next((run for paragraph in raw for run in paragraph["runs"] if run.get("text", "").strip()), None)
    defaults = {}
    if first_run is not None:
        for key in ("font_pt", "font_latin", "font_cjk", "font_cs"):
            if key in first_run:
                defaults[key] = first_run[key]
    size = defaults.get("font_pt", default_pt)
    bold = any(run.get("bold") for paragraph in raw for run in paragraph["runs"])
    for paragraph in raw:
        for run in paragraph["runs"]:
            for key, value in defaults.items():
                if run.get(key) == value:
                    del run[key]
            for key in ("bold", "italic"):
                if run.get(key) is False:
                    del run[key]
    payload = {"paragraphs": raw, "role": infer_role(raw, size, bold)}
    payload.update(defaults)
    for key in ("align", "bullet", "line_spacing_pct", "space_after_pt"):
        values = {str(paragraph.get(key)) for paragraph in raw}
        if len(values) == 1 and key in raw[0]:
            payload[key] = raw[0][key]
            for paragraph in raw:
                paragraph.pop(key, None)
    if align:
        payload["align"] = align
    if anchor:
        payload["anchor"] = anchor
    payload["insets_mm"] = insets
    return payload

def text_em(text):
    return sum(1.0 if unicodedata.east_asian_width(char) in "WF" else 0.5 for char in text)


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



def _box_mm(box, inner):
    return [inner[0] + box[0] * inner[2], inner[1] + box[1] * inner[3],
            box[2] * inner[2], box[3] * inner[3]]


def _boxes_overlap(first, second, tolerance=0.4):
    return (first[0] < second[0] + second[2] - tolerance
            and second[0] < first[0] + first[2] - tolerance
            and first[1] < second[1] + second[3] - tolerance
            and second[1] < first[1] + first[3] - tolerance)


def _resolve_collisions(entries, inner, images):
    """Nudge or shrink text that would visibly collide after size normalisation."""
    entries.sort(key=lambda entry: entry["payload"]["box"][1])
    placed = []
    for entry in entries:
        payload = entry["payload"]
        rect = _box_mm(payload["box"], inner)
        if payload.get("allow_overflow"):
            plain = payload_plain_text(payload)
            size = float(payload.get("font_pt") or 14)
            needed = text_height_mm(plain, rect[2] - 1.0, size)
            if needed > rect[3]:
                rect[3] = needed
        if payload.get("overlay"):
            placed.append(rect)
            continue
        obstacles = list(placed)
        for image in images:
            depth = min(rect[1] + rect[3], image[1] + image[3]) - max(rect[1], image[1])
            if depth > 0.3 and rect[1] < image[1] and depth < 7 and _boxes_overlap(rect, image):
                shift = depth + 0.6
                if rect[1] - shift >= inner[1]:
                    rect[1] -= shift
                    payload["box"] = [(rect[0] - inner[0]) / inner[2], (rect[1] - inner[1]) / inner[3],
                                      rect[2] / inner[2], rect[3] / inner[3]]
        attempts = [0.0]
        for step in range(1, 51):
            attempts.append(round(step * 0.5, 2))
        for step in range(1, 21):
            attempts.append(round(-step * 0.5, 2))
        for delta in attempts:
            candidate = [rect[0], rect[1] + delta, rect[2], rect[3]]
            if candidate[1] < inner[1] or candidate[1] + candidate[3] > inner[1] + inner[3]:
                continue
            if not any(_boxes_overlap(candidate, other) for other in obstacles):
                rect = candidate
                break
        else:
            size = float(payload.get("font_pt") or 14)
            for step in range(1, 5):
                smaller = max(9.0, size - step)
                plain = payload_plain_text(payload)
                needed = text_height_mm(plain, rect[2] - 1.0, smaller)
                candidate = [rect[0], rect[1], rect[2], min(rect[3], needed)]
                if not any(_boxes_overlap(candidate, other) for other in obstacles):
                    rect, size = candidate, smaller
                    payload["font_pt"] = smaller
                    break
            else:
                payload["overlay"] = True
        # Overflow height may exceed the module, but the stored box must stay
        # inside the inner area; the allow_overflow flag lets the renderer keep
        # the canonical font size and let glyphs spill instead.
        rect[0] = max(inner[0], min(rect[0], inner[0] + inner[2]))
        rect[1] = max(inner[1], min(rect[1], inner[1] + inner[3] - 0.5))
        rect[2] = max(0.5, min(rect[2], inner[0] + inner[2] - rect[0]))
        rect[3] = max(0.5, min(rect[3], inner[1] + inner[3] - rect[1]))
        if payload.get("allow_overflow"):
            pass
        payload["box"] = [(rect[0] - inner[0]) / inner[2], (rect[1] - inner[1]) / inner[3],
                          rect[2] / inner[2], rect[3] / inner[3]]
        placed.append(rect)
    # Final safety net: whatever still overlaps after nudging/shrinking is marked
    # as an intended overlay so the renderer keeps both items instead of failing.
    for first in range(len(entries)):
        for second in range(first + 1, len(entries)):
            a = entries[first]["payload"]
            b = entries[second]["payload"]
            if a.get("overlay") or b.get("overlay"):
                continue
            if _boxes_overlap(_box_mm(a["box"], inner), _box_mm(b["box"], inner), tolerance=0.0):
                b["overlay"] = True
    return entries


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
    """Keep the canonical size; expand the box and let the text overflow if needed.

    Shrinking per box is what made two headings of the same purpose end up at
    different sizes, so annotated text only shrinks when it is genuinely huge.
    """
    width = max(1.0, rect[2] - 1.0)
    plain = payload_plain_text(payload)
    target = max(1.0, rect[3] - 1.0)
    size = payload.get("font_pt", 18)
    if payload.get("annotation"):
        needed = text_height_mm(plain, width, size)
        if needed > target:
            grow = needed - target
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
    while size > 8 and text_height_mm(plain, width, size) > target:
        size -= 1
    payload["font_pt"] = size
    needed = text_height_mm(plain, width, size)
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


def analyse_slide(slide, number, colors, assets, asset_prefix, size_mm, padding=1.0, fonts=None,
                  catalog=None, palette=None):
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
            payload = text_payload(shape, colors, fonts or {}, catalog=catalog, palette=palette)
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
            in_inner = (module["_inner"][0] <= center[0] <= module["_inner"][0] + module["_inner"][2] and
                        module["_inner"][1] <= center[1] <= module["_inner"][1] + module["_inner"][3])
            in_frame = (module["rect"][0] <= center[0] <= module["rect"][0] + module["rect"][2] and
                        module["rect"][1] <= center[1] <= module["rect"][1] + module["rect"][3])
            # References sit right at the frame edge, so judge them by the frame
            # instead of the padded inner area; their box is clamped later.
            if in_inner or (item.get("annotation") == "citation" and in_frame):
                target = module
                break
        target["items"].append(item)
    result = []
    for module in [*modules, page_module]:
        inner = module.pop("_inner")
        images = [item["rect"] for item in module["items"] if "image" in item]
        boxes = []
        rendered = []
        for item in list(module["items"]):
            rect = clip(item["rect"], inner)
            if rect is None:
                continue
            if "text" in item or "paragraphs" in item or "citation" in item:
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
                for key in ("text", "paragraphs", "role", "annotation", "citation", "font_pt",
                            "allow_overflow",
                            "bold", "italic", "color", "color_role", "align", "anchor", "font_latin",
                            "font_cjk", "font_cs", "bullet", "line_spacing_pct", "space_after_pt",
                            "insets_mm", "word_wrap"):
                    if key in item:
                        payload[key] = item[key]
                payload.setdefault("font_pt", 18)
            for existing_payload, existing_box in boxes:
                if overlap(box, existing_box, tolerance=0.0005):
                    if "image" in existing_payload or "image" in payload:
                        payload["overlay"] = True
                    break
            boxes.append((payload, box))
            rendered.append(payload)
        module["items"] = _resolve_collisions(
            [{"payload": entry} for entry in rendered], inner, images)
        module["items"] = [entry["payload"] for entry in module["items"]]
        if module["items"] or module.get("fill"):
            result.append(module)
    return {"number": number, "title": title, "title_xml": title_xml,
            "layout_part": layout_part, "modules": result}


def build_spec(deck_path, slides, assets, asset_prefix, layout, template=None, report=None,
               catalog=None, palette=None):
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
                                 size_mm, fonts=fonts, catalog=catalog, palette=palette)
        page = {"layout": layout, "layout_part": analysis["layout_part"], "modules": analysis["modules"]}
        if analysis["title"].strip():
            page["title"] = analysis["title"]
            if analysis.get("title_xml"):
                page["title_xml"] = analysis["title_xml"]
        pages.append(page)
        raw.append(analysis)
    # Reproduction stays faithful to the deck: roles are recorded, but role
    # defaults must not add styling the source never set.
    spec = {"pages": pages, "base_dir": ".", "role_styles": False}
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
    parser.add_argument("--catalog", type=Path,
                        help="template.json whose text_styles normalise per-purpose formatting")
    args = parser.parse_args()
    assets = args.assets.expanduser().resolve()
    spec_parent = args.spec.expanduser().resolve().parent
    try:
        prefix = assets.relative_to(spec_parent).as_posix()
    except ValueError:
        prefix = str(assets)
    catalog = load_catalog(args.catalog) if args.catalog else None
    palette = None
    if args.catalog:
        import json as _json
        data = _json.loads(args.catalog.read_text(encoding="utf-8-sig"))
        palette = merge_text_colors(data.get("text_colors"))
    spec = build_spec(args.presentation.expanduser().resolve(), sorted(set(args.slide)),
                      assets, prefix, args.layout, args.template, args.report, catalog, palette)
    if catalog:
        spec["text_styles"] = catalog
        spec["text_colors"] = palette
    args.spec.write_text(json.dumps(spec, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"{args.spec}  ({len(spec['pages'])} page(s))")
    for page in spec["pages"]:
        modules = page.get("modules", [])
        items = sum(len(module["items"]) for module in modules)
        print(f"  title={page.get('title', '')!r:28} modules={len(modules)} items={items}")


if __name__ == "__main__":
    main()
