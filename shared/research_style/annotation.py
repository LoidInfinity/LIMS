"""Semantic annotations for text boxes and runs, plus per-purpose style catalogs.

Every text box gets an annotation describing what it is for (heading, body,
citation, ...). The catalog maps an annotation to one canonical format, so two
citation boxes from different pages end up with the same face, size and colour
even when the source deck was inconsistent. Run-level annotations cover the
narrower purposes inside a box: a citation's journal, year and volume, inline
emphasis, and super/subscripts.
"""
from __future__ import annotations

import re
import unicodedata
from collections import Counter

DEFAULT_TEXT_STYLES = {
    "display": {"font_pt": 28, "bold": True, "color": "#002060",
                "font_latin": "黑体", "font_cjk": "黑体", "font_cs": "Calibri"},
    "heading": {"font_pt": 16, "bold": True, "color": "#1F3864",
                "font_latin": "黑体", "font_cjk": "黑体", "font_cs": "Calibri"},
    "subheading": {"font_pt": 14, "bold": True, "color": "#1F3864",
                   "font_latin": "黑体", "font_cjk": "黑体", "font_cs": "Calibri"},
    "body": {"font_pt": 14, "color": "#000000",
             "font_latin": "黑体", "font_cjk": "黑体", "font_cs": "Calibri"},
    "body_large": {"font_pt": 18, "color": "#000000",
                   "font_latin": "Calibri", "font_cjk": "黑体", "font_cs": "Calibri"},
    "bullet": {"font_pt": 14, "color": "#000000",
               "font_latin": "黑体", "font_cjk": "黑体", "font_cs": "Calibri"},
    "caption": {"font_pt": 10.5, "italic": True, "color": "#404040", "align": "center",
                "font_latin": "Calibri", "font_cjk": "宋体", "font_cs": "Calibri"},
    "citation": {"font_pt": 12, "color": "#000000", "align": "center",
                 "font_latin": "Calibri", "font_cjk": "宋体", "font_cs": "Calibri",
                 "parts": {"journal": {"italic": True}, "year": {"bold": True},
                           "volume": {"italic": True}, "pages": {}, "separator": {}}},
    "badge": {"font_pt": 20, "bold": True, "color": "#FFFFFF", "align": "center", "anchor": "middle",
              "font_latin": "黑体", "font_cjk": "黑体", "font_cs": "Calibri"},
    "data_label": {"font_pt": 18, "bold": True, "align": "center",
                   "font_latin": "Calibri", "font_cjk": "Calibri", "font_cs": "Calibri"},
    "note": {"font_pt": 10, "color": "#595959",
             "font_latin": "Calibri", "font_cjk": "宋体", "font_cs": "Calibri"},
    "label": {"font_pt": 9, "color": "#404040",
              "font_latin": "Calibri", "font_cjk": "宋体", "font_cs": "Calibri"},
    "page_number": {"font_pt": 10, "color": "#8C8C8C", "align": "right"},
    "footer": {"font_pt": 9, "color": "#8C8C8C"},
    "emphasis": {"bold": True, "color": "#C00000"},
    "superscript": {"baseline": 30000},
    "subscript": {"baseline": -25000},
}


_FONT_PATHS = {
    "黑体": ("simhei.ttf",),
    "宋体": ("simsun.ttc",),
    "微软雅黑": ("msyh.ttc",),
    "Calibri": ("calibri.ttf", "calibrib.ttf"),
    "Arial": ("arial.ttf", "arialbd.ttf"),
    "Times New Roman": ("times.ttf", "timesbd.ttf"),
    "DejaVu Sans": ("DejaVuSans.ttf",),
}
_FONT_CACHE = {}


def _font(name, size=1000):
    import os
    from pathlib import Path as _Path
    from PIL import ImageFont
    key = (name, size)
    if key in _FONT_CACHE:
        return _FONT_CACHE[key]
    base = _Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts"
    font = None
    for directory in (base, _Path("/usr/share/fonts/truetype/dejavu")):
        for filename in _FONT_PATHS.get(name, ()):
            candidate = directory / filename
            if candidate.is_file():
                try:
                    font = ImageFont.truetype(str(candidate), size)
                    break
                except OSError:
                    continue
        if font is not None:
            break
    if font is None:
        try:
            font = ImageFont.truetype("DejaVuSans.ttf", size)
        except OSError:
            font = ImageFont.load_default()
    _FONT_CACHE[key] = font
    return font


def measure_text(text, font_pt, latin="Calibri", cjk="黑体"):
    """Ink size of a text block in mm: the actual glyph extent, not the box."""
    lines = text.split("\n") if text else [""]
    width_pt = 0.0
    for line in lines:
        total = 0.0
        for char in line:
            font = _font(cjk) if unicodedata.east_asian_width(char) in "WF" else _font(latin)
            total += font.getlength(char)
        width_pt = max(width_pt, total)
    scale = font_pt / 1000
    line_pt = font_pt * 1.22
    return width_pt * scale * 25.4 / 72, len(lines) * line_pt * 25.4 / 72


def _wrap_by_measure(text, width_pt, font_pt, latin, cjk):
    scale = font_pt / 1000
    limit = width_pt / scale if scale else width_pt
    lines = []
    for paragraph in (text or "").split("\n"):
        current, used = "", 0.0
        for char in paragraph:
            font = _font(cjk) if unicodedata.east_asian_width(char) in "WF" else _font(latin)
            advance = font.getlength(char)
            if used + advance > limit and current:
                lines.append(current)
                current, used = char, advance
            else:
                current += char
                used += advance
        lines.append(current)
    return lines


def text_ink_rect(box, text, style, insets=(0.0, 0.0, 0.0, 0.0), line_spacing_pct=None,
                  wrap=True):
    """Where the glyphs actually sit inside a text box (alignment + anchor)."""
    font_pt = float(style.get("font_pt") or style.get("size_pt") or 14)
    latin = style.get("font_latin") or "Calibri"
    cjk = style.get("font_cjk") or "黑体"
    left, top, box_w, box_h = box
    inner_left = left + float(insets[0])
    inner_top = top + float(insets[1])
    inner_w = max(0.1, box_w - float(insets[0]) - float(insets[2]))
    inner_h = max(0.1, box_h - float(insets[1]) - float(insets[3]))
    spacing = 1.0 if not line_spacing_pct else float(line_spacing_pct) / 100
    if wrap:
        wrapped = _wrap_by_measure(text, inner_w * 72 / 25.4, font_pt, latin, cjk)
        scale = font_pt / 1000
        width_mm = max((sum((_font(cjk) if unicodedata.east_asian_width(char) in "WF" else _font(latin))
                            .getlength(char) for char in line) for line in wrapped), default=0.0)
        width_mm = width_mm * scale * 25.4 / 72
        height_mm = max(1, len(wrapped)) * font_pt * 1.22 * 25.4 / 72 * spacing
    else:
        width_mm, height_mm = measure_text(text, font_pt, latin, cjk)
        height_mm *= spacing
    spacing = 1.0 if not line_spacing_pct else float(line_spacing_pct) / 100
    height_mm *= spacing
    align = style.get("align") or "left"
    anchor = style.get("anchor") or "top"
    left, top, box_w, box_h = box
    if align == "center":
        ink_left = inner_left + (inner_w - width_mm) / 2
    elif align == "right":
        ink_left = inner_left + inner_w - width_mm
    else:
        ink_left = inner_left
    if anchor == "middle":
        ink_top = inner_top + (inner_h - height_mm) / 2
    elif anchor == "bottom":
        ink_top = inner_top + inner_h - height_mm
    else:
        ink_top = inner_top
    return [ink_left, ink_top, width_mm, height_mm]


DEFAULT_TEXT_COLORS = {"normal": "#000000", "weak": "#1F3864", "strong": "#C00000",
                       "accent": "#1F8DCF", "corresponding": None}
COLOR_ROLES = ("normal", "weak", "strong", "accent")


def _rgb(value):
    value = value.lstrip("#")
    return tuple(int(value[index:index + 2], 16) for index in (0, 2, 4))


def color_role_for(color, palette):
    """Map an observed colour onto a role; None means 'keep the raw colour'
    (e.g. an orange correspondence colour that has no role yet)."""
    if not color or not str(color).startswith("#"):
        return None
    best, best_distance = None, None
    for role in COLOR_ROLES:
        reference = (palette or {}).get(role)
        if not reference:
            continue
        distance = sum(abs(a - b) for a, b in zip(_rgb(color), _rgb(reference)))
        if best_distance is None or distance < best_distance:
            best, best_distance = role, distance
    if best is not None and best_distance is not None and best_distance <= 120:
        return best
    return None


def merge_text_colors(override=None, base=None):
    merged = dict(base or DEFAULT_TEXT_COLORS)
    merged.update(override or {})
    return merged


def derive_text_colors(presentation, skip_indices=(), palette=None):
    """Learn normal/weak/strong/accent colours from how the deck uses colour."""
    palette = merge_text_colors(palette)
    weak, strong, accent = Counter(), Counter(), Counter()
    for index, slide in enumerate(presentation.slides, 1):
        if index in skip_indices:
            continue
        title = slide.shapes.title.text.strip() if slide.shapes.title is not None else ""
        for shape in slide.shapes:
            if not shape.has_text_frame or shape.is_placeholder:
                continue
            text = shape.text_frame.text.strip()
            if not text:
                continue
            first_run = next((run for p in shape.text_frame.paragraphs
                              for run in p.runs if run.text.strip()), None)
            if first_run is None:
                continue
            color = _run_color(first_run, palette)
            if not color:
                continue
            red, green, blue = _rgb(color)
            if red > 150 and red > green + 60 and red > blue + 60:
                strong[color] += 1
                continue
            if title == "目录" and blue > red + 40:
                accent[color] += 1
                continue
            size = first_run.font.size.pt if first_run.font.size else 18
            if first_run.font.bold and size >= 13:
                weak[color] += 1
    result = dict(DEFAULT_TEXT_COLORS)
    if weak:
        result["weak"] = _mode(weak)
    if strong:
        result["strong"] = _mode(strong)
    if accent:
        result["accent"] = _mode(accent)
    return result


# Page numbers are not always <= 6 digits: article-style references use 7-8
# digit article numbers (e.g. Adv. Funct. Mater. 2024, 34, 2314419), and some
# publishers use e-locators (e2304567).
PAGE_TOKEN = r"(?:\d{1,8}|[eE]\d{4,10})"
CITATION_RE = re.compile(
    r"^(?P<journal>[A-Z][A-Za-z&.'\- ]{1,60}?)\s*(?P<year>(?:19|20)\d{2})\s*,\s*"
    r"(?P<volume>[0-9A-Za-z]{1,12}?)\s*,\s*(?P<pages>" + PAGE_TOKEN +
    r"(?:\s*[\u2013\-]\s*" + PAGE_TOKEN + r")?)\.?$")


def parse_citation(text):
    """Return {journal, year, volume, pages} for a single reference string."""
    stripped = (text or "").strip()
    match = CITATION_RE.match(stripped)
    if not match:
        return None
    journal = match.group("journal").strip()
    if len(journal) < 3:
        return None
    return {"journal": journal, "year": match.group("year"),
            "volume": match.group("volume"), "pages": match.group("pages")}


def classify(text, size_pt, bold, bulleted=False, italic=False):
    """Annotation for a whole text box, chosen by what the text is used for."""
    stripped = (text or "").strip()
    if not stripped:
        return None
    if parse_citation(stripped):
        return "citation"
    if (len(stripped) <= 16 and re.search(r"\d", stripped)
            and re.fullmatch(r"[\d.,\s%+\-\u2013/]*\s?[A-Za-z\u00b5\u03bc\u00b0%/]{0,8}", stripped)):
        return "data_label"
    if bold and size_pt >= 20:
        return "display"
    if bold and size_pt >= 15:
        return "heading"
    if bold and size_pt >= 12.5:
        return "subheading"
    if size_pt <= 10:
        return "label"
    if size_pt <= 11.5:
        return "note"
    if bulleted:
        return "bullet"
    if italic and size_pt <= 12.5:
        return "caption"
    if size_pt >= 15:
        return "body_large"
    return "body"


def run_annotation(color, base_color, baseline, bold, base_bold):
    """Annotation for a run narrower than its box: emphasis, super/subscript."""
    if baseline is not None:
        if int(baseline) > 0:
            return "superscript"
        if int(baseline) < 0:
            return "subscript"
    if color and base_color and str(color).upper() != str(base_color).upper():
        return "emphasis"
    if bold and not base_bold:
        return "emphasis"
    return None


def load_catalog(path):
    """Load text_styles from a template.json (or a styles JSON) plus defaults."""
    import json
    with open(path, encoding="utf-8-sig") as stream:
        data = json.load(stream)
    override = data.get("text_styles", data) if isinstance(data, dict) else {}
    return merge_styles(override)



def payload_text(payload):
    if isinstance(payload.get("paragraphs"), list):
        parts = []
        for paragraph in payload["paragraphs"]:
            parts.append("".join(str(run.get("text", "")) for run in paragraph.get("runs", [])))
        return "\n".join(parts)
    return str(payload.get("text") or "")


def payload_collision_rect(payload, box):
    """Collision border of an element: glyph ink for text, placement for images.

    A badge is one element: its background plate is the collision box, the label
    inside it is contained and therefore never counts as intersecting.
    """
    if payload.get("annotation") == "badge":
        return list(box)
    if "image" in payload:
        return list(box)
    style = {
        "font_pt": payload.get("font_pt") or 14,
        "font_latin": payload.get("font_latin") or "Calibri",
        "font_cjk": payload.get("font_cjk") or "黑体",
        "align": payload.get("align") or "left",
        "anchor": payload.get("anchor") or "top",
    }
    insets = payload.get("insets_mm") or [0.0, 0.0, 0.0, 0.0]
    return text_ink_rect(list(box), payload_text(payload), style, insets,
                         payload.get("line_spacing_pct"),
                         wrap=payload.get("word_wrap", True) is not False)


def merge_styles(override=None, base=None):
    merged = {name: dict(style) for name, style in (base or DEFAULT_TEXT_STYLES).items()}
    for name, style in (override or {}).items():
        current = dict(merged.get(name, {}))
        current.update({key: value for key, value in style.items() if key != "parts"})
        if "parts" in style:
            parts = dict(current.get("parts", {}))
            for part, part_style in style["parts"].items():
                parts[part] = {**parts.get(part, {}), **part_style}
            current["parts"] = parts
        merged[name] = current
    return merged


def style_for(catalog, annotation, explicit=None):
    style = dict((catalog or {}).get(annotation or "", {}))
    style.pop("parts", None)
    for key, value in (explicit or {}).items():
        if key not in ("annotation", "citation", "parts"):
            style[key] = value
    # An explicit source colour beats the purpose's colour role.
    if (explicit or {}).get("color"):
        style.pop("color_role", None)
    return style


def part_style(catalog, annotation, part):
    base = (catalog or {}).get(annotation or "", {})
    return dict(base.get("parts", {}).get(part, {}))


def citation_segments(citation):
    """Ordered (part, text) segments so the renderer can style each part."""
    segments = []
    if citation.get("journal"):
        segments.append(("journal", str(citation["journal"])))
    if citation.get("year"):
        segments.append(("separator", " "))
        segments.append(("year", str(citation["year"])))
    if citation.get("volume"):
        segments.append(("separator", ", "))
        segments.append(("volume", str(citation["volume"])))
    if citation.get("pages"):
        segments.append(("separator", ", "))
        segments.append(("pages", str(citation["pages"])))
    if segments:
        segments.append(("separator", "."))
    return segments


def _mode(counter):
    return counter.most_common(1)[0][0] if counter else None


def _run_faces(run, fonts):
    from pptx.oxml.ns import qn
    properties = run._r.find(qn("a:rPr"))
    faces = {}
    if properties is None:
        return faces
    for tag, key in (("latin", "font_latin"), ("ea", "font_cjk"), ("cs", "font_cs")):
        node = properties.find(qn("a:" + tag))
        if node is not None and node.get("typeface"):
            value = node.get("typeface")
            if value.startswith("+"):
                continue
            faces[key] = value
    return faces


THEME_SLOTS = {"dk1": "dk1", "lt1": "lt1", "dk2": "dk2", "lt2": "lt2", "tx1": "dk1",
               "bg1": "lt1", "tx2": "dk2", "bg2": "lt2", "accent1": "accent1",
               "accent2": "accent2", "accent3": "accent3", "accent4": "accent4",
               "accent5": "accent5", "accent6": "accent6"}


def _modify(base, node):
    import colorsys
    value = base.lstrip("#")
    red, green, blue = [int(value[index:index + 2], 16) / 255 for index in (0, 2, 4)]
    lum_mod = lum_off = None
    for child in node:
        tag = child.tag.rsplit("}", 1)[-1]
        try:
            amount = int(child.get("val")) / 100000
        except (TypeError, ValueError):
            continue
        if tag == "lumMod":
            lum_mod = amount
        elif tag == "lumOff":
            lum_off = amount
    if lum_mod is not None or lum_off is not None:
        hue, lightness, saturation = colorsys.rgb_to_hls(red, green, blue)
        lightness = lightness * (lum_mod if lum_mod is not None else 1.0) + (lum_off or 0.0)
        red, green, blue = colorsys.hls_to_rgb(hue, max(0.0, min(1.0, lightness)), saturation)
    return "#%02X%02X%02X" % tuple(round(channel * 255) for channel in (red, green, blue))


def _node_color(node, colors):
    tag = node.tag.rsplit("}", 1)[-1]
    if tag == "srgbClr":
        return _modify("#" + (node.get("val") or "").upper(), node)
    if tag == "schemeClr":
        slot = THEME_SLOTS.get(node.get("val", ""))
        base = (colors or {}).get(slot) if slot else None
        if base:
            return _modify(base, node)
    return None


def _shape_font_color(shape, colors=None):
    """Resolve the inherited text colour from the shape's p:style/a:fontRef."""
    from pptx.oxml.ns import qn
    style = shape._element.find(qn("p:style"))
    if style is None:
        return None
    reference = style.find(qn("a:fontRef"))
    if reference is None:
        return None
    for child in reference:
        color = _node_color(child, colors)
        if color:
            return color
    return None


def _run_color(run, colors=None):
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
        color = _node_color(child, colors)
        if color:
            return color
    return None


def derive_text_styles(presentation, colors=None, skip_indices=()):
    """Learn one canonical format per annotation from the deck's own text boxes."""
    sizes, families, colors, boldness, italics, alignments = {}, {}, {}, {}, {}, {}
    part_italic, part_bold = {}, {}
    for index, slide in enumerate(presentation.slides, 1):
        if index in skip_indices:
            continue
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            text = shape.text_frame.text.strip()
            if not text or shape.is_placeholder:
                continue
            paragraphs = [p for p in shape.text_frame.paragraphs if p.text.strip()]
            if not paragraphs:
                continue
            first_run = next((run for p in paragraphs for run in p.runs if run.text.strip()), None)
            size = first_run.font.size.pt if first_run is not None and first_run.font.size else 18.0
            bold = bool(first_run.font.bold) if first_run is not None else False
            bulleted = any(p._p.find(".//" "{http://schemas.openxmlformats.org/drawingml/2006/main}buChar")
                           is not None for p in paragraphs)
            annotation = classify(text, size, bold, bulleted,
                                   bool(first_run.font.italic) if first_run is not None else False)
            if not annotation:
                continue
            faces = _run_faces(first_run, {}) if first_run is not None else {}
            citations = parse_citation(text)
            box_color = None
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    if not run.text.strip():
                        continue
                    box_color = _run_color(run, colors)
                    if box_color:
                        break
                if box_color:
                    break
            if box_color is None:
                box_color = _shape_font_color(shape, colors)
            if box_color:
                colors.setdefault(annotation, Counter())[box_color] += 1
            if citations:
                plain_italic, plain_bold = Counter(), Counter()
                for part, segment in citation_segments(citations):
                    if part == "separator":
                        continue
                    start = text.find(segment)
                    if start < 0:
                        continue
                    position = 0
                    for paragraph in paragraphs:
                        for run in paragraph.runs:
                            span = range(position, position + len(run.text))
                            if start in span:
                                italic = bool(run.font.italic)
                                run_bold = bool(run.font.bold)
                                part_italic.setdefault(part, Counter())[italic] += 1
                                part_bold.setdefault(part, Counter())[run_bold] += 1
                                if part in ("pages", "separator"):
                                    plain_italic[italic] += 1
                                    plain_bold[run_bold] += 1
                            position += len(run.text)
                        position += 1
                if plain_italic:
                    italics.setdefault(annotation, Counter())[_mode(plain_italic)] += 1
                if plain_bold:
                    boldness.setdefault(annotation, Counter())[_mode(plain_bold)] += 1
            sizes.setdefault(annotation, Counter())[round(size, 1)] += 1
            for key, value in faces.items():
                families.setdefault(key, {}).setdefault(annotation, Counter())[value] += 1
            if not citations:
                boldness.setdefault(annotation, Counter())[bold] += 1
                if first_run is not None:
                    italics.setdefault(annotation, Counter())[bool(first_run.font.italic)] += 1
            alignment = next((p.alignment for p in paragraphs if p.alignment is not None), None)
            if alignment is not None:
                alignments.setdefault(annotation, Counter())[alignment.name.lower()] += 1
    catalog = {}
    for annotation, counter in sizes.items():
        style = {"font_pt": _mode(counter)}
        if families.get("font_latin", {}).get(annotation):
            style["font_latin"] = _mode(families["font_latin"][annotation])
        if families.get("font_cjk", {}).get(annotation):
            style["font_cjk"] = _mode(families["font_cjk"][annotation])
        if families.get("font_cs", {}).get(annotation):
            style["font_cs"] = _mode(families["font_cs"][annotation])
        if boldness.get(annotation) and _mode(boldness[annotation]):
            style["bold"] = True
        if italics.get(annotation) and _mode(italics[annotation]):
            style["italic"] = True

        if alignments.get(annotation):
            style["align"] = _mode(alignments[annotation])
        if annotation == "citation":
            parts = {}
            for part in ("journal", "year", "volume", "pages"):
                part_spec = {}
                if part_italic.get(part) and _mode(part_italic[part]):
                    part_spec["italic"] = True
                if part_bold.get(part) and _mode(part_bold[part]):
                    part_spec["bold"] = True
                if part_spec:
                    parts[part] = part_spec
            if parts:
                style["parts"] = parts
        catalog[annotation] = style
    return catalog
