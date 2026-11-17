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

CITATION_RE = re.compile(
    r"^(?P<journal>[A-Z][A-Za-z&.'\- ]{1,60}?)\s*(?P<year>(?:19|20)\d{2})\s*,\s*"
    r"(?P<volume>[0-9A-Za-z]{1,12}?)\s*,\s*(?P<pages>\d{1,6}(?:\s*[\u2013\-]\s*\d{1,6})?)\.?$")


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
