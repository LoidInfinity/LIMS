"""Compose a deck from a template: prototype pages plus layout-driven pages.

A prototype keeps the source slide's exact shape tree (design, pictures, text
formatting) and only rewrites text, which is what makes a cover or TOC page
reproduce pixel for pixel. Layout pages are new slides created from the
template's layouts and filled through their placeholders.
"""
from __future__ import annotations

import argparse
import copy
import io
import math
from pathlib import Path
import unicodedata

from .annotation import (citation_segments, measure_text, merge_styles, merge_text_colors,
                         part_style, payload_collision_rect, payload_text, style_for)
from .core import atomic_bytes, preflight_outputs, read_json, write_json
from .template import DESCRIPTOR_NAME, descriptor_sha256, load_descriptor, template_root

PROTOTYPE_KINDS = ("prototype", "layout")
DASH_STYLES = {"dash": "DASH", "round_dot": "ROUND_DOT", "square_dot": "SQUARE_DOT",
               "dash_dot": "DASH_DOT", "long_dash": "LONG_DASH", "solid": "SOLID"}
ALIGNMENTS = {"left": "LEFT", "center": "CENTER", "right": "RIGHT", "justify": "JUSTIFY"}


def compose_manifest_path(output):
    return Path(str(output) + ".compose.json")


def _find_shape(slide, key):
    text = str(key).strip()
    if text.isdigit():
        for shape in slide.shapes:
            if shape.shape_id == int(text):
                return shape
        raise ValueError(f"No shape with id {text} on this prototype")
    matches = [shape for shape in slide.shapes if shape.name == text]
    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise ValueError(f"No shape named {text!r} on this prototype")
    raise ValueError(f"Shape name {text!r} is ambiguous; address it by shape id: "
                     + ", ".join(str(shape.shape_id) for shape in matches))


def _set_shape_text(shape, text):
    """Replace text while keeping the first run's formatting and paragraph properties."""
    frame = shape.text_frame
    lines = text.split("\n")
    paragraphs = list(frame.paragraphs)
    while len(paragraphs) < len(lines):
        frame._txBody.append(copy.deepcopy(paragraphs[-1]._p))
        paragraphs = list(frame.paragraphs)
    from pptx.oxml.ns import qn
    for paragraph, line in zip(paragraphs, lines):
        # A slide-number/footer field keeps its own text; drop it so the new
        # literal does not get concatenated with the field value.
        for field in paragraph._p.findall(qn("a:fld")):
            paragraph._p.remove(field)
        runs = paragraph.runs
        if not runs:
            paragraph.add_run().text = line
            continue
        runs[0].text = line
        for run in runs[1:]:
            run._r.getparent().remove(run._r)
    for paragraph in paragraphs[len(lines):]:
        paragraph._p.getparent().remove(paragraph._p)


def _text_em(text):
    """Approximate text width in em: CJK glyphs are square, latin roughly half."""
    return sum(1.0 if unicodedata.east_asian_width(char) in "WF" else 0.5 for char in text)


def _text_width_mm(text, font_pt):
    return _text_em(text) * font_pt * 25.4 / 72


def _text_height_mm(text, width_mm, font_pt, line_factor=1.35, wrap=True):
    if not wrap:
        lines = max(1, len(text.split("\n")))
    else:
        per_line = max(1.0, width_mm / (font_pt * 25.4 / 72))
        lines = sum(max(1, math.ceil(_text_em(line) / per_line)) for line in text.split("\n"))
    return lines * font_pt * line_factor * 25.4 / 72


def _text_fills_box(shape):
    """True when the original text was sized to fill its box width.

    Such boxes were left-aligned but sized to the text, so the design reads as
    centered. Rewriting text in them needs explicit centering, otherwise shorter
    text drifts to the left edge.
    """
    if shape.is_placeholder or not shape.has_text_frame or not shape.width:
        return False
    frame = shape.text_frame
    text = frame.text.strip()
    if not text:
        return False
    run = next((run for paragraph in frame.paragraphs for run in paragraph.runs if run.font.size), None)
    size = run.font.size.pt if run else 18.0
    usable = (shape.width - frame.margin_left - frame.margin_right) / 36000
    if usable <= 0:
        return False
    return 0.72 <= _text_width_mm(text, size) / usable <= 1.3


def _clone_slide(presentation, source):
    """Duplicate a prototype slide, keeping shapes, text and picture relationships."""
    import copy as _copy
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    clone = presentation.slides.add_slide(source.slide_layout)
    for shape in list(clone.shapes):
        shape._element.getparent().remove(shape._element)
    remap = {}
    for relationship in source.part.rels.values():
        if relationship.reltype.endswith("/slideLayout"):
            continue
        if relationship.is_external:
            remap[relationship.rId] = clone.part.relate_to(relationship.target_ref, relationship.reltype,
                                                           is_external=True)
        else:
            remap[relationship.rId] = clone.part.relate_to(relationship.target_part, relationship.reltype)
    for element in source.shapes._spTree:
        tag = element.tag.rsplit("}", 1)[-1]
        if tag in ("nvGrpSpPr", "grpSpPr"):
            continue
        copied = _copy.deepcopy(element)
        for node in copied.iter():
            for attribute in list(node.attrib):
                value = node.get(attribute)
                if attribute.startswith("{") and value in remap and attribute.endswith(("embed", "id", "link")):
                    node.set(attribute, remap[value])
        clone.shapes._spTree.append(copied)
    return clone


def _apply_texts(slide, texts, prototype):
    if not isinstance(texts, dict):
        raise ValueError("Prototype page 'texts' must be an object of shape id/name to text")
    current = {slot["shape_id"]: slot["text"] for slot in prototype.get("text_slots", [])}
    changed = []
    for key, value in texts.items():
        overrides = {}
        if isinstance(value, dict):
            if not isinstance(value.get("text"), str):
                raise ValueError(f"Text for shape {key} must provide a string 'text'")
            new_text = value["text"]
            align = value.get("align")
            overrides = {name: value[name] for name in ("color", "bold", "italic", "font_pt")
                         if name in value}
        else:
            if not isinstance(value, str):
                raise ValueError(f"Text for shape {key} must be a string")
            new_text, align = value, None
        shape = _find_shape(slide, key)
        if current.get(shape.shape_id) == new_text and align is None and not overrides:
            continue
        text_changed = current.get(shape.shape_id) != new_text
        centering = align == "center" or (align is None and text_changed and _text_fills_box(shape))
        _set_shape_text(shape, new_text)
        if overrides:
            for paragraph in shape.text_frame.paragraphs:
                for run in paragraph.runs:
                    if "color" in overrides:
                        from pptx.dml.color import RGBColor
                        run.font.color.rgb = RGBColor.from_string(str(overrides["color"]).lstrip("#"))
                    if "bold" in overrides:
                        run.font.bold = bool(overrides["bold"])
                    if "italic" in overrides:
                        run.font.italic = bool(overrides["italic"])
                    if "font_pt" in overrides:
                        from pptx.util import Pt
                        run.font.size = Pt(float(overrides["font_pt"]))
        if centering or align:
            alignment = {"left": "LEFT", "center": "CENTER", "right": "RIGHT"}.get(align or "center")
            from pptx.enum.text import PP_ALIGN
            for paragraph in shape.text_frame.paragraphs:
                paragraph.alignment = getattr(PP_ALIGN, alignment)
        changed.append(shape.shape_id)
    return changed


def _all_layouts(prs):
    return [layout for master in prs.slide_masters for layout in master.slide_layouts]


def _layout(prs, key, descriptor):
    """Resolve a layout by descriptor index (disambiguates masters) or unique name."""
    layouts = _all_layouts(prs)
    by_part = {str(layout.part.partname).lstrip("/"): layout for layout in layouts}
    text = str(key).strip()
    if text.isdigit():
        index = int(text)
        recorded = [layout for layout in descriptor.get("layouts", []) if layout["index"] == index]
        if not recorded:
            raise ValueError(f"No layout with index {index} in the template descriptor")
        layout = by_part.get(recorded[0].get("part", ""))
        if layout is None:
            raise ValueError(f"Layout index {index} ({recorded[0]['name']!r}) is not in the template package")
        return layout
    matches = [layout for layout in layouts if layout.name == text]
    if len(matches) == 1:
        return matches[0]
    if not matches:
        raise ValueError(f"No layout named {text!r}; available: "
                         + ", ".join(f"{layout.name!r}" for layout in layouts))
    indexes = [recorded["index"] for recorded in descriptor.get("layouts", []) if recorded["name"] == text]
    raise ValueError(f"Layout name {text!r} exists on more than one master; pass one of the indexes: "
                     + ", ".join(str(index) for index in indexes))


def _replace_title_xml(slide, title_xml):
    """Swap the generated title placeholder for the template's captured title shape."""
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import qn
    placeholder = slide.shapes.title
    if placeholder is None:
        raise ValueError("Layout has no title placeholder to receive the captured title")
    element = parse_xml(title_xml.encode("utf-8"))
    tree = slide.shapes._spTree
    used = {int(node.get("id")) for node in tree.iter(qn("p:cNvPr")) if node.get("id")}
    identifier = element.find(".//" + qn("p:cNvPr"))
    if identifier is not None and int(identifier.get("id", "0")) in used:
        identifier.set("id", str(max(used | {1}) + 1))
    tree.replace(placeholder._element, element)
    return next(shape for shape in slide.shapes if shape._element is element)


def _module_rect(module, index):
    rect = module.get("rect")
    if not isinstance(rect, list) or len(rect) != 4 or not all(isinstance(value, (int, float)) for value in rect):
        raise ValueError(f"Module {index}: rect must be [x_mm, y_mm, w_mm, h_mm]")
    if rect[2] <= 0 or rect[3] <= 0:
        raise ValueError(f"Module {index}: rect needs a positive size")
    return [float(value) for value in rect]


def _item_box(item, module_index, item_index):
    box = item.get("box", [0.0, 0.0, 1.0, 1.0])
    if not isinstance(box, list) or len(box) != 4 or not all(isinstance(value, (int, float)) for value in box):
        raise ValueError(f"Module {module_index} item {item_index}: box must be [x, y, w, h] fractions")
    return [float(value) for value in box]


def _check_inside(box, module_index, item_index):
    x, y, width, height = box
    if x < -0.001 or y < -0.001 or x + width > 1.001 or y + height > 1.001:
        raise ValueError(f"Module {module_index} item {item_index}: box {box} leaves the module's inner area; "
                         "content must not cross the dashed border")


def _overlaps(first, second, tolerance=0.002):
    return (first[0] < second[0] + second[2] - tolerance and second[0] < first[0] + first[2] - tolerance
            and first[1] < second[1] + second[3] - tolerance and second[1] < first[1] + first[3] - tolerance)


def _set_run_style(run, item):
    from pptx.dml.color import RGBColor
    from pptx.util import Pt
    if "font_pt" in item:
        run.font.size = Pt(item["font_pt"])
    if item.get("bold"):
        run.font.bold = True
    if item.get("italic"):
        run.font.italic = True
    if item.get("color"):
        run.font.color.rgb = RGBColor.from_string(item["color"].lstrip("#"))


ROLE_PROFILES = {
    "title": {"font_pt": 40, "bold": True, "color": "#002060", "align": "center",
              "font_latin": "黑体", "font_cjk": "黑体", "font_cs": "Calibri"},
    "heading": {"font_pt": 16, "bold": True, "color": "#1F3864",
                "font_latin": "黑体", "font_cjk": "黑体", "font_cs": "Calibri"},
    "subheading": {"font_pt": 14, "bold": True, "color": "#1F3864",
                   "font_latin": "黑体", "font_cjk": "黑体", "font_cs": "Calibri"},
    "body": {"font_pt": 14, "color": "#000000",
             "font_latin": "黑体", "font_cjk": "黑体", "font_cs": "Calibri"},
    "caption": {"font_pt": 10.5, "color": "#404040", "italic": True, "align": "center",
                "font_latin": "Calibri", "font_cjk": "宋体", "font_cs": "Calibri"},
    "citation": {"font_pt": 10.5, "color": "#000000", "align": "center",
                 "font_latin": "Calibri", "font_cjk": "宋体", "font_cs": "Calibri"},
    "note": {"font_pt": 9, "color": "#595959",
             "font_latin": "Calibri", "font_cjk": "宋体", "font_cs": "Calibri"},
    "page_number": {"font_pt": 10, "color": "#8C8C8C", "align": "right"},
}
CITATION_KEYS = ("journal", "year", "volume", "pages")
BASELINE_KEYS = ("baseline", "font_pt", "bold", "italic", "color",
                 "font_latin", "font_cjk", "font_cs")


def effective_style(item, role_styles=True):
    """Role defaults first, then explicit item properties, then per-run overrides.

    Reproduction specs disable role styles: the source deck's explicit runs are
    authoritative, including the properties it does NOT set.
    """
    style = dict(ROLE_PROFILES.get(str(item.get("role") or ""), {})) if role_styles else {}
    for key in ("font_pt", "bold", "italic", "color", "align", "anchor",
                "font_latin", "font_cjk", "font_cs", "line_spacing_pct", "space_after_pt", "baseline"):
        if key in item:
            style[key] = item[key]
    return style


def citation_paragraph(citation, style):
    """Render a reference as journal-italic, year-bold, volume-italic + pages."""
    runs = []
    if citation.get("journal"):
        runs.append({"text": str(citation["journal"]), "italic": True})
    if citation.get("year"):
        runs.append({"text": " "})
        runs.append({"text": str(citation["year"]), "bold": True})
    if citation.get("volume"):
        runs.append({"text": ", "})
        runs.append({"text": str(citation["volume"]), "italic": True})
    if citation.get("pages"):
        runs.append({"text": ", "})
        runs.append({"text": str(citation["pages"])})
    if not runs:
        raise ValueError("citation needs journal/year/volume/pages")
    runs.append({"text": "."})
    return {"runs": runs, "align": style.get("align", "center")}


def _paragraph_runs(paragraph_spec):
    if isinstance(paragraph_spec.get("runs"), list) and paragraph_spec["runs"]:
        return paragraph_spec["runs"]
    return [{"text": str(paragraph_spec.get("text", ""))}]


def _paragraph_text(paragraph_spec):
    return "".join(str(run.get("text", "")) for run in _paragraph_runs(paragraph_spec))


def _text_items(item, style):
    """Normalize 'text', 'citation' and 'paragraphs' forms into a paragraph list."""
    if isinstance(item.get("paragraphs"), list) and item["paragraphs"]:
        return item["paragraphs"]
    if isinstance(item.get("citation"), dict):
        return [citation_paragraph(item["citation"], style)]
    text = item.get("text")
    if not isinstance(text, str) or not text.strip():
        raise ValueError("text item needs 'text', 'citation' or 'paragraphs'")
    paragraphs = []
    for line in text.split("\n"):
        paragraph = {"runs": [{"text": line}]}
        if isinstance(item.get("bullet"), dict):
            paragraph["bullet"] = item["bullet"]
        paragraphs.append(paragraph)
    return paragraphs


def _apply_fonts(run, item):
    """Write the latin (English/digits), CJK and complex-script typefaces."""
    from xml.sax.saxutils import escape
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls, qn
    latin = item.get("font_latin")
    cjk = item.get("font_cjk")
    cs = item.get("font_cs")
    if latin:
        run.font.name = latin
    properties = run._r.get_or_add_rPr()
    for tag, value in (("ea", cjk), ("cs", cs)):
        if not value:
            continue
        existing = properties.find(qn("a:" + tag))
        if existing is not None:
            properties.remove(existing)
        properties.append(parse_xml('<a:%s %s typeface="%s"/>' % (tag, nsdecls("a"), escape(str(value)))))


def _apply_baseline(run, value):
    """Superscript/subscript is a run property (normal font, raised/lowered baseline)."""
    try:
        offset = int(round(float(value)))
    except (TypeError, ValueError):
        raise ValueError(f"baseline must be a number like 30000 or -25000, got {value!r}") from None
    run._r.get_or_add_rPr().set("baseline", str(offset))


def _apply_bullet(paragraph, bullet):
    from xml.sax.saxutils import escape
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls
    properties = paragraph._p.get_or_add_pPr()
    for key, attribute in (("mar_l_mm", "marL"), ("indent_mm", "indent")):
        if bullet.get(key) is not None:
            properties.set(attribute, str(int(round(float(bullet[key]) * 36000))))
    fragments = []
    if bullet.get("color"):
        fragments.append('<a:buClr %s><a:srgbClr val="%s"/></a:buClr>'
                         % (nsdecls("a"), escape(str(bullet["color"]).lstrip("#"))))
    if bullet.get("size_pct"):
        fragments.append('<a:buSzPct %s val="%d"/>' % (nsdecls("a"), int(round(float(bullet["size_pct"]) * 1000))))
    if bullet.get("font"):
        fragments.append('<a:buFont %s typeface="%s"/>' % (nsdecls("a"), escape(str(bullet["font"]))))
    if bullet.get("char"):
        fragments.append('<a:buChar %s char="%s"/>' % (nsdecls("a"), escape(str(bullet["char"]))))
    for fragment in fragments:
        properties.append(parse_xml(fragment))


def _role_color(style, palette):
    """Replace a colour role with the template's concrete colour."""
    role = style.pop("color_role", None) if style else None
    if role and palette.get(role):
        style["color"] = palette[role]
    return style


BADGE_PAD_X_MM = 3.5
BADGE_PAD_Y_MM = 1.6


def _badge_rect(box, item, style):
    """Badges own a background: the plate grows to fit the canonical font."""
    from pptx.util import Mm
    text = "\n".join(_paragraph_text(paragraph) for paragraph in _text_items(item, style)).strip()
    size = float(style.get("font_pt") or 20.0)
    text_w = _text_em(text) * size * 25.4 / 72
    line_h = size * 1.35 * 25.4 / 72
    width = max(box[2] * 1.0, text_w + 2 * BADGE_PAD_X_MM)
    height = max(box[3] * 1.0, line_h + 2 * BADGE_PAD_Y_MM)
    center_x = box[0] + box[2] / 2
    center_y = box[1] + box[3] / 2
    return [center_x - width / 2, center_y - height / 2, width, height]


def _add_badge_background(slide, rect, badge):
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.oxml.ns import qn
    from pptx.util import Mm
    radius = float(badge.get("radius_mm") or 0)
    if radius > 0:
        shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Mm(rect[0]), Mm(rect[1]),
                                       Mm(rect[2]), Mm(rect[3]))
        shape.adjustments[0] = max(0.0, min(0.5, radius / min(rect[2], rect[3])))
    else:
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Mm(rect[0]), Mm(rect[1]),
                                       Mm(rect[2]), Mm(rect[3]))
    fill = badge.get("fill") or {"color": "#1F8DCF", "alpha_pct": 50}
    specification = fill if isinstance(fill, dict) else {"color": fill}
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor.from_string(str(specification["color"]).lstrip("#"))
    if specification.get("alpha_pct") is not None:
        solid = shape._element.spPr.find(qn("a:solidFill"))
        srgb = solid.find(qn("a:srgbClr")) if solid is not None else None
        if srgb is not None:
            srgb.append(srgb.makeelement(qn("a:alpha"),
                                         {"val": str(int(round(float(specification["alpha_pct"]) * 1000)))}))
    shape.line.fill.background()
    return shape


def _add_text_item(slide, left_mm, top_mm, box, item, module_index, item_index,
                   role_styles=True, catalog=None, bounds=None):
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Mm, Pt
    catalog = catalog or {}
    palette = catalog.get("_palette", {})
    explicit = {key: item[key] for key in (*BASELINE_KEYS, "align", "anchor",
                                           "line_spacing_pct", "space_after_pt") if key in item}
    citation = item.get("citation") if isinstance(item.get("citation"), dict) else None
    annotation = item.get("annotation")
    if citation:
        style = style_for(catalog, "citation", explicit)
        paragraphs = [{"runs": [{"text": segment_text, "annotation": part}
                                for part, segment_text in citation_segments(citation)]}]
    elif annotation and annotation in catalog:
        style = style_for(catalog, annotation, explicit)
        paragraphs = _text_items(item, style)
    else:
        style = effective_style(item, role_styles)
        paragraphs = _text_items(item, style)
    _role_color(style, palette)
    badge = item.get("badge") if isinstance(item.get("badge"), dict) else None
    if badge:
        rect = _badge_rect([left_mm, top_mm, box[2], box[3]], item, style)
        if bounds is not None:
            rect[0] = max(bounds[0], min(rect[0], bounds[0] + bounds[2] - rect[2]))
            rect[1] = max(bounds[1], min(rect[1], bounds[1] + bounds[3] - rect[3]))
            rect[2] = min(rect[2], bounds[2])
            rect[3] = min(rect[3], bounds[3])
        box = [rect[0], rect[1], rect[2], rect[3]]
        left_mm, top_mm = rect[0], rect[1]
        style.setdefault("align", "center")
        style.setdefault("anchor", "middle")
    width_mm, height_mm = box[2], box[3]
    size = style.get("font_pt", 16)
    wrap = item.get("word_wrap", True) is not False
    if citation:
        # References under figures are single-line labels; wrapping them would
        # break the tight caption band the source design relies on.
        wrap = False
    joined = "\n".join(_paragraph_text(paragraph) for paragraph in paragraphs)
    needed = _text_height_mm(joined, width_mm - 1.0, size, wrap=wrap)
    if needed > height_mm + 0.5 and not (item.get("allow_overflow") or badge):
        raise ValueError(f"Module {module_index} item {item_index}: text needs {needed:.1f} mm but its box is "
                         f"{height_mm:.1f} mm high; enlarge the box, shorten the text or lower font_pt")
    if badge:
        _add_badge_background(slide, [left_mm, top_mm, width_mm, height_mm], badge)
    shape = slide.shapes.add_textbox(Mm(left_mm), Mm(top_mm), Mm(width_mm), Mm(height_mm))
    frame = shape.text_frame
    frame.word_wrap = wrap
    insets = item.get("insets_mm")
    if isinstance(insets, list) and len(insets) == 4:
        frame.margin_left, frame.margin_top = Mm(float(insets[0])), Mm(float(insets[1]))
        frame.margin_right, frame.margin_bottom = Mm(float(insets[2])), Mm(float(insets[3]))
    else:
        frame.margin_left = frame.margin_right = Mm(2.54)
        frame.margin_top = frame.margin_bottom = Mm(1.27)
    frame.vertical_anchor = getattr(MSO_ANCHOR, (style.get("anchor") or "top").upper())
    item_bullet = item.get("bullet") if isinstance(item.get("bullet"), dict) else None
    for position, paragraph_spec in enumerate(paragraphs):
        paragraph = frame.paragraphs[0] if position == 0 else frame.add_paragraph()
        alignment = paragraph_spec.get("align") or style.get("align", "left")
        paragraph.alignment = getattr(PP_ALIGN, ALIGNMENTS.get(alignment, "LEFT"))
        line_spacing = paragraph_spec.get("line_spacing_pct", style.get("line_spacing_pct"))
        if line_spacing:
            paragraph.line_spacing = float(line_spacing) / 100
        space_after = paragraph_spec.get("space_after_pt", style.get("space_after_pt"))
        if space_after is not None:
            paragraph.space_after = Pt(float(space_after))
        bullet = paragraph_spec.get("bullet") or item_bullet
        if bullet:
            _apply_bullet(paragraph, bullet)
        for run_spec in _paragraph_runs(paragraph_spec):
            run = paragraph.add_run()
            run.text = str(run_spec.get("text", ""))
            merged = dict(style)
            run_annotation = run_spec.get("annotation")
            if run_annotation:
                if citation:
                    merged.update(part_style(catalog, "citation", run_annotation))
                else:
                    merged.update(catalog.get(run_annotation, {}))
            _role_color(merged, palette)
            for key in BASELINE_KEYS:
                if key in run_spec:
                    merged[key] = run_spec[key]
            _set_run_style(run, merged)
            _apply_fonts(run, merged)
            if merged.get("baseline") is not None:
                _apply_baseline(run, merged["baseline"])
    return shape

def _add_image_item(slide, left_mm, top_mm, box, item, module_index, item_index, base_dir):
    from PIL import Image
    from pptx.util import Mm
    path = item.get("image")
    if not isinstance(path, str) or not path:
        raise ValueError(f"Module {module_index} item {item_index}: image item needs 'image' path")
    source = (Path(base_dir) / path).expanduser().resolve()
    if not source.is_file():
        raise ValueError(f"Module {module_index} item {item_index}: missing image {source}")
    if source.suffix.lower() not in (".png", ".jpg", ".jpeg"):
        raise ValueError(f"Module {module_index} item {item_index}: image must be PNG/JPEG: {source}")
    with Image.open(source) as bitmap:
        ratio = bitmap.width / bitmap.height
    box_w, box_h = box[2], box[3]
    width = min(box_w, box_h * ratio)
    height = width / ratio
    picture = slide.shapes.add_picture(str(source), Mm(left_mm + (box_w - width) / 2),
                                       Mm(top_mm + (box_h - height) / 2), Mm(width), Mm(height))
    crop = item.get("crop") or {}
    for side in ("left", "top", "right", "bottom"):
        if crop.get(side):
            setattr(picture, "crop_" + side, float(crop[side]))
    return picture


PANEL_RATIOS = {"equal": 1.0, "golden": 1.618, "sqrt2": 1.414, "silver": 2.414, "double": 2.0}


def _panel_rects(area, direction, ratio, gap):
    x, y, width, height = area
    if direction == "horizontal":
        first = (width - gap) * ratio / (1 + ratio)
        return [[x, y, first, height], [x + first + gap, y, width - first - gap, height]]
    if direction == "vertical":
        first = (height - gap) * ratio / (1 + ratio)
        return [[x, y, width, first], [x, y + first + gap, width, height - first - gap]]
    raise ValueError(f"panels.direction must be 'horizontal' or 'vertical', got {direction!r}")


def _content_fraction(items):
    """Fraction of the inner area covered by the union box of the items."""
    boxes = [item.get("box", [0.0, 0.0, 1.0, 1.0]) for item in items]
    if not boxes:
        return 0.0
    left = max(0.0, min(box[0] for box in boxes))
    top = max(0.0, min(box[1] for box in boxes))
    right = min(1.0, max(box[0] + box[2] for box in boxes))
    bottom = min(1.0, max(box[1] + box[3] for box in boxes))
    return max(0.0, (right - left) * (bottom - top))


def _add_panels(slide, panels, base_dir, index, role_styles=True, catalog=None):
    area = panels.get("area") or [8.2, 26.0, 322.3, 156.0]
    if not isinstance(area, list) or len(area) != 4:
        raise ValueError(f"Page {index}: panels.area must be [x_mm, y_mm, w_mm, h_mm]")
    direction = panels.get("direction", "horizontal")
    ratio_key = panels.get("ratio", "equal")
    if isinstance(ratio_key, (int, float)):
        ratio = float(ratio_key)
    elif str(ratio_key).replace(".", "", 1).replace("-", "", 1).isdigit():
        ratio = float(ratio_key)
    else:
        ratio = PANEL_RATIOS.get(str(ratio_key))
        if ratio is None:
            raise ValueError(f"Page {index}: unknown panels.ratio {ratio_key!r}; use equal, golden, sqrt2, silver, "
                             "double or a number")
    if ratio <= 0:
        raise ValueError(f"Page {index}: panels.ratio must be positive")
    gap = float(panels.get("gap_mm", 6.0))
    definitions = panels.get("panels")
    if not isinstance(definitions, list) or len(definitions) != 2:
        raise ValueError(f"Page {index}: panels.panels must describe exactly two panels")
    rects = _panel_rects(area, direction, ratio, gap)
    padding = float(panels.get("padding_mm", 4.0))
    tolerance = float(panels.get("whitespace_tolerance", 0.08))
    modules, whitespaces = [], []
    for position, definition in enumerate(definitions):
        if not isinstance(definition, dict):
            raise ValueError(f"Page {index}: panel {position + 1} must be an object")
        items = definition.get("items") or []
        modules.append({"rect": rects[position], "padding_mm": padding,
                        "border": definition.get("border"),
                        "items": items})
        whitespaces.append(1.0 - _content_fraction(items))
    if abs(whitespaces[0] - whitespaces[1]) > tolerance:
        raise ValueError(f"Page {index}: panel whitespace differs too much "
                         f"({whitespaces[0]:.3f} vs {whitespaces[1]:.3f}, tolerance {tolerance}); "
                         "balance the content or adjust panels.ratio")
    summary = _add_modules(slide, modules, base_dir, index, role_styles, catalog)
    for position, entry in enumerate(summary):
        entry["panel"] = position + 1
        entry["whitespace"] = round(whitespaces[position], 4)
    return summary


CITATION_INSET_MM = 2.0
CITATION_GAP_MM = 6.0
CITATION_IMAGE_GAP_MM = 1.2


def _citation_line_height(item):
    return float(item.get("font_pt") or 12) * 1.35 * 25.4 / 72


def _citation_plan(boxes, inner, layout):
    """Where each reference goes: under its figure, or merged at the frame corner."""
    citations = [(index, item, box) for index, (item, box) in enumerate(boxes)
                 if item.get("annotation") == "citation"]
    images = [(item, box) for item, box in boxes if "image" in item]
    targets, leftovers = {}, []
    for index, item, box in citations:
        line_h = _citation_line_height(item)
        if layout == "below_image" and images:
            cite_rect = [inner[0] + box[0] * inner[2], inner[1] + box[1] * inner[3],
                         box[2] * inner[2], box[3] * inner[3]]
            best, best_score = None, None
            for image_item, image_box in images:
                image_rect = [inner[0] + image_box[0] * inner[2], inner[1] + image_box[1] * inner[3],
                              image_box[2] * inner[2], image_box[3] * inner[3]]
                gap = cite_rect[1] - (image_rect[1] + image_rect[3])
                centre_gap = abs((cite_rect[0] + cite_rect[2] / 2) - (image_rect[0] + image_rect[2] / 2))
                if -3.0 <= gap <= 25.0 and centre_gap <= max(25.0, image_rect[2] * 0.7):
                    score = abs(gap) + centre_gap * 0.5
                    if best_score is None or score < best_score:
                        best, best_score = image_rect, score
            if best is not None:
                width = min(best[2], inner[0] + inner[2] - best[0] - CITATION_INSET_MM)
                top = best[1] + best[3] + CITATION_IMAGE_GAP_MM
                if top + line_h <= inner[1] + inner[3] - 0.5:
                    targets[id(item)] = [best[0], top, max(12.0, width), line_h]
                    continue
        leftovers.append((index, item))
    if leftovers:
        row = _citation_boxes([(float(item.get("font_pt") or 12), id(item))
                               for _, item in leftovers], inner)
        for _, item in leftovers:
            targets.setdefault(id(item), row[id(item)])
    return targets


def _merge_citations(items):
    """One text box holding every reference line, anchored to the frame corner."""
    paragraphs = []
    for item in items:
        segments = citation_segments(item.get("citation") or {})
        paragraphs.append({"runs": [{"text": text_value, "annotation": part}
                                    for part, text_value in segments]})
    return {"annotation": "citation", "citation": items[0].get("citation"),
            "paragraphs": paragraphs, "font_pt": items[0].get("font_pt", 12),
            "word_wrap": False, "allow_overflow": True}


def _citation_boxes(citations, inner):
    """References share one row anchored to the module's bottom-left corner.

    Every citation keeps the same offset from the corner and they are spread
    evenly, so a module with several figure references stays tidy instead of
    stacking them in the same spot."""
    if not citations:
        return {}
    rows = [citations]
    line_h = max(size for size, _ in citations) * 1.35 * 25.4 / 72
    available = inner[2] - 2 * CITATION_INSET_MM
    minimum = 20.0
    per_row = max(1, int((available + CITATION_GAP_MM) // (minimum + CITATION_GAP_MM)))
    if per_row < len(citations):
        rows = [citations[index:index + per_row] for index in range(0, len(citations), per_row)]
    layout = {}
    row_heights = [max(size for size, _ in row) * 1.35 * 25.4 / 72 for row in rows]
    top_offset = inner[1] + inner[3] - CITATION_INSET_MM
    for row_index, row in enumerate(rows):
        top_offset -= row_heights[row_index]
        count = len(row)
        width = (available - (count - 1) * CITATION_GAP_MM) / count
        for column, (_, item_id) in enumerate(row):
            left = inner[0] + CITATION_INSET_MM + column * (width + CITATION_GAP_MM)
            layout[item_id] = [left, top_offset, width, row_heights[row_index]]
        top_offset -= CITATION_GAP_MM
    return layout


COLLISION_EPSILON = 0.05


def _intersects(first, second, epsilon=COLLISION_EPSILON):
    return (first[0] < second[0] + second[2] - epsilon
            and second[0] < first[0] + first[2] - epsilon
            and first[1] < second[1] + second[3] - epsilon
            and second[1] < first[1] + first[3] - epsilon)


def _is_exempt(payload):
    return bool(payload.get("overlay") or payload.get("special"))


def _hard_check_page(records, index, strict=True, warnings=None):
    """Every element's content border must be disjoint; special items are exempt."""
    violations = []
    for position, (payload, rect) in enumerate(records):
        if _is_exempt(payload):
            continue
        for other_payload, other_rect in records[position + 1:]:
            if _is_exempt(other_payload):
                continue
            if _intersects(rect, other_rect):
                violations.append((payload_text(payload)[:24] or "<image>",
                                   payload_text(other_payload)[:24] or "<image>", rect, other_rect))
    if violations:
        detail = "; ".join(f"{first!r} x {second!r}" for first, second, _, _ in violations[:6])
        message = (f"Page {index}: content borders intersect ({len(violations)} pair(s)): {detail}. "
                   "Mark intentionally overlapping items with \"overlay\": true or fix the layout.")
        if strict:
            raise ValueError(message)
        if warnings is not None:
            warnings.append(message)


def _repair_collisions(entries, notes=None):
    """Move non-exempt elements vertically/horizontally until content borders clear."""
    from pptx.util import Mm as _Mm
    for _ in range(60):
        collision = None
        for position, (payload, rect, shape, inner) in enumerate(entries):
            if _is_exempt(payload):
                continue
            for other_payload, other_rect, _, _ in entries[position + 1:]:
                if _is_exempt(other_payload):
                    continue
                if _intersects(rect, other_rect):
                    collision = (payload, rect, shape, inner, other_rect)
                    break
            if collision:
                break
        if collision is None:
            return
        payload, rect, shape, inner, other_rect = collision
        second_record = next((record for record in entries
                              if record[1] is other_rect), None)
        candidates = [(payload, rect, shape, inner)]
        if second_record is not None:
            candidates.append(second_record)
        moved = False
        for payload, rect, shape, inner in candidates:
            attempts = []
            for step in range(1, 9):
                attempts.append((0.0, step * 0.5))
                attempts.append((0.0, -step * 0.5))
            for step in range(1, 5):
                attempts.append((step * 0.5, 0.0))
                attempts.append((-step * 0.5, 0.0))
            for dx, delta in attempts:
                candidate = [rect[0] + dx, rect[1] + delta, rect[2], rect[3]]
                if candidate[1] < inner[1] or candidate[1] + candidate[3] > inner[1] + inner[3]:
                    continue
                if candidate[0] < inner[0] or candidate[0] + candidate[2] > inner[0] + inner[2]:
                    continue
                if not _intersects(candidate, other_rect):
                    left = shape.left + _Mm(candidate[0] - rect[0])
                    top = shape.top + _Mm(candidate[1] - rect[1])
                    shape.left, shape.top = left, top
                    payload["box"][0] += (candidate[0] - rect[0]) / inner[2]
                    payload["box"][1] += (candidate[1] - rect[1]) / inner[3]
                    rect[0], rect[1] = candidate[0], candidate[1]
                    moved = True
                    break
            if moved:
                break
        if not moved:
            # Leave the element where it is; the hard check below reports the
            # violation instead of silently exempting it.
            return


def _add_modules(slide, modules, base_dir, index, role_styles=True, catalog=None,
                 citation_layout="below_image", warnings=None, page=None, strict=True):
    from pptx.dml.color import RGBColor
    from pptx.enum.dml import MSO_LINE_DASH_STYLE
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.util import Mm, Pt
    if not isinstance(modules, list) or not modules:
        raise ValueError(f"Page {index}: modules must be a nonempty list")
    placed, summary, collision_records = [], [], []
    for module_index, module in enumerate(modules, 1):
        if not isinstance(module, dict):
            raise ValueError(f"Page {index}: module {module_index} must be an object")
        rect = _module_rect(module, module_index)
        border = module.get("border")
        fill = module.get("fill")
        framed = isinstance(border, dict)
        if framed:
            width_pt = float(border.get("width_pt", 2.0))
            padding = float(module.get("padding_mm", 2.0))
            if padding < width_pt * 25.4 / 72 / 2 + 0.5:
                raise ValueError(f"Page {index} module {module_index}: padding_mm {padding} is too small for a "
                                 f"{width_pt} pt border; content would touch the line")
        else:
            width_pt = 0.0
            padding = float(module.get("padding_mm", 0.0))
        radius = float(module.get("radius_mm", 0) or 0)
        for key, value in (("x", rect[0]), ("y", rect[1])):
            if value < 0:
                raise ValueError(f"Page {index} module {module_index}: negative {key} {value}")
        if rect[0] + rect[2] > 338.667 + 0.5 or rect[1] + rect[3] > 190.5 + 0.5:
            raise ValueError(f"Page {index} module {module_index}: rect leaves the slide")
        if not framed and not fill:
            shape = None
        elif radius > 0:
            shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Mm(rect[0]), Mm(rect[1]),
                                           Mm(rect[2]), Mm(rect[3]))
            # roundRect's own default is a large radius; set the requested one explicitly.
            shape.adjustments[0] = max(0.0, min(0.5, radius / min(rect[2], rect[3])))
        else:
            shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Mm(rect[0]), Mm(rect[1]),
                                           Mm(rect[2]), Mm(rect[3]))
        if shape is not None:
            if fill:
                from pptx.oxml.ns import qn
                specification = fill if isinstance(fill, dict) else {"color": fill}
                shape.fill.solid()
                shape.fill.fore_color.rgb = RGBColor.from_string(str(specification["color"]).lstrip("#"))
                alpha = specification.get("alpha_pct")
                if alpha is not None:
                    solid = shape._element.spPr.find(qn("a:solidFill"))
                    srgb = solid.find(qn("a:srgbClr")) if solid is not None else None
                    if srgb is not None:
                        node = srgb.makeelement(qn("a:alpha"), {"val": str(int(round(float(alpha) * 1000)))})
                        srgb.append(node)
            else:
                shape.fill.background()
            if framed:
                shape.line.color.rgb = RGBColor.from_string(str(border.get("color", "#B8B8B7")).lstrip("#"))
                shape.line.width = Pt(width_pt)
                shape.line.dash_style = getattr(MSO_LINE_DASH_STYLE,
                                                DASH_STYLES.get(str(border.get("dash", "round_dot")), "ROUND_DOT"))
            else:
                shape.line.fill.background()
        inner = [rect[0] + padding, rect[1] + padding,
                 rect[2] - 2 * padding, rect[3] - 2 * padding]
        if inner[2] <= 0 or inner[3] <= 0:
            raise ValueError(f"Page {index} module {module_index}: padding leaves no content area")
        items = module.get("items") or []
        boxes = []
        for item_index, item in enumerate(items, 1):
            if not isinstance(item, dict):
                raise ValueError(f"Module {module_index} item {item_index} must be an object")
            box = _item_box(item, module_index, item_index)
            _check_inside(box, module_index, item_index)
            boxes.append((item, box))
        mode = module.get("citation_layout") or citation_layout
        if mode == "frame_corner" and any(item.get("annotation") == "citation" for item, _ in boxes):
            citation_items = [item for item, _ in boxes if item.get("annotation") == "citation"]
            boxes = [(item, box) for item, box in boxes if item.get("annotation") != "citation"]
            merged = _merge_citations(citation_items)
            boxes.append((merged, [0.0, 0.0, 0.0, 0.0]))
            merged_lines = len(citation_items)
            line_h = _citation_line_height(merged)
            width = max(measure_text("".join(_paragraph_text(p) for p in merged["paragraphs"]),
                                     float(merged.get("font_pt") or 12))[0] + 4.0, 20.0)
            top = inner[1] + inner[3] - CITATION_INSET_MM - merged_lines * line_h
            cite_target = (inner[0] + CITATION_INSET_MM, top,
                           min(width, inner[2] - 2 * CITATION_INSET_MM), merged_lines * line_h)
        else:
            plan = _citation_plan(boxes, inner, mode)
            cite_target = None
        for item_index, (item, box) in enumerate(boxes, 1):
            if item.get("annotation") == "citation":
                if cite_target is not None:
                    left, top, width_mm, height_mm = cite_target
                else:
                    left, top, width_mm, height_mm = plan[id(item)]
                target = (left, top, width_mm, height_mm)
            else:
                left = inner[0] + box[0] * inner[2]
                top = inner[1] + box[1] * inner[3]
                size = (box[2] * inner[2], box[3] * inner[3])
                target = (left, top, size[0], size[1])
            if "text" in item or "paragraphs" in item or "citation" in item:
                placed_shape = _add_text_item(slide, left, top, target, item, module_index, item_index,
                                              role_styles, catalog, inner)
            elif "image" in item:
                placed_shape = _add_image_item(slide, left, top, target, item, module_index, item_index, base_dir)
            else:
                raise ValueError(f"Module {module_index} item {item_index}: needs 'text' or 'image'")
            placed.append((module_index, item_index, placed_shape, inner))
            collision_records.append({
                "payload": item,
                "rect": payload_collision_rect(item, target),
                "shape": placed_shape,
                "inner": inner,
            })
        summary.append({"index": module_index, "rect_mm": rect, "items": len(items)})
    _repair_collisions([(record["payload"], record["rect"], record["shape"], record["inner"])
                        for record in collision_records], warnings)
    for record in collision_records:
        record["rect"] = payload_collision_rect(record["payload"],
                                                [record["shape"].left / 36000, record["shape"].top / 36000,
                                                 record["shape"].width / 36000, record["shape"].height / 36000])
    _hard_check_page([(record["payload"], record["rect"]) for record in collision_records], index,
                     strict, warnings)
    page_records = []
    for record in collision_records:
        payload = record["payload"]
        if "image" in payload:
            kind = "image"
        elif payload.get("annotation") == "badge":
            kind = "badge"
        else:
            kind = "text"
        page_records.append({
            "kind": kind,
            "exempt": bool(payload.get("overlay") or payload.get("special")),
            "label": (payload_text(payload)[:20] or "<image>"),
            "rect_mm": [round(value, 2) for value in record["rect"]],
        })
    if page is not None:
        page["_collision_boxes"] = page_records
    for record in collision_records:
        inner = record["inner"]
        x, y, width, height = record["rect"]
        crosses = (x < inner[0] - 0.4 or y < inner[1] - 0.4
                   or x + width > inner[0] + inner[2] + 0.4
                   or y + height > inner[1] + inner[3] + 0.4)
        if crosses and warnings is not None:
            warnings.append(f"border:{payload_text(record['payload'])[:20]}")
    return summary


def _apply_title_style(shape, style):
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from pptx.util import Mm, Pt
    if all(key in style for key in ("x_mm", "y_mm", "w_mm", "h_mm")):
        shape.left, shape.top = Mm(style["x_mm"]), Mm(style["y_mm"])
        shape.width, shape.height = Mm(style["w_mm"]), Mm(style["h_mm"])
    alignment = {"left": PP_ALIGN.LEFT, "center": PP_ALIGN.CENTER,
                 "right": PP_ALIGN.RIGHT, "justify": PP_ALIGN.JUSTIFY}.get(style.get("align"))
    for paragraph in shape.text_frame.paragraphs:
        if alignment is not None:
            paragraph.alignment = alignment
        for run in paragraph.runs:
            if "font_pt" in style:
                run.font.size = Pt(style["font_pt"])
            if style.get("bold"):
                run.font.bold = True
            if style.get("color"):
                run.font.color.rgb = RGBColor.from_string(style["color"].lstrip("#"))


def _add_layout_page(prs, layout, page, index, title_style=None, title_xml=None, base_dir=None,
                     role_styles=True, catalog=None, citation_layout="below_image", warnings=None,
                     strict_collisions=True):
    from pptx.enum.shapes import PP_PLACEHOLDER
    from pptx.util import Mm, Pt
    slide = prs.slides.add_slide(layout)
    title = page.get("title")
    page_title_xml = page.get("title_xml") or title_xml
    if title is not None:
        if not isinstance(title, str):
            raise ValueError(f"Page {index}: title must be a string")
        if page_title_xml:
            shape = _replace_title_xml(slide, page_title_xml)
            # Rewriting runs would drop per-slide autofit/geometry specifics, so
            # only touch the text when the requested title really differs.
            if shape.text_frame.text != title:
                _set_shape_text(shape, title)
        else:
            placeholder = slide.shapes.title
            if placeholder is not None:
                _set_shape_text(placeholder, title)
                _apply_title_style(placeholder, title_style or {})
            else:
                box = slide.shapes.add_textbox(Mm(20), Mm(15), prs.slide_width - Mm(40), Mm(30))
                _set_shape_text(box, title)
                for run in box.text_frame.paragraphs[0].runs:
                    run.font.size = Pt(28)
                _apply_title_style(box, title_style or {})
    body = page.get("body") or []
    if body:
        if not isinstance(body, list) or not all(isinstance(line, str) for line in body):
            raise ValueError(f"Page {index}: body must be a list of strings")
        size = page.get("font_pt", 14)
        target = next((shape for shape in slide.placeholders
                       if shape.placeholder_format.type in (PP_PLACEHOLDER.BODY, PP_PLACEHOLDER.OBJECT)), None)
        if target is not None:
            _set_shape_text(target, "\n".join(body))
            for paragraph in target.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(size)
        else:
            box = slide.shapes.add_textbox(Mm(20), Mm(48), prs.slide_width - Mm(40), prs.slide_height - Mm(70))
            box.text_frame.word_wrap = True
            _set_shape_text(box, "\n".join(body))
            for paragraph in box.text_frame.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(size)
    modules = page.get("modules") or []
    panels = page.get("panels")
    if panels:
        if modules or body:
            raise ValueError(f"Page {index}: use 'panels', 'modules' or 'body' alone, not a combination")
        page["_module_summary"] = _add_panels(slide, panels, base_dir or Path.cwd(), index, role_styles, catalog)
    elif modules:
        if body:
            raise ValueError(f"Page {index}: use 'modules' or 'body', not both")
        page["_module_summary"] = _add_modules(slide, modules, base_dir or Path.cwd(), index,
                                               role_styles, catalog, citation_layout, warnings, page,
                                               strict_collisions)
    return slide


def _delete_slide(prs, slide):
    sld_id_list = prs.slides._sldIdLst
    for sld_id in list(sld_id_list):
        if int(sld_id.get("id")) == slide.slide_id:
            prs.part.drop_rel(sld_id.rId)
            sld_id_list.remove(sld_id)
            return
    raise ValueError("Cannot find the slide id to delete")


def _reorder(prs, order):
    sld_id_list = prs.slides._sldIdLst
    # Resolve ids before touching the list: Slide.slide_id looks the slide up in it.
    elements = {int(sld_id.get("id")): sld_id for sld_id in sld_id_list}
    desired = [elements[slide.slide_id] for slide in order]
    if len(desired) != len(elements):
        raise ValueError(f"Composed order covers {len(desired)} of {len(elements)} slide(s)")
    for sld_id in list(sld_id_list):
        sld_id_list.remove(sld_id)
    for sld_id in desired:
        sld_id_list.append(sld_id)


def build(template_dir, spec, output, force=False):
    from pptx import Presentation
    template_dir = Path(template_dir).expanduser().resolve()
    output = Path(output).expanduser().resolve()
    descriptor = load_descriptor(template_dir / DESCRIPTOR_NAME)
    package = template_dir / descriptor["files"]["package"]
    document = spec if isinstance(spec, dict) else read_json(spec)
    spec_path = None if isinstance(spec, dict) else Path(spec)
    reference = document.get("base_dir")
    if reference:
        base_dir = Path(reference).expanduser().resolve()
    else:
        base_dir = spec_path.expanduser().resolve().parent if spec_path else Path.cwd()
    pages = document.get("pages")
    if not isinstance(pages, list) or not pages:
        raise ValueError("Compose spec requires a nonempty 'pages' array")
    if document.get("template") not in (None, descriptor["id"]):
        raise ValueError(f"Spec targets template {document['template']!r}, but {template_dir.name!r} was supplied")
    manifest_path_file = compose_manifest_path(output)
    preflight_outputs([output, manifest_path_file], force,
                      [package, spec_path] if spec_path else [package])
    presentation = Presentation(str(package))
    prototypes = {prototype["id"]: prototype for prototype in descriptor.get("prototypes", [])}
    if len(presentation.slides) != len(prototypes):
        raise ValueError(f"Template package has {len(presentation.slides)} prototype slide(s) but the descriptor "
                         f"declares {len(prototypes)}; re-ingest the template")
    slides = {prototype_id: slide for prototype_id, slide in
              zip([prototype["id"] for prototype in descriptor.get("prototypes", [])], presentation.slides)}
    used, order, manifest_pages = set(), [], []
    collision_warnings = []
    default_role_styles = bool(document.get("role_styles", True))
    catalog = merge_styles(document.get("text_styles") or descriptor.get("text_styles"))
    palette = merge_text_colors(document.get("text_colors") or descriptor.get("text_colors"),
                                document.get("text_colors") or descriptor.get("text_colors"))
    for annotation, alignment in (("display", "center"), ("heading", "left"), ("subheading", "left"),
                                  ("body", "left"), ("body_large", "left"), ("bullet", "left"),
                                  ("caption", "center"), ("citation", "left"), ("note", "left"),
                                  ("label", "left"), ("data_label", "center")):
        entry = catalog.setdefault(annotation, {})
        entry["align"] = alignment
    for annotation, role in (("heading", "weak"), ("subheading", "weak"), ("body", "normal"),
                             ("body_large", "normal"), ("bullet", "normal"), ("caption", "normal"),
                             ("note", "normal"), ("label", "normal"), ("data_label", "strong"),
                             ("emphasis", "strong")):
        entry = catalog.setdefault(annotation, {})
        entry.setdefault("color_role", role)
    catalog["_palette"] = palette
    for index, page in enumerate(pages, 1):
        if not isinstance(page, dict):
            raise ValueError(f"Page {index} must be an object")
        kinds = [kind for kind in PROTOTYPE_KINDS if kind in page]
        if len(kinds) != 1:
            raise ValueError(f"Page {index} needs exactly one of 'prototype' or 'layout'")
        if kinds[0] == "prototype":
            prototype_id = page["prototype"]
            if prototype_id not in slides:
                raise ValueError(f"Page {index}: unknown prototype {prototype_id!r}; available: "
                                 + ", ".join(sorted(slides)))
            reused = prototype_id in used
            used.add(prototype_id)
            slide = _clone_slide(presentation, slides[prototype_id]) if reused else slides[prototype_id]
            changed = _apply_texts(slide, page.get("texts", {}), prototypes[prototype_id])
            order.append(slide)
            manifest_pages.append({"page": index, "kind": "prototype", "prototype": prototype_id,
                                   "layout_index": prototypes[prototype_id]["layout_index"],
                                   "clone": reused, "texts_changed": changed})
        else:
            layout = _layout(presentation, page["layout"], descriptor)
            content_style = descriptor.get("content_style", {})
            page_warnings = []
            slide = _add_layout_page(presentation, layout, page, index, content_style.get("title"),
                                     content_style.get("title_xml"), base_dir,
                                     bool(page.get("role_styles", default_role_styles)), catalog,
                                     document.get("citation_layout", "below_image"), page_warnings,
                                     bool(document.get("strict_collisions", True)))
            order.append(slide)
            record = {"page": index, "kind": "layout", "layout": layout.name,
                      "title": page.get("title", ""),
                      "collision_boxes": page.get("_collision_boxes", [])}
            if page_warnings:
                record["collision_exemptions"] = page_warnings
                collision_warnings.extend(f"page {index}: {name}" for name in page_warnings)
            if page.get("_module_summary"):
                record["modules"] = page["_module_summary"]
            manifest_pages.append(record)
    for prototype_id, slide in slides.items():
        if prototype_id not in used:
            _delete_slide(presentation, slide)
    _reorder(presentation, order)
    buffer = io.BytesIO()
    presentation.save(buffer)
    atomic_bytes(output, buffer.getvalue())
    manifest = {
        "template": {"id": descriptor["id"], "school": descriptor["school"],
                     "descriptor_sha256": descriptor_sha256(template_dir / DESCRIPTOR_NAME),
                     "package": package.name},
        "pages": manifest_pages,
        "warnings": [
            *collision_warnings,
            "Prototype pages reproduce the template's design verbatim; only listed text slots were rewritten.",
            "Layout pages use the template's layout and placeholders, so free-form decoration from the source deck is not reproduced.",
        ],
    }
    write_json(manifest_path_file, manifest)
    return manifest, output


def main(skill_dir):
    parser = argparse.ArgumentParser(prog="compose_deck.py",
                                     description="Compose a PPTX from a template: prototype pages and layout pages.")
    source = parser.add_mutually_exclusive_group(required=True)
    source.add_argument("--template", type=Path, help="Template directory containing template.json")
    source.add_argument("--template-id", help="Template id resolved through the library catalog")
    parser.add_argument("--root", type=Path, help="Library root for --template-id")
    parser.add_argument("--spec", type=Path, required=True, help="Compose spec JSON")
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()
    if args.template:
        directory = args.template
    else:
        from .template import load_catalog
        root = template_root(skill_dir, args.root)
        catalog = load_catalog(root)
        matches = [entry for entry in catalog["templates"] if entry["id"] == args.template_id]
        if not matches:
            raise ValueError(f"No template with id {args.template_id!r}; run query_template.py list")
        directory = root / matches[0]["path"]
    manifest, output = build(directory, args.spec, args.output, args.force)
    print(f"{output}  ({len(manifest['pages'])} pages)")
    for page in manifest["pages"]:
        if page["kind"] == "prototype":
            detail = f"prototype {page['prototype']!r}, {len(page['texts_changed'])} text slot(s) rewritten"
        else:
            detail = f"layout {page['layout']!r}, title {page['title']!r}"
        print(f"  {page['page']:2d}  {detail}")
