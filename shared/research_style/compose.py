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
from pathlib import Path

from .core import atomic_bytes, preflight_outputs, read_json, write_json
from .template import DESCRIPTOR_NAME, descriptor_sha256, load_descriptor, template_root

PROTOTYPE_KINDS = ("prototype", "layout")


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
    for paragraph, line in zip(paragraphs, lines):
        runs = paragraph.runs
        if not runs:
            paragraph.add_run().text = line
            continue
        runs[0].text = line
        for run in runs[1:]:
            run._r.getparent().remove(run._r)
    for paragraph in paragraphs[len(lines):]:
        paragraph._p.getparent().remove(paragraph._p)


def _apply_texts(slide, texts, prototype):
    if not isinstance(texts, dict):
        raise ValueError("Prototype page 'texts' must be an object of shape id/name to text")
    current = {slot["shape_id"]: slot["text"] for slot in prototype.get("text_slots", [])}
    changed = []
    for key, value in texts.items():
        if not isinstance(value, str):
            raise ValueError(f"Text for shape {key} must be a string")
        shape = _find_shape(slide, key)
        if current.get(shape.shape_id) == value:
            continue
        _set_shape_text(shape, value)
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


def _add_layout_page(prs, layout, page, index, title_style=None):
    from pptx.enum.shapes import PP_PLACEHOLDER
    from pptx.util import Mm, Pt
    slide = prs.slides.add_slide(layout)
    title = page.get("title")
    if title is not None:
        if not isinstance(title, str):
            raise ValueError(f"Page {index}: title must be a string")
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
            if prototype_id in used:
                raise ValueError(f"Page {index}: prototype {prototype_id!r} cannot be used twice")
            used.add(prototype_id)
            changed = _apply_texts(slides[prototype_id], page.get("texts", {}), prototypes[prototype_id])
            order.append(slides[prototype_id])
            manifest_pages.append({"page": index, "kind": "prototype", "prototype": prototype_id,
                                   "layout_index": prototypes[prototype_id]["layout_index"],
                                   "texts_changed": changed})
        else:
            layout = _layout(presentation, page["layout"], descriptor)
            order.append(_add_layout_page(presentation, layout, page, index,
                                          descriptor.get("content_style", {}).get("title")))
            manifest_pages.append({"page": index, "kind": "layout", "layout": layout.name,
                                   "title": page.get("title", "")})
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
