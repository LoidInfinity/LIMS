"""Conservative extraction from explicit OOXML/SVG formatting, with evidence."""
from collections import Counter, defaultdict
from copy import deepcopy
from pathlib import Path
import argparse
import re
import xml.etree.ElementTree as ET
import zipfile

from .core import load_style, fingerprint, preflight_outputs, validate, write_json

NS = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
      "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
W = "{" + NS["w"] + "}"


def leaves(value, prefix=""):
    for key, item in value.items():
        name = f"{prefix}.{key}" if prefix else key
        if isinstance(item, dict):
            yield from leaves(item, name)
        else:
            yield name


class Observations:
    def __init__(self):
        self.values = defaultdict(list)
        self.warnings = []

    def add(self, field, value, source):
        if value is not None and value != "":
            self.values[field].append((value, str(source)))


def xml_from(archive, name):
    info = archive.getinfo(name)
    if info.file_size > 20_000_000:
        raise ValueError(f"XML part too large: {name}")
    data = archive.read(name)
    if b"<!DOCTYPE" in data or b"<!ENTITY" in data:
        raise ValueError(f"XML entities are not supported: {name}")
    return ET.fromstring(data)


def theme(archive, prefix, obs, source):
    name = f"{prefix}/theme/theme1.xml"
    if name not in archive.namelist():
        return
    root = xml_from(archive, name)
    colors = []
    for i in range(1, 7):
        elem = root.find(f".//a:clrScheme/a:accent{i}/a:srgbClr", NS)
        if elem is not None:
            colors.append("#" + elem.attrib["val"])
    if colors:
        obs.add("palette.series", tuple(colors), source)
        obs.add("palette.accent", colors[0], source)
    for field, tag in (("latin", "latin"), ("cjk", "ea")):
        node = root.find(f".//a:fontScheme/a:minorFont/a:{tag}", NS)
        if node is not None:
            obs.add(f"typography.{field}", node.get("typeface"), source)


def pptx(path, obs):
    with zipfile.ZipFile(path) as archive:
        root = xml_from(archive, "ppt/presentation.xml")
        size = root.find("p:sldSz", NS)
        if size is not None:
            obs.add("slides.width_mm", round(int(size.get("cx")) / 36000, 3), path)
            obs.add("slides.height_mm", round(int(size.get("cy")) / 36000, 3), path)
        theme(archive, "ppt", obs, path)
        for name in sorted(archive.namelist()):
            if not re.fullmatch(r"ppt/slides/slide\d+\.xml", name):
                continue
            slide = xml_from(archive, name)
            for shape in slide.findall(".//p:sp", NS):
                ph = shape.find(".//p:ph", NS)
                # Only placeholders have unambiguous semantic roles.
                role = ph.get("type", "body") if ph is not None else None
                field = "slides.title_pt" if role in ("title", "ctrTitle") else "slides.body_pt" if role in ("body", "obj") else None
                for props in shape.findall(".//a:rPr", NS):
                    if field and props.get("sz"):
                        obs.add(field, int(props.get("sz")) / 100, path)
                    for key, tag in (("latin", "latin"), ("cjk", "ea")):
                        font = props.find(f"a:{tag}", NS)
                        if font is not None and not font.get("typeface", "").startswith("+"):
                            obs.add(f"typography.{key}", font.get("typeface"), path)
        obs.warnings.append(f"{path.name}: PPT master/layout inheritance and non-placeholder text roles are not inferred")


def docx(path, obs):
    with zipfile.ZipFile(path) as archive:
        root = xml_from(archive, "word/document.xml")
        for section in root.findall(".//w:sectPr", NS):
            size, margins = section.find("w:pgSz", NS), section.find("w:pgMar", NS)
            if size is not None:
                for key, attr in (("width_mm", "w"), ("height_mm", "h")):
                    if size.get(W + attr):
                        obs.add(f"paper.{key}", round(int(size.get(W + attr)) * 25.4 / 1440, 3), path)
            if margins is not None:
                for key in ("top", "bottom", "left", "right"):
                    if margins.get(W + key):
                        obs.add(f"paper.margins_mm.{key}", round(int(margins.get(W + key)) * 25.4 / 1440, 3), path)
        if "word/styles.xml" in archive.namelist():
            styles = xml_from(archive, "word/styles.xml")
            for style in styles.findall("w:style", NS):
                sid = style.get(W + "styleId", "").lower()
                role = {"normal": "body", "title": "title", "heading1": "heading"}.get(sid)
                if not role:
                    continue
                sz = style.find("w:rPr/w:sz", NS)
                if sz is not None:
                    obs.add(f"paper.{role}_pt", int(sz.get(W + "val")) / 2, path)
                if role == "body":
                    fonts = style.find("w:rPr/w:rFonts", NS)
                    if fonts is not None:
                        obs.add("typography.latin", fonts.get(W + "ascii"), path)
                        obs.add("typography.cjk", fonts.get(W + "eastAsia"), path)
                    spacing = style.find("w:pPr/w:spacing", NS)
                    if spacing is not None and spacing.get(W + "lineRule", "auto") == "auto" and spacing.get(W + "line"):
                        obs.add("paper.line_spacing", int(spacing.get(W + "line")) / 240, path)
        obs.warnings.append(f"{path.name}: only explicit Normal/Title/Heading1 properties are extracted; citation conventions require confirmation")


def svg_mm(value):
    match = re.fullmatch(r"([\d.]+)(mm|cm|in|pt|px)?", value or "")
    if not match:
        return None
    scale = {"mm": 1, "cm": 10, "in": 25.4, "pt": 25.4 / 72, "px": 25.4 / 96, None: 25.4 / 96}
    return round(float(match[1]) * scale[match[2]], 3)


def svg(path, obs):
    data = path.read_bytes()
    # Common SVG exporters include the external SVG 1.1 DOCTYPE. ElementTree
    # does not fetch it. Reject internal subsets/entities, which can expand data.
    if len(data) > 20_000_000 or b"<!ENTITY" in data or re.search(rb"<!DOCTYPE[^>]*\[", data):
        raise ValueError("SVG must be under 20 MB without internal DTD/entity declarations")
    root = ET.fromstring(data)
    if root.tag.rsplit("}", 1)[-1] != "svg":
        raise ValueError("Expected an SVG root element")
    for dimension in ("width", "height"):
        obs.add(f"chart.{dimension}_mm", svg_mm(root.get(dimension)), path)
    colors = Counter()
    for node in root.iter():
        props = dict(node.attrib)
        for part in node.get("style", "").split(";"):
            if ":" in part:
                key, value = part.split(":", 1)
                props[key.strip()] = value.strip()
        color = props.get("stroke", "")
        if re.fullmatch(r"#[0-9a-fA-F]{6}", color) and color.lower() not in ("#ffffff", "#000000"):
            colors[color.upper()] += 1
        if node.tag.rsplit("}", 1)[-1] == "text":
            if props.get("font-family"):
                obs.add("typography.latin", props["font-family"].split(",")[0].strip("'\" "), path)
            size = svg_mm(props.get("font-size"))
            if size:
                obs.add("chart.font_pt", round(size * 72 / 25.4, 2), path)
    if colors:
        palette = tuple(color for color, count in colors.most_common(12))
        obs.add("palette.series", palette, path)
        obs.add("palette.accent", palette[0], path)
    obs.warnings.append(f"{path.name}: SVG CSS cascade, transforms and outlined text are not resolved; axis/curve widths are not guessed")


def extract(paths, base):
    obs = Observations()
    parsers = {".pptx": pptx, ".docx": docx, ".svg": svg}
    for path in map(Path, paths):
        if path.suffix.lower() not in parsers:
            raise ValueError(f"Unsupported input {path.suffix}: use PPTX, DOCX or SVG; PDF/raster/OPJU need manual style entry")
        try:
            parsers[path.suffix.lower()](path, obs)
        except (zipfile.BadZipFile, KeyError, ET.ParseError) as exc:
            raise ValueError(f"Cannot parse {path}: {exc}") from exc
    result, fields = deepcopy(base), {}
    for field, values in sorted(obs.values.items()):
        counts = Counter(value for value, source in values)
        winner, count = counts.most_common(1)[0]
        parts, candidate = field.split("."), deepcopy(result)
        target = candidate
        for part in parts[:-1]:
            target = target[part]
        target[parts[-1]] = list(winner) if isinstance(winner, tuple) else winner
        try:
            validate(candidate)
        except ValueError:
            obs.warnings.append(f"{field}: observed value {winner!r} is outside supported schema; retained base")
            continue
        result = candidate
        fields[field] = {"value": target[parts[-1]], "support": count, "observations": len(values),
                         "sources": sorted({source for value, source in values if value == winner}),
                         "ambiguous": len(counts) > 1}
    report = {"base_sha256": fingerprint(base), "fields": fields,
              "retained_from_base": sorted(set(leaves(base)) - set(fields)),
              "warnings": obs.warnings,
              "method": "Most frequent explicit observation; ties follow input order. No visual inference."}
    return validate(result), report


def main(skill_dir):
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("inputs", nargs="+", type=Path)
    parser.add_argument("--style", help="Base style; otherwise current style")
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--report", type=Path)
    parser.add_argument("--id", default="extracted-style")
    parser.add_argument("--name", default="Extracted research style")
    parser.add_argument("--force", action="store_true")
    args = parser.parse_args()
    base, source = load_style(skill_dir, args.style)
    report_path = args.report or Path(str(args.output) + ".evidence.json")
    preflight_outputs([args.output, report_path], args.force, [*args.inputs, source])
    result, report = extract(args.inputs, base)
    result.update(id=args.id, name=args.name)
    validate(result)
    report["style_sha256"] = fingerprint(result)
    write_json(args.output, result)
    write_json(report_path, report)
    print(f"Extracted {len(report['fields'])} fields -> {args.output}; evidence -> {report_path}")
