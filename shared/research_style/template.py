"""Template library: descriptor validation, catalog maintenance and queries."""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import sys
import unicodedata
from datetime import datetime, timezone
from pathlib import Path

from .core import read_json, write_json

CATALOG_NAME = "index.json"
DESCRIPTOR_NAME = "template.json"
SCHEMA_PATH = Path(__file__).with_name("template.schema.json")
SLUG_RE = re.compile(r"^[a-z0-9][a-z0-9-]{0,63}$")
ID_RE = re.compile(r"^[a-z0-9][a-z0-9-]{1,63}$")


def slugify(value):
    """ASCII slug; Chinese names need an explicit pinyin slug."""
    text = unicodedata.normalize("NFKD", value or "")
    text = text.encode("ascii", "ignore").decode("ascii").lower()
    text = re.sub(r"[^a-z0-9]+", "-", text).strip("-")
    return re.sub(r"-{2,}", "-", text)[:64].strip("-")


def now_iso():
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def template_root(skill_dir=None, explicit=None):
    if explicit:
        return Path(explicit).expanduser().resolve()
    candidates = []
    if skill_dir:
        skill = Path(skill_dir).resolve()
        candidates += [skill.parent / "templates", skill / "assets" / "templates"]
    candidates.append(Path.cwd() / "templates")
    for candidate in candidates:
        if (candidate / CATALOG_NAME).is_file():
            return candidate.resolve()
    for candidate in candidates:
        if candidate.is_dir():
            return candidate.resolve()
    raise ValueError("Missing templates/ library; pass --root or ingest a template first")


def validate_template(doc):
    try:
        from jsonschema import Draft202012Validator
    except ImportError as exc:
        raise ValueError("Missing jsonschema; install requirements.txt in the active Python environment") from exc
    errors = sorted(Draft202012Validator(read_json(SCHEMA_PATH)).iter_errors(doc), key=lambda e: str(e.path))
    if errors:
        message = "; ".join(f"{'.'.join(map(str, e.path)) or '<root>'}: {e.message}" for e in errors[:8])
        raise ValueError(f"Invalid template descriptor: {message}")
    for value in _walk(doc):
        if isinstance(value, float) and not math.isfinite(value):
            raise ValueError("Template descriptor contains a non-finite number")
    indices = [layout["index"] for layout in doc["layouts"]]
    if len(indices) != len(set(indices)):
        raise ValueError("Template descriptor repeats a layout index")
    if doc["files"]["package"] == doc["files"]["preview"]:
        raise ValueError("Template package and preview must be different files")
    return doc


def _walk(value):
    yield value
    if isinstance(value, dict):
        for item in value.values():
            yield from _walk(item)
    elif isinstance(value, list):
        for item in value:
            yield from _walk(item)


def descriptor_sha256(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def load_descriptor(path):
    doc = validate_template(read_json(path))
    directory = Path(path).parent
    if directory.name != doc["id"]:
        raise ValueError(f"{path}: directory name must equal descriptor id {doc['id']!r}")
    if directory.parent.name != doc["school"]["slug"]:
        raise ValueError(f"{path}: parent directory must equal school slug {doc['school']['slug']!r}")
    for key in ("package", "preview"):
        name = doc["files"][key]
        if not (directory / name).is_file():
            raise ValueError(f"{path}: missing {key} file {name}")
    if not SLUG_RE.match(doc["school"]["slug"]) or not ID_RE.match(doc["id"]):
        raise ValueError(f"{path}: invalid slug or id")
    return doc


def iter_descriptor_paths(root):
    return sorted(Path(root).glob(f"*/*/{DESCRIPTOR_NAME}"))


def catalog_entry(root, path):
    doc = load_descriptor(path)
    entry = {
        "id": doc["id"],
        "name": doc["name"],
        "school": doc["school"],
        "tags": doc["tags"],
        "aspect": doc["aspect"],
        "size": doc["size"],
        "layout_types": sorted({layout["type"] for layout in doc["layouts"]}),
        "layout_names": list(dict.fromkeys(layout["name"] for layout in doc["layouts"])),
        "path": Path(path).parent.relative_to(Path(root)).as_posix(),
        "preview": (Path(path).parent / doc["files"]["preview"]).relative_to(Path(root)).as_posix(),
        "ingested_at": doc["provenance"]["ingested_at"],
        "descriptor_sha256": descriptor_sha256(path),
    }
    return entry


def rebuild_catalog(root, write=True):
    root = Path(root).resolve()
    entries = [catalog_entry(root, path) for path in iter_descriptor_paths(root)]
    catalog = {
        "schema_version": 1,
        "generated_at": now_iso(),
        "templates": sorted(entries, key=lambda e: (e["school"]["slug"], e["id"])),
    }
    if write:
        write_json(root / CATALOG_NAME, catalog)
    return catalog


def load_catalog(root):
    path = Path(root) / CATALOG_NAME
    if not path.is_file():
        raise ValueError(f"Missing catalog {path}; run: query_template.py index --rebuild")
    return read_json(path)


def catalog_drift(root, catalog=None):
    """Return human-readable differences between the catalog and the library."""
    root = Path(root).resolve()
    catalog = catalog or load_catalog(root)
    recorded = {entry["id"]: entry["descriptor_sha256"] for entry in catalog["templates"]}
    actual = {}
    for path in iter_descriptor_paths(root):
        try:
            actual[path.parent.name] = descriptor_sha256(path)
        except OSError as exc:
            actual[path.parent.name] = f"unreadable: {exc}"
    problems = []
    for template_id in sorted(set(recorded) | set(actual)):
        if template_id not in actual:
            problems.append(f"catalog lists {template_id!r} but the directory is gone")
        elif template_id not in recorded:
            problems.append(f"{template_id!r} exists on disk but is missing from the catalog")
        elif recorded[template_id] != actual[template_id]:
            problems.append(f"{template_id!r} descriptor changed since the catalog was written")
    return problems


def _normalize(text):
    return (text or "").casefold().strip()


def query(entries, school=None, tags=(), aspect=None, layout_type=None, text=None):
    school = _normalize(school)
    aspect = _normalize(aspect)
    layout_type = _normalize(layout_type)
    words = [word for word in _normalize(text).split() if word]
    wanted = [_normalize(tag) for tag in tags if _normalize(tag)]
    results = []
    for entry in entries:
        if school and school not in _normalize(entry["school"]["name"]) and school not in _normalize(entry["school"]["slug"]):
            continue
        if aspect and aspect != _normalize(entry["aspect"]):
            continue
        if layout_type and layout_type not in {_normalize(t) for t in entry["layout_types"]}:
            continue
        if wanted:
            available = [_normalize(tag) for tag in entry["tags"]]
            if not all(any(want in have for have in available) for want in wanted):
                continue
        if words:
            haystack = " ".join([entry["id"], entry["name"], entry["school"]["name"], entry["school"]["slug"],
                                 " ".join(entry["tags"]), " ".join(entry["layout_names"])]).casefold()
            if not all(word in haystack for word in words):
                continue
        results.append(entry)
    return results


def _display_width(text):
    return sum(2 if unicodedata.east_asian_width(char) in "WF" else 1 for char in text)


def _cell(text, width):
    text = str(text)
    if _display_width(text) > width:
        cut, used = [], 0
        for char in text:
            step = 2 if unicodedata.east_asian_width(char) in "WF" else 1
            if used + step > width - 1:
                break
            cut.append(char)
            used += step
        text = "".join(cut) + "…"
    return text + " " * max(0, width - _display_width(text))


def render_table(entries):
    headers = ("ID", "SCHOOL", "ASPECT", "LAYOUTS", "TAGS", "NAME")
    rows = [[entry["id"], entry["school"]["name"], entry["aspect"],
             str(len(entry["layout_names"])), ", ".join(entry["tags"]), entry["name"]] for entry in entries]
    widths = [min(max([_display_width(str(row[i])) for row in [headers, *rows]]), 34) for i in range(len(headers))]
    lines = ["  ".join(_cell(value, widths[i]) for i, value in enumerate(headers)).rstrip(),
             "  ".join("-" * widths[i] for i in range(len(headers)))]
    for row in rows:
        lines.append("  ".join(_cell(value, widths[i]) for i, value in enumerate(row)).rstrip())
    return "\n".join(lines)


def render_detail(root, entry):
    path = Path(root) / entry["path"] / DESCRIPTOR_NAME
    doc = load_descriptor(path)
    lines = [
        f"id        {doc['id']}",
        f"name      {doc['name']}",
        f"school    {doc['school']['name']} ({doc['school']['slug']})",
        f"aspect    {doc['aspect']}   size {doc['size']['width_mm']}x{doc['size']['height_mm']} mm",
        f"tags      {', '.join(doc['tags']) or '-'}",
        f"package   {(Path(root) / entry['path'] / doc['files']['package']).as_posix()}",
        f"preview   {(Path(root) / entry['preview']).as_posix()}",
        f"source    {doc['provenance']['source_filename']}  ({doc['provenance']['contributed_by']}, {doc['provenance']['license']})",
        f"ingested  {doc['provenance']['ingested_at']} by {doc['provenance']['tool']}",
        "",
        _cell("#", 4) + _cell("TYPE", 14) + _cell("PLACEHOLDERS", 13) + "NAME",
    ]
    for layout in doc["layouts"]:
        lines.append(_cell(layout["index"], 4) + _cell(layout["type"], 14)
                     + _cell(len(layout.get("placeholders", [])), 13) + layout["name"])
    if doc.get("warnings"):
        lines += ["", "warnings:"] + [f"  - {item}" for item in doc["warnings"]]
    return "\n".join(lines)


def main(skill_dir):
    parser = argparse.ArgumentParser(prog="query_template.py", description="Query the local PPT template library.")
    parser.add_argument("--root", type=Path, help="Template library root (default: templates/ beside the skills)")
    sub = parser.add_subparsers(dest="command", required=True)

    listing = sub.add_parser("list", help="List templates matching filters")
    listing.add_argument("--school", help="School name or slug substring")
    listing.add_argument("--tag", action="append", default=[], help="Required tag (repeatable, all must match)")
    listing.add_argument("--aspect", help="Exact aspect, e.g. 16:9")
    listing.add_argument("--layout", help="Layout kind the template must provide, e.g. content, two-content")
    listing.add_argument("--text", help="Free text over id/name/school/tags/layout names")
    listing.add_argument("--json", action="store_true")

    detail = sub.add_parser("show", help="Show one template and its layouts")
    detail.add_argument("id")
    detail.add_argument("--json", action="store_true")

    index = sub.add_parser("index", help="Check or rebuild the catalog")
    index.add_argument("--rebuild", action="store_true")
    index.add_argument("--check", action="store_true")

    args = parser.parse_args()
    root = template_root(skill_dir, args.root)
    if args.command == "index":
        if args.rebuild and args.check:
            raise ValueError("Choose --rebuild or --check, not both")
        if args.rebuild:
            catalog = rebuild_catalog(root)
            print(f"Rebuilt {root / CATALOG_NAME}: {len(catalog['templates'])} template(s)")
            return
        problems = catalog_drift(root)
        if problems:
            raise ValueError("Catalog is out of date: " + "; ".join(problems) + ". Run: query_template.py index --rebuild")
        print(f"Catalog is up to date: {len(load_catalog(root)['templates'])} template(s)")
        return
    catalog = load_catalog(root)
    if args.command == "show":
        matches = [entry for entry in catalog["templates"] if entry["id"] == args.id]
        if not matches:
            raise ValueError(f"No template with id {args.id!r}; run list to see available ids")
        entry = matches[0]
        print(json.dumps(entry, ensure_ascii=False, indent=2) if args.json else render_detail(root, entry))
        return
    entries = query(catalog["templates"], school=args.school, tags=args.tag, aspect=args.aspect,
                    layout_type=args.layout, text=args.text)
    if args.json:
        print(json.dumps(entries, ensure_ascii=False, indent=2))
        return
    if not entries:
        print("No matching templates.", file=sys.stderr)
        return
    print(render_table(entries))
