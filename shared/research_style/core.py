"""Validated style loading, deterministic fingerprints and safe output writes."""
from __future__ import annotations

import hashlib
import json
import math
import os
from pathlib import Path
import tempfile

SKILLS = ("style-hub", "origin-chart", "ppt-deck", "paper-draft")
SUITE_ID = "research-style-suite"


def read_json(path):
    def reject(value):
        raise ValueError(f"Non-finite JSON number: {value}")
    def unique(pairs):
        result = {}
        for key, value in pairs:
            if key in result:
                raise ValueError(f"Duplicate JSON key: {key}")
            result[key] = value
        return result
    return json.loads(Path(path).read_text(encoding="utf-8-sig"),
                      parse_constant=reject, object_pairs_hook=unique)


def json_bytes(value):
    return (json.dumps(value, ensure_ascii=False, indent=2, allow_nan=False) + "\n").encode("utf-8")


def fingerprint(style):
    payload = json.dumps(style, sort_keys=True, ensure_ascii=False,
                         separators=(",", ":"), allow_nan=False).encode("utf-8")
    return hashlib.sha256(payload).hexdigest()


def validate(style):
    try:
        from jsonschema import Draft202012Validator
    except ImportError as exc:
        raise ValueError("Missing jsonschema; install requirements.txt in the active Python environment") from exc
    schema = read_json(Path(__file__).with_name("style.schema.json"))
    errors = sorted(Draft202012Validator(schema).iter_errors(style), key=lambda e: str(e.path))
    if errors:
        message = "; ".join(f"{'.'.join(map(str, e.path)) or '<root>'}: {e.message}" for e in errors[:8])
        raise ValueError(f"Invalid style: {message}")
    # JSON Schema treats Python float infinity as a number; reject it explicitly.
    def finite(value):
        if isinstance(value, float) and not math.isfinite(value):
            raise ValueError("Style contains a non-finite number")
        if isinstance(value, dict):
            for item in value.values():
                finite(item)
        elif isinstance(value, list):
            for item in value:
                finite(item)
    finite(style)
    for section, margin_name in (("slides", "margin_mm"), ("paper", "margins_mm")):
        cfg = style[section]
        margins = cfg[margin_name]
        if isinstance(margins, (float, int)):
            width_margin = height_margin = 2 * margins
        else:
            width_margin = margins["left"] + margins["right"]
            height_margin = margins["top"] + margins["bottom"]
        if cfg["width_mm"] - width_margin < 40 or cfg["height_mm"] - height_margin < 40:
            raise ValueError(f"{section}: margins leave less than 40 mm of usable space")
    return style


def default_style_path(skill_dir):
    skill_dir = Path(skill_dir).resolve()
    suite = skill_dir.parent / "shared" / "styles" / "current.json"
    bundled = skill_dir / "assets" / "styles" / "current.json"
    if suite.is_file():
        return suite
    if bundled.is_file():
        return bundled
    raise ValueError("Missing style assets. Install with tools/install.py, or pass --style FILE")


def load_style(skill_dir, explicit=None):
    selected = explicit or os.environ.get("RESEARCH_STYLE_FILE")
    path = Path(selected).expanduser().resolve() if selected else default_style_path(skill_dir)
    return validate(read_json(path)), path


def preflight_outputs(paths, force=False, inputs=()):
    if any(Path(p).expanduser().is_symlink() for p in paths):
        raise ValueError("Refusing an output file symlink")
    resolved = [Path(p).expanduser().resolve() for p in paths]
    if len(resolved) != len(set(resolved)):
        raise ValueError("Output paths must be distinct")
    sources = {Path(p).resolve() for p in inputs if p is not None}
    for path in resolved:
        if path in sources:
            raise ValueError(f"Output would overwrite an input: {path}")
        if path.exists() and not path.is_file():
            raise ValueError(f"Output must be a regular file: {path}")
        if path.exists() and not force:
            raise ValueError(f"Output exists: {path}. Use --force to replace it")
    return resolved


def atomic_bytes(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=f".{path.name}.", dir=path.parent)
    try:
        with os.fdopen(fd, "wb") as stream:
            stream.write(data)
        os.replace(temporary, path)
    finally:
        Path(temporary).unlink(missing_ok=True)


def write_json(path, value):
    atomic_bytes(path, json_bytes(value))


def artifact_manifest(style, style_path, backend, **extra):
    return {"suite": SUITE_ID, "style_id": style["id"], "style_sha256": fingerprint(style),
            "style_source": str(style_path), "backend": backend, **extra}


def manifest_path(output):
    return Path(str(output) + ".style.json")


def main_guard(main):
    try:
        main()
    except (ValueError, OSError, ImportError) as exc:
        raise SystemExit(f"Error: {exc}") from None
