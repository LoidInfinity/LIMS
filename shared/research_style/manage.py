"""Validate, inspect, compare and synchronise style snapshots."""
import argparse
import difflib
import json
from pathlib import Path

from .core import (SKILLS, SUITE_ID, atomic_bytes, default_style_path, fingerprint,
                   json_bytes, load_style, read_json, validate)


def installed_targets(root):
    root = Path(root).resolve()
    targets = []
    for name in SKILLS:
        skill = root / name
        marker = skill / "assets" / "suite.json"
        if not skill.exists():
            continue
        if skill.is_symlink() or not marker.is_file() or read_json(marker).get("suite") != SUITE_ID:
            raise ValueError(f"Refusing unowned skill directory: {skill}")
        target = skill / "assets" / "styles" / "current.json"
        if not target.resolve().is_relative_to(skill.resolve()):
            raise ValueError(f"Style target escapes skill: {target}")
        targets.append(target)
    if not targets:
        raise ValueError(f"No packaged suite skills found under {root}; use tools/install.py")
    return targets


def synchronize(style, targets):
    """Atomic per-file replacement, with rollback on ordinary I/O failures.

    This is deliberately not advertised as a cross-directory crash transaction.
    Do not sync while another process is editing styles.
    """
    validate(style)
    targets = list(dict.fromkeys(Path(p).resolve() for p in targets))
    previous = {}
    for path in targets:
        if path.exists() and not path.is_file():
            raise ValueError(f"Style target is not a file: {path}")
        previous[path] = path.read_bytes() if path.exists() else None
    changed = []
    try:
        for path in targets:
            atomic_bytes(path, json_bytes(style))
            changed.append(path)
    except OSError:
        for path in reversed(changed):
            if previous[path] is None:
                path.unlink(missing_ok=True)
            else:
                atomic_bytes(path, previous[path])
        raise
    return targets


def main(skill_dir):
    p = argparse.ArgumentParser(description=__doc__)
    sub = p.add_subparsers(dest="command", required=True)
    for command in ("validate", "show"):
        item = sub.add_parser(command)
        item.add_argument("--style")
    diff = sub.add_parser("diff")
    diff.add_argument("before", type=Path)
    diff.add_argument("after", type=Path)
    activate = sub.add_parser("activate", help="Set current style and refresh explicitly selected installed snapshots")
    activate.add_argument("style", type=Path)
    activate.add_argument("--skills-dir", action="append", type=Path, default=[])
    activate.add_argument("--dry-run", action="store_true")
    check = sub.add_parser("check", help="Check installed snapshots against the selected source style")
    check.add_argument("--style")
    check.add_argument("--skills-dir", action="append", type=Path, required=True)
    a = p.parse_args()
    if a.command == "diff":
        before, after = validate(read_json(a.before)), validate(read_json(a.after))
        print("".join(difflib.unified_diff(json_bytes(before).decode().splitlines(True),
              json_bytes(after).decode().splitlines(True), fromfile=str(a.before), tofile=str(a.after))), end="")
        return
    if a.command == "activate":
        style = validate(read_json(a.style))
        current = default_style_path(skill_dir)
        roots = a.skills_dir
        # A packaged hub owns the sibling packaged skills; source checkout only owns shared/current.
        if not roots and (Path(skill_dir) / "assets" / "suite.json").is_file():
            roots = [Path(skill_dir).parent]
        targets = [current, *(path for root in roots for path in installed_targets(root))]
        targets = list(dict.fromkeys(path.resolve() for path in targets))
        if not a.dry_run:
            synchronize(style, targets)
        print(json.dumps({"dry_run": a.dry_run, "style_sha256": fingerprint(style), "targets": list(map(str, targets))}, indent=2))
        return
    style, path = load_style(skill_dir, a.style)
    if a.command == "show":
        print(json_bytes(style).decode(), end="")
    elif a.command == "validate":
        print(f"Valid: {path}\nSHA-256: {fingerprint(style)}")
    else:
        drift = []
        for root in a.skills_dir:
            for target in installed_targets(root):
                if fingerprint(validate(read_json(target))) != fingerprint(style):
                    drift.append(str(target))
        if drift:
            raise ValueError("Style drift detected: " + ", ".join(drift))
        print(f"All installed snapshots match {fingerprint(style)}")
