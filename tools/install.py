#!/usr/bin/env python3
"""Install portable skill bundles without symlinks or dependencies outside each skill."""
import argparse
import os
from pathlib import Path
import shutil
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "shared"))
from research_style.core import SKILLS, SUITE_ID, load_style, main_guard, write_json


def install(destination, names=SKILLS):
    destination = Path(destination).expanduser().resolve()
    for name in SKILLS:
        if destination == ROOT / name or destination.is_relative_to(ROOT / name):
            raise ValueError("Install destination cannot be inside a source skill")
    if destination == ROOT or destination.is_relative_to(ROOT / "shared"):
        raise ValueError("Install destination must not overlap suite sources")
    names = tuple(dict.fromkeys(names))
    if not names or set(names) - set(SKILLS):
        raise ValueError("Unknown or empty skill selection")
    for name in names:
        if (destination / name).exists() or (destination / name).is_symlink():
            raise ValueError(f"Destination exists: {destination / name}. Choose a new directory; use sync_style.py for style changes")
    style, _ = load_style(ROOT / "style-hub")
    destination.mkdir(parents=True, exist_ok=True)
    moved = []
    with tempfile.TemporaryDirectory(prefix=".research-style-", dir=destination) as temp:
        staging = Path(temp)
        for name in names:
            target = staging / name
            shutil.copytree(ROOT / name, target, ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
            shutil.copytree(ROOT / "shared" / "research_style", target / "assets" / "runtime" / "research_style",
                            ignore=shutil.ignore_patterns("__pycache__", "*.pyc"))
            shutil.copytree(ROOT / "shared" / "styles", target / "assets" / "styles")
            shutil.copy2(ROOT / "requirements.txt", target / "assets" / "requirements.txt")
            write_json(target / "assets" / "styles" / "current.json", style)
            write_json(target / "assets" / "suite.json", {"suite": SUITE_ID, "version": "0.1.0", "skill": name})
        try:
            for name in names:
                os.rename(staging / name, destination / name)
                moved.append(name)
        except OSError:
            for name in reversed(moved):
                os.rename(destination / name, staging / name)
            raise
    return [destination / name for name in names]


def main():
    p = argparse.ArgumentParser(description=__doc__)
    group = p.add_mutually_exclusive_group(required=True)
    group.add_argument("--dest", type=Path)
    group.add_argument("--host", choices=["codex", "claude", "dsh"])
    p.add_argument("--scope", choices=["project", "user"], default="project")
    p.add_argument("--project", type=Path, default=Path.cwd())
    p.add_argument("--skills", nargs="+", choices=SKILLS, default=list(SKILLS))
    p.add_argument("--dry-run", action="store_true")
    a = p.parse_args()
    folder = {"codex": ".agents", "claude": ".claude", "dsh": ".dsh"}
    if a.dest:
        destination = a.dest
    elif a.host == "dsh" and a.scope == "user" and os.environ.get("DSH_HOME"):
        destination = Path(os.environ["DSH_HOME"]) / "skills"
    else:
        destination = (Path.home() if a.scope == "user" else a.project) / folder[a.host] / "skills"
    if a.dry_run:
        print("\n".join(str(destination.expanduser().resolve() / name) for name in a.skills))
    else:
        for path in install(destination, a.skills):
            print(path)


if __name__ == "__main__":
    main_guard(main)
