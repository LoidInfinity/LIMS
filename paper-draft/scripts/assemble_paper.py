#!/usr/bin/env python3
"""Research Style Skill Suite entrypoint; works in source and packaged installations."""
from pathlib import Path
import sys

SKILL = Path(__file__).resolve().parents[1]
for runtime in (SKILL / "assets" / "runtime", SKILL.parent / "shared"):
    if (runtime / "research_style" / "__init__.py").is_file():
        sys.path.insert(0, str(runtime))
        break
else:
    raise SystemExit("Missing shared runtime. Install with tools/install.py; raw individual skill copies are incomplete.")

from research_style.core import main_guard
from research_style.paper import main

if __name__ == "__main__":
    main_guard(lambda: main(SKILL))
