#!/usr/bin/env python3
"""Create distributable, independently usable skill ZIP files."""
import argparse
from io import BytesIO
from pathlib import Path
import tempfile
from zipfile import ZipFile, ZIP_DEFLATED

from install import ROOT, SKILLS, install, main_guard
from research_style.core import atomic_bytes, preflight_outputs


def main():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--output-dir", type=Path, default=ROOT / "dist")
    p.add_argument("--force", action="store_true")
    a = p.parse_args()
    a.output_dir.mkdir(parents=True, exist_ok=True)
    outputs = [a.output_dir / f"{name}.zip" for name in (*SKILLS, "research-style-suite")]
    preflight_outputs(outputs, a.force)
    with tempfile.TemporaryDirectory(prefix="research-style-package-") as directory:
        root = Path(directory)
        install(root)
        for name, output in zip((*SKILLS, "research-style-suite"), outputs):
            buffer = BytesIO()
            archive_root = root if name == "research-style-suite" else root / name
            with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
                for file in sorted(archive_root.rglob("*")):
                    if file.is_file():
                        archive.write(file, file.relative_to(root))
            atomic_bytes(output, buffer.getvalue())
            print(output)


if __name__ == "__main__":
    main_guard(main)
