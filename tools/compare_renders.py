#!/usr/bin/env python3
"""Compare two directories of rendered slide PNGs pixel by pixel.

Each directory holds slides named 01.png, 02.png, ... in presentation order.
Exit code is 1 when any page exceeds the allowed difference, so the tool can
gate a visual regression check in CI or a pre-commit step.
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from PIL import Image, ImageChops


def page_files(directory):
    files = sorted(Path(directory).glob("*.png"))
    if not files:
        raise ValueError(f"No PNG files in {directory}")
    return files


def compare_pair(baseline, candidate, tolerance):
    first = Image.open(baseline).convert("RGB")
    second = Image.open(candidate).convert("RGB")
    record = {"baseline": Path(baseline).name, "candidate": Path(candidate).name,
              "size": list(first.size), "same_size": first.size == second.size}
    if first.size != second.size:
        record.update(differing_pixels=None, ratio=1.0, max_channel_delta=255, mean_abs=255.0)
        return record, None
    difference = ImageChops.difference(first, second)
    channels = difference.split()
    combined = ImageChops.lighter(ImageChops.lighter(channels[0], channels[1]), channels[2])
    histogram = combined.histogram()
    total = first.size[0] * first.size[1]
    differing = sum(histogram[tolerance + 1:])
    record.update(differing_pixels=differing, ratio=differing / total,
                  max_channel_delta=max(channel.getextrema()[1] for channel in channels),
                  mean_abs=round(sum(index * count for index, count in enumerate(combined.histogram())) / total, 4))
    return record, combined


def main():
    parser = argparse.ArgumentParser(description="Compare rendered slide PNGs pixel by pixel.")
    parser.add_argument("baseline", type=Path, help="Directory of baseline renders")
    parser.add_argument("candidate", type=Path, help="Directory of candidate renders")
    parser.add_argument("--tolerance", type=int, default=0,
                        help="Per-channel difference ignored as identical (0 = exact)")
    parser.add_argument("--max-diff-ratio", type=float, default=0.0,
                        help="Allowed fraction of differing pixels per page (default 0 = exact)")
    parser.add_argument("--diff-dir", type=Path, help="Write amplified difference images here")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    baseline_pages = page_files(args.baseline)
    candidate_pages = page_files(args.candidate)
    if len(baseline_pages) != len(candidate_pages):
        print(f"Page count differs: {len(baseline_pages)} baseline vs {len(candidate_pages)} candidate",
              file=sys.stderr)
        return 1
    if args.diff_dir:
        args.diff_dir.mkdir(parents=True, exist_ok=True)
    records, failures = [], []
    for baseline, candidate in zip(baseline_pages, candidate_pages):
        record, difference = compare_pair(baseline, candidate, args.tolerance)
        records.append(record)
        if record["ratio"] > args.max_diff_ratio and difference is not None:
            failures.append(record)
            if args.diff_dir:
                amplified = difference.point(lambda value: min(255, value * 8))
                amplified.save(args.diff_dir / f"diff-{Path(baseline).name}")
    summary = {"pages": len(records), "failures": len(failures),
               "max_diff_ratio": max((record["ratio"] or 0) for record in records),
               "records": records}
    if args.json:
        print(json.dumps(summary, ensure_ascii=False, indent=2))
    else:
        print(f"{'PAGE':6}{'DIFF PX':>12}{'RATIO':>10}{'MAXDELTA':>10}{'MEANABS':>10}  STATUS")
        for record in records:
            status = "OK" if record["ratio"] <= args.max_diff_ratio else "DIFF"
            differing = record["differing_pixels"]
            print(f"{record['baseline']:6}{differing if differing is not None else '-':>12}"
                  f"{record['ratio']:>10.6f}{record['max_channel_delta']:>10}"
                  f"{record['mean_abs']:>10.3f}  {status}")
        print(f"\n{len(records) - len(failures)}/{len(records)} page(s) within tolerance "
              f"({args.max_diff_ratio:.6f} differing pixels, channel tolerance {args.tolerance})")
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
