"""Render a CSV XY chart with an explicit Origin, Matplotlib or plan backend."""
import argparse
import csv
from io import BytesIO
import math
import os
from pathlib import Path
import sys
import tempfile

from .core import (artifact_manifest, atomic_bytes, load_style, manifest_path, preflight_outputs, write_json)


def read_data(path, x, ys, errors=()):
    if not ys or len(set([x, *ys, *errors])) != len([x, *ys, *errors]):
        raise ValueError("X, Y and error columns must be distinct; at least one Y is required")
    if errors and len(errors) != len(ys):
        raise ValueError("Supply one --y-error column for each --y column")
    with Path(path).open(encoding="utf-8-sig", newline="") as stream:
        reader = csv.DictReader(stream)
        headers = reader.fieldnames or []
        if len(set(headers)) != len(headers):
            raise ValueError("Duplicate CSV column names")
        names = [x, *ys, *errors]
        missing = set(names) - set(headers)
        if missing:
            raise ValueError(f"Missing CSV columns: {sorted(missing)}")
        data = {key: [] for key in names}
        for row_index, row in enumerate(reader, 2):
            for name in names:
                try:
                    value = float(row[name])
                except (ValueError, TypeError) as exc:
                    raise ValueError(f"CSV row {row_index}, {name}: expected a number") from exc
                if not math.isfinite(value) or (name in errors and value < 0):
                    raise ValueError(f"CSV row {row_index}, {name}: invalid finite value/error magnitude")
                data[name].append(value)
    if not data[x]:
        raise ValueError("CSV has no data rows")
    return data


def render_matplotlib(output, data, x, ys, errors, style, xlabel, ylabel, title):
    os.environ.setdefault("MPLCONFIGDIR", str(Path(tempfile.gettempdir()) / "research-style-matplotlib"))
    import matplotlib
    matplotlib.use("Agg")
    from matplotlib import pyplot as plt
    from matplotlib import font_manager
    cfg, palette, fonts = style["chart"], style["palette"], style["typography"]
    available = {font.name for font in font_manager.fontManager.ttflist}
    warnings = [f"Font unavailable; renderer may substitute: {font}" for font in fonts.values() if font not in available]
    with plt.rc_context({"font.family": [fonts["latin"], fonts["cjk"], "DejaVu Sans"],
                         "font.size": cfg["font_pt"], "text.color": palette["text"],
                         "axes.labelcolor": palette["text"], "axes.edgecolor": palette["text"],
                         "xtick.color": palette["text"], "ytick.color": palette["text"],
                         "svg.fonttype": "none", "pdf.fonttype": 42}):
        fig, ax = plt.subplots(figsize=(cfg["width_mm"] / 25.4, cfg["height_mm"] / 25.4), layout="constrained")
        fig.set_facecolor(palette["background"])
        ax.set_facecolor(palette["background"])
        marker = {"none": "", "circle": "o", "square": "s", "triangle": "^"}[cfg["marker"]]
        for index, name in enumerate(ys):
            ax.errorbar(data[x], data[name], yerr=data[errors[index]] if errors else None,
                        color=palette["series"][index % len(palette["series"])],
                        linewidth=cfg["line_width_pt"], linestyle=cfg["line_style"],
                        marker=marker, markersize=cfg["marker_size_pt"], label=name,
                        capsize=3 if errors else 0)
        ax.set(xlabel=xlabel, ylabel=ylabel, title=title)
        ax.grid(cfg["grid"], color=palette["muted"], alpha=.2) if cfg["grid"] else ax.grid(False)
        ax.spines[["top", "right"]].set_visible(False)
        if len(ys) > 1:
            ax.legend(frameon=False)
        try:
            stream = BytesIO()
            fig.savefig(stream, format=output.suffix.lower()[1:], dpi=cfg["dpi"])
            atomic_bytes(output, stream.getvalue())
        finally:
            plt.close(fig)
    return warnings


def render_origin(output, data, x, ys, errors, style, template=None):
    if sys.platform != "win32":
        raise ValueError("Origin backend requires Windows and a licensed local Origin installation")
    try:
        import originpro as op
    except ImportError as exc:
        raise ValueError("Origin backend requires originpro and Origin 2021+. Use --backend matplotlib explicitly for portable output") from exc
    if not op.oext:
        raise ValueError("Run this adapter in external Python, not an embedded/live Origin session")
    cfg = style["chart"]
    # Only documented Plot properties are automated; the supplied template owns layout.
    try:
        sheet = op.new_sheet("w", lname="Research data")
        columns = [x, *ys, *errors]
        for i, name in enumerate(columns):
            sheet.from_list(i, data[name], lname=name)
        graph = op.new_graph(template=str(Path(template).resolve()) if template else "")
        layer = graph[0]
        for i, name in enumerate(ys):
            plot = layer.add_plot(sheet, i + 1, 0, type=200 if cfg["marker"] == "none" else 202,
                                  colyerr=1 + len(ys) + i if errors else -1)
            plot.color = style["palette"]["series"][i % len(style["palette"]["series"])]
            # LabTalk -w uses 1/500 pt units (not points).
            plot.set_cmd(f"-w {cfg['line_width_pt'] * 500}",
                         f"-d {dict(solid=0, dashed=1, dotted=2, dashdot=3)[cfg['line_style']]}")
            if cfg["marker"] != "none":
                plot.symbol_size = cfg["marker_size_pt"]
                plot.symbol_kind = {"circle": 2, "square": 1, "triangle": 3}[cfg["marker"]]
        layer.rescale()
        if not op.save(str(output.resolve())):
            raise ValueError("Origin could not save the OPJU project")
    finally:
        op.exit()
    return ["Origin adapter applies series color, line style/width and marker kind/size. "
            "Page size, fonts, axes, grid, background, title and export DPI remain template-controlled. "
            "Inspect the OPJU in Origin before publication; native backend has not been validated on this build machine."]


def main(skill_dir):
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("data", type=Path)
    p.add_argument("--x", required=True)
    p.add_argument("--y", nargs="+", required=True)
    p.add_argument("--y-error", nargs="+", default=[])
    p.add_argument("--style")
    p.add_argument("--backend", choices=["origin", "matplotlib", "plan"], default="origin")
    p.add_argument("--output", type=Path, required=True)
    p.add_argument("--template", type=Path)
    p.add_argument("--xlabel")
    p.add_argument("--ylabel", default="Value")
    p.add_argument("--title", default="")
    p.add_argument("--force", action="store_true")
    a = p.parse_args()
    style, source = load_style(skill_dir, a.style)
    extensions = {"origin": {".opju"}, "matplotlib": {".png", ".svg", ".pdf"}, "plan": {".json"}}
    if a.output.suffix.lower() not in extensions[a.backend]:
        raise ValueError(f"{a.backend} output extension must be one of {sorted(extensions[a.backend])}")
    if a.template and (a.backend != "origin" or not a.template.is_file()):
        raise ValueError("--template requires Origin and an existing template file")
    data = read_data(a.data, a.x, a.y, a.y_error)
    preflight_outputs([a.output, manifest_path(a.output)], a.force, [a.data, source, a.template])
    a.output.parent.mkdir(parents=True, exist_ok=True)
    warnings = []
    if a.backend == "plan":
        write_json(a.output, {"style": style, "data": data, "x": a.x, "y": a.y, "y_error": a.y_error,
                              "xlabel": a.xlabel or a.x, "ylabel": a.ylabel, "title": a.title})
    elif a.backend == "matplotlib":
        warnings = render_matplotlib(a.output, data, a.x, a.y, a.y_error, style, a.xlabel or a.x, a.ylabel, a.title)
    else:
        warnings = render_origin(a.output, data, a.x, a.y, a.y_error, style, a.template)
    write_json(manifest_path(a.output), artifact_manifest(style, source, a.backend, warnings=warnings,
               rows=len(data[a.x]), x=a.x, y=a.y, output=str(a.output),
               applied_scope="series-only; layout from template" if a.backend == "origin" else "chart, typography, palette"))
    for warning in warnings:
        print(f"Warning: {warning}", file=sys.stderr)
    print(a.output)
