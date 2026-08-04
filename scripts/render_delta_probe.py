#!/usr/bin/env python3
"""Measure simple visual deltas between a source image and a rendered preview."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


def load_rgb(path: Path):
    try:
        from PIL import Image
    except ImportError as exc:
        raise SystemExit("Pillow is required for render_delta_probe.py") from exc
    return Image.open(path).convert("RGB")


def fit_size(w: int, h: int, max_w: int) -> tuple[int, int]:
    if w <= max_w:
        return w, h
    return max_w, max(1, round(h * max_w / w))


def mean_abs_delta(a, b) -> tuple[float, float]:
    pa = list(a.getdata())
    pb = list(b.getdata())
    total = 0
    worst = 0
    for x, y in zip(pa, pb):
        d = abs(x[0] - y[0]) + abs(x[1] - y[1]) + abs(x[2] - y[2])
        total += d
        if d > worst:
            worst = d
    denom = len(pa) * 3 * 255
    return total / denom, worst / (3 * 255)


def optional_ssim(a, b):
    try:
        import numpy as np
        from skimage.metrics import structural_similarity
    except Exception:
        return None
    arr_a = np.asarray(a)
    arr_b = np.asarray(b)
    score = structural_similarity(arr_a, arr_b, channel_axis=2)
    return float(score)


def main() -> int:
    parser = argparse.ArgumentParser(description="Compare a source image with a rendered reconstruction preview.")
    parser.add_argument("source_image")
    parser.add_argument("rendered_image")
    parser.add_argument("--sample-width", type=int, default=480)
    parser.add_argument("--pretty", action="store_true")
    args = parser.parse_args()

    source_path = Path(args.source_image).expanduser()
    rendered_path = Path(args.rendered_image).expanduser()
    if not source_path.is_file() or not rendered_path.is_file():
        print(json.dumps({"passed": False, "errors": ["source or rendered image is missing"]}))
        return 2

    source = load_rgb(source_path)
    rendered = load_rgb(rendered_path)
    size = fit_size(source.width, source.height, args.sample_width)
    source_small = source.resize(size)
    rendered_small = rendered.resize(size)
    mad, worst = mean_abs_delta(source_small, rendered_small)
    result = {
        "passed": True,
        "source_size": {"width": source.width, "height": source.height},
        "rendered_size": {"width": rendered.width, "height": rendered.height},
        "sample_size": {"width": size[0], "height": size[1]},
        "mean_abs_delta_0_to_1": round(mad, 6),
        "max_pixel_delta_0_to_1": round(worst, 6),
        "ssim": optional_ssim(source_small, rendered_small),
    }
    print(json.dumps(result, indent=2 if args.pretty else None))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
