#!/usr/bin/env python3
"""Map source image pixels onto an Office point canvas."""

from __future__ import annotations

import argparse
import json


def positive_float(raw: str) -> float:
    try:
        value = float(raw)
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"not a number: {raw}") from exc
    if value <= 0:
        raise argparse.ArgumentTypeError("value must be positive")
    return value


def mapping(source_w: float, source_h: float, canvas_w: float, canvas_h: float) -> dict:
    sx = canvas_w / source_w
    sy = canvas_h / source_h
    scale = sx if sx < sy else sy
    fit_w = source_w * scale
    fit_h = source_h * scale
    left_pad = (canvas_w - fit_w) / 2
    top_pad = (canvas_h - fit_h) / 2

    sample = {
        "x": round(source_w / 10, 3),
        "y": round(source_h / 10, 3),
        "width": round(source_w / 4, 3),
        "height": round(source_h / 5, 3),
    }
    return {
        "input": {
            "image_width": source_w,
            "image_height": source_h,
            "target_width_pt": canvas_w,
            "target_height_pt": canvas_h,
        },
        "scale_x": sx,
        "scale_y": sy,
        "suggested_uniform_scale": scale,
        "uniform_fit": {
            "width_pt": fit_w,
            "height_pt": fit_h,
            "offset_x_pt": left_pad,
            "offset_y_pt": top_pad,
        },
        "coordinate_formulas": {
            "uniform": "x_pt = offset_x_pt + x_px * suggested_uniform_scale; y_pt = offset_y_pt + y_px * suggested_uniform_scale",
            "stretched": "x_pt = x_px * scale_x; y_pt = y_px * scale_y",
        },
        "example_source_box_px": sample,
        "example_uniform_box_pt": {
            "left_pt": round(left_pad + sample["x"] * scale, 3),
            "top_pt": round(top_pad + sample["y"] * scale, 3),
            "width_pt": round(sample["width"] * scale, 3),
            "height_pt": round(sample["height"] * scale, 3),
        },
        "example_stretched_box_pt": {
            "left_pt": round(sample["x"] * sx, 3),
            "top_pt": round(sample["y"] * sy, 3),
            "width_pt": round(sample["width"] * sx, 3),
            "height_pt": round(sample["height"] * sy, 3),
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Compute image-pixel to Office-point placement values.")
    parser.add_argument("image_width", type=positive_float)
    parser.add_argument("image_height", type=positive_float)
    parser.add_argument("target_width_pt", type=positive_float)
    parser.add_argument("target_height_pt", type=positive_float)
    parser.add_argument("--pretty", action="store_true")
    args = parser.parse_args()

    result = mapping(args.image_width, args.image_height, args.target_width_pt, args.target_height_pt)
    print(json.dumps(result, indent=2 if args.pretty else None))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
