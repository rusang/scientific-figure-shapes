#!/usr/bin/env python3
"""Compare two renderer outputs to detect PowerPoint/LibreOffice drift."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


def _load_rgb(path: Path):
    try:
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError("dual_renderer_probe.py 需要 Pillow：pip install Pillow") from exc
    return Image.open(path).convert("RGB")


def _ssim(first, second) -> float | None:
    try:
        import numpy as np
        from skimage.metrics import structural_similarity
    except Exception:
        return None
    return float(
        structural_similarity(np.asarray(first), np.asarray(second), channel_axis=2)
    )


def compare_renderers(
    image_a: str | Path,
    image_b: str | Path,
    *,
    renderer_a: str = "renderer_a",
    renderer_b: str = "renderer_b",
    max_mean_delta: float = 0.015,
    min_ssim: float | None = None,
    heatmap_path: str | Path | None = None,
) -> dict[str, Any]:
    from PIL import ImageChops, ImageEnhance, ImageOps, ImageStat

    first_path, second_path = Path(image_a), Path(image_b)
    if not first_path.is_file() or not second_path.is_file():
        raise FileNotFoundError("renderer output is missing")
    if not 0 <= max_mean_delta <= 1:
        raise ValueError("max_mean_delta 必须位于 0..1")
    first = _load_rgb(first_path)
    second_original = _load_rgb(second_path)
    size_match = first.size == second_original.size
    second = second_original.resize(first.size)
    difference = ImageChops.difference(first, second)
    stat = ImageStat.Stat(difference)
    mean_delta = sum(stat.mean) / (3 * 255)
    ssim = _ssim(first, second)
    issues: list[dict[str, Any]] = []
    if not size_match:
        issues.append(
            {
                "code": "renderer_size_mismatch",
                "size_a": list(first.size),
                "size_b": list(second_original.size),
            }
        )
    if mean_delta > max_mean_delta:
        issues.append(
            {
                "code": "renderer_visual_delta",
                "maximum": max_mean_delta,
                "actual": round(mean_delta, 6),
            }
        )
    if min_ssim is not None and ssim is not None and ssim < min_ssim:
        issues.append(
            {"code": "renderer_ssim_too_low", "minimum": min_ssim, "actual": round(ssim, 6)}
        )
    if heatmap_path:
        heat = ImageOps.colorize(
            ImageEnhance.Contrast(difference.convert("L")).enhance(3.0),
            black=(255, 255, 255),
            white=(220, 0, 0),
        )
        target = Path(heatmap_path)
        target.parent.mkdir(parents=True, exist_ok=True)
        heat.save(target)
    return {
        "passed": not issues,
        "verdict": "PASS" if not issues else "FAIL",
        "renderer_a": {"name": renderer_a, "image": str(first_path.resolve()), "size_px": list(first.size)},
        "renderer_b": {"name": renderer_b, "image": str(second_path.resolve()), "size_px": list(second_original.size)},
        "mean_abs_delta": round(mean_delta, 6),
        "ssim": round(ssim, 6) if ssim is not None else None,
        "heatmap_path": str(Path(heatmap_path).resolve()) if heatmap_path else None,
        "issues": issues,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Compare two PPTX renderer PNG outputs.")
    parser.add_argument("image_a")
    parser.add_argument("image_b")
    parser.add_argument("--renderer-a", default="PowerPoint")
    parser.add_argument("--renderer-b", default="LibreOffice")
    parser.add_argument("--max-mean-delta", type=float, default=0.015)
    parser.add_argument("--min-ssim", type=float)
    parser.add_argument("--heatmap")
    parser.add_argument("--json-out")
    parser.add_argument("--pretty", action="store_true")
    args = parser.parse_args()
    try:
        result = compare_renderers(
            args.image_a,
            args.image_b,
            renderer_a=args.renderer_a,
            renderer_b=args.renderer_b,
            max_mean_delta=args.max_mean_delta,
            min_ssim=args.min_ssim,
            heatmap_path=args.heatmap,
        )
    except Exception as exc:
        result = {"passed": False, "verdict": "FAIL", "issues": [{"code": "runtime_error", "message": str(exc)}]}
    payload = json.dumps(result, ensure_ascii=False, indent=2 if args.pretty else None)
    print(payload)
    if args.json_out:
        target = Path(args.json_out)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(payload + "\n", encoding="utf-8")
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
