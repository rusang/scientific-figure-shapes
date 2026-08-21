#!/usr/bin/env python3
"""Trace simple raster icons into editable polygon paths, SVG, or PPTX freeforms."""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
from typing import Any


def _load_cv2():
    try:
        import cv2
    except ImportError as exc:
        raise RuntimeError("vector_trace.py 需要 OpenCV：pip install opencv-python") from exc
    return cv2


def _write_svg(result: dict[str, Any], path: Path, fill: str) -> None:
    width, height = result["source_size_px"]
    commands = []
    for item in result["paths"]:
        points = item["points_px"]
        if len(points) < 3:
            continue
        command = "M " + " L ".join(f"{x:.3f} {y:.3f}" for x, y in points) + " Z"
        commands.append(command)
    compound = " ".join(commands)
    payload = "\n".join(
        [
            '<?xml version="1.0" encoding="UTF-8"?>',
            f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {width} {height}">',
            f'  <path id="trace_compound" d="{compound}" fill="{fill}" fill-rule="evenodd"/>',
            "</svg>",
        ]
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(payload + "\n", encoding="utf-8")


def trace_image(
    image_path: str | Path,
    *,
    threshold: int = 245,
    invert: bool = False,
    min_area: float = 12.0,
    epsilon_ratio: float = 0.01,
    svg_path: str | Path | None = None,
    svg_fill: str = "#123A9A",
) -> dict[str, Any]:
    source = Path(image_path)
    if not source.is_file():
        raise FileNotFoundError(source)
    if not 0 <= threshold <= 255:
        raise ValueError("threshold 必须位于 0..255")
    if min_area < 0 or not 0 <= epsilon_ratio <= 0.2:
        raise ValueError("min_area/epsilon_ratio 参数无效")

    cv2 = _load_cv2()
    image = cv2.imread(str(source), cv2.IMREAD_UNCHANGED)
    if image is None:
        raise ValueError(f"无法读取图像：{source}")
    if len(image.shape) == 2:
        gray = image
    elif image.shape[2] == 4:
        color = image[:, :, :3].astype("float32")
        alpha = image[:, :, 3:4].astype("float32") / 255.0
        background = 0.0 if invert else 255.0
        composited = (color * alpha + background * (1.0 - alpha)).astype("uint8")
        gray = cv2.cvtColor(composited, cv2.COLOR_BGR2GRAY)
    else:
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    threshold_mode = cv2.THRESH_BINARY if invert else cv2.THRESH_BINARY_INV
    _, mask = cv2.threshold(gray, threshold, 255, threshold_mode)
    contours, hierarchy = cv2.findContours(mask, cv2.RETR_CCOMP, cv2.CHAIN_APPROX_SIMPLE)
    hierarchy_rows = hierarchy[0] if hierarchy is not None else []
    paths: list[dict[str, Any]] = []
    for index, contour in enumerate(contours):
        area = abs(float(cv2.contourArea(contour)))
        if area < min_area:
            continue
        perimeter = float(cv2.arcLength(contour, True))
        approximation = cv2.approxPolyDP(contour, perimeter * epsilon_ratio, True)
        points = [[float(point[0][0]), float(point[0][1])] for point in approximation]
        if len(points) < 3:
            continue
        parent = int(hierarchy_rows[index][3]) if len(hierarchy_rows) else -1
        paths.append(
            {
                "id": "trace_path_pending",
                "points_px": points,
                "area_px2": round(area, 3),
                "hole": parent >= 0,
            }
        )
    paths.sort(key=lambda item: item["area_px2"], reverse=True)
    for index, item in enumerate(paths, start=1):
        item["id"] = f"trace_path_{index:03d}"
    height, width = gray.shape[:2]
    result: dict[str, Any] = {
        "passed": bool(paths),
        "source_image": str(source.resolve()),
        "source_size_px": [width, height],
        "threshold": threshold,
        "invert": invert,
        "paths": paths,
        "issues": [] if paths else [{"code": "no_traceable_contours"}],
    }
    if svg_path:
        target = Path(svg_path)
        _write_svg(result, target, svg_fill)
        result["svg_path"] = str(target.resolve())
    return result


def _load_shape_canvas():
    path = Path(__file__).with_name("office_shape_canvas.py")
    spec = importlib.util.spec_from_file_location("office_shape_canvas_for_trace", path)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    import sys

    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module.ShapeCanvas


def materialize_pptx(
    trace_result: dict[str, Any],
    output_path: str | Path,
    *,
    width_pt: float | None = None,
    height_pt: float | None = None,
    fill: tuple[int, int, int] = (18, 58, 154),
    line: tuple[int, int, int] | None = None,
    hole_fill: tuple[int, int, int] = (255, 255, 255),
) -> Path:
    source_width, source_height = trace_result["source_size_px"]
    width_pt = float(width_pt or source_width)
    height_pt = float(height_pt or source_height)
    scale_x, scale_y = width_pt / source_width, height_pt / source_height
    ShapeCanvas = _load_shape_canvas()
    canvas = ShapeCanvas(width_pt, height_pt, name_prefix="TRACE")
    with canvas.group("traced_icon"):
        for index, item in enumerate(trace_result.get("paths", []), start=1):
            points = [(x * scale_x, y * scale_y) for x, y in item["points_px"]]
            canvas.freeform(
                points,
                fill=hole_fill if item.get("hole") else fill,
                line=line,
                name=f"TRACE_E_{'hole' if item.get('hole') else 'path'}_{index:03d}",
            )
    target = Path(output_path)
    canvas.save(target)
    return target


def main() -> int:
    parser = argparse.ArgumentParser(description="Trace simple raster icons into editable vectors.")
    parser.add_argument("image")
    parser.add_argument("--threshold", type=int, default=245)
    parser.add_argument("--invert", action="store_true")
    parser.add_argument("--min-area", type=float, default=12.0)
    parser.add_argument("--epsilon-ratio", type=float, default=0.01)
    parser.add_argument("--svg-out")
    parser.add_argument("--pptx-out")
    parser.add_argument("--json-out")
    parser.add_argument("--pretty", action="store_true")
    args = parser.parse_args()
    try:
        result = trace_image(
            args.image,
            threshold=args.threshold,
            invert=args.invert,
            min_area=args.min_area,
            epsilon_ratio=args.epsilon_ratio,
            svg_path=args.svg_out,
        )
        if args.pptx_out and result["passed"]:
            materialize_pptx(result, args.pptx_out)
            result["pptx_path"] = str(Path(args.pptx_out).resolve())
    except Exception as exc:
        result = {"passed": False, "issues": [{"code": "runtime_error", "message": str(exc)}]}
    payload = json.dumps(result, ensure_ascii=False, indent=2 if args.pretty else None)
    print(payload)
    if args.json_out:
        target = Path(args.json_out)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(payload + "\n", encoding="utf-8")
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
