#!/usr/bin/env python3
"""Create local crops for raster-preserved figure regions."""

from __future__ import annotations

import argparse
import json
import re
from pathlib import Path


def import_pillow():
    try:
        from PIL import Image
    except ImportError as exc:
        raise SystemExit("Pillow is required for preserve_cropper.py") from exc
    return Image


def clean_filename(text: str) -> str:
    name = re.sub(r"[^A-Za-z0-9_.-]+", "_", text.strip()).strip("._-")
    return name or "crop"


def parse_region(text: str) -> dict:
    fields = text.split(":")
    if len(fields) != 5:
        raise argparse.ArgumentTypeError("region must be name:x:y:w:h")
    name, x, y, w, h = fields
    try:
        box = {"name": name, "x": float(x), "y": float(y), "width": float(w), "height": float(h)}
    except ValueError as exc:
        raise argparse.ArgumentTypeError(f"bad numeric value in {text!r}") from exc
    if box["width"] <= 0 or box["height"] <= 0:
        raise argparse.ArgumentTypeError("region width and height must be positive")
    return box


def regions_from_file(path: Path) -> list[dict]:
    data = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(data, list):
        raise ValueError("regions JSON must be a list")
    out: list[dict] = []
    for item in data:
        if not isinstance(item, dict):
            raise ValueError("each region must be an object")
        for key in ("name", "x", "y", "width", "height"):
            if key not in item:
                raise ValueError(f"region missing {key}")
        out.append(
            {
                "name": str(item["name"]),
                "x": float(item["x"]),
                "y": float(item["y"]),
                "width": float(item["width"]),
                "height": float(item["height"]),
                "reason": str(item.get("reason", "")),
            }
        )
    return out


def crop_box(region: dict, image_w: int, image_h: int, padding: float) -> tuple[int, int, int, int]:
    left = round(region["x"] - padding)
    top = round(region["y"] - padding)
    right = round(region["x"] + region["width"] + padding)
    bottom = round(region["y"] + region["height"] + padding)
    left = max(0, min(image_w, int(left)))
    right = max(0, min(image_w, int(right)))
    top = max(0, min(image_h, int(top)))
    bottom = max(0, min(image_h, int(bottom)))
    if right <= left or bottom <= top:
        raise ValueError(f"region {region['name']} is outside the source image")
    return left, top, right, bottom


def main() -> int:
    parser = argparse.ArgumentParser(description="Crop preserved raster regions from a source image.")
    parser.add_argument("source_image")
    parser.add_argument("output_dir")
    parser.add_argument("--region", action="append", type=parse_region, default=[])
    parser.add_argument("--regions-json")
    parser.add_argument("--padding", type=float, default=0.0)
    parser.add_argument("--manifest")
    parser.add_argument("--pretty", action="store_true")
    args = parser.parse_args()

    source = Path(args.source_image).expanduser()
    if not source.is_file():
        print(json.dumps({"passed": False, "errors": [f"missing source image: {source}"]}))
        return 2

    regions = list(args.region)
    if args.regions_json:
        try:
            regions.extend(regions_from_file(Path(args.regions_json).expanduser()))
        except Exception as exc:
            print(json.dumps({"passed": False, "errors": [str(exc)]}, indent=2 if args.pretty else None))
            return 2
    if not regions:
        print(json.dumps({"passed": False, "errors": ["no crop regions supplied"]}))
        return 2

    Image = import_pillow()
    outdir = Path(args.output_dir).expanduser()
    outdir.mkdir(parents=True, exist_ok=True)
    image = Image.open(source).convert("RGBA")
    width, height = image.size
    assets = []

    try:
        for index, region in enumerate(regions, start=1):
            left, top, right, bottom = crop_box(region, width, height, args.padding)
            crop = image.crop((left, top, right, bottom))
            filename = f"{index:02d}_{clean_filename(region['name'])}.png"
            asset_path = outdir / filename
            crop.save(asset_path)
            assets.append(
                {
                    "name": region["name"],
                    "reason": region.get("reason", ""),
                    "asset_path": str(asset_path.resolve()),
                    "source_box_px": {"x": left, "y": top, "width": right - left, "height": bottom - top},
                }
            )
    except Exception as exc:
        print(json.dumps({"passed": False, "errors": [str(exc)]}, indent=2 if args.pretty else None))
        return 2

    result = {
        "passed": True,
        "source_image": str(source.resolve()),
        "source_size_px": {"width": width, "height": height},
        "output_dir": str(outdir.resolve()),
        "assets": assets,
    }
    if args.manifest:
        manifest = Path(args.manifest).expanduser()
        manifest.parent.mkdir(parents=True, exist_ok=True)
        manifest.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
        result["manifest_path"] = str(manifest.resolve())
    print(json.dumps(result, indent=2 if args.pretty else None, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
