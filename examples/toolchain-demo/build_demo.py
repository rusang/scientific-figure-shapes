#!/usr/bin/env python3
"""Build a small editable deck with the reusable Shapes runtime."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT / "scripts"))

from office_shape_canvas import ShapeCanvas  # noqa: E402


def build(outdir: Path) -> dict[str, str]:
    outdir.mkdir(parents=True, exist_ok=True)
    canvas = ShapeCanvas(480, 270, name_prefix="DEMO")
    canvas.rect(0, 0, 480, 270, fill=(250, 252, 255), line=None, name="DEMO_E_background")
    canvas.text(
        30,
        14,
        420,
        30,
        "Scientific Figure Shapes — Editable Demo",
        size=20,
        color=(10, 48, 125),
        bold=True,
        name="DEMO_T_title",
    )
    with canvas.group("feature_block"):
        front = canvas.rect(
            75,
            105,
            90,
            75,
            fill=(180, 210, 250),
            line=(40, 105, 220),
            name="DEMO_E_feature_front",
        )
        canvas.gradient_fill(front, (220, 238, 255), (105, 155, 230), angle=90)
        canvas.freeform(
            [(75, 105), (88, 92), (178, 92), (165, 105)],
            fill=(225, 241, 255),
            line=(40, 105, 220),
            name="DEMO_E_feature_top",
        )
        canvas.freeform(
            [(165, 105), (178, 92), (178, 167), (165, 180)],
            fill=(92, 140, 215),
            line=(40, 105, 220),
            name="DEMO_E_feature_right",
        )
        canvas.text(80, 126, 80, 28, "C5\n20×C", size=13, bold=True, name="DEMO_T_feature")
    canvas.line(
        178,
        136,
        255,
        136,
        color=(20, 20, 20),
        weight=1.4,
        arrow_end=True,
        name="DEMO_L_feature_to_fusion",
    )
    canvas.round_rect(
        255,
        112,
        110,
        48,
        fill=(223, 241, 251),
        line=(20, 135, 170),
        name="DEMO_E_fusion_box",
    )
    canvas.text(255, 112, 110, 48, "CGAFusion", size=15, bold=True, name="DEMO_T_fusion")

    pptx = outdir / "editable-demo.pptx"
    vba = outdir / "editable-demo.bas"
    scene = outdir / "scene-manifest.json"
    audit_manifest = outdir / "figure-manifest.json"
    canvas.save(pptx)
    canvas.emit_vba(vba)
    canvas.export_scene_manifest(scene)
    audit_manifest.write_text(
        json.dumps(
            {
                "schema_version": "1.0",
                "canvas": {"width_px": 960, "height_px": 540, "width_pt": 480, "height_pt": 270},
                "elements": [
                    {"id": "feature", "bbox_px": [150, 184, 206, 176], "status": "matched", "salience": "high"}
                ],
                "routing_audit": {
                    "routes": [
                        {
                            "connector": "DEMO_L_feature_to_fusion",
                            "target": "DEMO_E_fusion_box",
                            "target_edge": "left",
                            "tolerance_pt": 1.0,
                        }
                    ]
                },
                "text_audit": {
                    "items": [
                        {"name": "DEMO_T_title", "min_font_pt": 20, "align": "center", "max_lines": 1}
                    ]
                },
                "editability_audit": {
                    "allowed_raster_shapes": [],
                    "max_unregistered_raster_ratio": 0,
                    "min_editable_shape_count": 8,
                },
            },
            ensure_ascii=False,
            indent=2,
        )
        + "\n",
        encoding="utf-8",
    )
    return {
        "pptx": str(pptx),
        "vba": str(vba),
        "scene_manifest": str(scene),
        "figure_manifest": str(audit_manifest),
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--outdir", type=Path, required=True)
    args = parser.parse_args()
    print(json.dumps(build(args.outdir), ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
