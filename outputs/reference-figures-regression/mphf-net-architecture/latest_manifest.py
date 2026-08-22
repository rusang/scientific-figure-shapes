from __future__ import annotations

import json
from pathlib import Path


def write_manifest(
    here: Path,
    orthogonal_routes: dict,
    cuboid_geometry: dict,
) -> dict:
    routing = json.loads((here / "routing_manifest.json").read_text(encoding="utf-8"))
    routing["routing_audit"]["routes"] = [
        route for route in routing["routing_audit"]["routes"]
        if route.get("connector") not in {
            "SUMMER_L_backbone_neck", "SUMMER_L_neck_head"
        }
    ]
    style_segments = []
    for route_name, route in orthogonal_routes.items():
        for index, ((start_x, start_y), (end_x, end_y)) in enumerate(
            zip(route["points"], route["points"][1:]), start=1
        ):
            orientation = "horizontal" if start_y == end_y else "vertical"
            style_segments.append({
                "connector": f"{route_name}_seg{index}",
                "dash": route["dash"],
                "orientation": orientation,
                "tolerance_pt": 0.2,
            })
    routing["routing_audit"]["segments"] = style_segments
    layering = json.loads((here / "layering_manifest.json").read_text(encoding="utf-8"))
    layering["layering_audit"]["overlap_pairs"] = []
    for spec in layering["layering_audit"]["cuboids"]:
        prefix = spec["front"].removesuffix("_front")
        geometry = cuboid_geometry.get(prefix)
        if geometry:
            spec.pop("expected_face_bounds_pt", None)
            spec["expected_face_vertices_pt"] = geometry["vertices"]
            spec["vertex_tolerance_pt"] = 0.8
            spec["expected_gradients"] = {"front": geometry["front_gradient"]}
    unified = {
        "schema_version": "1.0",
        "canvas": {"width_px": 1536, "height_px": 1024,
                   "width_pt": 960, "height_pt": 640},
        "elements": [
            {"id": "input_pcb", "label": "PCB input image",
             "bbox_px": [25, 162, 171, 244], "status": "preserved",
             "salience": "high"},
            {"id": "backbone_pyramid", "bbox_px": [230, 145, 240, 335],
             "status": "matched", "salience": "normal"},
            {"id": "neck", "bbox_px": [498, 76, 574, 460],
             "status": "matched", "salience": "normal"},
            {"id": "head", "bbox_px": [1086, 76, 240, 460],
             "status": "matched", "salience": "normal"},
            {"id": "advantages", "bbox_px": [1148, 852, 378, 164],
             "status": "matched", "salience": "normal"}
        ],
        **routing,
        **layering,
        "text_audit": {
            "items": [
                {"name": "SUMMER_T_banner", "contains": "MPHF-Net",
                 "min_font_pt": 17, "align": "center", "max_lines": 1},
                {"name": "SUMMER_T_backbone_title", "contains": "CSP-MEEM",
                 "min_font_pt": 11, "max_lines": 1},
                {"name": "SUMMER_T_neck_title", "contains": "FPN-PAN",
                 "min_font_pt": 11, "max_lines": 1},
                {"name": "SUMMER_T_head_title", "contains": "RT-DETR",
                 "min_font_pt": 10, "max_lines": 1,
                 "estimate_overflow": True},
                {"name": "SUMMER_T_m1_title", "contains": "CSP-MEEM",
                 "min_font_pt": 10, "max_lines": 1},
                {"name": "SUMMER_T_m2_title", "contains": "LOSC",
                 "min_font_pt": 10, "max_lines": 1},
                {"name": "SUMMER_T_m3_title", "contains": "CGAFusion",
                 "min_font_pt": 10, "max_lines": 1},
                {"name": "SUMMER_T_b4_title", "text": "三模块协同优势",
                 "min_font_pt": 10, "max_lines": 1}
            ]
        },
        "editability_audit": {
            "allowed_raster_shapes": [
                "SUMMER_R1_pcb",
                "SUMMER_E_b4_icon0",
                "SUMMER_E_b4_icon1",
                "SUMMER_E_b4_icon2",
            ],
            "reference_slide_shapes": ["SUMMER_R_reference_full"],
            "max_unregistered_raster_ratio": 0,
            "max_preserved_raster_ratio": 0.04,
            "min_editable_shape_count": 280
        }
    }
    (here / "figure-manifest.json").write_text(
        json.dumps(unified, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return unified
