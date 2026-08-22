#!/usr/bin/env python3
"""Run the MPHF-Net scene through the latest scientific-figure-shapes runtime."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path

from pptx.enum.dml import MSO_LINE_DASH_STYLE


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]
REFERENCE = ROOT / "assets/reference-figures/mphf-net-architecture.png"
LEGACY_BUILDER = ROOT / "examples/mphf-net/build_mphf_deck.py"
sys.path.insert(0, str(ROOT / "scripts"))

from office_shape_canvas import ShapeCanvas  # noqa: E402


def load_legacy():
    spec = importlib.util.spec_from_file_location("mphf_scene_definition", LEGACY_BUILDER)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


legacy = load_legacy()
SCALE = legacy.SCALE
CUBOID_GEOMETRY: dict[str, dict] = {}
from latest_scene_config import (  # noqa: E402
    ICON_TRACE_FILES,
    LINE_ENDPOINT_OVERRIDES,
    ORTHOGONAL_ROUTES,
    RECT_OVERRIDES,
)
from latest_manifest import write_manifest  # noqa: E402
from latest_proxies import SlideProxy, SlidesProxy  # noqa: E402

def pt(value: float) -> float:
    return float(value) * SCALE


def color(value):
    return None if value is None else tuple(int(channel) for channel in value)


def line_dash(value) -> bool:
    return value is not None and value is not False


class PresentationProxy:
    def __init__(self) -> None:
        self.canvas = ShapeCanvas(1536 * SCALE, 1024 * SCALE, name_prefix="SUMMER")
        self.slide = SlideProxy(self.canvas)
        self.slides = SlidesProxy(self.slide)
        self.slide_layouts = [None] * 7
        self.slide_width = None
        self.slide_height = None

    def save(self, path) -> None:
        self.canvas.add_reference_slide(REFERENCE, name="SUMMER_R_reference_full")
        target = Path(path)
        self.canvas.save(target)
        self.canvas.emit_vba(target.with_suffix(".bas"))
        self.canvas.export_scene_manifest(HERE / "scene-manifest.json")


def rect_adapter(slide, name, x, y, w, h, *, fill=None, edge=None, edge_w=1.0,
                 shape=None, radius=0.08, dash=None):
    canvas = slide.canvas
    if name in RECT_OVERRIDES:
        x, y, w, h = RECT_OVERRIDES[name]
    if name == "SUMMER_B_banner":
        fill = legacy.RGBColor(1, 23, 73)
    if name == "SUMMER_E_neck_bu":
        fill = legacy.RGBColor(226, 244, 230)
        edge = legacy.GREEN
        edge_w = 0.8
    if name == "SUMMER_E_neck_td":
        fill = legacy.RGBColor(218, 239, 210)
        edge = None
    if name == "SUMMER_E_neck_bu":
        fill = legacy.RGBColor(218, 239, 210)
        edge = None
    if name == "SUMMER_E_neck_note_cga":
        fill = legacy.RGBColor(224, 236, 252)
    if name == "SUMMER_E_neck_note_losc":
        fill = legacy.RGBColor(252, 230, 220)
    if name.startswith("SUMMER_E_panel_"):
        fill = legacy.RGBColor(253, 253, 253)
        edge = legacy.RGBColor(192, 203, 224)
        edge_w = 0.8
    if name == "SUMMER_E_panel_m2":
        edge = legacy.RGBColor(255, 105, 55)
    m1_styles = {
        "SUMMER_E_m1_csp": (legacy.RGBColor(228, 238, 253), legacy.BLUE),
        "SUMMER_E_m1_b0": (legacy.RGBColor(232, 245, 225), legacy.GREEN),
        "SUMMER_E_m1_b1": (legacy.RGBColor(232, 245, 225), legacy.GREEN),
        "SUMMER_E_m1_b2": (legacy.RGBColor(232, 245, 225), legacy.GREEN),
        "SUMMER_E_m1_fuse": (legacy.RGBColor(226, 237, 252), legacy.BLUE),
        "SUMMER_E_m1_ema": (legacy.RGBColor(255, 237, 215), legacy.RGBColor(230, 125, 35)),
    }
    if name in m1_styles:
        fill, edge = m1_styles[name]
        edge_w = 0.8
    if name in {"SUMMER_E_head_iou", "SUMMER_E_head_dec", "SUMMER_E_head_out"}:
        fill = legacy.RGBColor(255, 255, 255)
        edge = legacy.RGBColor(205, 196, 235)
        edge_w = 0.8
    stage_styles = {
        "SUMMER_E_bb_c5": (legacy.RGBColor(245, 248, 255), legacy.BLUE),
        "SUMMER_E_bb_c4": (legacy.RGBColor(245, 252, 242), legacy.GREEN),
        "SUMMER_E_bb_c3": (legacy.RGBColor(255, 248, 243), legacy.RED),
    }
    if name in stage_styles:
        fill, edge = stage_styles[name]
        edge_w = 1.0
    neck_stage_styles = {
        "SUMMER_E_neck_c4": (legacy.RGBColor(238, 249, 233), legacy.GREEN),
        "SUMMER_E_neck_c3": (legacy.RGBColor(255, 239, 224), legacy.RED),
    }
    if name in neck_stage_styles:
        fill, edge = neck_stage_styles[name]
        edge_w = 0.9
    shape = shape or legacy.MSO_SHAPE.ROUNDED_RECTANGLE
    kwargs = {
        "fill": color(fill),
        "line": color(edge),
        "weight": float(edge_w),
        "name": name,
    }
    if shape == legacy.MSO_SHAPE.ROUNDED_RECTANGLE:
        result = canvas.round_rect(pt(x), pt(y), pt(w), pt(h), **kwargs)
        result.adjustments[0] = radius
        if name == "SUMMER_E_head_bbox":
            result.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        if name == "SUMMER_E_b2_last":
            line_adapter(slide, "SUMMER_L_b2_last_top", 835, 904, 835, 918,
                         color=legacy.PURPLE, width=0.7)
            line_adapter(slide, "SUMMER_L_b2_last_bottom", 835, 950, 835, 964,
                         color=legacy.PURPLE, width=0.7)
        if name == "SUMMER_E_m3_high":
            canvas.line(
                pt(1002), pt(640), pt(1062), pt(640),
                color=color_tuple(legacy.INK), weight=0.8,
                arrow_end=True, name="SUMMER_L_m3_input_high",
            )
        if name == "SUMMER_E_m3_low":
            canvas.line(
                pt(1002), pt(724), pt(1062), pt(724),
                color=color_tuple(legacy.INK), weight=0.8,
                arrow_end=True, name="SUMMER_L_m3_input_low",
            )
        return result
    if shape == legacy.MSO_SHAPE.RECTANGLE:
        result = canvas.rect(pt(x), pt(y), pt(w), pt(h), **kwargs)
        if name == "SUMMER_E_head_bbox":
            result.line.dash_style = MSO_LINE_DASH_STYLE.DASH
        if name == "SUMMER_E_b2_last":
            line_adapter(slide, "SUMMER_L_b2_last_top", 835, 904, 835, 918,
                         color=legacy.PURPLE, width=0.7)
            line_adapter(slide, "SUMMER_L_b2_last_bottom", 835, 950, 835, 964,
                         color=legacy.PURPLE, width=0.7)
        return result
    if shape == legacy.MSO_SHAPE.OVAL:
        return canvas.ellipse(pt(x), pt(y), pt(w), pt(h), **kwargs)
    return canvas.shape(
        shape, "msoShapeRectangle", pt(x), pt(y), pt(w), pt(h), **kwargs
    )


def label_adapter(slide, name, x, y, w, h, text, *, size=8, color=legacy.INK,
                  bold=False, align=None, anchor=None):
    text_overrides = {
        "SUMMER_T_backbone_title": "Backbone: CSP-MEEM",
        "SUMMER_T_neck_title": "Neck: FPN-PAN （LOSC + CGAFusion）",
        "SUMMER_T_head_title": "Head: RT-DETR 原始检测头",
    }
    text = text_overrides.get(name, text)
    if name == "SUMMER_T_banner":
        x, y, w, h, size = 10, 8, 1516, 50, 26
    if name == "SUMMER_T_neck_bu":
        color = legacy.GREEN
    if name == "SUMMER_T_m1_csp":
        text = "CSP\n分割"
    if name == "SUMMER_T_b3_box":
        x, y, w, h, text = 1008, 958, 124, 18, "归一化边界框坐标"
    if name == "SUMMER_T_b3_bars":
        x, y, w, h = 916, 958, 100, 18
    if name == "SUMMER_T_b3_title":
        color = legacy.BLUE
    if name == "SUMMER_T_head_title":
        x, y, w, h, size = 1076, 92, 251, 28, 10
    if name in {"SUMMER_T_head_iou", "SUMMER_T_head_dec"}:
        bold = False
    if name in {"SUMMER_T_neck_td", "SUMMER_T_neck_bu"}:
        color = legacy.INK
    if name == "SUMMER_T_head_iou":
        x, y, w, h, size = 1087, 158, 229, 42, 9
    if name == "SUMMER_T_head_dec":
        x, y, w, h, size = 1087, 306, 229, 24, 9
    if name == "SUMMER_T_head_out":
        x, y, w, h, size, align = 1098, 447, 207, 38, 8.5, legacy.PP_ALIGN.LEFT
    if name == "SUMMER_T_b4_title":
        x, y, w, h, size = 1150, 851, 372, 28, 11
        color = legacy.RGBColor(0, 55, 210)
    if name.startswith("SUMMER_T_b4_head"):
        index = int(name[-1])
        x, y, w, h, size = 1225, 884 + index * 40, 72, 28, 9.8
        color = [
            legacy.RGBColor(0, 55, 210),
            legacy.RGBColor(0, 125, 48),
            legacy.RGBColor(255, 55, 30),
        ][index]
    if name.startswith("SUMMER_T_b4_tail"):
        index = int(name[-1])
        x, y, w, h, size = 1292, 884 + index * 40, 222, 28, 10.5
        color = legacy.RGBColor(18, 18, 18)
    if name == "SUMMER_T_b2_title":
        x, y, w, h, size = 350, 856, 529, 24, 10
        color = legacy.BLUE
    if name == "SUMMER_T_b2_iou":
        x, y, w, h, size = 382, 884, 142, 42, 8
    if name == "SUMMER_T_b2_dec":
        x, y, w, h, size = 590, 875, 270, 22, 8.5
        color = legacy.BLUE
    if name == "SUMMER_T_b2_dots2":
        x, y, w, h, size = 752, 923, 48, 24, 9
    neck_labels = {
        "SUMMER_T_neck_c5": (560, 156, 40, 20),
        "SUMMER_T_neck_c4": (504, 270, 40, 20),
        "SUMMER_T_neck_c3": (504, 412, 40, 20),
        "SUMMER_T_neck_fusion1": (650, 270, 92, 20),
        "SUMMER_T_neck_fusion2": (650, 410, 92, 20),
    }
    if name in neck_labels:
        x, y, w, h = neck_labels[name]
    if name == "SUMMER_T_neck_c4":
        color = legacy.GREEN
    if name == "SUMMER_T_neck_c3":
        color = legacy.RED
    stage_labels = {
        "SUMMER_T_bb_c5": (395, 204, 61, 24, legacy.BLUE),
        "SUMMER_T_bb_c5_dim": (395, 228, 61, 22, legacy.BLUE),
        "SUMMER_T_bb_c4": (395, 298, 61, 24, legacy.GREEN),
        "SUMMER_T_bb_c4_dim": (395, 322, 61, 22, legacy.GREEN),
        "SUMMER_T_bb_c3": (395, 391, 61, 24, legacy.RED),
        "SUMMER_T_bb_c3_dim": (395, 415, 61, 22, legacy.RED),
        "SUMMER_T_bb_outbar": (240, 504, 216, 24, legacy.INK),
    }
    if name in stage_labels:
        x, y, w, h, color = stage_labels[name]
    bottom_labels = {
        "SUMMER_E_out_p3_label": (34, 944, 74, 50),
        "SUMMER_E_out_p4_label": (124, 948, 70, 48),
        "SUMMER_E_out_p5_label": (204, 952, 88, 44),
    }
    if name in bottom_labels:
        x, y, w, h = bottom_labels[name]
    if name == "SUMMER_T_m2_out":
        x, y, w, h = 520, 806, 120, 18
    m1_labels = {
        "SUMMER_T_m1_in": (95, 594, 90, 20),
        "SUMMER_T_m1_csp": (110, 622, 60, 30),
        "SUMMER_T_m1_b0": (40, 680, 45, 19),
        "SUMMER_T_m1_b1": (116, 680, 45, 19),
        "SUMMER_T_m1_b2": (188, 680, 45, 19),
        "SUMMER_T_m1_fuse": (84, 720, 112, 20),
        "SUMMER_T_m1_ema": (84, 761, 112, 20),
        "SUMMER_T_m1_out": (95, 795, 90, 20),
    }
    if name in m1_labels:
        x, y, w, h = m1_labels[name]
    if name == "SUMMER_T_m1_bdots":
        x, y, w, h, size = 161, 678, 27, 21, 10
        color = legacy.INK
    if name.startswith("SUMMER_T_m1_p"):
        x += 34
        size = 9
        color = legacy.RGBColor(12, 12, 12)
    if name.startswith("SUMMER_T_m2_p"):
        index = int(name[-1])
        x, y = 675, 624 + index * 40
        size = 9.2
        color = legacy.RGBColor(12, 12, 12)
    m3_labels = {
        "SUMMER_T_m3_high": (1062, 624, 94, 32),
        "SUMMER_T_m3_low": (1062, 708, 94, 32),
        "SUMMER_T_m3_out": (1164, 812, 100, 18),
    }
    if name in m3_labels:
        x, y, w, h = m3_labels[name]
    if name.startswith("SUMMER_T_m3_p"):
        index = int(name[-1])
        x, y = 1278, 622 + index * 50
        size = 9
        color = legacy.RGBColor(12, 12, 12)
    legend_labels = {
        "SUMMER_T_lg_p5": (1395, 116, 125, 24),
        "SUMMER_T_lg_p4": (1395, 152, 125, 24),
        "SUMMER_T_lg_p3": (1395, 188, 125, 24),
        "SUMMER_T_lg_up": (1398, 238, 122, 24),
        "SUMMER_T_lg_down": (1398, 280, 122, 24),
        "SUMMER_T_lg_add": (1398, 322, 122, 24),
        "SUMMER_T_lg_losc_tag": (1356, 365, 38, 16),
        "SUMMER_T_lg_losc": (1400, 360, 120, 24),
        "SUMMER_T_lg_cga_tag": (1354, 405, 42, 16),
        "SUMMER_T_lg_cga": (1400, 400, 120, 24),
        "SUMMER_T_lg_solid": (1400, 459, 120, 24),
        "SUMMER_T_lg_dash": (1400, 495, 120, 24),
    }
    if name in legend_labels:
        x, y, w, h = legend_labels[name]
    if name == "SUMMER_T_lg_cga_tag":
        size = 3
    if name == "SUMMER_T_head_iou":
        slide.canvas.text(
            pt(1087), pt(160), pt(229), pt(20), "IoU-Aware",
            size=9, color=color_tuple(legacy.INK), align="center",
            valign="middle", margin=0, name="SUMMER_T_head_iou_line1",
        )
        return slide.canvas.text(
            pt(1087), pt(181), pt(229), pt(20), "Query Selection",
            size=9, color=color_tuple(legacy.INK), align="center",
            valign="middle", margin=0, name=name,
        )
    align_name = {
        legacy.PP_ALIGN.LEFT: "left",
        legacy.PP_ALIGN.RIGHT: "right",
        legacy.PP_ALIGN.CENTER: "center",
        None: "center",
    }.get(align, "center")
    valign = {
        legacy.MSO_ANCHOR.TOP: "top",
        legacy.MSO_ANCHOR.BOTTOM: "bottom",
        legacy.MSO_ANCHOR.MIDDLE: "middle",
        None: "middle",
    }.get(anchor, "middle")
    return slide.canvas.text(
        pt(x), pt(y), pt(w), pt(h), text,
        size=float(size), color=color_tuple(color), bold=bool(bold),
        align=align_name, valign=valign,
        margin=0 if name.endswith("_title") or name == "SUMMER_T_banner" else 0.5,
        font_name="Arial" if name == "SUMMER_T_banner" else None,
        name=name,
    )


def color_tuple(value):
    return tuple(int(channel) for channel in value)


def line_adapter(slide, name, x1, y1, x2, y2, *, color=legacy.INK,
                 width=1.4, dash=False, arrow=True):
    if name in {"SUMMER_L_backbone_neck", "SUMMER_L_neck_head", "SUMMER_LD_skip2"}:
        return None
    if name == "SUMMER_LD_skip1":
        name = "SUMMER_L_neck_fusion1_up2"
    if name in LINE_ENDPOINT_OVERRIDES:
        x1, y1, x2, y2 = LINE_ENDPOINT_OVERRIDES[name]
    if name == "SUMMER_L_m1_f2":
        canvas = slide.canvas
        canvas.line(pt(62), pt(662), pt(211), pt(662),
                    color=color_tuple(legacy.INK), weight=0.8,
                    name="SUMMER_L_m1_branch_bus")
        for index, branch_x in enumerate((62, 140, 211)):
            canvas.line(pt(branch_x), pt(662), pt(branch_x), pt(677),
                        color=color_tuple(legacy.INK), weight=0.8,
                        arrow_end=True, name=f"SUMMER_L_m1_branch_drop_{index}")
        for suffix, branch_x, target_x in (("left", 62, 84), ("right", 211, 196)):
            canvas.line(pt(branch_x), pt(702), pt(branch_x), pt(730),
                        color=color_tuple(legacy.INK), weight=0.8,
                        name=f"SUMMER_L_m1_fanin_{suffix}_v")
            canvas.line(pt(branch_x), pt(730), pt(target_x), pt(730),
                        color=color_tuple(legacy.INK), weight=0.8,
                        arrow_end=True, name=f"SUMMER_L_m1_fanin_{suffix}_h")
    if name == "SUMMER_L_bb_to_c5":
        canvas = slide.canvas
        canvas.line(pt(327), pt(183), pt(371), pt(183),
                    color=color_tuple(legacy.INK), weight=0.8,
                    name="SUMMER_L_bb_to_spine")
        canvas.line(pt(371), pt(183), pt(371), pt(414),
                    color=color_tuple(legacy.INK), weight=0.8,
                    name="SUMMER_L_bb_output_spine")
        for index, (start_y, end_y) in enumerate(
            [(205, 239), (300, 324), (351, 380)], start=1
        ):
            canvas.line(pt(304), pt(start_y), pt(304), pt(end_y),
                        color=color_tuple(legacy.INK), weight=0.8,
                        arrow_end=True, name=f"SUMMER_L_bb_down_{index}")
    if name in ORTHOGONAL_ROUTES:
        if name == "SUMMER_L_neck_fusion2_add":
            extra_name = "SUMMER_L_neck_fusion2_p4"
            extra = ORTHOGONAL_ROUTES[extra_name]
            extra_segments = list(zip(extra["points"], extra["points"][1:]))
            for extra_index, ((sx, sy), (ex, ey)) in enumerate(extra_segments, start=1):
                slide.canvas.line(
                    pt(sx), pt(sy), pt(ex), pt(ey),
                    color=(55, 58, 68), weight=1.1,
                    dash=True, arrow_end=extra_index == len(extra_segments),
                    name=f"{extra_name}_seg{extra_index}",
                )
        route = ORTHOGONAL_ROUTES[name]
        segments = list(zip(route["points"], route["points"][1:]))
        result = None
        for index, ((start_x, start_y), (end_x, end_y)) in enumerate(segments, start=1):
            result = slide.canvas.line(
                pt(start_x), pt(start_y), pt(end_x), pt(end_y),
                color=(55, 58, 68), weight=1.1,
                dash=route["dash"], arrow_end=index == len(segments),
                name=f"{name}_seg{index}",
            )
        return result
    return slide.canvas.line(
        pt(x1), pt(y1), pt(x2), pt(y2),
        color=color_tuple(color), weight=float(width), dash=bool(dash),
        arrow_end=bool(arrow), name=name,
    )


def badge_adapter(slide, name, cx, cy, r, fill, text, *, size=6.5,
                  text_color=None):
    if name in ICON_TRACE_FILES:
        icon_index = int(name[-1])
        icon_path = [
            HERE / "assets/02_adv_precision.png",
            HERE / "assets/03_adv_robust.png",
            HERE / "assets/04_adv_efficient.png",
        ][icon_index]
        return slide.canvas.picture(
            icon_path,
            pt(1160),
            pt(874 + icon_index * 40),
            pt(42),
            pt(42),
            name=name,
        )
    if name in {
        "SUMMER_E_neck_up1", "SUMMER_E_neck_up2",
        "SUMMER_E_neck_add_td1", "SUMMER_E_neck_add_td2",
    }:
        cx = 582
    if "_up" in name:
        text = "UP\n↑"
        size = 5.5
        text_color = legacy.RGBColor(0, 0, 0)
        fill = legacy.RGBColor(209, 193, 246)
    if "_down" in name or name == "SUMMER_E_m2_down":
        text = "Down\n↓"
        size = 4.0
        text_color = legacy.RGBColor(0, 0, 0)
    if name == "SUMMER_E_m2_down":
        cx, cy, r, size = 578, 716, 22, 6.5
    if name == "SUMMER_E_m3_add":
        cx, cy, r = 1214, 724, 20
    if name in {"SUMMER_E_lg_up", "SUMMER_E_lg_down", "SUMMER_E_lg_add"}:
        r = 13
    if name == "SUMMER_E_lg_up":
        cx, cy = 1368, 248
    if name == "SUMMER_E_lg_down":
        cx, cy = 1368, 292
    if name == "SUMMER_E_lg_add":
        cx, cy = 1368, 334
    if name in {"SUMMER_E_neck_up1", "SUMMER_E_neck_up2"}:
        r = 20
    if name in {"SUMMER_E_neck_down1", "SUMMER_E_neck_down2"}:
        r = 20
    if name in {
        "SUMMER_E_neck_add_td1", "SUMMER_E_neck_add_td2",
        "SUMMER_E_neck_add_p5", "SUMMER_E_neck_add_p4",
        "SUMMER_E_neck_add_p3",
    }:
        r = 20
    is_plus = color_tuple(fill) == color_tuple(legacy.BADGE_PLUS)
    if is_plus:
        fill = legacy.RGBColor(158, 226, 235)
        text_color = legacy.RGBColor(0, 0, 0)
    text_color = text_color or legacy.RGBColor(255, 255, 255)
    result = slide.canvas.ellipse(
        pt(cx - r), pt(cy - r), pt(2 * r), pt(2 * r),
        fill=color_tuple(fill),
        line=(25, 150, 170) if is_plus else (105, 55, 190) if "_up" in name else None,
        weight=0.8,
        name=name,
    )
    label_adapter(
        slide, f"{name}_text", cx - r, cy - r, 2 * r, 2 * r, text,
        size=size, color=text_color, bold=True,
    )
    return result


def cuboid_adapter(slide, prefix, x, y, w, h, base, depth=8):
    canvas = slide.canvas
    bottom_geometry = {
        "SUMMER_E_out_p3": (47, 894, 34, 32, 8),
        "SUMMER_E_out_p4": (138, 900, 32, 27, 8),
        "SUMMER_E_out_p5": (228, 906, 28, 23, 7),
    }
    if prefix in bottom_geometry:
        x, y, w, h, depth = bottom_geometry[prefix]
    backbone_geometry = {
        "SUMMER_E_bb1": (280, 161, 39, 36, 8),
        "SUMMER_E_bb2": (274, 239, 46, 60, 15),
        "SUMMER_E_bb3": (264, 291, 60, 60, 21),
        "SUMMER_E_bb4": (250, 350, 68, 63, 27),
    }
    if prefix in backbone_geometry:
        x, y, w, h, depth = backbone_geometry[prefix]
    neck_output_geometry = {
        "SUMMER_E_neck_p5": (972, 162, 24, 32, 8),
        "SUMMER_E_neck_p4": (972, 317, 24, 31, 8),
        "SUMMER_E_neck_p3": (972, 444, 24, 30, 8),
    }
    if prefix in neck_output_geometry:
        x, y, w, h, depth = neck_output_geometry[prefix]
    legend_geometry = {
        "SUMMER_E_lg_p5": (1360, 118, 16, 15, 5),
        "SUMMER_E_lg_p4": (1360, 154, 16, 15, 5),
        "SUMMER_E_lg_p3": (1360, 190, 16, 15, 5),
    }
    if prefix in legend_geometry:
        x, y, w, h, depth = legend_geometry[prefix]
    red, green, blue = base
    if base == legacy.CUBE_BLUE:
        red, green, blue = 142, 180, 238
    elif base == legacy.CUBE_GREEN:
        red, green, blue = 145, 202, 132
    elif base == legacy.CUBE_ORANGE:
        red, green, blue = 255, 118, 62
    front_color = (red, green, blue)
    top_color = (min(red + 32, 255), min(green + 32, 255), min(blue + 32, 255))
    right_color = (max(red - 48, 0), max(green - 48, 0), max(blue - 48, 0))
    edge = tuple(max(channel - 70, 0) for channel in front_color)
    front_points = [(pt(x), pt(y + depth)), (pt(x + w), pt(y + depth)),
                    (pt(x + w), pt(y + depth + h)), (pt(x), pt(y + depth + h))]
    top_points = [(pt(x), pt(y + depth)), (pt(x + depth), pt(y)),
                  (pt(x + w + depth), pt(y)), (pt(x + w), pt(y + depth))]
    right_points = [(pt(x + w), pt(y + depth)), (pt(x + w + depth), pt(y)),
                    (pt(x + w + depth), pt(y + h)),
                    (pt(x + w), pt(y + depth + h))]
    canvas.freeform(top_points, fill=top_color, line=edge, weight=0.8,
                    name=f"{prefix}_top")
    front = canvas.freeform(front_points, fill=front_color, line=edge, weight=0.8,
                            name=f"{prefix}_front")
    canvas.gradient_fill(
        front,
        tuple(min(channel + 24, 255) for channel in front_color),
        tuple(max(channel - 18, 0) for channel in front_color),
        angle=90,
    )
    canvas.freeform(right_points, fill=right_color, line=edge, weight=0.8,
                    name=f"{prefix}_right")
    CUBOID_GEOMETRY[prefix] = {
        "vertices": {
            "front": [[round(a, 3), round(b, 3)] for a, b in front_points],
            "top": [[round(a, 3), round(b, 3)] for a, b in top_points],
            "right": [[round(a, 3), round(b, 3)] for a, b in right_points],
        },
        "front_gradient": {
            "colors": [
                list(tuple(min(channel + 24, 255) for channel in front_color)),
                list(tuple(max(channel - 18, 0) for channel in front_color)),
            ],
            "angle": 90,
            "color_tolerance": 2,
            "angle_tolerance": 1,
        },
    }


def grid_adapter(slide, prefix, x, y, w, h, cols=4, rows=3):
    grid_overrides = {
        "SUMMER_G_m2_in": (541, 617, 74, 48),
        "SUMMER_G_m2_out": (550, 770, 58, 34),
        "SUMMER_G_m3_fh": (946, 618, 56, 46),
        "SUMMER_G_m3_fl": (946, 702, 56, 46),
        "SUMMER_G_m3_out": (1188, 770, 52, 40),
    }
    if prefix in grid_overrides:
        x, y, w, h = grid_overrides[prefix]
    grid_fill = legacy.GRID_BG
    grid_line = legacy.GRID_LINE
    if prefix == "SUMMER_G_m3_fl":
        grid_fill = legacy.RGBColor(224, 242, 216)
        grid_line = legacy.RGBColor(112, 176, 92)
    rect_adapter(slide, f"{prefix}_bg", x, y, w, h, fill=grid_fill,
                 edge=grid_line, edge_w=0.75,
                 shape=legacy.MSO_SHAPE.RECTANGLE)
    for column in range(1, cols):
        line_adapter(slide, f"{prefix}_v{column}", x + w * column / cols, y,
                     x + w * column / cols, y + h, color=grid_line,
                     width=0.5, arrow=False)
    for row in range(1, rows):
        line_adapter(slide, f"{prefix}_h{row}", x, y + h * row / rows,
                     x + w, y + h * row / rows, color=grid_line,
                     width=0.5, arrow=False)


def square_row_adapter(slide, prefix, x, y, count, size=13, gap=6,
                       fill=None):
    fill = fill or legacy.RGBColor(150, 120, 220)
    if prefix == "SUMMER_E_head_iou_sq":
        x, y, count, size, gap = 1120, 218, 6, 17, 8
        fill = legacy.RGBColor(198, 178, 239)
    if prefix == "SUMMER_E_head_dec_sq":
        node_xs = [1102, 1148, 1194, 1282]
        node_y, node_w, node_h = 369, 20, 28
        line_adapter(
            slide, "SUMMER_L_head_dec_top_bus",
            1112, 354, 1292, 354,
            color=legacy.PURPLE, width=0.7, arrow=False,
        )
        for index, node_x in enumerate(node_xs):
            rect_adapter(
                slide, f"{prefix}_{index}", node_x, node_y, node_w, node_h,
                fill=fill, edge=legacy.PURPLE, edge_w=0.8,
                shape=legacy.MSO_SHAPE.ROUNDED_RECTANGLE,
            )
            line_adapter(
                slide, f"SUMMER_L_head_dec_stem_top_{index}",
                node_x + node_w / 2, 354, node_x + node_w / 2, node_y,
                color=legacy.PURPLE, width=0.7,
            )
            line_adapter(
                slide, f"SUMMER_L_head_dec_stem_bottom_{index}",
                node_x + node_w / 2, node_y + node_h,
                node_x + node_w / 2, 411,
                color=legacy.PURPLE, width=0.7,
            )
        for index, (start_x, end_x) in enumerate(
            [(1122, 1148), (1168, 1194), (1214, 1230), (1260, 1282)]
        ):
            line_adapter(
                slide, f"SUMMER_L_head_dec_chain_{index}",
                start_x, 383, end_x, 383,
                color=legacy.PURPLE, width=0.7,
                arrow=index != 2,
            )
        label_adapter(
            slide, f"{prefix}_dots", 1228, 373, 34, 20,
            "••••", size=8, color=legacy.INK, bold=True,
        )
        return
    if prefix == "SUMMER_E_b2_sq1":
        x, y, count, size, gap = 384, 934, 6, 16, 7
        fill = legacy.RGBColor(198, 178, 239)
    if prefix == "SUMMER_E_b2_sq2":
        node_xs = [594, 654, 714]
        node_y, node_w, node_h = 918, 22, 32
        for index, node_x in enumerate(node_xs):
            rect_adapter(
                slide, f"{prefix}_{index}", node_x, node_y, node_w, node_h,
                fill=fill, edge=legacy.PURPLE, edge_w=0.8,
                shape=legacy.MSO_SHAPE.ROUNDED_RECTANGLE,
            )
            line_adapter(
                slide, f"SUMMER_L_b2_stem_top_{index}",
                node_x + node_w / 2, 904, node_x + node_w / 2, node_y,
                color=legacy.PURPLE, width=0.7,
            )
            line_adapter(
                slide, f"SUMMER_L_b2_stem_bottom_{index}",
                node_x + node_w / 2, node_y + node_h,
                node_x + node_w / 2, 964,
                color=legacy.PURPLE, width=0.7,
            )
        for index, (start_x, end_x) in enumerate(
            [(616, 654), (676, 714), (736, 748), (800, 824)]
        ):
            line_adapter(
                slide, f"SUMMER_L_b2_chain_{index}",
                start_x, 934, end_x, 934,
                color=legacy.PURPLE, width=0.7,
                arrow=index not in {2},
            )
        return
    for index in range(count):
        rect_adapter(slide, f"{prefix}_{index}", x + index * (size + gap), y,
                     size, size, fill=fill,
                     edge=legacy.PURPLE if prefix == "SUMMER_E_head_iou_sq" else None,
                     edge_w=0.8, shape=legacy.MSO_SHAPE.RECTANGLE)
    label_adapter(slide, f"{prefix}_dots", x + count * (size + gap) + 2,
                  y - 2, 30, size + 4, "····", size=8, color=legacy.GRAY,
                  bold=True)


def mini_bars_adapter(slide, prefix, x, y):
    if prefix == "SUMMER_E_head_bars":
        heights = (32, 18, 26, 42)
        x, y = 1134, 493
        line_adapter(slide, f"{prefix}_axis_x", x - 4, y + 44, x + 48, y + 44,
                     color=legacy.BLUE, width=1.0, arrow=False)
        line_adapter(slide, f"{prefix}_axis_y", x - 4, y + 44, x - 4, y,
                     color=legacy.BLUE, width=1.0, arrow=False)
        for index, height in enumerate(heights):
            rect_adapter(
                slide, f"{prefix}_{index}", x + index * 11, y + 44 - height,
                7, height, fill=legacy.RGBColor(90, 118, 216),
                shape=legacy.MSO_SHAPE.RECTANGLE,
            )
        return
    if prefix == "SUMMER_E_b3_bars":
        heights = (25, 50, 18, 36, 12)
        x, y = 944, 898
        line_adapter(slide, f"{prefix}_axis_x", x - 4, y + 54, x + 48, y + 54,
                     color=legacy.BLUE, width=1.0, arrow=False)
        line_adapter(slide, f"{prefix}_axis_y", x - 4, y + 54, x - 4, y,
                     color=legacy.BLUE, width=1.0, arrow=False)
        for index, height in enumerate(heights):
            rect_adapter(
                slide, f"{prefix}_{index}", x + index * 9, y + 54 - height,
                6, height, fill=legacy.RGBColor(70, 95, 205),
                shape=legacy.MSO_SHAPE.RECTANGLE,
            )
        return
    for index, height in enumerate((18, 30, 12, 24)):
        rect_adapter(slide, f"{prefix}_{index}", x + index * 10,
                     y + 32 - height, 7, height,
                     fill=legacy.RGBColor(90, 118, 216),
                     shape=legacy.MSO_SHAPE.RECTANGLE)



def build() -> Path:
    legacy.HERE = HERE
    legacy.Presentation = PresentationProxy
    legacy.rect = rect_adapter
    legacy.label = label_adapter
    legacy.line = line_adapter
    legacy.badge = badge_adapter
    legacy.cuboid = cuboid_adapter
    legacy.grid = grid_adapter
    legacy.square_row = square_row_adapter
    legacy.mini_bars = mini_bars_adapter
    output = legacy.build()
    write_manifest(HERE, ORTHOGONAL_ROUTES, CUBOID_GEOMETRY)
    return output


if __name__ == "__main__":
    print(build())
