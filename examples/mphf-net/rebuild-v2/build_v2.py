#!/usr/bin/env python3
"""MPHF-Net 结构图从零重建 v2 —— ShapeCanvas 运行时 + 等轴测样式。

源图 assets/reference-figures/mphf-net-architecture.png（1536x1024 px）
-> slide 960x640 pt（uniform 0.625 pt/px）。场景坐标以源图 px 表达。
前缀 FIG_；PCB 板图按 Crop Contract 保留为 assets/01_R1_pcb.png。
"""
from __future__ import annotations

import importlib.util
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[2]
SCALE = 0.625

spec = importlib.util.spec_from_file_location(
    "office_shape_canvas", ROOT / "scripts/office_shape_canvas.py")
canvas_module = importlib.util.module_from_spec(spec)
sys.modules["office_shape_canvas"] = canvas_module
spec.loader.exec_module(canvas_module)
ShapeCanvas = canvas_module.ShapeCanvas

NAVY = (24, 33, 74)
BLUE = (37, 80, 216)
RED = (224, 58, 38)
PURPLE = (112, 62, 200)
GREEN = (28, 148, 84)
INK = (40, 44, 60)
GRAY = (120, 126, 140)
PANEL_EDGE = (203, 210, 228)
WHITE = (255, 255, 255)
CUBE_BLUE = (142, 180, 238)
CUBE_GREEN = (145, 202, 132)
CUBE_ORANGE = (255, 118, 62)
GRID_BG = (214, 228, 252)
GRID_LINE = (120, 156, 228)
GRID_BG_GREEN = (224, 242, 216)
GRID_LINE_GREEN = (112, 176, 92)

CUBOID_GEOMETRY: dict[str, dict] = {}
ROUTES: list[dict] = []
IGNORE_TEXT: list[str] = []
IGNORE_DANGLING: list[str] = []


def px(value: float) -> float:
    return value * SCALE


def cuboid(canvas: ShapeCanvas, prefix: str, x: float, y: float, w: float,
           h: float, base: tuple[int, int, int], depth: float = 8) -> None:
    """等轴测三面块：top/right 斜面 freeform，front 垂直渐变。"""
    red, green, blue = base
    front_color = (red, green, blue)
    top_color = (min(red + 32, 255), min(green + 32, 255), min(blue + 32, 255))
    right_color = (max(red - 48, 0), max(green - 48, 0), max(blue - 48, 0))
    edge = tuple(max(channel - 70, 0) for channel in front_color)
    front_points = [(px(x), px(y + depth)), (px(x + w), px(y + depth)),
                    (px(x + w), px(y + depth + h)), (px(x), px(y + depth + h))]
    top_points = [(px(x), px(y + depth)), (px(x + depth), px(y)),
                  (px(x + w + depth), px(y)), (px(x + w), px(y + depth))]
    right_points = [(px(x + w), px(y + depth)), (px(x + w + depth), px(y)),
                    (px(x + w + depth), px(y + h)),
                    (px(x + w), px(y + depth + h))]
    canvas.freeform(top_points, fill=top_color, line=edge, weight=0.8,
                    name=f"{prefix}_top")
    front = canvas.freeform(front_points, fill=front_color, line=edge,
                            weight=0.8, name=f"{prefix}_front")
    gradient_start = tuple(min(channel + 24, 255) for channel in front_color)
    gradient_end = tuple(max(channel - 18, 0) for channel in front_color)
    canvas.gradient_fill(front, gradient_start, gradient_end, angle=90)
    canvas.freeform(right_points, fill=right_color, line=edge, weight=0.8,
                    name=f"{prefix}_right")
    CUBOID_GEOMETRY[prefix] = {
        "front": [[round(a, 3), round(b, 3)] for a, b in front_points],
        "top": [[round(a, 3), round(b, 3)] for a, b in top_points],
        "right": [[round(a, 3), round(b, 3)] for a, b in right_points],
        "gradient": {"colors": [list(gradient_start), list(gradient_end)],
                     "angle": 90},
    }


def badge(canvas: ShapeCanvas, name: str, cx: float, cy: float, r: float,
          fill: tuple, edge: tuple, text: str, text_color: tuple,
          size: float = 6.5) -> None:
    canvas.ellipse(px(cx - r), px(cy - r), px(2 * r), px(2 * r), fill=fill,
                   line=edge, weight=1.0, name=name)
    canvas.text(px(cx - r), px(cy - r), px(2 * r), px(2 * r), text, size=size,
                color=text_color, bold=True, name=f"{name}_text")
    IGNORE_TEXT.append(f"{name}_text")


def arrow(canvas: ShapeCanvas, name: str, x1: float, y1: float, x2: float,
          y2: float, *, color: tuple = INK, weight: float = 1.2,
          dash: bool = False) -> None:
    canvas.line(px(x1), px(y1), px(x2), px(y2), color=color, weight=weight,
                dash=dash, arrow_end=True, name=name)


def ortho(canvas: ShapeCanvas, stem: str, points: list[tuple[float, float]],
          *, color: tuple = GRAY, weight: float = 1.0,
          dash: bool = True) -> None:
    """正交折线：末段带箭头，其余为过渡段。"""
    scaled = [(px(a), px(b)) for a, b in points]
    for index, ((x1, y1), (x2, y2)) in enumerate(zip(scaled, scaled[1:])):
        canvas.line(x1, y1, x2, y2, color=color, weight=weight, dash=dash,
                    arrow_end=index == len(scaled) - 2,
                    name=f"{stem}_{index + 1}")


def grid(canvas: ShapeCanvas, prefix: str, x: float, y: float, w: float,
         h: float, cols: int = 4, rows: int = 3, *, fill: tuple = GRID_BG,
         line: tuple = GRID_LINE) -> None:
    canvas.rect(px(x), px(y), px(w), px(h), fill=fill, line=line, weight=0.75,
                name=f"{prefix}_bg")
    for col in range(1, cols):
        canvas.line(px(x + w * col / cols), px(y), px(x + w * col / cols),
                    px(y + h), color=line, weight=0.5,
                    name=f"{prefix}_v{col}")
    for row in range(1, rows):
        canvas.line(px(x), px(y + h * row / rows), px(x + w),
                    px(y + h * row / rows), color=line, weight=0.5,
                    name=f"{prefix}_h{row}")


def square_row(canvas: ShapeCanvas, prefix: str, x: float, y: float,
               count: int, size: float = 13, gap: float = 6,
               fill: tuple = (150, 120, 220)) -> None:
    for index in range(count):
        canvas.rect(px(x + index * (size + gap)), px(y), px(size), px(size),
                    fill=fill, line=(110, 82, 178), weight=0.6,
                    name=f"{prefix}_{index}")
    canvas.text(px(x + count * (size + gap) + 2), px(y - 2), px(30),
                px(size + 4), "····", size=8, color=GRAY, bold=True,
                name=f"{prefix}_dots")


def mini_bars(canvas: ShapeCanvas, prefix: str, x: float, y: float) -> None:
    heights = (18, 30, 12, 24)
    for index, height in enumerate(heights):
        canvas.rect(px(x + index * 10), px(y + 32 - height), px(7), px(height),
                    fill=(90, 118, 216), line=None, name=f"{prefix}_{index}")
    canvas.line(px(x - 3), px(y + 33), px(x + 42), px(y + 33), color=INK,
                weight=0.8, name=f"{prefix}_axis")
    IGNORE_DANGLING.append(f"{prefix}_axis")


def route(connector: str, target: str, target_edge: str, *, source=None,
          source_edge=None, tolerance: float = 1.5) -> None:
    entry = {"connector": connector, "target": target,
             "target_edge": target_edge, "tolerance_pt": tolerance}
    if source:
        entry["source"] = source
        entry["source_edge"] = source_edge
    ROUTES.append(entry)


def build() -> Path:
    canvas = ShapeCanvas(960, 640, name_prefix="FIG")

    # ---- 标题横幅 ----
    with canvas.group("banner"):
        banner = canvas.round_rect(px(10), px(8), px(1516), px(54), fill=NAVY,
                                   name="FIG_E_banner")
        banner.adjustments[0] = 0.18
        canvas.text(px(10), px(12), px(1516), px(46),
                    "MPHF-Net 工业表面缺陷检测网络结构设计图", size=17,
                    color=WHITE, bold=True, name="FIG_T_banner")

    # ---- 面板底框 ----
    panels = {
        "input": (12, 76, 200, 460), "backbone": (228, 76, 258, 460),
        "neck": (498, 76, 574, 460), "head": (1086, 76, 240, 460),
        "legend": (1340, 76, 186, 460),
        "m1": (12, 556, 474, 284), "m2": (498, 556, 408, 284),
        "m3": (916, 556, 610, 284),
        "b1": (12, 852, 322, 164), "b2": (348, 852, 542, 164),
        "b3": (908, 852, 226, 164), "b4": (1148, 852, 378, 164),
    }
    with canvas.group("panels"):
        for key, (x, y, w, h) in panels.items():
            canvas.round_rect(px(x), px(y), px(w), px(h), fill=(252, 252, 254),
                              line=PANEL_EDGE, weight=1.2,
                              name=f"FIG_E_panel_{key}")

    # ---- P1 输入图像 ----
    with canvas.group("input"):
        canvas.text(px(12), px(86), px(200), px(22), "输入图像", size=11,
                    bold=True, color=INK, name="FIG_T_input_title")
        canvas.picture(HERE / "assets/01_R1_pcb.png", px(25), px(162),
                       px(171), px(244), name="FIG_R_pcb")
        canvas.text(px(12), px(430), px(200), px(24), "640 × 640 × 3",
                    size=11, color=INK, name="FIG_T_input_size")

    # ---- P2 Backbone ----
    with canvas.group("backbone"):
        canvas.text(px(228), px(86), px(258), px(20), "Backbone：CSP-MEEM",
                    size=11, color=BLUE, bold=True, name="FIG_T_backbone_title")
        canvas.text(px(228), px(108), px(258), px(16), "多尺度细粒度特征提取",
                    size=8, color=GRAY, name="FIG_T_backbone_sub")
        # 原图拓扑：4 块分离垂直排列 + 块间向下箭头（特征流），
        # 右侧垂直 bus 干线一干三支分发到 C5/C4/C3（trunk+branch）
        pyramid = [
            ("FIG_E_bb1", 300, 150, 34, 30, 10),
            ("FIG_E_bb2", 284, 224, 46, 40, 12),
            ("FIG_E_bb3", 266, 300, 62, 50, 16),
            ("FIG_E_bb4", 240, 388, 84, 64, 22),
        ]
        for name, x, y, w, h, depth in pyramid:
            cuboid(canvas, name, x, y, w, h, CUBE_BLUE, depth)
        # 块间向下箭头：上块 front 底中心 -> 下块 top 上缘
        chain = [(317, 190, 224), (307, 276, 300), (297, 366, 388)]
        for index, (cx, y1, y2) in enumerate(chain, start=1):
            arrow(canvas, f"FIG_L_bb_chain{index}", cx, y1, cx, y2)
        # bus 干线：块1 右缘引出，垂直贯穿到 C3 行；块3/块4 右缘并入
        bus_x = 372
        canvas.line(px(344), px(175), px(bus_x), px(175), color=INK,
                    weight=1.2, name="FIG_L_bb_bus_in1")
        canvas.line(px(bus_x), px(175), px(bus_x), px(414), color=INK,
                    weight=1.2, name="FIG_L_bb_bus")
        canvas.line(px(344), px(341), px(bus_x), px(341), color=INK,
                    weight=1.2, name="FIG_L_bb_bus_in3")
        stages = [("c5", "C5", "20×C", BLUE, 227),
                  ("c4", "C4", "40×C", BLUE, 321),
                  ("c3", "C3", "80×C", RED, 414)]
        for key, label_text, dim, color, center_y in stages:
            canvas.round_rect(px(396), px(center_y - 22), px(58), px(44),
                              fill=(245, 248, 255), line=color, weight=1.4,
                              name=f"FIG_E_bb_{key}")
            canvas.text(px(396), px(center_y - 18), px(58), px(20),
                        label_text, size=10, color=color, bold=True,
                        name=f"FIG_T_bb_{key}")
            canvas.text(px(396), px(center_y), px(58), px(18), dim, size=8,
                        color=color, name=f"FIG_T_bb_{key}_dim")
            start_x = 346 if key == "c3" else bus_x  # C3 行与块4并入线共线直达
            arrow(canvas, f"FIG_L_bb_to_{key}", start_x, center_y, 396,
                  center_y)
            route(f"FIG_L_bb_to_{key}", f"FIG_E_bb_{key}", "left")
        canvas.text(px(270), px(474), px(40), px(16), "⋮", size=11,
                    color=GRAY, bold=True, name="FIG_T_backbone_dots")
        canvas.round_rect(px(240), px(498), px(216), px(30),
                          fill=(235, 241, 255), line=PANEL_EDGE, weight=0.8,
                          name="FIG_E_bb_outbar")
        canvas.text(px(240), px(502), px(216), px(22), "输出多尺度特征",
                    size=9, color=INK, name="FIG_T_bb_outbar")

    # ---- P3 Neck ----
    with canvas.group("neck"):
        canvas.text(px(498), px(86), px(574), px(20),
                    "Neck：FPN-PAN （LOSC + CGAFusion）", size=11,
                    color=GREEN, bold=True, name="FIG_T_neck_title")
        canvas.round_rect(px(540), px(116), px(130), px(26),
                          fill=(226, 244, 230), line=GREEN, weight=0.9,
                          name="FIG_E_neck_td")
        canvas.text(px(540), px(119), px(130), px(20), "Top-Down 路径",
                    size=8.5, color=GREEN, bold=True, name="FIG_T_neck_td")
        canvas.round_rect(px(828), px(116), px(130), px(26),
                          fill=(252, 236, 226), line=(226, 120, 60),
                          weight=0.9, name="FIG_E_neck_bu")
        canvas.text(px(828), px(119), px(130), px(20), "Bottom-Up 路径",
                    size=8.5, color=(206, 100, 40), bold=True,
                    name="FIG_T_neck_bu")

        for key, x, y, color in (("c5", 552, 152, BLUE), ("c4", 508, 266, BLUE),
                                 ("c3", 508, 408, BLUE)):
            canvas.round_rect(px(x), px(y), px(40), px(28),
                              fill=(245, 248, 255), line=color, weight=1.2,
                              name=f"FIG_E_neck_{key}")
            canvas.text(px(x), px(y + 4), px(40), px(20), key.upper(),
                        size=9, color=color, bold=True,
                        name=f"FIG_T_neck_{key}")

        badge(canvas, "FIG_E_neck_up1", 602, 222, 14, (222, 205, 248),
              (140, 96, 216), "UP\n↑", PURPLE, size=5)
        badge(canvas, "FIG_E_neck_up2", 602, 344, 14, (222, 205, 248),
              (140, 96, 216), "UP\n↑", PURPLE, size=5)
        badge(canvas, "FIG_E_neck_add_td1", 602, 280, 13, (52, 58, 78),
              (30, 34, 48), "+", WHITE, size=9)
        badge(canvas, "FIG_E_neck_add_td2", 602, 420, 13, (52, 58, 78),
              (30, 34, 48), "+", WHITE, size=9)
        for key, y in (("fusion1", 265), ("fusion2", 405)):
            canvas.round_rect(px(660), px(y), px(92), px(30),
                              fill=(240, 245, 255), line=BLUE, weight=1.2,
                              name=f"FIG_E_neck_{key}")
            canvas.text(px(660), px(y + 5), px(92), px(20), "CGAFusion",
                        size=8.5, color=BLUE, bold=True,
                        name=f"FIG_T_neck_{key}")
        for key, cy in (("add_p5", 182), ("add_p4", 342), ("add_p3", 468)):
            badge(canvas, f"FIG_E_neck_{key}", 892, cy, 13, (52, 58, 78),
                  (30, 34, 48), "+", WHITE, size=9)
        for key, cy in (("down1", 242), ("down2", 404)):
            badge(canvas, f"FIG_E_neck_{key}", 892, cy, 14, (250, 214, 170),
                  (206, 120, 40), "Down\n↓", (150, 74, 12), size=4.5)
        canvas.round_rect(px(856), px(288), px(72), px(26),
                          fill=(253, 240, 232), line=(226, 110, 50),
                          weight=1.2, name="FIG_E_neck_losc")
        canvas.text(px(856), px(291), px(72), px(20), "LOSC", size=9,
                    color=(206, 90, 30), bold=True, name="FIG_T_neck_losc")

        for key, y, base, color, dim in (
                ("p5", 158, CUBE_BLUE, BLUE, "20×C"),
                ("p4", 318, CUBE_GREEN, GREEN, "40×C"),
                ("p3", 444, CUBE_ORANGE, RED, "80×C")):
            cuboid(canvas, f"FIG_E_neck_{key}", 964, y, 30, 26, base, 8)
            canvas.text(px(1014), px(y - 4), px(52), px(20), key.upper(),
                        size=10, color=color, bold=True,
                        name=f"FIG_T_neck_{key}")
            canvas.text(px(1014), px(y + 14), px(52), px(16), dim, size=8,
                        color=color, name=f"FIG_T_neck_{key}_dim")

        arrow(canvas, "FIG_L_neck_c5_up", 572, 180, 597, 207)
        arrow(canvas, "FIG_L_neck_up1_add", 602, 236, 602, 267)
        arrow(canvas, "FIG_L_neck_c4_add", 548, 280, 589, 280)
        arrow(canvas, "FIG_L_neck_add1_fusion", 615, 280, 660, 280)
        route("FIG_L_neck_add1_fusion", "FIG_E_neck_fusion1", "left")
        arrow(canvas, "FIG_L_neck_up2_add", 602, 358, 602, 407)
        arrow(canvas, "FIG_L_neck_c3_add", 548, 421, 589, 421)
        arrow(canvas, "FIG_L_neck_add2_fusion", 615, 420, 660, 420)
        route("FIG_L_neck_add2_fusion", "FIG_E_neck_fusion2", "left")
        arrow(canvas, "FIG_L_neck_addp5_cube", 905, 182, 964, 182)
        arrow(canvas, "FIG_L_neck_addp4_cube", 905, 342, 964, 342)
        arrow(canvas, "FIG_L_neck_addp3_cube", 905, 468, 964, 468)
        arrow(canvas, "FIG_L_neck_addp5_down", 892, 195, 892, 228)
        arrow(canvas, "FIG_L_neck_down_losc", 892, 256, 892, 288)
        arrow(canvas, "FIG_L_neck_losc_addp4", 892, 314, 892, 329)
        arrow(canvas, "FIG_L_neck_addp4_down2", 892, 355, 892, 390)
        arrow(canvas, "FIG_L_neck_down2_addp3", 892, 418, 892, 455)
        # CGAFusion 输出 -> 右侧拼接点：正交虚线（贴原图跨层连接）
        ortho(canvas, "FIG_LD_fusion1_addp5",
              [(752, 280), (790, 280), (790, 182), (879, 182)])
        ortho(canvas, "FIG_LD_fusion2_addp3",
              [(752, 420), (790, 420), (790, 468), (879, 468)])
        ortho(canvas, "FIG_LD_fusion1_fusion2", [(706, 295), (706, 405)])

        canvas.round_rect(px(506), px(508), px(250), px(28),
                          fill=(235, 242, 255), line=BLUE, weight=0.8,
                          name="FIG_E_neck_note_cga")
        canvas.text(px(506), px(512), px(250), px(20),
                    "CGAFusion（高低频自适应融合）", size=8.5, color=BLUE,
                    name="FIG_T_neck_note_cga")
        canvas.round_rect(px(774), px(508), px(250), px(28),
                          fill=(253, 238, 232), line=RED, weight=0.8,
                          name="FIG_E_neck_note_losc")
        canvas.text(px(774), px(512), px(250), px(20),
                    "LOSC（方向性下采样建模）", size=8.5, color=RED,
                    name="FIG_T_neck_note_losc")

    # ---- P4 Head ----
    with canvas.group("head"):
        canvas.text(px(1086), px(86), px(240), px(20),
                    "Head：RT-DETR 原始检测头", size=9.5, color=PURPLE,
                    bold=True, name="FIG_T_head_title")
        canvas.round_rect(px(1098), px(142), px(216), px(96), fill=WHITE,
                          line=PANEL_EDGE, weight=1.0, name="FIG_E_head_iou")
        canvas.text(px(1098), px(152), px(216), px(36),
                    "IoU-Aware\nQuery Selection", size=9, color=INK,
                    bold=True, name="FIG_T_head_iou")
        square_row(canvas, "FIG_E_head_iou_sq", 1116, 202, 7, 12, 5)
        arrow(canvas, "FIG_L_head_iou_dec", 1206, 238, 1206, 292)
        route("FIG_L_head_iou_dec", "FIG_E_head_dec", "top",
              source="FIG_E_head_iou", source_edge="bottom")
        canvas.round_rect(px(1098), px(292), px(216), px(108), fill=WHITE,
                          line=PANEL_EDGE, weight=1.0, name="FIG_E_head_dec")
        canvas.text(px(1098), px(300), px(216), px(22),
                    "Transformer Decoder × 6 层", size=8.5, color=INK,
                    bold=True, name="FIG_T_head_dec")
        square_row(canvas, "FIG_E_head_dec_sq", 1114, 340, 6, 14, 7)
        arrow(canvas, "FIG_L_head_dec_out", 1206, 400, 1206, 432)
        route("FIG_L_head_dec_out", "FIG_E_head_out", "top",
              source="FIG_E_head_dec", source_edge="bottom")
        canvas.round_rect(px(1098), px(432), px(216), px(92), fill=WHITE,
                          line=PANEL_EDGE, weight=1.0, name="FIG_E_head_out")
        canvas.text(px(1098), px(438), px(216), px(34),
                    "输出：\n类别概率 + 归一化边界框坐标", size=8.5,
                    color=INK, name="FIG_T_head_out")
        mini_bars(canvas, "FIG_E_head_bars", 1136, 478)
        bbox = canvas.rect(px(1216), px(478), px(40), px(32), fill=None,
                           line=BLUE, weight=1.2, name="FIG_E_head_bbox")
        bbox.line.dash_style = canvas_module.MSO_LINE_DASH_STYLE.DASH

    # ---- P5 图例 ----
    with canvas.group("legend"):
        canvas.text(px(1340), px(86), px(186), px(20), "图例说明", size=10.5,
                    bold=True, color=INK, name="FIG_T_legend_title")
        for key, text, base, y in (("p5", "P5 / C5 （20×C）", CUBE_BLUE, 132),
                                   ("p4", "P4 / C4 （40×C）", CUBE_GREEN, 168),
                                   ("p3", "P3 / C3 （80×C）", CUBE_ORANGE, 204)):
            cuboid(canvas, f"FIG_E_lg_{key}", 1354, y, 18, 13, base, 5)
            canvas.text(px(1388), px(y - 2), px(136), px(22), text, size=7.5,
                        align="left", color=INK, name=f"FIG_T_lg_{key}")
        badge(canvas, "FIG_E_lg_up", 1364, 250, 10, (222, 205, 248),
              (140, 96, 216), "UP", PURPLE, size=5)
        canvas.text(px(1388), px(238), px(136), px(22), "上采样 （× 2）",
                    size=7.5, align="left", color=INK, name="FIG_T_lg_up")
        badge(canvas, "FIG_E_lg_down", 1364, 286, 10, (250, 214, 170),
              (206, 120, 40), "Down", (150, 74, 12), size=4)
        canvas.text(px(1388), px(274), px(136), px(22), "下采样 （× 2）",
                    size=7.5, align="left", color=INK, name="FIG_T_lg_down")
        badge(canvas, "FIG_E_lg_add", 1364, 322, 10, (52, 58, 78),
              (30, 34, 48), "+", WHITE, size=7)
        canvas.text(px(1388), px(310), px(136), px(22), "拼接 （Concat）",
                    size=7.5, align="left", color=INK, name="FIG_T_lg_add")
        canvas.round_rect(px(1352), px(350), px(32), px(16),
                          fill=(253, 240, 232), line=(226, 110, 50),
                          weight=0.8, name="FIG_E_lg_losc")
        canvas.text(px(1352), px(351), px(32), px(14), "LOSC", size=5.5,
                    color=(206, 90, 30), bold=True, name="FIG_T_lg_losc_tag")
        canvas.text(px(1388), px(347), px(136), px(22), "方向性下采样建模",
                    size=7.5, align="left", color=INK, name="FIG_T_lg_losc")
        canvas.round_rect(px(1350), px(384), px(44), px(16),
                          fill=(235, 242, 255), line=BLUE, weight=0.8,
                          name="FIG_E_lg_cga")
        canvas.text(px(1350), px(385), px(44), px(14), "CGAFusion", size=4.5,
                    color=BLUE, bold=True, name="FIG_T_lg_cga_tag")
        canvas.text(px(1388), px(381), px(136), px(22), "高低频自适应融合",
                    size=7.5, align="left", color=INK, name="FIG_T_lg_cga")
        canvas.line(px(1354), px(428), px(1384), px(428), color=INK,
                    weight=1.2, arrow_end=True, name="FIG_LD_lg_solid")
        IGNORE_DANGLING.append("FIG_LD_lg_solid")
        canvas.text(px(1388), px(417), px(136), px(22), "数据流向", size=7.5,
                    align="left", color=INK, name="FIG_T_lg_solid")
        canvas.line(px(1354), px(462), px(1384), px(462), color=INK,
                    weight=1.2, dash=True, arrow_end=True,
                    name="FIG_LD_lg_dash")
        IGNORE_DANGLING.append("FIG_LD_lg_dash")
        canvas.text(px(1388), px(451), px(136), px(22), "跨层连接", size=7.5,
                    align="left", color=INK, name="FIG_T_lg_dash")

    # ---- 顶行主流程箭头 ----
    with canvas.group("flow"):
        for name, x1, x2, target in (
                ("FIG_L_input_backbone", 212, 228, "FIG_E_panel_backbone"),
                ("FIG_L_backbone_neck", 486, 498, "FIG_E_panel_neck"),
                ("FIG_L_neck_head", 1072, 1086, "FIG_E_panel_head")):
            arrow(canvas, name, x1, 296, x2, 296, color=BLUE, weight=4.0)
            route(name, target, "left")

    # ---- 模块一 ----
    with canvas.group("m1"):
        canvas.text(px(12), px(566), px(474), px(20),
                    "模块一：CSP-MEEM（多尺度细粒度特征提取）", size=10,
                    color=BLUE, bold=True, name="FIG_T_m1_title")
        canvas.text(px(60), px(596), px(120), px(20), "输入图像", size=8,
                    color=INK, name="FIG_T_m1_in")
        canvas.round_rect(px(84), px(622), px(72), px(30),
                          fill=(235, 242, 255), line=BLUE, weight=0.9,
                          name="FIG_E_m1_csp")
        canvas.text(px(84), px(626), px(72), px(22), "CSP 分割", size=8,
                    color=INK, name="FIG_T_m1_csp")
        for index, item in enumerate(("分支1", "分支2", "分支n")):
            x = 28 + index * 66
            canvas.round_rect(px(x), px(674), px(52), px(26),
                              fill=(240, 246, 255), line=BLUE, weight=0.8,
                              name=f"FIG_E_m1_b{index}")
            canvas.text(px(x), px(677), px(52), px(20), item, size=7.5,
                        color=INK, name=f"FIG_T_m1_b{index}")
        canvas.text(px(156), px(677), px(30), px(20), "···", size=8,
                    color=GRAY, name="FIG_T_m1_bdots")
        canvas.round_rect(px(70), px(716), px(100), px(26),
                          fill=(235, 242, 255), line=BLUE, weight=0.9,
                          name="FIG_E_m1_fuse")
        canvas.text(px(70), px(719), px(100), px(20), "融合", size=8,
                    color=INK, name="FIG_T_m1_fuse")
        canvas.round_rect(px(62), px(754), px(116), px(26),
                          fill=(255, 240, 226), line=(226, 130, 50),
                          weight=0.9, name="FIG_E_m1_ema")
        canvas.text(px(62), px(757), px(116), px(20), "EMA 注意力", size=8,
                    color=(200, 100, 30), name="FIG_T_m1_ema")
        canvas.text(px(60), px(790), px(120), px(20), "输出特征", size=8,
                    color=INK, name="FIG_T_m1_out")
        for index, (y1, y2) in enumerate(((616, 622), (652, 674), (700, 716),
                                          (742, 754), (780, 790)), start=1):
            arrow(canvas, f"FIG_L_m1_f{index}", 120, y1, 120, y2, weight=1.0)
        for index, item in enumerate(("多尺度并行感知，捕获不同大小的缺陷信息",
                                      "增强细节特征表达，保留边缘和纹理",
                                      "EMA 注意力增强关键特征，抑制冗余信息")):
            canvas.text(px(226), px(606 + index * 66), px(246), px(56),
                        f"•  {item}", size=8.5, align="left", valign="top",
                        color=INK, name=f"FIG_T_m1_p{index}")

    # ---- 模块二 ----
    with canvas.group("m2"):
        canvas.text(px(498), px(566), px(408), px(20),
                    "模块二：LOSC（方向性下采样建模）", size=10, color=RED,
                    bold=True, name="FIG_T_m2_title")
        canvas.text(px(520), px(594), px(120), px(18), "输入特征", size=8,
                    color=INK, name="FIG_T_m2_in")
        grid(canvas, "FIG_G_m2_in", 540, 614, 76, 56)
        badge(canvas, "FIG_E_m2_down", 578, 706, 15, (250, 214, 170),
              (206, 120, 40), "Down\n↓", (150, 74, 12), size=4.5)
        grid(canvas, "FIG_G_m2_out", 552, 748, 52, 40, 3, 2)
        canvas.text(px(520), px(794), px(120), px(18), "输出特征", size=8,
                    color=INK, name="FIG_T_m2_out")
        arrow(canvas, "FIG_L_m2_f1", 578, 670, 578, 690, weight=1.0)
        arrow(canvas, "FIG_L_m2_f2", 578, 722, 578, 748, weight=1.0)
        for index, item in enumerate(("引入方向性卷积核", "沿关键方向提取特征",
                                      "抑制无关干扰，保留方向信息",
                                      "增强对工业表面纹理与缺陷形态的建模能力")):
            canvas.text(px(660), px(600 + index * 52), px(236), px(48),
                        f"•  {item}", size=8.5, align="left", valign="top",
                        color=INK, name=f"FIG_T_m2_p{index}")

    # ---- 模块三 ----
    with canvas.group("m3"):
        canvas.text(px(916), px(566), px(610), px(20),
                    "模块三：CGAFusion（高低频自适应融合）", size=10,
                    color=BLUE, bold=True, name="FIG_T_m3_title")
        canvas.text(px(930), px(598), px(130), px(18), "输入特征 Fh（高频）",
                    size=7.5, align="left", color=INK, name="FIG_T_m3_fh")
        grid(canvas, "FIG_G_m3_fh", 946, 618, 64, 46)
        canvas.round_rect(px(1046), px(620), px(96), px(40),
                          fill=(235, 242, 255), line=BLUE, weight=0.9,
                          name="FIG_E_m3_high")
        canvas.text(px(1046), px(624), px(96), px(32), "高频分支\n（细节信息）",
                    size=7.5, color=BLUE, name="FIG_T_m3_high")
        canvas.text(px(930), px(682), px(130), px(18), "输入特征 Fl（低频）",
                    size=7.5, align="left", color=INK, name="FIG_T_m3_fl")
        grid(canvas, "FIG_G_m3_fl", 946, 702, 64, 46, fill=GRID_BG_GREEN,
             line=GRID_LINE_GREEN)
        canvas.round_rect(px(1046), px(704), px(96), px(40),
                          fill=(235, 248, 238), line=GREEN, weight=0.9,
                          name="FIG_E_m3_low")
        canvas.text(px(1046), px(708), px(96), px(32), "低频分支\n（结构信息）",
                    size=7.5, color=GREEN, name="FIG_T_m3_low")
        badge(canvas, "FIG_E_m3_add", 1190, 682, 14, (52, 58, 78),
              (30, 34, 48), "+", WHITE, size=9)
        grid(canvas, "FIG_G_m3_out", 1166, 730, 52, 40, 3, 2)
        canvas.text(px(1140), px(776), px(110), px(18), "输出特征", size=8,
                    color=INK, name="FIG_T_m3_out")
        ortho(canvas, "FIG_L_m3_f1", [(1142, 640), (1176, 640), (1176, 672)],
              color=INK, dash=False, weight=1.0)
        arrow(canvas, "FIG_L_m3_f2", 1142, 724, 1176, 690, weight=1.0)
        arrow(canvas, "FIG_L_m3_f3", 1190, 696, 1190, 730, weight=1.0)
        for index, item in enumerate(("分离高频细节与低频结构信息",
                                      "空间联合注意力自适应融合",
                                      "突出关键细节，抑制噪声干扰",
                                      "提升特征表达与泛化能力")):
            canvas.text(px(1268), px(600 + index * 52), px(250), px(44),
                        f"•  {item}", size=8.5, align="left", valign="top",
                        color=INK, name=f"FIG_T_m3_p{index}")

    # ---- 底行 ----
    with canvas.group("bottom"):
        canvas.text(px(12), px(860), px(322), px(20), "输出多尺度特征金字塔",
                    size=10, color=BLUE, bold=True, name="FIG_T_b1_title")
        for name, x, y, w, h, base, text, color in (
                ("FIG_E_out_p3", 46, 896, 56, 44, CUBE_ORANGE,
                 "P3\n（80×C）", RED),
                ("FIG_E_out_p4", 142, 904, 46, 36, CUBE_GREEN,
                 "P4\n（40×C）", GREEN),
                ("FIG_E_out_p5", 226, 912, 38, 28, CUBE_BLUE,
                 "P5\n（20×C）", BLUE)):
            cuboid(canvas, name, x, y, w, h, base, 8)
            canvas.text(px(x - 12), px(y + h + 22), px(w + 34), px(32), text,
                        size=8, color=color, bold=True, name=f"{name}_label")
        canvas.text(px(348), px(860), px(542), px(20), "输入到 RT-DETR Head",
                    size=10, color=INK, bold=True, name="FIG_T_b2_title")
        canvas.text(px(380), px(892), px(130), px(34),
                    "IoU-Aware\nQuery Selection", size=8, color=INK,
                    name="FIG_T_b2_iou")
        square_row(canvas, "FIG_E_b2_sq1", 386, 934, 6, 12, 5)
        canvas.text(px(560), px(880), px(240), px(18),
                    "Transformer Decoder × 6 层", size=8.5, color=INK,
                    name="FIG_T_b2_dec")
        square_row(canvas, "FIG_E_b2_sq2", 584, 922, 5, 16, 9,
                   fill=(140, 110, 214))
        canvas.rect(px(792), px(922), px(16), px(16), fill=(140, 110, 214),
                    line=(110, 82, 178), weight=0.6, name="FIG_E_b2_last")
        arrow(canvas, "FIG_L_b2_flow", 516, 930, 584, 930, weight=1.2)
        arrow(canvas, "FIG_L_b1_b2", 334, 934, 348, 934, color=BLUE,
              weight=4.0)
        route("FIG_L_b1_b2", "FIG_E_panel_b2", "left")
        arrow(canvas, "FIG_L_b2_b3", 890, 934, 908, 934, color=BLUE,
              weight=4.0)
        route("FIG_L_b2_b3", "FIG_E_panel_b3", "left")
        canvas.text(px(908), px(860), px(226), px(20), "检测结果", size=10,
                    color=INK, bold=True, name="FIG_T_b3_title")
        mini_bars(canvas, "FIG_E_b3_bars", 936, 896)
        canvas.text(px(916), px(946), px(100), px(18), "类别概率", size=7.5,
                    color=INK, name="FIG_T_b3_bars")
        canvas.rect(px(1030), px(892), px(46), px(36), fill=None, line=RED,
                    weight=1.2, name="FIG_E_b3_box1")
        canvas.rect(px(1048), px(908), px(46), px(36), fill=None, line=GREEN,
                    weight=1.2, name="FIG_E_b3_box2")
        canvas.text(px(1010), px(946), px(120), px(30), "归一化边界\n框坐标",
                    size=7.5, color=INK, name="FIG_T_b3_box")
        canvas.text(px(1148), px(860), px(378), px(20), "三模块协同优势",
                    size=10, color=BLUE, bold=True, name="FIG_T_b4_title")
        for index, (head, tail, color) in enumerate(
                (("更精细：", "保留小目标细粒度特征", BLUE),
                 ("更鲁棒：", "增强方向性缺陷感知", GREEN),
                 ("更高效：", "自适应融合多尺度特征", RED))):
            y = 890 + index * 40
            canvas.ellipse(px(1170), px(y - 2), px(20), px(20), fill=color,
                           line=None, name=f"FIG_E_b4_icon{index}")
            canvas.text(px(1200), px(y - 2), px(70), px(22), head, size=8.5,
                        color=color, bold=True, align="left",
                        name=f"FIG_T_b4_head{index}")
            canvas.text(px(1268), px(y - 2), px(250), px(22), tail, size=8.5,
                        align="left", color=INK, name=f"FIG_T_b4_tail{index}")

    out = HERE / "mphf_net_v2.pptx"
    canvas.save(out)
    canvas.export_scene_manifest(HERE / "scene-manifest.json")
    canvas.emit_vba(HERE / "mphf_net_v2.bas")
    _write_manifests()
    return out


def _write_manifests() -> None:
    cuboids = []
    for prefix, geometry in CUBOID_GEOMETRY.items():
        entry = {
            "id": prefix.removeprefix("FIG_E_"),
            "front": f"{prefix}_front",
            "top": f"{prefix}_top",
            "right": f"{prefix}_right",
        }
        cuboids.append(entry)
    layering = {
        "layering_audit": {
            "cuboids": cuboids,
            "size_order": ["bb1", "bb2", "bb3", "bb4"],
            "z_order": ["bb1", "bb2", "bb3", "bb4"],
        }
    }
    (HERE / "layering_manifest.json").write_text(
        json.dumps(layering, ensure_ascii=False, indent=2), encoding="utf-8")
    routing = {
        "routing_audit": {
            "routes": ROUTES,
            "ignore_text_shapes": IGNORE_TEXT,
            "ignore_dangling": IGNORE_DANGLING,
        }
    }
    (HERE / "routing_manifest.json").write_text(
        json.dumps(routing, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    print(build())
