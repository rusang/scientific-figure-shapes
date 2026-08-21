"""
scipilot-figure-skill :: fidelity_audit.py
================================================
对“参考图 → 重绘图”做分区域高保真审计。

全图平均差异会稀释小图标、局部标注和单个 panel 的缺失。本脚本同时检查：

- 全图颜色差异与边缘差异；
- 网格分区差异，定位局部异常；
- manifest 中逐项登记的显著元素，拦截遗漏和未经批准的简化。

该脚本只用于用户提供参考图、要求复刻布局/风格或明确要求高保真时；普通数据图
不需要参考图差异审计。
"""
from __future__ import annotations

import argparse
import json
import os
import sys
from pathlib import Path
from typing import Any


ALLOWED_STATUSES = {"matched", "preserved", "simplified", "omitted"}
PASS_STATUSES = {"matched", "preserved"}


def _load_rgb(path: str):
    try:
        from PIL import Image
    except ImportError as exc:
        raise RuntimeError("fidelity_audit 需要 Pillow：pip install Pillow") from exc
    return Image.open(path).convert("RGB")


def _mae(a, b) -> float:
    from PIL import ImageChops, ImageStat

    stat = ImageStat.Stat(ImageChops.difference(a, b))
    return sum(stat.mean) / (3 * 255)


def _edge_mae(a, b) -> float:
    from PIL import ImageChops, ImageFilter, ImageStat

    edge_a = a.convert("L").filter(ImageFilter.FIND_EDGES)
    edge_b = b.convert("L").filter(ImageFilter.FIND_EDGES)
    return ImageStat.Stat(ImageChops.difference(edge_a, edge_b)).mean[0] / 255


def _bbox(value: Any, size: tuple[int, int]) -> tuple[int, int, int, int]:
    if not isinstance(value, (list, tuple)) or len(value) != 4:
        raise ValueError("bbox_px 必须是 [x, y, width, height]")
    x, y, w, h = (int(round(float(v))) for v in value)
    if w <= 0 or h <= 0:
        raise ValueError("bbox_px 的 width/height 必须为正数")
    image_w, image_h = size
    left = max(0, min(image_w, x))
    top = max(0, min(image_h, y))
    right = max(0, min(image_w, x + w))
    bottom = max(0, min(image_h, y + h))
    if right <= left or bottom <= top:
        raise ValueError("bbox_px 位于画布之外")
    return left, top, right, bottom


def _severity(value: float, warn: float, fail: float) -> str:
    if value >= fail:
        return "FAIL"
    if value >= warn:
        return "WARN"
    return "PASS"


def _tile_metrics(reference, rendered, rows: int, cols: int,
                  warn: float, fail: float) -> list[dict[str, Any]]:
    width, height = reference.size
    tiles: list[dict[str, Any]] = []
    for row in range(rows):
        top = round(row * height / rows)
        bottom = round((row + 1) * height / rows)
        for col in range(cols):
            left = round(col * width / cols)
            right = round((col + 1) * width / cols)
            box = (left, top, right, bottom)
            value = _mae(reference.crop(box), rendered.crop(box))
            tiles.append({
                "row": row,
                "col": col,
                "bbox_px": [left, top, right - left, bottom - top],
                "mae": round(value, 6),
                "severity": _severity(value, warn, fail),
            })
    return tiles


def _load_manifest(path: str | None) -> dict[str, Any] | None:
    if path is None:
        return None
    data = json.loads(Path(path).read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ValueError("manifest 顶层必须是 JSON object")
    return data


def _mask_approved_exceptions(manifest: dict[str, Any] | None,
                              reference, rendered) -> tuple[Any, list[str]]:
    """从全图/分区差异中排除用户已批准的简化或省略区域。"""
    adjusted = reference.copy()
    approved_ids: list[str] = []
    if manifest is None or not isinstance(manifest.get("elements"), list):
        return adjusted, approved_ids
    for item in manifest["elements"]:
        if not isinstance(item, dict):
            continue
        status = str(item.get("status", "")).strip().lower()
        if status in PASS_STATUSES or not bool(item.get("approved", False)):
            continue
        try:
            ref_box = _bbox(item.get("bbox_px"), reference.size)
            rendered_box = _bbox(
                item.get("rendered_bbox_px", item.get("bbox_px")), rendered.size
            )
        except (TypeError, ValueError):
            continue
        ref_size = (ref_box[2] - ref_box[0], ref_box[3] - ref_box[1])
        rendered_roi = rendered.crop(rendered_box).resize(ref_size)
        adjusted.paste(rendered_roi, (ref_box[0], ref_box[1]))
        approved_ids.append(str(item.get("id", "")))
    return adjusted, approved_ids


def _audit_elements(manifest: dict[str, Any] | None, reference, rendered,
                    warn: float, fail: float) -> tuple[list[dict], list[dict]]:
    if manifest is None:
        return [], [{
            "severity": "WARN",
            "code": "manifest_missing",
            "message": "未提供显著元素 manifest；小元素遗漏只能依靠分区差异和人工读图发现。",
        }]

    elements = manifest.get("elements")
    if not isinstance(elements, list) or not elements:
        return [], [{
            "severity": "FAIL",
            "code": "elements_missing",
            "message": "高保真 manifest 必须包含非空 elements 列表。",
        }]

    reports: list[dict] = []
    issues: list[dict] = []
    seen: set[str] = set()
    for index, item in enumerate(elements):
        if not isinstance(item, dict):
            issues.append({
                "severity": "FAIL", "code": "element_invalid",
                "message": f"elements[{index}] 必须是 object。",
            })
            continue
        element_id = str(item.get("id", "")).strip()
        if not element_id:
            issues.append({
                "severity": "FAIL", "code": "element_id_missing",
                "message": f"elements[{index}] 缺少 id。",
            })
            continue
        if element_id in seen:
            issues.append({
                "severity": "FAIL", "code": "element_id_duplicate",
                "message": f"显著元素 id 重复：{element_id}",
            })
            continue
        seen.add(element_id)

        status = str(item.get("status", "")).strip().lower()
        approved = bool(item.get("approved", False))
        salience = str(item.get("salience", "normal")).strip().lower()
        if status not in ALLOWED_STATUSES:
            issues.append({
                "severity": "FAIL", "code": "element_status_invalid",
                "message": f"{element_id} 的 status 必须是 {sorted(ALLOWED_STATUSES)} 之一。",
            })
            continue
        if status not in PASS_STATUSES and not approved:
            issues.append({
                "severity": "FAIL", "code": "unapproved_substitution",
                "message": f"{element_id} 被标记为 {status}，但没有用户批准；语义相近不等于高保真。",
            })

        try:
            ref_box = _bbox(item.get("bbox_px"), reference.size)
            rendered_box = _bbox(item.get("rendered_bbox_px", item.get("bbox_px")), rendered.size)
        except (TypeError, ValueError) as exc:
            issues.append({
                "severity": "FAIL", "code": "element_bbox_invalid",
                "message": f"{element_id}: {exc}",
            })
            continue

        ref_roi = reference.crop(ref_box)
        rendered_roi = rendered.crop(rendered_box).resize(ref_roi.size)
        value = _mae(ref_roi, rendered_roi)
        edge_value = _edge_mae(ref_roi, rendered_roi)
        element_warn = warn * (0.75 if salience == "high" else 1.0)
        element_fail = fail * (0.75 if salience == "high" else 1.0)
        severity = max(
            (_severity(value, element_warn, element_fail),
             _severity(edge_value, element_warn, element_fail)),
            key={"PASS": 0, "WARN": 1, "FAIL": 2}.get,
        )
        if status not in PASS_STATUSES and approved:
            severity = "PASS"
        report = {
            "id": element_id,
            "label": item.get("label", element_id),
            "status": status,
            "approved": approved,
            "salience": salience,
            "mae": round(value, 6),
            "edge_mae": round(edge_value, 6),
            "severity": severity,
        }
        if status not in PASS_STATUSES and approved:
            report["exception"] = "user_approved"
        reports.append(report)
        if severity != "PASS":
            issues.append({
                "severity": severity,
                "code": "element_visual_delta",
                "message": f"{element_id} 局部差异过大：mae={value:.4f}, edge={edge_value:.4f}",
            })
    return reports, issues


def _write_heatmap(reference, rendered, tiles: list[dict], out_path: str) -> str:
    from PIL import Image, ImageChops, ImageDraw, ImageEnhance, ImageOps

    diff = ImageChops.difference(reference, rendered).convert("L")
    diff = ImageEnhance.Contrast(diff).enhance(3.0)
    heat = ImageOps.colorize(diff, black=(255, 255, 255), white=(220, 20, 20)).convert("RGB")
    overlay = Image.blend(reference, heat, 0.45)
    draw = ImageDraw.Draw(overlay)
    colors = {"WARN": (255, 165, 0), "FAIL": (220, 0, 0)}
    for tile in tiles:
        if tile["severity"] == "PASS":
            continue
        x, y, w, h = tile["bbox_px"]
        draw.rectangle((x, y, x + w - 1, y + h - 1), outline=colors[tile["severity"]], width=3)
    Path(out_path).parent.mkdir(parents=True, exist_ok=True)
    overlay.save(out_path)
    return str(Path(out_path))


def audit_reference(reference_path: str, rendered_path: str, *,
                    manifest_path: str | None = None,
                    rows: int = 4, cols: int = 4,
                    warn_mae: float = 0.08, fail_mae: float = 0.16,
                    heatmap_path: str | None = None) -> dict[str, Any]:
    if rows <= 0 or cols <= 0:
        raise ValueError("rows/cols 必须为正整数")
    if not (0 <= warn_mae < fail_mae <= 1):
        raise ValueError("阈值必须满足 0 <= warn_mae < fail_mae <= 1")
    if not os.path.isfile(reference_path) or not os.path.isfile(rendered_path):
        raise FileNotFoundError("参考图或重绘图不存在")

    reference = _load_rgb(reference_path)
    rendered_original = _load_rgb(rendered_path)
    issues: list[dict[str, Any]] = []
    if rendered_original.size != reference.size:
        issues.append({
            "severity": "WARN",
            "code": "size_mismatch",
            "message": f"重绘图尺寸 {rendered_original.size} 与参考图 {reference.size} 不一致，已缩放后比较。",
        })
    rendered = rendered_original.resize(reference.size)
    manifest = _load_manifest(manifest_path)
    comparison_reference, approved_exceptions = _mask_approved_exceptions(
        manifest, reference, rendered
    )
    global_mae = _mae(comparison_reference, rendered)
    edge_mae = _edge_mae(comparison_reference, rendered)
    global_severity = max(
        (_severity(global_mae, warn_mae, fail_mae),
         _severity(edge_mae, warn_mae, fail_mae)),
        key={"PASS": 0, "WARN": 1, "FAIL": 2}.get,
    )
    if global_severity != "PASS":
        issues.append({
            "severity": global_severity,
            "code": "global_visual_delta",
            "message": f"全图差异：mae={global_mae:.4f}, edge={edge_mae:.4f}",
        })

    tiles = _tile_metrics(
        comparison_reference, rendered, rows, cols, warn_mae, fail_mae
    )
    for tile in tiles:
        if tile["severity"] != "PASS":
            issues.append({
                "severity": tile["severity"],
                "code": "tile_visual_delta",
                "message": f"分区 r{tile['row']}c{tile['col']} 差异 mae={tile['mae']:.4f}",
            })

    elements, element_issues = _audit_elements(
        manifest, reference, rendered, warn_mae, fail_mae
    )
    issues.extend(element_issues)

    ranks = {"PASS": 0, "WARN": 1, "FAIL": 2}
    verdict = "PASS"
    if issues:
        verdict = max((issue["severity"] for issue in issues), key=ranks.get)
    result: dict[str, Any] = {
        "passed": verdict != "FAIL",
        "verdict": verdict,
        "reference_size_px": list(reference.size),
        "rendered_size_px": list(rendered_original.size),
        "global": {
            "mae": round(global_mae, 6),
            "edge_mae": round(edge_mae, 6),
            "severity": global_severity,
        },
        "grid": {"rows": rows, "cols": cols, "tiles": tiles},
        "elements": elements,
        "approved_exceptions": approved_exceptions,
        "issues": issues,
    }
    if heatmap_path:
        result["heatmap_path"] = _write_heatmap(
            comparison_reference, rendered, tiles, heatmap_path
        )
    return result


def _cli() -> int:
    parser = argparse.ArgumentParser(description="参考图分区域高保真审计")
    parser.add_argument("reference")
    parser.add_argument("rendered")
    parser.add_argument("--manifest")
    parser.add_argument("--rows", type=int, default=4)
    parser.add_argument("--cols", type=int, default=4)
    parser.add_argument("--warn-mae", type=float, default=0.08)
    parser.add_argument("--fail-mae", type=float, default=0.16)
    parser.add_argument("--heatmap")
    parser.add_argument("--json-out")
    parser.add_argument("--pretty", action="store_true")
    parser.add_argument("--strict", action="store_true", help="WARN 也返回非零退出码")
    args = parser.parse_args()

    try:
        result = audit_reference(
            args.reference,
            args.rendered,
            manifest_path=args.manifest,
            rows=args.rows,
            cols=args.cols,
            warn_mae=args.warn_mae,
            fail_mae=args.fail_mae,
            heatmap_path=args.heatmap,
        )
    except Exception as exc:
        result = {"passed": False, "verdict": "FAIL", "issues": [{
            "severity": "FAIL", "code": "runtime_error", "message": str(exc)
        }]}
    payload = json.dumps(result, ensure_ascii=False, indent=2 if args.pretty else None)
    print(payload)
    if args.json_out:
        Path(args.json_out).parent.mkdir(parents=True, exist_ok=True)
        Path(args.json_out).write_text(payload + "\n", encoding="utf-8")
    if result.get("verdict") == "FAIL":
        return 2
    if args.strict and result.get("verdict") == "WARN":
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(_cli())
