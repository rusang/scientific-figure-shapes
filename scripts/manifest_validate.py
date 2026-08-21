#!/usr/bin/env python3
"""Validate the unified scientific-figure reconstruction manifest."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


ALLOWED_ELEMENT_STATUSES = {"matched", "preserved", "simplified", "omitted"}
ALLOWED_ROUTE_EDGES = {"left", "right", "top", "bottom", "center"}


def _error(errors: list[dict[str, Any]], code: str, path: str, message: str) -> None:
    errors.append({"code": code, "path": path, "message": message})


def _valid_bbox(value: Any) -> bool:
    if not isinstance(value, list) or len(value) != 4:
        return False
    try:
        _, _, width, height = (float(item) for item in value)
    except (TypeError, ValueError):
        return False
    return width > 0 and height > 0


def validate_manifest(data: Any) -> dict[str, Any]:
    errors: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    if not isinstance(data, dict):
        return {
            "passed": False,
            "schema_version": None,
            "errors": [{"code": "manifest_not_object", "path": "$", "message": "manifest 顶层必须是 object"}],
            "warnings": [],
        }
    version = data.get("schema_version")
    if version != "1.0":
        _error(errors, "schema_version_invalid", "$.schema_version", "当前仅支持 schema_version=1.0")

    canvas = data.get("canvas")
    if canvas is not None:
        if not isinstance(canvas, dict):
            _error(errors, "canvas_invalid", "$.canvas", "canvas 必须是 object")
        else:
            for key in ("width_px", "height_px"):
                try:
                    valid = float(canvas.get(key, 0)) > 0
                except (TypeError, ValueError):
                    valid = False
                if not valid:
                    _error(errors, "canvas_dimension_invalid", f"$.canvas.{key}", f"{key} 必须为正数")
    else:
        warnings.append({"code": "canvas_missing", "path": "$.canvas", "message": "建议登记参考画布尺寸"})

    elements = data.get("elements", [])
    seen: set[str] = set()
    if not isinstance(elements, list):
        _error(errors, "elements_invalid", "$.elements", "elements 必须是 list")
        elements = []
    for index, item in enumerate(elements):
        path = f"$.elements[{index}]"
        if not isinstance(item, dict):
            _error(errors, "element_invalid", path, "element 必须是 object")
            continue
        element_id = str(item.get("id", "")).strip()
        if not element_id:
            _error(errors, "element_id_missing", f"{path}.id", "element id 不能为空")
        elif element_id in seen:
            _error(errors, "element_id_duplicate", f"{path}.id", f"element id 重复：{element_id}")
        else:
            seen.add(element_id)
        if "bbox_px" in item and not _valid_bbox(item["bbox_px"]):
            _error(errors, "element_bbox_invalid", f"{path}.bbox_px", "bbox_px 必须是 [x,y,w,h] 且 w/h>0")
        status = item.get("status")
        if status is not None and status not in ALLOWED_ELEMENT_STATUSES:
            _error(errors, "element_status_invalid", f"{path}.status", "status 值无效")
        if status in {"simplified", "omitted"} and not bool(item.get("approved", False)):
            warnings.append(
                {"code": "unapproved_visual_exception", "path": path, "message": "简化/省略项未标记用户批准"}
            )

    routing = data.get("routing_audit", {})
    if routing is not None and not isinstance(routing, dict):
        _error(errors, "routing_audit_invalid", "$.routing_audit", "routing_audit 必须是 object")
        routing = {}
    routes = routing.get("routes", []) if isinstance(routing, dict) else []
    if not isinstance(routes, list):
        _error(errors, "routes_invalid", "$.routing_audit.routes", "routes 必须是 list")
        routes = []
    for index, route in enumerate(routes):
        path = f"$.routing_audit.routes[{index}]"
        if not isinstance(route, dict):
            _error(errors, "route_invalid", path, "route 必须是 object")
            continue
        if not str(route.get("connector", "")).strip():
            _error(errors, "route_connector_missing", f"{path}.connector", "connector 不能为空")
        if not str(route.get("target", "")).strip():
            _error(errors, "route_target_missing", f"{path}.target", "target 不能为空")
        edge = route.get("target_edge")
        if edge is not None and edge not in ALLOWED_ROUTE_EDGES:
            _error(errors, "route_edge_invalid", f"{path}.target_edge", "target_edge 值无效")

    for section, list_key in (("text_audit", "items"), ("layering_audit", "cuboids")):
        value = data.get(section)
        if value is not None and not isinstance(value, dict):
            _error(errors, f"{section}_invalid", f"$.{section}", f"{section} 必须是 object")
        elif isinstance(value, dict) and list_key in value and not isinstance(value[list_key], list):
            _error(errors, f"{list_key}_invalid", f"$.{section}.{list_key}", f"{list_key} 必须是 list")
    editability = data.get("editability_audit")
    if editability is not None and not isinstance(editability, dict):
        _error(errors, "editability_audit_invalid", "$.editability_audit", "editability_audit 必须是 object")

    return {
        "passed": not errors,
        "schema_version": version,
        "errors": errors,
        "warnings": warnings,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate a scientific-figure manifest.")
    parser.add_argument("manifest")
    parser.add_argument("--json-out")
    parser.add_argument("--pretty", action="store_true")
    args = parser.parse_args()
    try:
        data = json.loads(Path(args.manifest).read_text(encoding="utf-8"))
        result = validate_manifest(data)
    except Exception as exc:
        result = {"passed": False, "errors": [{"code": "runtime_error", "message": str(exc)}], "warnings": []}
    payload = json.dumps(result, ensure_ascii=False, indent=2 if args.pretty else None)
    print(payload)
    if args.json_out:
        target = Path(args.json_out)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(payload + "\n", encoding="utf-8")
    return 0 if result["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
