from __future__ import annotations

import importlib.util
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "scripts" / "manifest_validate.py"


def load_module():
    if not MODULE_PATH.is_file():
        raise AssertionError("缺少 scripts/manifest_validate.py")
    spec = importlib.util.spec_from_file_location("manifest_validate", MODULE_PATH)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class ManifestValidateTests(unittest.TestCase):
    def test_unified_manifest_passes(self) -> None:
        module = load_module()
        manifest = {
            "schema_version": "1.0",
            "canvas": {"width_px": 1600, "height_px": 1000},
            "elements": [
                {
                    "id": "fusion",
                    "type": "shape",
                    "bbox_px": [600, 300, 120, 50],
                    "status": "matched",
                }
            ],
            "routing_audit": {
                "routes": [
                    {
                        "connector": "FIG_L_to_fusion",
                        "target": "FIG_E_fusion",
                        "target_edge": "left",
                    }
                ],
                "segments": [
                    {"connector": "FIG_L_skip_h", "dash": True, "orientation": "horizontal"}
                ],
            },
            "text_audit": {"items": [{"name": "FIG_T_fusion", "text": "融合"}]},
            "editability_audit": {"allowed_raster_shapes": []},
        }
        result = module.validate_manifest(manifest)
        self.assertTrue(result["passed"])
        self.assertEqual(result["errors"], [])

    def test_duplicate_elements_invalid_bbox_and_bad_route_fail(self) -> None:
        module = load_module()
        manifest = {
            "schema_version": "1.0",
            "canvas": {"width_px": 100, "height_px": 100},
            "elements": [
                {"id": "x", "bbox_px": [0, 0, -5, 10], "status": "matched"},
                {"id": "x", "bbox_px": [1, 1, 5, 5], "status": "matched"},
            ],
            "routing_audit": {
                "routes": [{"connector": "FIG_L_1", "target": "", "target_edge": "diagonal"}],
                "segments": [{"connector": "", "orientation": "zigzag"}],
            },
        }
        result = module.validate_manifest(manifest)
        codes = {error["code"] for error in result["errors"]}
        self.assertFalse(result["passed"])
        self.assertEqual(
            {
                "element_bbox_invalid", "element_id_duplicate",
                "route_target_missing", "route_edge_invalid",
                "segment_connector_missing", "segment_orientation_invalid",
            },
            codes,
        )


if __name__ == "__main__":
    unittest.main()
