from __future__ import annotations

import importlib.util
import tempfile
import unittest
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.util import Pt


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "scripts" / "pptx_layering_audit.py"


def load_module():
    if not MODULE_PATH.is_file():
        raise AssertionError("缺少 scripts/pptx_layering_audit.py")
    spec = importlib.util.spec_from_file_location("pptx_layering_audit", MODULE_PATH)
    module = importlib.util.module_from_spec(spec)
    assert spec and spec.loader
    spec.loader.exec_module(module)
    return module


class LayeringAuditTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    @staticmethod
    def _face(slide, name, x, y, w, h, color):
        shape = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Pt(x), Pt(y), Pt(w), Pt(h)
        )
        shape.name = name
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(*color)
        return shape

    def _add_cuboid(self, slide, prefix, x, y, w, h, colors=None):
        colors = colors or {
            "front": (180, 210, 250),
            "top": (220, 235, 255),
            "right": (120, 165, 225),
        }
        self._face(slide, f"{prefix}_front", x, y + 10, w, h, colors["front"])
        self._face(slide, f"{prefix}_top", x, y, w, 10, colors["top"])
        self._face(slide, f"{prefix}_right", x + w, y + 10, 10, h, colors["right"])

    def _manifest(self):
        return {
            "cuboids": [
                {
                    "id": "small",
                    "front": "SUMMER_E_small_front",
                    "top": "SUMMER_E_small_top",
                    "right": "SUMMER_E_small_right",
                },
                {
                    "id": "large",
                    "front": "SUMMER_E_large_front",
                    "top": "SUMMER_E_large_top",
                    "right": "SUMMER_E_large_right",
                },
            ],
            "size_order": ["small", "large"],
            "z_order": ["large", "small"],
        }

    def test_valid_layered_cuboids_pass(self) -> None:
        module = load_module()
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        self._add_cuboid(slide, "SUMMER_E_large", 80, 80, 80, 70)
        self._add_cuboid(slide, "SUMMER_E_small", 100, 40, 40, 35)
        path = self.root / "valid.pptx"
        presentation.save(path)
        result = module.audit_presentation(str(path), manifest=self._manifest())
        self.assertTrue(result["passed"])

    def test_missing_face_fails(self) -> None:
        module = load_module()
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        self._add_cuboid(slide, "SUMMER_E_large", 80, 80, 80, 70)
        self._face(slide, "SUMMER_E_small_front", 100, 50, 40, 35, (180, 210, 250))
        self._face(slide, "SUMMER_E_small_top", 100, 40, 40, 10, (220, 235, 255))
        path = self.root / "missing.pptx"
        presentation.save(path)
        result = module.audit_presentation(str(path), manifest=self._manifest())
        self.assertFalse(result["passed"])
        self.assertEqual(result["missing_faces"][0]["face"], "right")

    def test_flat_same_color_faces_fail(self) -> None:
        module = load_module()
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        flat = {face: (180, 210, 250) for face in ("front", "top", "right")}
        self._add_cuboid(slide, "SUMMER_E_large", 80, 80, 80, 70, flat)
        self._add_cuboid(slide, "SUMMER_E_small", 100, 40, 40, 35)
        path = self.root / "flat.pptx"
        presentation.save(path)
        result = module.audit_presentation(str(path), manifest=self._manifest())
        self.assertFalse(result["passed"])
        self.assertEqual(result["face_color_errors"][0]["cuboid"], "large")

    def test_wrong_z_order_fails(self) -> None:
        module = load_module()
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        self._add_cuboid(slide, "SUMMER_E_small", 100, 40, 40, 35)
        self._add_cuboid(slide, "SUMMER_E_large", 80, 80, 80, 70)
        path = self.root / "z-order.pptx"
        presentation.save(path)
        result = module.audit_presentation(str(path), manifest=self._manifest())
        self.assertFalse(result["passed"])
        self.assertEqual(result["z_order_errors"][0]["before"], "large")

    def test_expected_face_bounds_mismatch_fails(self) -> None:
        module = load_module()
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        self._add_cuboid(slide, "SUMMER_E_large", 80, 80, 80, 70)
        self._add_cuboid(slide, "SUMMER_E_small", 100, 40, 40, 35)
        path = self.root / "geometry-reference.pptx"
        presentation.save(path)
        manifest = self._manifest()
        manifest["cuboids"][0]["expected_face_bounds_pt"] = {
            "front": [110, 50, 40, 35],
            "top": [100, 40, 40, 10],
            "right": [140, 50, 10, 35],
        }
        manifest["cuboids"][0]["bounds_tolerance_pt"] = 0.5
        result = module.audit_presentation(str(path), manifest=manifest)
        self.assertFalse(result["passed"])
        self.assertEqual(result["geometry_reference_errors"][0]["face"], "front")

    def test_required_gradient_face_fails_when_solid(self) -> None:
        module = load_module()
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        self._add_cuboid(slide, "SUMMER_E_large", 80, 80, 80, 70)
        self._add_cuboid(slide, "SUMMER_E_small", 100, 40, 40, 35)
        path = self.root / "gradient-required.pptx"
        presentation.save(path)
        manifest = self._manifest()
        manifest["cuboids"][0]["gradient_faces"] = ["front"]
        result = module.audit_presentation(str(path), manifest=manifest)
        self.assertFalse(result["passed"])
        self.assertEqual(result["gradient_errors"][0]["face"], "front")


if __name__ == "__main__":
    unittest.main()
