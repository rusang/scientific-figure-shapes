from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image, ImageDraw


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "scripts" / "dual_renderer_probe.py"


def load_module():
    if not MODULE_PATH.is_file():
        raise AssertionError("缺少 scripts/dual_renderer_probe.py")
    spec = importlib.util.spec_from_file_location("dual_renderer_probe", MODULE_PATH)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class DualRendererProbeTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_identical_renderer_outputs_pass(self) -> None:
        module = load_module()
        first = self.root / "powerpoint.png"
        second = self.root / "libreoffice.png"
        Image.new("RGB", (120, 80), "white").save(first)
        Image.new("RGB", (120, 80), "white").save(second)

        result = module.compare_renderers(
            first, second, renderer_a="PowerPoint", renderer_b="LibreOffice"
        )

        self.assertTrue(result["passed"])
        self.assertEqual(result["mean_abs_delta"], 0)

    def test_material_renderer_delta_fails_and_writes_heatmap(self) -> None:
        module = load_module()
        first = self.root / "powerpoint.png"
        second = self.root / "libreoffice.png"
        heatmap = self.root / "renderer-delta.png"
        Image.new("RGB", (120, 80), "white").save(first)
        changed = Image.new("RGB", (120, 80), "white")
        ImageDraw.Draw(changed).rectangle((10, 10, 110, 70), fill=(0, 0, 120))
        changed.save(second)

        result = module.compare_renderers(
            first,
            second,
            max_mean_delta=0.02,
            heatmap_path=heatmap,
        )

        self.assertFalse(result["passed"])
        self.assertEqual(result["verdict"], "FAIL")
        self.assertTrue(heatmap.is_file())


if __name__ == "__main__":
    unittest.main()
