from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "scripts" / "preserve_cropper.py"


def load_module():
    spec = importlib.util.spec_from_file_location("preserve_cropper", MODULE_PATH)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class PreserveCropperTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.source = self.root / "source.png"
        Image.new("RGBA", (100, 80), (255, 255, 255, 255)).save(self.source)
        self.regions = [
            {"name": "target icon", "x": 10, "y": 12, "width": 20, "height": 18}
        ]

    def tearDown(self) -> None:
        self.temp.cleanup()

    def test_start_index_is_stable_and_manifest_records_asset_id(self) -> None:
        module = load_module()
        result = module.crop_regions(
            self.source, self.root / "assets", self.regions, start_index=5
        )
        asset = result["assets"][0]
        self.assertEqual(Path(asset["asset_path"]).name, "05_target_icon.png")
        self.assertEqual(asset["asset_id"], "preserved_005")

    def test_existing_assets_refuse_overwrite_unless_explicit(self) -> None:
        module = load_module()
        output = self.root / "assets"
        module.crop_regions(self.source, output, self.regions)
        with self.assertRaises(FileExistsError):
            module.crop_regions(self.source, output, self.regions)
        result = module.crop_regions(self.source, output, self.regions, overwrite=True)
        self.assertTrue(result["passed"])


if __name__ == "__main__":
    unittest.main()
