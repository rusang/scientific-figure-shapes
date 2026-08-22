from __future__ import annotations

import importlib.util
import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from PIL import Image, ImageDraw, ImageFilter


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "scripts" / "fidelity_audit.py"


def load_module():
    if not MODULE_PATH.is_file():
        raise AssertionError("缺少 scripts/fidelity_audit.py")
    spec = importlib.util.spec_from_file_location("fidelity_audit", MODULE_PATH)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class FidelityAuditTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def _manifest(self, bbox: list[int]) -> Path:
        path = self.root / "manifest.json"
        path.write_text(
            json.dumps(
                {
                    "schema_version": "1.0",
                    "elements": [
                        {
                            "id": "icon",
                            "label": "small icon",
                            "bbox_px": bbox,
                            "status": "matched",
                            "salience": "high",
                        }
                    ],
                }
            ),
            encoding="utf-8",
        )
        return path

    def test_identical_images_pass_and_emit_heatmap(self) -> None:
        module = load_module()
        reference = self.root / "reference.png"
        rendered = self.root / "rendered.png"
        heatmap = self.root / "heatmap.png"
        Image.new("RGB", (64, 64), "white").save(reference)
        Image.new("RGB", (64, 64), "white").save(rendered)

        result = module.audit_reference(
            str(reference),
            str(rendered),
            manifest_path=str(self._manifest([10, 10, 12, 12])),
            heatmap_path=str(heatmap),
        )

        self.assertTrue(result["passed"])
        self.assertEqual(result["verdict"], "PASS")
        self.assertTrue(heatmap.is_file())

    def test_salient_roi_failure_is_not_diluted_by_global_average(self) -> None:
        module = load_module()
        reference = self.root / "reference.png"
        rendered = self.root / "rendered.png"
        ref = Image.new("RGB", (200, 200), "white")
        ImageDraw.Draw(ref).rectangle((98, 98, 103, 103), fill=(0, 20, 160))
        ref.save(reference)
        Image.new("RGB", (200, 200), "white").save(rendered)

        result = module.audit_reference(
            str(reference),
            str(rendered),
            manifest_path=str(self._manifest([98, 98, 6, 6])),
            rows=2,
            cols=2,
            warn_mae=0.08,
            fail_mae=0.16,
        )

        self.assertLess(result["global"]["mae"], 0.01)
        self.assertFalse(result["passed"])
        self.assertEqual(result["elements"][0]["severity"], "FAIL")
        self.assertIn(
            "element_visual_delta",
            {issue["code"] for issue in result["issues"]},
        )

    def test_unapproved_omission_fails_even_when_pixels_match(self) -> None:
        module = load_module()
        reference = self.root / "reference.png"
        rendered = self.root / "rendered.png"
        manifest = self.root / "manifest.json"
        Image.new("RGB", (32, 32), "white").save(reference)
        Image.new("RGB", (32, 32), "white").save(rendered)
        manifest.write_text(
            json.dumps(
                {
                    "elements": [
                        {
                            "id": "badge",
                            "bbox_px": [2, 2, 8, 8],
                            "status": "omitted",
                            "approved": False,
                        }
                    ]
                }
            ),
            encoding="utf-8",
        )

        result = module.audit_reference(
            str(reference), str(rendered), manifest_path=str(manifest)
        )

        self.assertFalse(result["passed"])
        self.assertIn(
            "unapproved_substitution",
            {issue["code"] for issue in result["issues"]},
        )

    def test_cli_accepts_pretty_output_flag(self) -> None:
        reference = self.root / "reference.png"
        rendered = self.root / "rendered.png"
        Image.new("RGB", (16, 16), "white").save(reference)
        Image.new("RGB", (16, 16), "white").save(rendered)
        process = subprocess.run(
            [sys.executable, str(MODULE_PATH), str(reference), str(rendered), "--pretty"],
            capture_output=True,
            text=True,
        )
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertIn("\n  \"passed\"", process.stdout)

    def test_renderer_antialiasing_does_not_become_a_false_failure(self) -> None:
        module = load_module()
        reference = self.root / "reference.png"
        rendered = self.root / "rendered.png"
        source = Image.new("RGB", (80, 40), "white")
        draw = ImageDraw.Draw(source)
        for x in range(5, 75, 6):
            draw.rectangle((x, 8, x + 2, 32), fill="black")
        source.save(reference)
        source.filter(ImageFilter.GaussianBlur(0.75)).save(rendered)
        manifest = self._manifest([0, 0, 80, 40])

        result = module.audit_reference(
            str(reference),
            str(rendered),
            manifest_path=str(manifest),
            warn_mae=0.08,
            fail_mae=0.16,
            rows=1,
            cols=1,
        )

        self.assertTrue(result["passed"])
        self.assertNotEqual(result["elements"][0]["severity"], "FAIL")


if __name__ == "__main__":
    unittest.main()
