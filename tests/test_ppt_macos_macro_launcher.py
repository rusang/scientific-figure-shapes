from __future__ import annotations

import importlib.util
import subprocess
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = ROOT / "scripts" / "ppt_macos_macro_launcher.py"


def load_module():
    spec = importlib.util.spec_from_file_location("ppt_macos_macro_launcher", MODULE_PATH)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


class MacMacroLauncherTests(unittest.TestCase):
    def test_default_macro_matches_generated_entrypoint(self) -> None:
        module = load_module()
        self.assertEqual(module.DEFAULT_MACRO, "BuildFinal")

    def test_visual_basic_capability_failure_has_editable_fallback(self) -> None:
        module = load_module()
        failure = subprocess.CompletedProcess(
            ["osascript", "-"],
            1,
            "",
            "Microsoft PowerPoint got an error: doesn't understand do Visual Basic",
        )
        result = module.classify_capability(failure)

        self.assertFalse(result["available"])
        self.assertEqual(result["status"], "visual_basic_unavailable")
        self.assertIn("python-pptx", result["fallback"])

    def test_capability_probe_is_non_destructive(self) -> None:
        module = load_module()
        program = module.visual_basic_probe_program()
        self.assertNotIn("Presentations.Add", program)
        self.assertNotIn("Save", program)
        self.assertIn("do Visual Basic", program)

    def test_probe_only_does_not_require_a_vba_file_argument(self) -> None:
        module = load_module()
        args = module.build_parser().parse_args(["--probe-only"])
        self.assertTrue(args.probe_only)
        self.assertIsNone(args.vba_file)


if __name__ == "__main__":
    unittest.main()
