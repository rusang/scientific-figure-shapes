#!/usr/bin/env python3
"""Best-effort macOS launcher for a generated PowerPoint VBA file."""

from __future__ import annotations

import argparse
import json
import platform
import shutil
import subprocess
import tempfile
from pathlib import Path


PPT_APP = Path("/Applications/Microsoft PowerPoint.app")
DEFAULT_MACRO = "BuildFinal"
EDITABLE_FALLBACK = (
    "Use scripts/office_shape_canvas.py (python-pptx) to materialize the same "
    "scene as an editable PPTX, or import the .bas module manually."
)


def q(value: str) -> str:
    return value.replace("\\", "\\\\").replace('"', '\\"')


def apple_program(vba_file: Path) -> str:
    return "\n".join(
        [
            f'set figureMacroFile to POSIX file "{q(str(vba_file))}"',
            "set figureMacroSource to read figureMacroFile",
            'tell application "Microsoft PowerPoint"',
            "    activate",
            "    do Visual Basic figureMacroSource",
            "end tell",
        ]
    )


def visual_basic_probe_program() -> str:
    """Return a non-destructive AppleScript capability probe."""
    return "\n".join(
        [
            'tell application "Microsoft PowerPoint"',
            "    activate",
            '    do Visual Basic "Debug.Print \\\"codex-vba-probe\\\""',
            "end tell",
        ]
    )


def probe_visual_basic() -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["osascript", "-"],
        input=visual_basic_probe_program(),
        capture_output=True,
        text=True,
    )


def classify_capability(proc: subprocess.CompletedProcess[str]) -> dict:
    if proc.returncode == 0:
        return {
            "available": True,
            "status": "visual_basic_available",
            "returncode": proc.returncode,
        }
    return {
        "available": False,
        "status": "visual_basic_unavailable",
        "returncode": proc.returncode,
        "stderr": proc.stderr.strip(),
        "fallback": EDITABLE_FALLBACK,
    }


def invoke(vba_file: Path) -> subprocess.CompletedProcess[str]:
    return subprocess.run(["osascript", "-"], input=apple_program(vba_file), capture_output=True, text=True)


def read_module(path: Path) -> str:
    return path.read_text(encoding="utf-8", errors="replace")


def procedure_body(module_text: str, proc_name: str) -> str | None:
    wanted = proc_name.lower()
    lines = module_text.splitlines()
    start = None
    for i, line in enumerate(lines):
        compact = " ".join(line.strip().lower().split())
        if compact.startswith("sub ") or compact.startswith("public sub ") or compact.startswith("private sub "):
            name_part = compact.split("sub ", 1)[1].split("(", 1)[0].split()[0]
            if name_part == wanted:
                start = i + 1
                break
    if start is None:
        return None
    for end in range(start, len(lines)):
        if lines[end].strip().lower().startswith("end sub"):
            body = "\n".join(lines[start:end]).strip()
            return body or None
    return None


def record(kind: str, proc: subprocess.CompletedProcess[str], temp: Path | None = None) -> dict:
    data = {"mode": kind, "returncode": proc.returncode, "stdout": proc.stdout.strip(), "stderr": proc.stderr.strip()}
    if temp is not None:
        data["temp_file"] = str(temp)
    return data


def emit(payload: dict, pretty: bool, exit_code: int) -> int:
    print(json.dumps(payload, indent=2 if pretty else None))
    return exit_code


def build_parser() -> argparse.ArgumentParser:
    ap = argparse.ArgumentParser(description="Try to execute a generated VBA module in PowerPoint for Mac.")
    ap.add_argument("vba_file", nargs="?")
    ap.add_argument("--macro", default=DEFAULT_MACRO)
    ap.add_argument("--no-body-fallback", action="store_true")
    ap.add_argument("--skip-capability-probe", action="store_true")
    ap.add_argument("--probe-only", action="store_true")
    ap.add_argument("--pretty", action="store_true")
    return ap


def main() -> int:
    ns = build_parser().parse_args()

    if platform.system() != "Darwin":
        return emit({"status": "unsupported_platform", "platform": platform.system()}, ns.pretty, 2)
    if not PPT_APP.exists():
        return emit({"status": "powerpoint_not_found", "expected_path": str(PPT_APP)}, ns.pretty, 2)
    if shutil.which("osascript") is None:
        return emit({"status": "osascript_not_found"}, ns.pretty, 2)

    if ns.probe_only or not ns.skip_capability_probe:
        capability = classify_capability(probe_visual_basic())
        if ns.probe_only or not capability["available"]:
            return emit(capability, ns.pretty, 0 if capability["available"] else 1)

    if not ns.vba_file:
        return emit({"status": "vba_file_required"}, ns.pretty, 2)

    module_path = Path(ns.vba_file).expanduser().resolve()
    if not module_path.is_file():
        return emit({"status": "vba_file_not_found", "path": str(module_path)}, ns.pretty, 2)

    attempts = []
    whole = invoke(module_path)
    attempts.append(record("module_source", whole))
    if whole.returncode == 0:
        return emit({"status": "ran", "vba_file": str(module_path), "attempts": attempts}, ns.pretty, 0)

    if not ns.no_body_fallback:
        body = procedure_body(read_module(module_path), ns.macro)
        if body:
            with tempfile.NamedTemporaryFile("w", suffix=".vba", encoding="utf-8", delete=False) as tmp:
                tmp.write(body)
                tmp_path = Path(tmp.name)
            body_run = invoke(tmp_path)
            attempts.append(record("procedure_body", body_run, tmp_path))
            if body_run.returncode == 0:
                return emit(
                    {"status": "ran_with_body", "vba_file": str(module_path), "macro": ns.macro, "attempts": attempts},
                    ns.pretty,
                    0,
                )

    return emit(
        {
            "status": "automation_failed",
            "vba_file": str(module_path),
            "macro": ns.macro,
            "attempts": attempts,
            "fallback": EDITABLE_FALLBACK,
        },
        ns.pretty,
        1,
    )


if __name__ == "__main__":
    raise SystemExit(main())
