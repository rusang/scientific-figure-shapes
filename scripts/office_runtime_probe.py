#!/usr/bin/env python3
"""Report local presentation software useful for figure-shape materialization."""

from __future__ import annotations

import json
import platform
import shutil
import subprocess
from pathlib import Path


def app_exists(path: str) -> bool:
    return Path(path).exists()


def mac_apps() -> list[dict]:
    candidates = [
        ("Microsoft PowerPoint", "/Applications/Microsoft PowerPoint.app"),
        ("WPS Office", "/Applications/WPS Office.app"),
        ("Keynote", "/Applications/Keynote.app"),
        ("LibreOffice", "/Applications/LibreOffice.app"),
    ]
    return [{"name": name, "paths": [path]} for name, path in candidates if app_exists(path)]


def windows_apps() -> list[dict]:
    result = []
    try:
        import winreg  # type: ignore
    except Exception:
        return result
    app_paths = {
        "Microsoft PowerPoint": "POWERPNT.EXE",
        "WPS Presentation": "wpp.exe",
    }
    for label, exe in app_paths.items():
        key_name = rf"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\{exe}"
        for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(root, key_name) as key:
                    path, _ = winreg.QueryValueEx(key, None)
                result.append({"name": label, "paths": [path]})
                break
            except OSError:
                continue
    return result


def linux_apps() -> list[dict]:
    found = []
    for label, exe in (("LibreOffice", "libreoffice"), ("LibreOffice", "soffice"), ("WPS Office", "wpp")):
        path = shutil.which(exe)
        if path:
            found.append({"name": label, "paths": [path]})
    return found


def tool(name: str) -> dict | None:
    path = shutil.which(name)
    return {"name": name, "path": path} if path else None


def applescript_smoke() -> bool | None:
    if platform.system() != "Darwin" or not shutil.which("osascript"):
        return None
    try:
        proc = subprocess.run(["osascript", "-e", "return 1"], capture_output=True, text=True, timeout=5)
    except Exception:
        return False
    return proc.returncode == 0


def main() -> int:
    system = platform.system()
    if system == "Darwin":
        apps = mac_apps()
        preferred = "macos_powerpoint_applescript" if any(a["name"] == "Microsoft PowerPoint" for a in apps) else "manual_or_fallback"
    elif system == "Windows":
        apps = windows_apps()
        preferred = "windows_powerpoint_com" if any(a["name"] == "Microsoft PowerPoint" for a in apps) else "manual_or_fallback"
    else:
        apps = linux_apps()
        preferred = "fallback_pptx_or_manual"

    tools = [item for item in (tool("osascript"), tool("powershell"), tool("pwsh"), tool("soffice"), tool("libreoffice")) if item]
    notes = []
    if applescript_smoke() is False:
        notes.append("osascript exists but a simple AppleScript smoke test failed")
    if not apps:
        notes.append("no presentation application detected from common locations")

    print(
        json.dumps(
            {
                "platform": system,
                "presentation_apps": apps,
                "automation_tools": tools,
                "recommended_path": preferred,
                "notes": notes,
            },
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
