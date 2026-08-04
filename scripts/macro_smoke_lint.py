#!/usr/bin/env python3
"""Sanity-check a generated scientific figure VBA file before trying to run it."""

from __future__ import annotations

import argparse
import json
from pathlib import Path


def code_lines(source: str):
    for raw in source.splitlines():
        buf = []
        quoted = False
        pos = 0
        while pos < len(raw):
            ch = raw[pos]
            if ch == '"':
                buf.append(ch)
                if quoted and pos + 1 < len(raw) and raw[pos + 1] == '"':
                    buf.append('"')
                    pos += 2
                    continue
                quoted = not quoted
            elif ch == "'" and not quoted:
                break
            else:
                buf.append(ch)
            pos += 1
        yield "".join(buf), quoted


def lower_lines(source: str) -> list[str]:
    return [line.lower() for line, _ in code_lines(source)]


def delimiter_score(source: str) -> int:
    score = 0
    for line, _ in code_lines(source):
        quoted = False
        pos = 0
        while pos < len(line):
            ch = line[pos]
            if ch == '"':
                if quoted and pos + 1 < len(line) and line[pos + 1] == '"':
                    pos += 2
                    continue
                quoted = not quoted
            elif not quoted:
                if ch == "(":
                    score += 1
                elif ch == ")":
                    score -= 1
            pos += 1
    return score


def check(source: str) -> dict:
    lowered = lower_lines(source)
    joined = "\n".join(lowered)
    errors: list[str] = []
    warnings: list[str] = []

    if not any(line.strip().startswith(("sub ", "public sub ", "private sub ")) for line in lowered):
        errors.append("VBA procedure entry not found.")
    if "end sub" not in joined:
        errors.append("VBA procedure terminator not found.")
    if delimiter_score(source) != 0:
        errors.append("Parenthesis count does not return to zero.")
    if any(open_quote for _, open_quote in code_lines(source)):
        errors.append("At least one line ends inside a string literal.")

    shape_signals = ("shapes.addshape", "shapes.addtextbox", "shapes.addline", "shapes.addpicture", "buildfreeform")
    if not any(signal in joined for signal in shape_signals):
        warnings.append("No common shape-building API call was detected.")
    if ".name" not in joined:
        warnings.append("Shape naming assignments were not detected.")
    if "rgb(" not in joined and "forecolor" not in joined:
        warnings.append("Color formatting was not detected.")
    if "shapes.count" not in joined and ".delete" not in joined:
        warnings.append("Slide cleanup logic was not detected.")

    return {"passed": len(errors) == 0, "warnings": warnings, "errors": errors}


def load(path: Path) -> str:
    for enc in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            return path.read_text(encoding=enc)
        except UnicodeDecodeError:
            pass
    return path.read_text(encoding="utf-8", errors="replace")


def main() -> int:
    ap = argparse.ArgumentParser(description="Smoke-lint a PowerPoint VBA macro file.")
    ap.add_argument("vba_file")
    ap.add_argument("--pretty", action="store_true")
    ns = ap.parse_args()

    target = Path(ns.vba_file).expanduser()
    if not target.exists():
        print(json.dumps({"passed": False, "warnings": [], "errors": [f"missing file: {target}"]}))
        return 2
    verdict = check(load(target))
    print(json.dumps(verdict, indent=2 if ns.pretty else None))
    return 0 if verdict["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
