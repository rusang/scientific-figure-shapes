# Scientific Figure Shapes

`scientific-figure-shapes` is a Codex skill for quickly reconstructing scientific figures, mechanism diagrams, screenshots, and academic visuals as editable PowerPoint VBA Shapes.

It favors practical editability over pixel-perfect tracing: text, labels, arrows, boxes, tables, simple icons, axes, and diagram structure are rebuilt as Office shapes, while complex biological artwork or texture-heavy regions can be preserved as small traceable image crops.

## What It Does

- Converts uploaded scientific images into editable PowerPoint-oriented VBA modules.
- Rebuilds layout, labels, callouts, connectors, legends, and simple visual elements as shapes.
- Preserves high-detail elements as local crops when redrawing them would be slow or lower fidelity.
- Produces runnable `.bas` macros plus manifests, validation notes, and optional editable `.pptx` fallbacks.
- Supports a fast workflow first, with stricter fidelity checks available when requested.

## Typical Use

Ask Codex something like:

```text
Use scientific-figure-shapes to quickly convert this uploaded academic image into editable PowerPoint VBA Shapes.
```

Chinese prompts work well too:

```text
用 scientific-figure-shapes，把这张科研机制图快速还原成可编辑 PowerPoint VBA Shapes。
```

## Installation

Clone the repository into your Codex skills directory:

```bash
mkdir -p ~/.codex/skills
git clone git@github-summer-ai-lab:summer-ai-lab/scientific-figure-shapes.git ~/.codex/skills/scientific-figure-shapes
```

Restart Codex, or reload skills if your environment supports it.

## Repository Structure

```text
scientific-figure-shapes/
├── SKILL.md
├── agents/
│   └── scientific-figure-shapes.yaml
├── references/
│   ├── delivery-note-format.md
│   ├── fidelity-review-gates.md
│   └── office-shape-recipes.md
└── scripts/
    ├── canvas_point_mapper.py
    ├── macro_smoke_lint.py
    ├── office_runtime_probe.py
    ├── ppt_macos_macro_launcher.py
    ├── ppt_windows_macro_runner.ps1
    ├── preserve_cropper.py
    └── render_delta_probe.py
```

## Outputs

A normal run may produce:

- `*.bas` - runnable VBA module with a `BuildFinal` macro.
- `assets/*.png` - preserved source crops for complex visual regions.
- `manifest.md` - editable-vs-preserved reconstruction map.
- `*.pptx` - editable fallback deck when direct Office automation is unavailable.
- `preview.png` - diagnostic preview image.
- `run_report.md` - validation and automation notes.

## Design Notes

This skill is an independent rewrite focused on a speed-first scientific figure reconstruction workflow. It is meant for practical academic editing tasks, not automated publication-grade vector tracing. For high-fidelity recreation, ask explicitly for stricter review, preview comparison, and correction passes.

## Requirements

- Codex with local skills support.
- Python 3 for helper scripts.
- Microsoft PowerPoint or WPS Presentation for direct macro execution when available.

The skill can still generate VBA and supporting files when Office automation is not available.
