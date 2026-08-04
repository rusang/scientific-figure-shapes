---
name: scientific-figure-shapes
description: Rapid, practical reconstruction of scientific figures, mechanism diagrams, screenshots, and academic visuals into editable PowerPoint VBA Shapes. Use when the user invokes scientific-figure-shapes, asks for 图片转 VBA, 可编辑 PPT, scientific image-to-shapes reconstruction, mechanism figure recreation, or a faster hybrid alternative to strict pixel-level tracing.
---

# Scientific Figure Shapes

Scientific Figure Shapes turns a source image into an editable PowerPoint-oriented reconstruction. It is designed for speed: rebuild the parts that benefit from editing, and preserve complex artwork as small local crops when redrawing would waste time or reduce fidelity.

## Language

Reply, write manifests, and write run notes in Simplified Chinese by default. Keep code identifiers, Office API names, file names, and `RGB(...)` calls in English.

## Operating Style

Default to PowerPoint and a 16:9 canvas. If the source image has a different aspect ratio, keep the source ratio and set the slide long side to 960 pt unless the user gives a target size.

Prefer these editable objects:

- text, titles, labels, legends, tables, boxes, callouts, arrows, axes, simple charts
- repeated simple marks such as dots, rounded bars, cells, icons, and connector scaffolds
- layout backgrounds, frames, bands, simple gradients, and diagram structure

Preserve these as local image crops:

- anatomical or biological illustrations, microscopy, photos, 3D renderings, screenshots inside devices, logos
- texture-heavy, painterly, shaded, or high-detail elements
- any one element that would need roughly more than 15 shapes to redraw cleanly
- anything the user explicitly asks to keep as image

Never hide the whole source image behind the reconstruction unless the user explicitly asks for a background-assisted draft.

Use generated names with the prefix `SUMMER_`.

## Fast Path

1. Inspect the source dimensions and split the figure into coarse regions.
2. Probe local Office support with `scripts/office_runtime_probe.py`.
3. Draft a compact element map:
   - `B*` base/background regions
   - `E*` editable shape/text regions
   - `R*` preserved crop regions
   - `L*` major lines and arrows
4. Crop preserved regions with `scripts/preserve_cropper.py`; use no padding unless the bbox intentionally includes shadow or glow.
5. Calculate the slide mapping with `scripts/canvas_point_mapper.py`.
6. Generate a complete `.bas` module with `BuildFinal`. Add `BuildSkeleton` when it is quick.
7. Smoke-check the macro with `scripts/macro_smoke_lint.py`.
8. Try materialization once when possible:
   - macOS PowerPoint: `scripts/ppt_macos_macro_launcher.py`
   - Windows PowerPoint: `scripts/ppt_windows_macro_runner.ps1`
9. If automation is blocked, stop there and deliver the macro, crops, a manifest, and an editable `.pptx` fallback when feasible.

## Fidelity Escalation

Use the slower path only when the user asks for high fidelity, pixel-level matching, 1:1 recreation, publication-grade output, or repeated corrections.

For that path:

- expand the map to individual visible elements
- record exact crop boxes and key arrow endpoints
- render a preview when possible
- compare source and preview with `scripts/render_delta_probe.py`
- use `references/fidelity-review-gates.md` for the review checklist

Fast mode does not require an SSIM threshold. Report whatever comparison data is available without pretending a diagnostic preview is the editable deliverable.

## Crop Contract

For each preserved crop, keep the bbox traceable. The crop should be as small as useful, source aspect ratio must be preserved, and the inserted image must not be stretched independently on x/y. Recreate labels and arrows around the crop as editable objects.

## Deliverables

For a normal request, create as many of these as practical:

- `*.bas`: a complete runnable VBA module
- `assets/*.png`: preserved crops
- `manifest.md`: concise editable-vs-preserved map
- `*.pptx`: editable fallback deck
- `preview.png`: diagnostic visual preview
- `run_report.md`: short automation and validation note

Always distinguish editable deliverables from preview images in the final response.

## Resource Guide

Use bundled resources directly. Load long references only when needed.

- `scripts/office_runtime_probe.py`: detect presentation apps and automation helpers.
- `scripts/canvas_point_mapper.py`: calculate source-pixel to Office-point mapping.
- `scripts/preserve_cropper.py`: crop raster-preserved regions and emit a JSON crop manifest.
- `scripts/macro_smoke_lint.py`: catch common generated VBA mistakes.
- `scripts/ppt_macos_macro_launcher.py`: one-shot macOS PowerPoint macro attempt.
- `scripts/ppt_windows_macro_runner.ps1`: one-shot Windows PowerPoint macro attempt.
- `scripts/render_delta_probe.py`: optional source-vs-preview image diagnostics.
- `references/office-shape-recipes.md`: concise VBA shape patterns.
- `references/delivery-note-format.md`: expanded delivery report format.
- `references/fidelity-review-gates.md`: stricter review gates for high-fidelity work.
