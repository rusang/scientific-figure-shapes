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
9. For every materialized `.pptx` containing `SUMMER_L_` connectors, run `scripts/pptx_connector_audit.py`. Use a routing manifest for required target edges and explicit intentional exceptions.
10. If the source uses multi-face cuboids, feature pyramids, or deliberate overlap, run `scripts/pptx_layering_audit.py` with a layering manifest.
11. If automation is blocked, stop there and deliver the macro, crops, a manifest, and an editable `.pptx` fallback when feasible.

## Fidelity Escalation

Use the slower path only when the user asks for high fidelity, pixel-level matching, 1:1 recreation, publication-grade output, or repeated corrections.

For that path:

- expand the map to individual visible elements
- record exact crop boxes, key arrow endpoints, and the intended target edge for each fan-in/fan-out route
- render a preview when possible
- compare source and preview with `scripts/render_delta_probe.py`
- audit the materialized deck with `scripts/pptx_connector_audit.py`
- audit explicit faces, size hierarchy, overlap, and z-order with `scripts/pptx_layering_audit.py` when depth is visually salient
- use `references/fidelity-review-gates.md` for the review checklist

Fast mode does not require an SSIM threshold. Report whatever comparison data is available without pretending a diagnostic preview is the editable deliverable.

## Crop Contract

For each preserved crop, keep the bbox traceable. The crop should be as small as useful, source aspect ratio must be preserved, and the inserted image must not be stretched independently on x/y. Recreate labels and arrows around the crop as editable objects.

## Connector Routing Contract

Treat connector routing as geometry, not decoration:

- Give merge/split boxes, their text overlays, and final connector segments stable `SUMMER_` names.
- Draw shared fan-out geometry once: one trunk, one horizontal/vertical bus, then separate branch drops. Repeating the trunk inside a loop creates darker strokes and duplicate geometry.
- Do not emit zero-length connector segments when a branch is already aligned with the trunk.
- For fan-in routes, choose the target edge deliberately. Left, center, and right sources should normally enter through left, top/bottom, and right edges rather than sharing one hard-coded endpoint.
- A connector may touch a text-bearing box at its boundary, but must not enter the shrunken text rectangle or cross the label.
- Record required edges in a routing manifest. Intentional line-to-text cases must be listed explicitly with `ignore_text_shapes` or `ignore_pairs`; do not disable the audit globally.
- Run the audit on the materialized `.pptx`; connector/text collisions, duplicate segments, zero-length segments, and target-edge mismatches are failures. Then inspect the rendered preview. A clean source macro is not evidence that the final arrows are routed correctly.

Read `references/office-shape-recipes.md` for the fan-in pattern and manifest example.

## Depth and Layering Contract

Do not use one `msoShapeCube` when the source relies on visible face shading, stepped scale, or occlusion:

- Rebuild each important cuboid as separate `front`, `top`, and `right` faces with stable names.
- Keep one consistent depth vector across all layers; top faces are lighter and side faces darker than front faces.
- Create back layers first and foreground layers later so z-order matches the source.
- Record size order and required overlap pairs in a layering manifest.
- Run `pptx_layering_audit.py` on the materialized deck. Missing faces, flat face colors, detached faces, wrong size order, missing overlap, or wrong z-order are failures.

Read `references/office-shape-recipes.md` for the three-face cuboid recipe.

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
- `scripts/pptx_connector_audit.py`: catch connector/text collisions and verify named connectors enter the required target edge.
- `scripts/pptx_layering_audit.py`: verify explicit cuboid faces, depth contrast, size hierarchy, overlap, and z-order.
- `scripts/ppt_macos_macro_launcher.py`: one-shot macOS PowerPoint macro attempt.
- `scripts/ppt_windows_macro_runner.ps1`: one-shot Windows PowerPoint macro attempt.
- `scripts/render_delta_probe.py`: optional source-vs-preview image diagnostics.
- `references/office-shape-recipes.md`: concise VBA shape patterns.
- `references/delivery-note-format.md`: expanded delivery report format.
- `references/fidelity-review-gates.md`: stricter review gates for high-fidelity work.
