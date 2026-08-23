---
name: scientific-figure-shapes
description: Rapid, fidelity-audited reconstruction of scientific figures, mechanism diagrams, screenshots, and academic visuals into editable PowerPoint Shapes and VBA. Use when the user invokes scientific-figure-shapes, asks for 图片转可编辑 PPT/VBA, scientific image-to-shapes reconstruction, mechanism figure recreation, high-fidelity layout/arrow restoration, or an editable hybrid alternative to raster tracing.
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

Use stable generated names with the default prefix `FIG_`. A project-specific prefix is allowed, but one deck must use one prefix consistently. Use `_E_` for editable shapes, `_T_` for text, `_L_` for connectors, `_R_` for preserved raster regions, and `_G_` for logical groups.

## Fast Path

1. Inspect the source dimensions and split the figure into coarse regions.
2. Draft one `figure-manifest.json` from `references/figure-manifest-example.json`, then validate it with `scripts/manifest_validate.py`.
3. Probe local Office support with `scripts/office_runtime_probe.py`.
4. Draft a compact element map:
   - `B*` base/background regions
   - `E*` editable shape/text regions
   - `R*` preserved crop regions
   - `L*` major lines and arrows
5. Before preserving a simple icon as raster, try `scripts/vector_trace.py`. Keep the trace only when its contour count and preview remain clean.
6. Crop truly complex regions with `scripts/preserve_cropper.py`; use stable `--start-index`, and never pass `--overwrite` without an intentional replacement.
7. Calculate the slide mapping with `scripts/canvas_point_mapper.py`.
8. Build the editable deck with `scripts/office_shape_canvas.py`. Use logical groups and stable names; emit both `.pptx` and a complete `.bas` module with `BuildFinal` when VBA is useful.
9. Smoke-check the macro with `scripts/macro_smoke_lint.py`.
10. Try native macro materialization once when possible:
   - macOS PowerPoint: `scripts/ppt_macos_macro_launcher.py`
   - Windows PowerPoint: `scripts/ppt_windows_macro_runner.ps1`
    The macOS launcher first probes whether `do Visual Basic` is available. If it is blocked, treat the `python-pptx` output from `office_shape_canvas.py` as the first-class editable deliverable; do not keep retrying AppleScript.
11. Run `scripts/pptx_text_audit.py` for required labels, font size, alignment, explicit line count, and opt-in overflow estimation.
12. For every materialized `.pptx` containing `_L_` connectors, run `scripts/pptx_connector_audit.py` with required target edges and explicit intentional exceptions.
13. If the source uses multi-face cuboids, feature pyramids, gradients, or deliberate overlap, run `scripts/pptx_layering_audit.py`.
14. Run `scripts/pptx_editability_audit.py`; unregistered full-slide or large raster pictures are failures.
15. If automation is blocked, deliver the editable `.pptx`, `.bas`, crops, validated manifest, and audit reports. Never substitute a preview PNG for the editable deck.

## Fidelity Escalation

Use the slower path only when the user asks for high fidelity, pixel-level matching, 1:1 recreation, publication-grade output, or repeated corrections.

For that path:

- expand the map to individual visible elements
- zoom into each panel of the source and take an explicit arrow inventory (every connector's start element, end element, solid/dashed style, and direction) before drawing; a connector present in the source but absent from the rebuild is a coverage defect that no per-connector audit can catch
- record the inventoried minimum as `routing_audit.min_connector_count`; the connector audit fails with `connector_inventory_shortfall` when the materialized deck carries fewer audited connectors
- record exact crop boxes, key arrow endpoints, and the intended target edge for each fan-in/fan-out route
- render a preview when possible
- compare source and preview with `scripts/fidelity_audit.py`; register every salient icon, label, face group, and output symbol in `elements`
- audit the materialized deck with `scripts/pptx_connector_audit.py`
- audit explicit face vertices, gradient colors/angles, size hierarchy, overlap, and z-order with `scripts/pptx_layering_audit.py` when depth is visually salient
- compare PowerPoint and LibreOffice previews with `scripts/dual_renderer_probe.py` when both renderers are available; renderer drift is a compatibility defect, not source fidelity evidence
- use `references/fidelity-review-gates.md` for the review checklist

Fast mode does not require an SSIM threshold. High-fidelity mode must use local ROI/element gates so that a missing small icon cannot be diluted by a low global average. Report whatever comparison data is available without pretending a diagnostic preview is the editable deliverable.

## Crop Contract

For each preserved crop, keep the bbox and `asset_id` traceable. The crop should be as small as useful, source aspect ratio must be preserved, and the inserted image must not be stretched independently on x/y. Recreate labels and arrows around the crop as editable objects. Repeated crop runs must use a stable index range or an explicit overwrite decision; silent replacement is forbidden.

## Connector Routing Contract

Treat connector routing as geometry, not decoration:

- Give merge/split boxes, their text overlays, and final connector segments stable generated names.
- Draw shared fan-out geometry once: one trunk, one horizontal/vertical bus, then separate branch drops. Repeating the trunk inside a loop creates darker strokes and duplicate geometry.
- Do not emit zero-length connector segments when a branch is already aligned with the trunk.
- For fan-in routes, choose the target edge deliberately. Left, center, and right sources should normally enter through left, top/bottom, and right edges rather than sharing one hard-coded endpoint.
- A connector may touch a text-bearing box at its boundary, but must not enter the shrunken text rectangle or cross the label.
- Both endpoints of every audited connector must be anchored: touching or entering a shape, or continuing another connector's endpoint. A connector starting or ending in blank canvas (including a visible gap left to "avoid" a target) is a failure — the audit reports it as `dangling_endpoints` by default, with `ignore_dangling` as the narrow manifest exemption for deliberate free leaders.
- For principal routes, record `source`/`source_edge` alongside `target`/`target_edge` in the routing manifest so an arrow provably departs from the intended shape, not just arrives at one.
- Record required edges in a routing manifest. Intentional line-to-text cases must be listed explicitly with `ignore_text_shapes` or `ignore_pairs`; do not disable the audit globally.
- Record cross-scale or semantically styled segments in `routing_audit.segments`; declare `dash` and `orientation` so a dashed orthogonal link cannot silently become a solid diagonal line.
- Run the audit on the materialized `.pptx`; connector/text collisions, duplicate segments, zero-length segments, and target-edge mismatches are failures. Then inspect the rendered preview. A clean source macro is not evidence that the final arrows are routed correctly.

Read `references/office-shape-recipes.md` for the fan-in pattern and manifest example.

## Depth and Layering Contract

Do not use one `msoShapeCube` when the source relies on visible face shading, stepped scale, or occlusion:

- Rebuild each important cuboid as separate `front`, `top`, and `right` faces with stable names.
- For high-fidelity work, trace source face polygons in source pixels and record both `expected_face_bounds_pt` and `expected_face_vertices_pt` with explicit tolerances. A structural three-face match is not enough.
- Keep one consistent depth vector across all layers; top faces are lighter and side faces darker than front faces.
- Sample face colors from the source and use controlled gradients when the source contains directional shading; record expected gradient stop colors and angle in the manifest.
- Create back layers first and foreground layers later so z-order matches the source.
- Record size order and required overlap pairs in a layering manifest.
- Run `pptx_layering_audit.py` on the materialized deck. Missing faces, flat face colors, detached faces, reference-bound errors, missing gradients, wrong size order, missing overlap, or wrong z-order are failures.

Read `references/office-shape-recipes.md` for the three-face cuboid recipe.

## Deliverables

For a normal request, create as many of these as practical:

- `*.bas`: a complete runnable VBA module
- `assets/*.png`: preserved crops
- `figure-manifest.json`: machine-checkable scene, routing, text, layering, crop, and editability contract
- `manifest.md`: optional human-readable summary of editable-vs-preserved regions
- `*.pptx`: editable fallback deck
- `preview.png`: diagnostic visual preview
- `run_report.md`: short automation and validation note
- `audit/*.json` and `audit/*.png`: fidelity, routing, layering, text, editability, and renderer-delta evidence

Always distinguish editable deliverables from preview images in the final response.

## Resource Guide

Use bundled resources directly. Load long references only when needed.

- `scripts/office_runtime_probe.py`: detect presentation apps and automation helpers.
- `scripts/office_shape_canvas.py`: reusable editable Shapes runtime with gradients, freeforms, stable names, logical groups, incremental updates, PPTX output, and VBA emission.
- `scripts/canvas_point_mapper.py`: calculate source-pixel to Office-point mapping.
- `scripts/preserve_cropper.py`: crop raster-preserved regions with stable IDs and collision-safe output.
- `scripts/vector_trace.py`: trace simple raster icons into editable SVG/PPTX polygon paths.
- `scripts/macro_smoke_lint.py`: catch common generated VBA mistakes.
- `scripts/pptx_connector_audit.py`: catch connector/text collisions and verify named connectors enter the required target edge.
- `scripts/pptx_layering_audit.py`: verify explicit face vertices, gradients, depth contrast, size hierarchy, overlap, and z-order.
- `scripts/pptx_text_audit.py`: verify required text, typography, alignment, line count, and opt-in overflow estimates.
- `scripts/pptx_editability_audit.py`: score editable shapes versus registered and unregistered raster area.
- `scripts/ppt_macos_macro_launcher.py`: capability-probed macOS PowerPoint macro attempt with an explicit editable fallback.
- `scripts/ppt_windows_macro_runner.ps1`: one-shot Windows PowerPoint macro attempt.
- `scripts/render_delta_probe.py`: optional source-vs-preview image diagnostics.
- `scripts/fidelity_audit.py`: strict global, grid, edge, and salient-element source-vs-preview audit with heatmap output.
- `scripts/dual_renderer_probe.py`: compare PowerPoint and LibreOffice PNG renders.
- `scripts/manifest_validate.py`: validate the unified reconstruction manifest before building.
- `references/office-shape-recipes.md`: concise VBA shape patterns.
- `references/delivery-note-format.md`: expanded delivery report format.
- `references/fidelity-review-gates.md`: stricter review gates for high-fidelity work.
- `references/audit-toolchain.md`: command-level audit workflow and thresholds.
- `references/figure-manifest.schema.json`: unified manifest JSON Schema.
