# Scientific Figure Shapes Fidelity Review Gates

Use this checklist only for high-fidelity or repeated-correction work.

## 1. Canvas Gate

- Source size and target slide size are recorded.
- Mapping mode is explicit: uniform fit or exact same-ratio scale.
- No full-source hidden background remains in the final editable deliverable unless requested.

## 2. Crop Gate

For every `R*` crop:

- bbox and stable `asset_id` are recorded in source pixels.
- crop file exists and has nonzero size.
- crop aspect ratio matches the recorded bbox.
- inserted image keeps aspect ratio.
- rotation and flips are intentional and documented.
- nearby labels, arrows, and highlights are rebuilt as editable objects when practical.
- repeated crop runs cannot silently overwrite an earlier asset; `--overwrite` is an explicit decision.

## 3. Editable Structure Gate

Check that these are editable when present:

- titles and labels
- panels, boxes, callouts, dividers
- arrows and connector lines
- legends, axes, simple charts, tables
- repeated simple marks and cells

## 4. Depth and Layering Gate

- Visually important 3D blocks use explicit `front`, `top`, and `right` faces instead of one uncontrollable auto-cube.
- Each face has source-derived expected bounds and, when perspective matters, expected vertices; materialized geometry stays within the declared tolerance (normally no more than 1 pt after Office quantization).
- All layers share one depth vector and consistent face shading direction.
- Faces marked as gradient retain expected stop colors and angle after PPTX materialization.
- Feature-pyramid sizes increase in the intended order.
- Required layer pairs overlap, and back-to-front z-order matches the source.
- Run `scripts/pptx_layering_audit.py`; missing faces, flat colors, detached faces, reference-geometry errors, missing gradients, hierarchy errors, overlap errors, or z-order errors fail the gate.

## 5. Text Gate

- Text fits inside its box on the rendered preview.
- Important labels are not hidden behind crops.
- Font size hierarchy is close enough for the intended use.
- Biological symbols, arrows, down/up marks, and abbreviations are not accidentally changed.
- Run `scripts/pptx_text_audit.py`; required text, font, alignment, explicit line count, and opted-in overflow estimates must pass.

## 6. Routing Gate

- Major arrows point to the right target region.
- An arrow inventory taken from the zoomed source (per panel: endpoints, style, direction) is recorded, and `min_connector_count` in the routing manifest reflects it; `connector_inventory_shortfall` findings fail the gate.
- Dashed zoom lines attach to the intended source and destination.
- Every important final connector segment has a stable name and a recorded target edge (`left`, `right`, `top`, or `bottom`).
- Shared fan-out trunks/buses are drawn once, not repeated once per branch.
- No connector is zero-length, and no two connectors have duplicate geometry unless one is explicitly excluded for a documented reason.
- Fan-in routes do not reuse one hard-coded endpoint when that would force a line through the merge label.
- Lines may touch target boundaries but do not enter the shrunken text rectangle or cross labels.
- Every connector endpoint is anchored to a shape or to another connector's endpoint; `dangling_endpoints` findings (blank-canvas starts, hovering arrowheads) fail the gate.
- Principal routes declare `source`/`source_edge` in addition to `target_edge` so departure anchoring is machine-checked, not eyeballed.
- Run `scripts/pptx_connector_audit.py` on the materialized deck. Any collision or target-edge mismatch is a failure.
- Intentional line-to-text cases are narrow manifest exceptions (`ignore_text_shapes` / `ignore_pairs`), not a global audit bypass.

## 7. Render Gate

When a preview exists:

- no large blank region appears unexpectedly.
- no crop is visibly stretched.
- main panels sit in the expected order.
- source and preview are compared with `scripts/fidelity_audit.py`; every salient small icon or label has an ROI entry so global averaging cannot hide an omission.
- when both renderers are available, PowerPoint and LibreOffice previews are compared with `scripts/dual_renderer_probe.py`.

## 8. Editability Gate

- Run `scripts/pptx_editability_audit.py` on the editable deck.
- Every preserved raster shape is registered by name in the manifest.
- A full-slide or large unregistered picture is a failure even when editable objects are placed above it.
- Reference-image slides are explicitly registered or use the `_R_reference_full` name and are excluded from the content-slide score.

## 9. Handoff Gate

- Editable PPTX, macro, assets, validated manifest, preview, and audit reports are clearly labeled.
- Automation failures are reported as local execution issues, not as successful macro runs.
- Remaining non-editable regions are listed honestly.
