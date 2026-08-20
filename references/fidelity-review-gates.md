# Scientific Figure Shapes Fidelity Review Gates

Use this checklist only for high-fidelity or repeated-correction work.

## 1. Canvas Gate

- Source size and target slide size are recorded.
- Mapping mode is explicit: uniform fit or exact same-ratio scale.
- No full-source hidden background remains in the final editable deliverable unless requested.

## 2. Crop Gate

For every `R*` crop:

- bbox is recorded in source pixels.
- crop file exists and has nonzero size.
- crop aspect ratio matches the recorded bbox.
- inserted image keeps aspect ratio.
- rotation and flips are intentional and documented.
- nearby labels, arrows, and highlights are rebuilt as editable objects when practical.

## 3. Editable Structure Gate

Check that these are editable when present:

- titles and labels
- panels, boxes, callouts, dividers
- arrows and connector lines
- legends, axes, simple charts, tables
- repeated simple marks and cells

## 4. Depth and Layering Gate

- Visually important 3D blocks use explicit `front`, `top`, and `right` faces instead of one uncontrollable auto-cube.
- Each face has source-derived expected bounds; materialized geometry stays within the declared tolerance (normally no more than 1 pt after Office quantization).
- All layers share one depth vector and consistent face shading direction.
- Faces marked as gradient in the manifest remain gradient after PPTX materialization.
- Feature-pyramid sizes increase in the intended order.
- Required layer pairs overlap, and back-to-front z-order matches the source.
- Run `scripts/pptx_layering_audit.py`; missing faces, flat colors, detached faces, reference-geometry errors, missing gradients, hierarchy errors, overlap errors, or z-order errors fail the gate.

## 5. Text Gate

- Text fits inside its box on the rendered preview.
- Important labels are not hidden behind crops.
- Font size hierarchy is close enough for the intended use.
- Biological symbols, arrows, down/up marks, and abbreviations are not accidentally changed.

## 6. Routing Gate

- Major arrows point to the right target region.
- Dashed zoom lines attach to the intended source and destination.
- Every important final connector segment has a stable name and a recorded target edge (`left`, `right`, `top`, or `bottom`).
- Shared fan-out trunks/buses are drawn once, not repeated once per branch.
- No connector is zero-length, and no two connectors have duplicate geometry unless one is explicitly excluded for a documented reason.
- Fan-in routes do not reuse one hard-coded endpoint when that would force a line through the merge label.
- Lines may touch target boundaries but do not enter the shrunken text rectangle or cross labels.
- Run `scripts/pptx_connector_audit.py` on the materialized deck. Any collision or target-edge mismatch is a failure.
- Intentional line-to-text cases are narrow manifest exceptions (`ignore_text_shapes` / `ignore_pairs`), not a global audit bypass.

## 7. Render Gate

When a preview exists:

- no large blank region appears unexpectedly.
- no crop is visibly stretched.
- main panels sit in the expected order.
- source and preview may be compared with `scripts/render_delta_probe.py`.

## 8. Handoff Gate

- Macro, assets, manifest, preview, and fallback deck are clearly labeled.
- Automation failures are reported as local execution issues, not as successful macro runs.
- Remaining non-editable regions are listed honestly.
