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

## 4. Text Gate

- Text fits inside its box on the rendered preview.
- Important labels are not hidden behind crops.
- Font size hierarchy is close enough for the intended use.
- Biological symbols, arrows, down/up marks, and abbreviations are not accidentally changed.

## 5. Routing Gate

- Major arrows point to the right target region.
- Dashed zoom lines attach to the intended source and destination.
- Lines do not cross text in a way that harms reading.

## 6. Render Gate

When a preview exists:

- no large blank region appears unexpectedly.
- no crop is visibly stretched.
- main panels sit in the expected order.
- source and preview may be compared with `scripts/render_delta_probe.py`.

## 7. Handoff Gate

- Macro, assets, manifest, preview, and fallback deck are clearly labeled.
- Automation failures are reported as local execution issues, not as successful macro runs.
- Remaining non-editable regions are listed honestly.
