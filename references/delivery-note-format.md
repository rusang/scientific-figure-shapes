# Scientific Figure Shapes Delivery Note Format

Use this file only when the user needs a fuller handoff than the normal final response.

## Short Report Shape

Write in Simplified Chinese unless the user asked otherwise.

Recommended sections:

1. **交付物**
   - List the editable `.bas` first.
   - List the editable `.pptx` fallback when created.
   - List crop assets and previews separately.

2. **可编辑范围**
   - State which regions are Office Shapes/TextBoxes/Lines.
   - State which regions are preserved local images.
   - Mention that preview PNG/PDF files are diagnostic only.

3. **运行与验证**
   - VBA smoke lint result.
   - Office automation result, if attempted.
   - Fallback path if automation was blocked.
   - Preview rendering result, if available.

4. **手动使用**
   - PowerPoint: import or paste the `.bas`, then run `BuildFinal`.
   - WPS: use the fallback `.pptx` if macro support is unavailable.
   - If crop assets are referenced by macro path, keep the assets folder beside the macro or update paths.

## Concision Rules

- Do not paste long macro code in the report.
- Do not overclaim visual identity if no rendered preview was checked.
- Do not describe a crop as editable.
- Do not call a PNG/PDF the final product.

## Example Snippet

```markdown
## 结果

- `figure_shapes.bas`: editable VBA reconstruction macro.
- `figure_editable_fallback.pptx`: editable fallback deck built from the same element plan.
- `assets/*.png`: local crops for complex artwork.
- `preview.png`: diagnostic preview only.

## 验证

- Macro smoke lint: passed.
- PowerPoint automation: blocked by local macro/automation policy.
- Fallback deck: created and opened/preview-rendered.
```
