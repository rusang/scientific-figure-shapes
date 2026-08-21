from __future__ import annotations

from pathlib import Path


def emit_vba(canvas, path: Path) -> None:
    lines = [
        "Attribute VB_Name = \"MPHFNetReconstruction\"",
        "Option Explicit",
        "",
        "Public Sub BuildFinal()",
        "    Dim pres As Presentation",
        "    Dim sld As Slide",
        "    Dim shp As Shape",
        "    Dim ff As FreeformBuilder",
        "    Dim assetBase As String",
        "    assetBase = ActivePresentation.Path",
        "    Set pres = Application.Presentations.Add(msoTrue)",
        f"    pres.PageSetup.SlideWidth = {canvas.width_pt:.3f}",
        f"    pres.PageSetup.SlideHeight = {canvas.height_pt:.3f}",
        "    Set sld = pres.Slides.Add(1, ppLayoutBlank)",
        "    Do While sld.Shapes.Count > 0",
        "        sld.Shapes(1).Delete",
        "    Loop",
    ]
    for op in canvas.ops:
        v = op.values
        if op.kind == "shape":
            lines.append(
                f"    Set shp = AddFigureShape(sld, {v['vba_shape']}, {v['x']:.3f}, {v['y']:.3f}, {v['w']:.3f}, {v['h']:.3f}, "
                f"{canvas._vba_color(v['fill'])}, {canvas._vba_color(v['line'])}, {v['weight']:.3f}, {v['transparency']}, \"{v['name']}\")"
            )
            if v["rotation"]:
                lines.append(f"    shp.Rotation = {v['rotation']:.3f}")
            if v.get("gradient"):
                gradient = v["gradient"]
                style = "msoGradientVertical" if abs(gradient["angle"] - 90) <= 45 else "msoGradientHorizontal"
                lines.append(f"    shp.Fill.TwoColorGradient {style}, 1")
                lines.append(f"    shp.Fill.ForeColor.RGB = {canvas._vba_color(gradient['start'])}")
                lines.append(f"    shp.Fill.BackColor.RGB = {canvas._vba_color(gradient['end'])}")
        elif op.kind == "freeform":
            first_x, first_y = v["points"][0]
            lines.append(
                f"    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, {first_x:.3f}, {first_y:.3f})"
            )
            for point_x, point_y in v["points"][1:]:
                lines.append(
                    f"    ff.AddNodes msoSegmentLine, msoEditingAuto, {point_x:.3f}, {point_y:.3f}"
                )
            lines.append(
                f"    ff.AddNodes msoSegmentLine, msoEditingAuto, {first_x:.3f}, {first_y:.3f}"
            )
            lines.append("    Set shp = ff.ConvertToShape")
            lines.append(f"    shp.Name = \"{v['name']}\"")
            lines.append("    shp.Shadow.Visible = msoFalse")
            if v["fill"] is None:
                lines.append("    shp.Fill.Visible = msoFalse")
            else:
                lines.append(
                    f"    shp.Fill.ForeColor.RGB = {canvas._vba_color(v['fill'])}: shp.Fill.Transparency = {v['transparency']} / 100#"
                )
            if v["line"] is None:
                lines.append("    shp.Line.Visible = msoFalse")
            else:
                lines.append(
                    f"    shp.Line.ForeColor.RGB = {canvas._vba_color(v['line'])}: shp.Line.Weight = {v['weight']:.3f}"
                )
            if v["rotation"]:
                lines.append(f"    shp.Rotation = {v['rotation']:.3f}")
            if v.get("gradient"):
                gradient = v["gradient"]
                style = "msoGradientVertical" if abs(gradient["angle"] - 90) <= 45 else "msoGradientHorizontal"
                lines.append(f"    shp.Fill.TwoColorGradient {style}, 1")
                lines.append(f"    shp.Fill.ForeColor.RGB = {canvas._vba_color(gradient['start'])}")
                lines.append(f"    shp.Fill.BackColor.RGB = {canvas._vba_color(gradient['end'])}")
        elif op.kind == "text":
            lines.append(
                f"    Set shp = AddFigureText(sld, {v['x']:.3f}, {v['y']:.3f}, {v['w']:.3f}, {v['h']:.3f}, "
                f"{canvas._vba_string(v['value'])}, {v['size']:.3f}, {canvas._vba_color(v['color'])}, "
                f"{str(v['bold']).lower()}, \"{v['align']}\", \"{v['valign']}\", {v['margin']:.3f}, "
                f"\"{v['font_name']}\", \"{v['name']}\")"
            )
            if v["rotation"]:
                lines.append(f"    shp.Rotation = {v['rotation']:.3f}")
        elif op.kind == "line":
            lines.append(
                f"    Set shp = AddFigureLine(sld, {v['x1']:.3f}, {v['y1']:.3f}, {v['x2']:.3f}, {v['y2']:.3f}, "
                f"{canvas._vba_color(v['color'])}, {v['weight']:.3f}, {str(v['dash']).lower()}, "
                f"{str(v['arrow_begin']).lower()}, {str(v['arrow_end']).lower()}, \"{v['name']}\")"
            )
        elif op.kind == "picture":
            filename = Path(v["path"]).name
            lines.append(
                f"    Set shp = AddFigurePicture(sld, ResolveAssetPath(assetBase, \"{filename}\", \"{v['path']}\"), "
                f"{v['x']:.3f}, {v['y']:.3f}, {v['w']:.3f}, {v['h']:.3f}, \"{v['name']}\")"
            )
        elif op.kind == "reference_slide_picture":
            filename = Path(v["path"]).name
            lines.append("    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)")
            lines.append(
                f"    Set shp = AddFigurePicture(sld, ResolveAssetPath(assetBase, \"{filename}\", \"{v['path']}\"), "
                f"{v['x']:.3f}, {v['y']:.3f}, {v['w']:.3f}, {v['h']:.3f}, \"{v['name']}\")"
            )
    grouped = {key: value for key, value in canvas._groups.items() if len(value) > 1}
    if grouped:
        lines.append("    Set sld = pres.Slides(1)")
    for group_id, members in grouped.items():
        member_list = ", ".join(f'\"{name}\"' for name in members)
        safe_group = "".join(
            character if character.isalnum() or character == "_" else "_"
            for character in group_id
        )
        lines.append(
            f"    Set shp = sld.Shapes.Range(Array({member_list})).Group"
        )
        lines.append(f"    shp.Name = \"{canvas.name_prefix}_G_{safe_group}\"")
    lines.extend(
        [
            "End Sub",
            "",
            "Private Function AddFigureShape(sld As Slide, kind As MsoAutoShapeType, x As Single, y As Single, w As Single, h As Single, fillRgb As Long, lineRgb As Long, lineWeight As Single, fillTransparency As Integer, shapeName As String) As Shape",
            "    Dim shp As Shape",
            "    Set shp = sld.Shapes.AddShape(kind, x, y, w, h)",
            "    shp.Name = shapeName",
            "    shp.Shadow.Visible = msoFalse",
            "    If kind = msoShapeRoundedRectangle Then shp.Adjustments(1) = 0.08",
            "    If fillRgb < 0 Then shp.Fill.Visible = msoFalse Else shp.Fill.ForeColor.RGB = fillRgb: shp.Fill.Transparency = fillTransparency / 100#",
            "    If lineRgb < 0 Then shp.Line.Visible = msoFalse Else shp.Line.ForeColor.RGB = lineRgb: shp.Line.Weight = lineWeight",
            "    Set AddFigureShape = shp",
            "End Function",
            "",
            "Private Function AddFigureText(sld As Slide, x As Single, y As Single, w As Single, h As Single, txt As String, fontSize As Single, fontRgb As Long, isBold As Boolean, hAlign As String, vAlign As String, margin As Single, fontName As String, shapeName As String) As Shape",
            "    Dim shp As Shape",
            "    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)",
            "    shp.Name = shapeName",
            "    shp.Shadow.Visible = msoFalse",
            "    With shp.TextFrame2",
            "        .MarginLeft = margin: .MarginRight = margin: .MarginTop = margin: .MarginBottom = margin",
            "        .WordWrap = msoTrue",
            "        If vAlign = \"top\" Then .VerticalAnchor = msoAnchorTop Else If vAlign = \"bottom\" Then .VerticalAnchor = msoAnchorBottom Else .VerticalAnchor = msoAnchorMiddle",
            "        .TextRange.Text = txt",
            "        .TextRange.Font.Name = fontName",
            "        .TextRange.Font.Size = fontSize",
            "        .TextRange.Font.Bold = isBold",
            "        .TextRange.Font.Fill.ForeColor.RGB = fontRgb",
            "        If hAlign = \"left\" Then .TextRange.ParagraphFormat.Alignment = msoAlignLeft Else If hAlign = \"right\" Then .TextRange.ParagraphFormat.Alignment = msoAlignRight Else .TextRange.ParagraphFormat.Alignment = msoAlignCenter",
            "    End With",
            "    Set AddFigureText = shp",
            "End Function",
            "",
            "Private Function AddFigureLine(sld As Slide, x1 As Single, y1 As Single, x2 As Single, y2 As Single, lineRgb As Long, lineWeight As Single, isDashed As Boolean, arrowBegin As Boolean, arrowEnd As Boolean, shapeName As String) As Shape",
            "    Dim shp As Shape",
            "    Set shp = sld.Shapes.AddLine(x1, y1, x2, y2)",
            "    shp.Name = shapeName",
            "    shp.Line.ForeColor.RGB = lineRgb: shp.Line.Weight = lineWeight",
            "    If isDashed Then shp.Line.DashStyle = msoLineDash",
            "    If arrowBegin Then shp.Line.BeginArrowheadStyle = msoArrowheadTriangle",
            "    If arrowEnd Then shp.Line.EndArrowheadStyle = msoArrowheadTriangle",
            "    Set AddFigureLine = shp",
            "End Function",
            "",
            "Private Function AddFigurePicture(sld As Slide, assetPath As String, x As Single, y As Single, w As Single, h As Single, shapeName As String) As Shape",
            "    Dim shp As Shape",
            "    Set shp = sld.Shapes.AddPicture(assetPath, msoFalse, msoTrue, x, y, w, h)",
            "    shp.Name = shapeName: shp.LockAspectRatio = msoTrue",
            "    Set AddFigurePicture = shp",
            "End Function",
            "",
            "Private Function ResolveAssetPath(baseFolder As String, assetFile As String, fallbackPath As String) As String",
            "    Dim sep As String",
            "    Dim candidate As String",
            "    sep = Application.PathSeparator",
            "    If InStr(baseFolder, \"/\") > 0 Then sep = \"/\"",
            "    If Len(baseFolder) > 0 Then",
            "        candidate = baseFolder & sep & \"assets\" & sep & assetFile",
            "        If Len(Dir$(candidate)) > 0 Then ResolveAssetPath = candidate: Exit Function",
            "    End If",
            "    ResolveAssetPath = fallbackPath",
            "End Function",
        ]
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
