# Scientific Figure Shapes Office Shape Recipes

Load this file when writing unfamiliar VBA shape code.

## Slide Setup

```vb
Dim pres As Presentation
Dim sld As Slide
Set pres = Application.Presentations.Add(msoTrue)
pres.PageSetup.SlideWidth = 960
pres.PageSetup.SlideHeight = 540
Set sld = pres.Slides.Add(1, ppLayoutBlank)
```

## Named Rectangle

```vb
Dim box As Shape
Set box = sld.Shapes.AddShape(msoShapeRectangle, 40, 40, 220, 90)
box.Name = "SUMMER_E_panel_01"
box.Fill.ForeColor.RGB = RGB(245, 248, 252)
box.Line.ForeColor.RGB = RGB(80, 110, 130)
box.Line.Weight = 1.25
```

## Text Box

```vb
Dim t As Shape
Set t = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, 60, 55, 180, 40)
t.Name = "SUMMER_E_label_01"
With t.TextFrame2
    .MarginLeft = 0
    .MarginRight = 0
    .MarginTop = 0
    .MarginBottom = 0
    .TextRange.Text = "Barrier dysfunction"
    .TextRange.Font.Name = "Arial"
    .TextRange.Font.Size = 14
    .TextRange.Font.Bold = msoTrue
End With
```

## Line With Arrow

```vb
Dim ln As Shape
Set ln = sld.Shapes.AddLine(100, 120, 210, 120)
ln.Name = "SUMMER_L_arrow_01"
ln.Line.ForeColor.RGB = RGB(0, 0, 0)
ln.Line.Weight = 1.6
ln.Line.EndArrowheadStyle = msoArrowheadTriangle
```

## Fan-In Routing Without Crossing Labels

Do not route every branch to the same hard-coded point. Name the merge box and each final segment, then enter through different edges:

```vb
Dim mergeBox As Shape
Dim branchLeft As Shape
Dim branchCenter As Shape
Dim branchRight As Shape

Set mergeBox = sld.Shapes.AddShape(msoShapeRoundedRectangle, 84, 117, 111, 25)
mergeBox.Name = "SUMMER_E_merge_box"

' Left branch enters the left edge.
Set branchLeft = sld.Shapes.AddLine(62, 129, 84, 129)
branchLeft.Name = "SUMMER_L_branch_left_to_merge"
branchLeft.Line.EndArrowheadStyle = msoArrowheadTriangle

' Center branch enters the top edge.
Set branchCenter = sld.Shapes.AddLine(139.5, 102, 139.5, 117)
branchCenter.Name = "SUMMER_L_branch_center_to_merge"
branchCenter.Line.EndArrowheadStyle = msoArrowheadTriangle

' Right branch enters the right edge.
Set branchRight = sld.Shapes.AddLine(216, 129, 195, 129)
branchRight.Name = "SUMMER_L_branch_right_to_merge"
branchRight.Line.EndArrowheadStyle = msoArrowheadTriangle
```

Record the expected edges in a routing manifest used by `pptx_connector_audit.py`:

```json
{
  "routing_audit": {
    "routes": [
      {
        "connector": "SUMMER_L_branch_left_to_merge",
        "target": "SUMMER_E_merge_box",
        "target_edge": "left"
      },
      {
        "connector": "SUMMER_L_branch_center_to_merge",
        "target": "SUMMER_E_merge_box",
        "target_edge": "top"
      },
      {
        "connector": "SUMMER_L_branch_right_to_merge",
        "target": "SUMMER_E_merge_box",
        "target_edge": "right"
      }
    ]
  }
}
```

Run:

```bash
python scripts/pptx_connector_audit.py rebuilt.pptx \
  --manifest routing_manifest.json --pretty
```

If an arrow intentionally points into a text-only label, list that exact text shape or connector/text pair in the manifest. Never ignore all collisions.

## Fan-Out Bus Without Duplicate Strokes

Do not build each branch with a full polyline from the source. That repeats the shared trunk and can create a zero-length segment for the centered branch. Draw the shared geometry once:

```vb
Dim trunk As Shape
Dim bus As Shape
Dim dropLeft As Shape
Dim dropCenter As Shape
Dim dropRight As Shape

' Shared trunk: draw once.
Set trunk = sld.Shapes.AddLine(139, 58, 139, 66)
trunk.Name = "SUMMER_L_split_trunk"

' Shared horizontal bus: draw once.
Set bus = sld.Shapes.AddLine(62, 66, 216, 66)
bus.Name = "SUMMER_L_split_bus"

' Three independent drops; no zero-length horizontal center segment.
Set dropLeft = sld.Shapes.AddLine(62, 66, 62, 77)
dropLeft.Name = "SUMMER_L_split_drop_left"
dropLeft.Line.EndArrowheadStyle = msoArrowheadTriangle

Set dropCenter = sld.Shapes.AddLine(139, 66, 139, 77)
dropCenter.Name = "SUMMER_L_split_drop_center"
dropCenter.Line.EndArrowheadStyle = msoArrowheadTriangle

Set dropRight = sld.Shapes.AddLine(216, 66, 216, 77)
dropRight.Name = "SUMMER_L_split_drop_right"
dropRight.Line.EndArrowheadStyle = msoArrowheadTriangle
```

`pptx_connector_audit.py` treats duplicate connector geometry and zero-length connectors as failures. If the bus looks darker than its drops, assume the trunk/bus was drawn more than once and inspect the generated object list.

## Three-Face Cuboid With Controlled Depth

Use three separately named faces when the source depends on depth, overlap, or feature-pyramid scale. Do not rely on `msoShapeCube` adjustments.

```vb
Dim frontFace As Shape
Dim topFace As Shape
Dim rightFace As Shape
Dim ff As FreeformBuilder

' Front face.
Set frontFace = sld.Shapes.AddShape(msoShapeRectangle, 100, 70, 60, 80)
frontFace.Name = "SUMMER_E_layer_c3_front"
frontFace.Fill.ForeColor.RGB = RGB(213, 229, 250)
frontFace.Line.ForeColor.RGB = RGB(46, 111, 239)

' Top face: one consistent depth vector (+14, -14).
Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 100, 70)
ff.AddNodes msoSegmentLine, msoEditingAuto, 114, 56
ff.AddNodes msoSegmentLine, msoEditingAuto, 174, 56
ff.AddNodes msoSegmentLine, msoEditingAuto, 160, 70
ff.AddNodes msoSegmentLine, msoEditingAuto, 100, 70
Set topFace = ff.ConvertToShape
topFace.Name = "SUMMER_E_layer_c3_top"
topFace.Fill.ForeColor.RGB = RGB(225, 238, 252)
topFace.Line.ForeColor.RGB = RGB(46, 111, 239)

' Right face: darker than the front face.
Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 160, 70)
ff.AddNodes msoSegmentLine, msoEditingAuto, 174, 56
ff.AddNodes msoSegmentLine, msoEditingAuto, 174, 136
ff.AddNodes msoSegmentLine, msoEditingAuto, 160, 150
ff.AddNodes msoSegmentLine, msoEditingAuto, 160, 70
Set rightFace = ff.ConvertToShape
rightFace.Name = "SUMMER_E_layer_c3_right"
rightFace.Fill.ForeColor.RGB = RGB(175, 188, 205)
rightFace.Line.ForeColor.RGB = RGB(46, 111, 239)
```

Create the largest/back layer first, then progressively smaller foreground layers. A layering manifest records face names and hierarchy:

```json
{
  "layering_audit": {
    "cuboids": [
      {
        "id": "c3",
        "front": "SUMMER_E_layer_c3_front",
        "top": "SUMMER_E_layer_c3_top",
        "right": "SUMMER_E_layer_c3_right"
      }
    ],
    "size_order": ["c5", "c4", "c3"],
    "z_order": ["c3", "c4", "c5"],
    "overlap_pairs": [["c4", "c3"]]
  }
}
```

Run `python scripts/pptx_layering_audit.py rebuilt.pptx --manifest layering_manifest.json --pretty` before handoff.

## Picture Crop Insert

```vb
Dim pic As Shape
Set pic = sld.Shapes.AddPicture( _
    FileName:=assetPath, _
    LinkToFile:=msoFalse, _
    SaveWithDocument:=msoTrue, _
    Left:=xPt, Top:=yPt, Width:=wPt, Height:=hPt)
pic.Name = "SUMMER_R_anatomy_01"
pic.LockAspectRatio = msoTrue
```

## Cleanup Existing Slide

```vb
Do While sld.Shapes.Count > 0
    sld.Shapes(1).Delete
Loop
```

## Practical Notes

- Prefer simple solid fills and lines for compatibility.
- Use freeforms sparingly; many small freeforms are hard to maintain.
- Name every generated object.
- Name merge boxes and final connector segments semantically so routing manifests remain stable across revisions.
- Keep helper procedures small: `AddText`, `AddShape`, `AddLine`, `AddPicture`.
- Use `TextFrame2` when PowerPoint is the target; avoid advanced typography when WPS compatibility matters.
