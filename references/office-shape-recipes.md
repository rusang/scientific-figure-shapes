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
- Keep helper procedures small: `AddText`, `AddShape`, `AddLine`, `AddPicture`.
- Use `TextFrame2` when PowerPoint is the target; avoid advanced typography when WPS compatibility matters.
