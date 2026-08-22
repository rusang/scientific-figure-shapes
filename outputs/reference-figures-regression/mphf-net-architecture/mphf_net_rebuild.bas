Attribute VB_Name = "MPHFNetReconstruction"
Option Explicit

Public Sub BuildFinal()
    Dim pres As Presentation
    Dim sld As Slide
    Dim shp As Shape
    Dim ff As FreeformBuilder
    Dim assetBase As String
    assetBase = ActivePresentation.Path
    Set pres = Application.Presentations.Add(msoTrue)
    pres.PageSetup.SlideWidth = 960.000
    pres.PageSetup.SlideHeight = 640.000
    Set sld = pres.Slides.Add(1, ppLayoutBlank)
    Do While sld.Shapes.Count > 0
        sld.Shapes(1).Delete
    Loop
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 2.500, 2.500, 955.000, 37.500, RGB(1, 23, 73), -1, 1.000, 0, "SUMMER_B_banner")
    Set shp = AddFigureText(sld, 6.250, 5.000, 947.500, 31.250, "MPHF-Net 工业表面缺陷检测网络结构设计图", 26.000, RGB(255, 255, 255), true, "center", "middle", 0.000, "Arial", "SUMMER_T_banner")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 7.500, 48.750, 123.750, 290.625, RGB(253, 253, 253), RGB(192, 203, 224), 0.800, 0, "SUMMER_E_panel_input")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 141.250, 48.750, 160.625, 290.625, RGB(253, 253, 253), RGB(192, 203, 224), 0.800, 0, "SUMMER_E_panel_backbone")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 310.625, 48.750, 353.125, 290.625, RGB(253, 253, 253), RGB(192, 203, 224), 0.800, 0, "SUMMER_E_panel_neck")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 672.500, 48.750, 156.875, 290.625, RGB(253, 253, 253), RGB(192, 203, 224), 0.800, 0, "SUMMER_E_panel_head")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 838.125, 48.750, 113.125, 290.625, RGB(253, 253, 253), RGB(192, 203, 224), 0.800, 0, "SUMMER_E_panel_legend")
    Set shp = AddFigureText(sld, 7.500, 53.750, 125.000, 13.750, "输入图像", 11.000, RGB(40, 44, 60), true, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_input_title")
    Set shp = AddFigurePicture(sld, ResolveAssetPath(assetBase, "01_R1_pcb.png", "outputs/reference-figures-regression/mphf-net-architecture/assets/01_R1_pcb.png"), 15.625, 101.250, 106.875, 152.500, "SUMMER_R1_pcb")
    Set shp = AddFigureText(sld, 7.500, 268.750, 125.000, 15.000, "640 × 640 × 3", 11.000, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_input_size")
    Set shp = AddFigureText(sld, 142.500, 53.750, 161.250, 12.500, "Backbone: CSP-MEEM", 11.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_backbone_title")
    Set shp = AddFigureText(sld, 142.500, 67.500, 161.250, 10.000, "多尺度细粒度特征提取", 8.000, RGB(120, 126, 140), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_backbone_sub")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 175.000, 105.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 180.000, 100.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 204.375, 100.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 199.375, 105.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 175.000, 105.625
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_bb1_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 175.000, 105.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 199.375, 105.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 199.375, 128.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 175.000, 128.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 175.000, 105.625
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_bb1_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 199.375, 105.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 204.375, 100.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 204.375, 123.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 199.375, 128.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 199.375, 105.625
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_bb1_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 171.250, 158.750)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 180.625, 149.375
    ff.AddNodes msoSegmentLine, msoEditingAuto, 209.375, 149.375
    ff.AddNodes msoSegmentLine, msoEditingAuto, 200.000, 158.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 171.250, 158.750
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_bb2_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 171.250, 158.750)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 200.000, 158.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 200.000, 196.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 171.250, 196.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 171.250, 158.750
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_bb2_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 200.000, 158.750)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 209.375, 149.375
    ff.AddNodes msoSegmentLine, msoEditingAuto, 209.375, 186.875
    ff.AddNodes msoSegmentLine, msoEditingAuto, 200.000, 196.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 200.000, 158.750
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_bb2_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 165.000, 195.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 178.125, 181.875
    ff.AddNodes msoSegmentLine, msoEditingAuto, 215.625, 181.875
    ff.AddNodes msoSegmentLine, msoEditingAuto, 202.500, 195.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 165.000, 195.000
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_bb3_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 165.000, 195.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 202.500, 195.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 202.500, 232.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 165.000, 232.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 165.000, 195.000
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_bb3_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 202.500, 195.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 215.625, 181.875
    ff.AddNodes msoSegmentLine, msoEditingAuto, 215.625, 219.375
    ff.AddNodes msoSegmentLine, msoEditingAuto, 202.500, 232.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 202.500, 195.000
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_bb3_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 156.250, 235.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 173.125, 218.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 215.625, 218.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 198.750, 235.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 156.250, 235.625
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_bb4_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 156.250, 235.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 198.750, 235.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 198.750, 275.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 156.250, 275.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 156.250, 235.625
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_bb4_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 198.750, 235.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 215.625, 218.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 215.625, 258.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 198.750, 275.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 198.750, 235.625
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_bb4_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 168.750, 282.500, 25.000, 11.250, "⋮", 11.000, RGB(120, 126, 140), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_backbone_dots")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 246.875, 125.000, 38.125, 33.750, RGB(245, 248, 255), RGB(37, 80, 216), 1.000, 0, "SUMMER_E_bb_c5")
    Set shp = AddFigureText(sld, 246.875, 127.500, 38.125, 15.000, "C5", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_bb_c5")
    Set shp = AddFigureText(sld, 246.875, 142.500, 38.125, 13.750, "20×C", 8.000, RGB(37, 80, 216), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_bb_c5_dim")
    Set shp = AddFigureLine(sld, 204.375, 114.375, 231.875, 114.375, RGB(40, 44, 60), 0.800, false, false, false, "SUMMER_L_bb_to_spine")
    Set shp = AddFigureLine(sld, 231.875, 114.375, 231.875, 258.750, RGB(40, 44, 60), 0.800, false, false, false, "SUMMER_L_bb_output_spine")
    Set shp = AddFigureLine(sld, 190.000, 128.125, 190.000, 149.375, RGB(40, 44, 60), 0.800, false, false, true, "SUMMER_L_bb_down_1")
    Set shp = AddFigureLine(sld, 190.000, 187.500, 190.000, 202.500, RGB(40, 44, 60), 0.800, false, false, true, "SUMMER_L_bb_down_2")
    Set shp = AddFigureLine(sld, 190.000, 219.375, 190.000, 237.500, RGB(40, 44, 60), 0.800, false, false, true, "SUMMER_L_bb_down_3")
    Set shp = AddFigureLine(sld, 231.875, 141.875, 246.875, 141.875, RGB(40, 44, 60), 1.200, false, false, true, "SUMMER_L_bb_to_c5")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 246.875, 183.750, 38.125, 33.750, RGB(245, 252, 242), RGB(28, 148, 84), 1.000, 0, "SUMMER_E_bb_c4")
    Set shp = AddFigureText(sld, 246.875, 186.250, 38.125, 15.000, "C4", 10.000, RGB(28, 148, 84), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_bb_c4")
    Set shp = AddFigureText(sld, 246.875, 201.250, 38.125, 13.750, "40×C", 8.000, RGB(28, 148, 84), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_bb_c4_dim")
    Set shp = AddFigureLine(sld, 231.875, 200.625, 246.875, 200.625, RGB(40, 44, 60), 1.200, false, false, true, "SUMMER_L_bb_to_c4")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 246.875, 241.875, 38.125, 33.750, RGB(255, 248, 243), RGB(224, 58, 38), 1.000, 0, "SUMMER_E_bb_c3")
    Set shp = AddFigureText(sld, 246.875, 244.375, 38.125, 15.000, "C3", 10.000, RGB(224, 58, 38), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_bb_c3")
    Set shp = AddFigureText(sld, 246.875, 259.375, 38.125, 13.750, "80×C", 8.000, RGB(224, 58, 38), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_bb_c3_dim")
    Set shp = AddFigureLine(sld, 231.875, 258.750, 246.875, 258.750, RGB(40, 44, 60), 1.200, false, false, true, "SUMMER_L_bb_to_c3")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 150.000, 312.500, 135.000, 20.000, RGB(235, 241, 255), RGB(214, 219, 232), 1.000, 0, "SUMMER_E_bb_outbar")
    Set shp = AddFigureText(sld, 150.000, 315.000, 135.000, 15.000, "输出多尺度特征", 9.000, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_bb_outbar")
    Set shp = AddFigureText(sld, 311.250, 53.750, 358.750, 12.500, "Neck: FPN-PAN （LOSC + CGAFusion）", 11.000, RGB(28, 148, 84), true, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_neck_title")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 337.500, 72.500, 81.250, 16.250, RGB(218, 239, 210), -1, 0.900, 0, "SUMMER_E_neck_td")
    Set shp = AddFigureText(sld, 337.500, 74.375, 81.250, 12.500, "Top-Down 路径", 8.500, RGB(40, 44, 60), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_td")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 517.500, 72.500, 81.250, 16.250, RGB(218, 239, 210), -1, 0.800, 0, "SUMMER_E_neck_bu")
    Set shp = AddFigureText(sld, 517.500, 74.375, 81.250, 12.500, "Bottom-Up 路径", 8.500, RGB(40, 44, 60), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_bu")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 350.000, 95.000, 25.000, 17.500, RGB(245, 248, 255), RGB(37, 80, 216), 1.200, 0, "SUMMER_E_neck_c5")
    Set shp = AddFigureText(sld, 350.000, 97.500, 25.000, 12.500, "C5", 9.000, RGB(37, 80, 216), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_c5")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 315.000, 166.250, 25.000, 17.500, RGB(238, 249, 233), RGB(28, 148, 84), 0.900, 0, "SUMMER_E_neck_c4")
    Set shp = AddFigureText(sld, 315.000, 168.750, 25.000, 12.500, "C4", 9.000, RGB(28, 148, 84), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_c4")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 315.000, 255.000, 25.000, 17.500, RGB(255, 239, 224), RGB(224, 58, 38), 0.900, 0, "SUMMER_E_neck_c3")
    Set shp = AddFigureText(sld, 315.000, 257.500, 25.000, 12.500, "C3", 9.000, RGB(224, 58, 38), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_c3")
    Set shp = AddFigureShape(sld, msoShapeOval, 351.250, 126.250, 25.000, 25.000, RGB(209, 193, 246), RGB(105, 55, 190), 0.800, 0, "SUMMER_E_neck_up1")
    Set shp = AddFigureText(sld, 351.250, 126.250, 25.000, 25.000, "UP" & vbLf & "↑", 5.500, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_neck_up1_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 351.250, 202.500, 25.000, 25.000, RGB(209, 193, 246), RGB(105, 55, 190), 0.800, 0, "SUMMER_E_neck_up2")
    Set shp = AddFigureText(sld, 351.250, 202.500, 25.000, 25.000, "UP" & vbLf & "↑", 5.500, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_neck_up2_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 351.250, 162.500, 25.000, 25.000, RGB(158, 226, 235), RGB(25, 150, 170), 0.800, 0, "SUMMER_E_neck_add_td1")
    Set shp = AddFigureText(sld, 351.250, 162.500, 25.000, 25.000, "+", 6.500, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_neck_add_td1_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 351.250, 250.000, 25.000, 25.000, RGB(158, 226, 235), RGB(25, 150, 170), 0.800, 0, "SUMMER_E_neck_add_td2")
    Set shp = AddFigureText(sld, 351.250, 250.000, 25.000, 25.000, "+", 6.500, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_neck_add_td2_text")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 406.250, 165.625, 57.500, 18.750, RGB(240, 245, 255), RGB(37, 80, 216), 1.200, 0, "SUMMER_E_neck_fusion1")
    Set shp = AddFigureText(sld, 406.250, 168.750, 57.500, 12.500, "CGAFusion", 8.500, RGB(37, 80, 216), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_fusion1")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 406.250, 253.125, 57.500, 18.750, RGB(240, 245, 255), RGB(37, 80, 216), 1.200, 0, "SUMMER_E_neck_fusion2")
    Set shp = AddFigureText(sld, 406.250, 256.250, 57.500, 12.500, "CGAFusion", 8.500, RGB(37, 80, 216), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_fusion2")
    Set shp = AddFigureShape(sld, msoShapeOval, 545.000, 101.250, 25.000, 25.000, RGB(158, 226, 235), RGB(25, 150, 170), 0.800, 0, "SUMMER_E_neck_add_p5")
    Set shp = AddFigureText(sld, 545.000, 101.250, 25.000, 25.000, "+", 6.500, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_neck_add_p5_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 545.000, 201.250, 25.000, 25.000, RGB(158, 226, 235), RGB(25, 150, 170), 0.800, 0, "SUMMER_E_neck_add_p4")
    Set shp = AddFigureText(sld, 545.000, 201.250, 25.000, 25.000, "+", 6.500, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_neck_add_p4_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 545.000, 280.000, 25.000, 25.000, RGB(158, 226, 235), RGB(25, 150, 170), 0.800, 0, "SUMMER_E_neck_add_p3")
    Set shp = AddFigureText(sld, 545.000, 280.000, 25.000, 25.000, "+", 6.500, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_neck_add_p3_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 545.000, 138.750, 25.000, 25.000, RGB(244, 172, 96), -1, 0.800, 0, "SUMMER_E_neck_down1")
    Set shp = AddFigureText(sld, 545.000, 138.750, 25.000, 25.000, "Down" & vbLf & "↓", 4.000, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_neck_down1_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 545.000, 240.000, 25.000, 25.000, RGB(244, 172, 96), -1, 0.800, 0, "SUMMER_E_neck_down2")
    Set shp = AddFigureText(sld, 545.000, 240.000, 25.000, 25.000, "Down" & vbLf & "↓", 4.000, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_neck_down2_text")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 535.000, 180.000, 45.000, 16.250, RGB(253, 240, 232), RGB(226, 110, 50), 1.200, 0, "SUMMER_E_neck_losc")
    Set shp = AddFigureText(sld, 535.000, 181.875, 45.000, 12.500, "LOSC", 9.000, RGB(206, 90, 30), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_losc")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 607.500, 106.250)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 612.500, 101.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 627.500, 101.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 106.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 607.500, 106.250
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_neck_p5_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 607.500, 106.250)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 106.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 126.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 607.500, 126.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 607.500, 106.250
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_neck_p5_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 622.500, 106.250)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 627.500, 101.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 627.500, 121.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 126.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 106.250
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_neck_p5_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 632.500, 96.250, 32.500, 12.500, "P5", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_p5")
    Set shp = AddFigureText(sld, 632.500, 107.500, 32.500, 10.000, "20×C", 8.000, RGB(37, 80, 216), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_p5_dim")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 607.500, 203.125)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 612.500, 198.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 627.500, 198.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 203.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 607.500, 203.125
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_neck_p4_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(177, 234, 164): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 607.500, 203.125)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 203.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 222.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 607.500, 222.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 607.500, 203.125
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_neck_p4_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(145, 202, 132): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(169, 226, 156)
    shp.Fill.BackColor.RGB = RGB(127, 184, 114)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 622.500, 203.125)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 627.500, 198.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 627.500, 217.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 222.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 203.125
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_neck_p4_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(97, 154, 84): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 632.500, 196.250, 32.500, 12.500, "P4", 10.000, RGB(28, 148, 84), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_p4")
    Set shp = AddFigureText(sld, 632.500, 207.500, 32.500, 10.000, "40×C", 8.000, RGB(28, 148, 84), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_p4_dim")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 607.500, 282.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 612.500, 277.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 627.500, 277.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 282.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 607.500, 282.500
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_neck_p3_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 150, 94): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 607.500, 282.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 282.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 301.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 607.500, 301.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 607.500, 282.500
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_neck_p3_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 118, 62): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(255, 142, 86)
    shp.Fill.BackColor.RGB = RGB(237, 100, 44)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 622.500, 282.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 627.500, 277.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 627.500, 296.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 301.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 622.500, 282.500
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_neck_p3_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(207, 70, 14): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 632.500, 275.000, 32.500, 12.500, "P3", 10.000, RGB(224, 58, 38), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_p3")
    Set shp = AddFigureText(sld, 632.500, 286.250, 32.500, 10.000, "80×C", 8.000, RGB(224, 58, 38), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_p3_dim")
    Set shp = AddFigureLine(sld, 363.750, 126.250, 362.500, 112.500, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_c5_up")
    Set shp = AddFigureLine(sld, 363.750, 151.250, 363.750, 162.500, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_up1_add")
    Set shp = AddFigureLine(sld, 350.000, 175.000, 340.000, 175.000, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_c4_add")
    Set shp = AddFigureLine(sld, 377.500, 175.000, 406.250, 175.000, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_add1_fusion")
    Set shp = AddFigureLine(sld, 363.750, 227.500, 363.750, 250.000, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_up2_add")
    Set shp = AddFigureLine(sld, 340.000, 263.125, 350.000, 263.125, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_c3_add")
    Set shp = AddFigureLine(sld, 377.500, 262.500, 406.250, 262.500, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_add2_fusion")
    Set shp = AddFigureLine(sld, 463.750, 175.000, 488.750, 175.000, RGB(55, 58, 68), 1.100, true, false, false, "SUMMER_L_neck_fusion1_add_seg1")
    Set shp = AddFigureLine(sld, 488.750, 175.000, 488.750, 113.750, RGB(55, 58, 68), 1.100, true, false, false, "SUMMER_L_neck_fusion1_add_seg2")
    Set shp = AddFigureLine(sld, 488.750, 113.750, 543.750, 113.750, RGB(55, 58, 68), 1.100, true, false, true, "SUMMER_L_neck_fusion1_add_seg3")
    Set shp = AddFigureLine(sld, 463.750, 262.500, 488.750, 262.500, RGB(55, 58, 68), 1.100, true, false, false, "SUMMER_L_neck_fusion2_p4_seg1")
    Set shp = AddFigureLine(sld, 488.750, 262.500, 488.750, 213.750, RGB(55, 58, 68), 1.100, true, false, false, "SUMMER_L_neck_fusion2_p4_seg2")
    Set shp = AddFigureLine(sld, 488.750, 213.750, 543.750, 213.750, RGB(55, 58, 68), 1.100, true, false, true, "SUMMER_L_neck_fusion2_p4_seg3")
    Set shp = AddFigureLine(sld, 435.000, 272.500, 435.000, 292.500, RGB(55, 58, 68), 1.100, true, false, false, "SUMMER_L_neck_fusion2_add_seg1")
    Set shp = AddFigureLine(sld, 435.000, 292.500, 543.750, 292.500, RGB(55, 58, 68), 1.100, true, false, true, "SUMMER_L_neck_fusion2_add_seg2")
    Set shp = AddFigureLine(sld, 571.250, 113.750, 607.500, 113.750, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_addp5_cube")
    Set shp = AddFigureLine(sld, 571.250, 213.750, 607.500, 213.750, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_addp4_cube")
    Set shp = AddFigureLine(sld, 571.250, 292.500, 607.500, 292.500, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_addp3_cube")
    Set shp = AddFigureLine(sld, 557.500, 126.250, 557.500, 138.750, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_addp5_down")
    Set shp = AddFigureLine(sld, 557.500, 164.375, 557.500, 180.000, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_down_losc")
    Set shp = AddFigureLine(sld, 557.500, 196.250, 557.500, 201.250, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_losc_addp4")
    Set shp = AddFigureLine(sld, 557.500, 226.250, 557.500, 240.000, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_addp4_down2")
    Set shp = AddFigureLine(sld, 557.500, 265.000, 557.500, 280.000, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_neck_down2_addp3")
    Set shp = AddFigureLine(sld, 435.000, 185.000, 435.000, 215.000, RGB(55, 58, 68), 1.100, true, false, false, "SUMMER_L_neck_fusion1_up2_seg1")
    Set shp = AddFigureLine(sld, 435.000, 215.000, 376.875, 215.000, RGB(55, 58, 68), 1.100, true, false, true, "SUMMER_L_neck_fusion1_up2_seg2")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 316.250, 317.500, 156.250, 17.500, RGB(224, 236, 252), RGB(37, 80, 216), 0.800, 0, "SUMMER_E_neck_note_cga")
    Set shp = AddFigureText(sld, 316.250, 320.000, 156.250, 12.500, "CGAFusion（高低频自适应融合）", 8.500, RGB(37, 80, 216), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_note_cga")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 483.750, 317.500, 156.250, 17.500, RGB(252, 230, 220), RGB(224, 58, 38), 0.800, 0, "SUMMER_E_neck_note_losc")
    Set shp = AddFigureText(sld, 483.750, 320.000, 156.250, 12.500, "LOSC（方向性下采样建模）", 8.500, RGB(224, 58, 38), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_neck_note_losc")
    Set shp = AddFigureText(sld, 672.500, 57.500, 156.875, 17.500, "Head: RT-DETR 原始检测头", 10.000, RGB(112, 62, 200), true, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_head_title")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 679.375, 90.000, 143.125, 75.000, RGB(255, 255, 255), RGB(205, 196, 235), 0.800, 0, "SUMMER_E_head_iou")
    Set shp = AddFigureText(sld, 679.375, 100.000, 143.125, 12.500, "IoU-Aware", 9.000, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_head_iou_line1")
    Set shp = AddFigureText(sld, 679.375, 113.125, 143.125, 12.500, "Query Selection", 9.000, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_head_iou")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 700.000, 136.250, 10.625, 10.625, RGB(198, 178, 239), RGB(112, 62, 200), 0.800, 0, "SUMMER_E_head_iou_sq_0")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 715.625, 136.250, 10.625, 10.625, RGB(198, 178, 239), RGB(112, 62, 200), 0.800, 0, "SUMMER_E_head_iou_sq_1")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 731.250, 136.250, 10.625, 10.625, RGB(198, 178, 239), RGB(112, 62, 200), 0.800, 0, "SUMMER_E_head_iou_sq_2")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 746.875, 136.250, 10.625, 10.625, RGB(198, 178, 239), RGB(112, 62, 200), 0.800, 0, "SUMMER_E_head_iou_sq_3")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 762.500, 136.250, 10.625, 10.625, RGB(198, 178, 239), RGB(112, 62, 200), 0.800, 0, "SUMMER_E_head_iou_sq_4")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 778.125, 136.250, 10.625, 10.625, RGB(198, 178, 239), RGB(112, 62, 200), 0.800, 0, "SUMMER_E_head_iou_sq_5")
    Set shp = AddFigureText(sld, 795.000, 135.000, 18.750, 13.125, "····", 8.000, RGB(120, 126, 140), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_head_iou_sq_dots")
    Set shp = AddFigureLine(sld, 753.750, 165.000, 753.750, 181.250, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_head_iou_dec")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 679.375, 181.250, 143.125, 83.125, RGB(255, 255, 255), RGB(205, 196, 235), 0.800, 0, "SUMMER_E_head_dec")
    Set shp = AddFigureText(sld, 679.375, 191.250, 143.125, 15.000, "Transformer Decoder × 6 层", 9.000, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_head_dec")
    Set shp = AddFigureLine(sld, 695.000, 221.250, 807.500, 221.250, RGB(112, 62, 200), 0.700, false, false, false, "SUMMER_L_head_dec_top_bus")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 688.750, 230.625, 12.500, 17.500, RGB(150, 120, 220), RGB(112, 62, 200), 0.800, 0, "SUMMER_E_head_dec_sq_0")
    Set shp = AddFigureLine(sld, 695.000, 221.250, 695.000, 230.625, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_head_dec_stem_top_0")
    Set shp = AddFigureLine(sld, 695.000, 248.125, 695.000, 256.875, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_head_dec_stem_bottom_0")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 717.500, 230.625, 12.500, 17.500, RGB(150, 120, 220), RGB(112, 62, 200), 0.800, 0, "SUMMER_E_head_dec_sq_1")
    Set shp = AddFigureLine(sld, 723.750, 221.250, 723.750, 230.625, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_head_dec_stem_top_1")
    Set shp = AddFigureLine(sld, 723.750, 248.125, 723.750, 256.875, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_head_dec_stem_bottom_1")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 746.250, 230.625, 12.500, 17.500, RGB(150, 120, 220), RGB(112, 62, 200), 0.800, 0, "SUMMER_E_head_dec_sq_2")
    Set shp = AddFigureLine(sld, 752.500, 221.250, 752.500, 230.625, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_head_dec_stem_top_2")
    Set shp = AddFigureLine(sld, 752.500, 248.125, 752.500, 256.875, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_head_dec_stem_bottom_2")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 801.250, 230.625, 12.500, 17.500, RGB(150, 120, 220), RGB(112, 62, 200), 0.800, 0, "SUMMER_E_head_dec_sq_3")
    Set shp = AddFigureLine(sld, 807.500, 221.250, 807.500, 230.625, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_head_dec_stem_top_3")
    Set shp = AddFigureLine(sld, 807.500, 248.125, 807.500, 256.875, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_head_dec_stem_bottom_3")
    Set shp = AddFigureLine(sld, 701.250, 239.375, 717.500, 239.375, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_head_dec_chain_0")
    Set shp = AddFigureLine(sld, 730.000, 239.375, 746.250, 239.375, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_head_dec_chain_1")
    Set shp = AddFigureLine(sld, 758.750, 239.375, 768.750, 239.375, RGB(112, 62, 200), 0.700, false, false, false, "SUMMER_L_head_dec_chain_2")
    Set shp = AddFigureLine(sld, 787.500, 239.375, 801.250, 239.375, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_head_dec_chain_3")
    Set shp = AddFigureText(sld, 767.500, 233.125, 21.250, 12.500, "••••", 8.000, RGB(40, 44, 60), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_head_dec_sq_dots")
    Set shp = AddFigureLine(sld, 753.750, 264.375, 753.750, 274.375, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_L_head_dec_out")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 679.375, 274.375, 143.125, 62.500, RGB(255, 255, 255), RGB(205, 196, 235), 0.800, 0, "SUMMER_E_head_out")
    Set shp = AddFigureText(sld, 686.250, 279.375, 129.375, 23.750, "输出：" & vbLf & "类别概率 + 归一化边界框坐标", 8.500, RGB(40, 44, 60), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_head_out")
    Set shp = AddFigureLine(sld, 706.250, 335.625, 738.750, 335.625, RGB(37, 80, 216), 1.000, false, false, false, "SUMMER_E_head_bars_axis_x")
    Set shp = AddFigureLine(sld, 706.250, 335.625, 706.250, 308.125, RGB(37, 80, 216), 1.000, false, false, false, "SUMMER_E_head_bars_axis_y")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 708.750, 315.625, 4.375, 20.000, RGB(90, 118, 216), -1, 1.000, 0, "SUMMER_E_head_bars_0")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 715.625, 324.375, 4.375, 11.250, RGB(90, 118, 216), -1, 1.000, 0, "SUMMER_E_head_bars_1")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 722.500, 319.375, 4.375, 16.250, RGB(90, 118, 216), -1, 1.000, 0, "SUMMER_E_head_bars_2")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 729.375, 309.375, 4.375, 26.250, RGB(90, 118, 216), -1, 1.000, 0, "SUMMER_E_head_bars_3")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 758.750, 306.250, 28.750, 26.250, -1, RGB(37, 80, 216), 1.200, 0, "SUMMER_E_head_bbox")
    Set shp = AddFigureText(sld, 837.500, 53.750, 116.250, 12.500, "图例说明", 10.500, RGB(40, 44, 60), true, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_legend_title")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 850.000, 76.875)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 853.125, 73.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 863.125, 73.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 76.875
    ff.AddNodes msoSegmentLine, msoEditingAuto, 850.000, 76.875
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_lg_p5_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 850.000, 76.875)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 76.875
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 86.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 850.000, 86.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 850.000, 76.875
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_lg_p5_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 860.000, 76.875)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 863.125, 73.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 863.125, 83.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 86.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 76.875
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_lg_p5_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 871.875, 72.500, 78.125, 15.000, "P5 / C5 （20×C）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_lg_p5")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 850.000, 99.375)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 853.125, 96.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 863.125, 96.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 99.375
    ff.AddNodes msoSegmentLine, msoEditingAuto, 850.000, 99.375
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_lg_p4_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(177, 234, 164): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 850.000, 99.375)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 99.375
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 108.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 850.000, 108.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 850.000, 99.375
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_lg_p4_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(145, 202, 132): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(169, 226, 156)
    shp.Fill.BackColor.RGB = RGB(127, 184, 114)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 860.000, 99.375)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 863.125, 96.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 863.125, 105.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 108.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 99.375
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_lg_p4_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(97, 154, 84): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 871.875, 95.000, 78.125, 15.000, "P4 / C4 （40×C）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_lg_p4")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 850.000, 121.875)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 853.125, 118.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 863.125, 118.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 121.875
    ff.AddNodes msoSegmentLine, msoEditingAuto, 850.000, 121.875
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_lg_p3_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 150, 94): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 850.000, 121.875)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 121.875
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 131.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 850.000, 131.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 850.000, 121.875
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_lg_p3_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 118, 62): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(255, 142, 86)
    shp.Fill.BackColor.RGB = RGB(237, 100, 44)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 860.000, 121.875)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 863.125, 118.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 863.125, 128.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 131.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.000, 121.875
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_lg_p3_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(207, 70, 14): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 871.875, 117.500, 78.125, 15.000, "P3 / C3 （80×C）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_lg_p3")
    Set shp = AddFigureShape(sld, msoShapeOval, 846.875, 146.875, 16.250, 16.250, RGB(209, 193, 246), RGB(105, 55, 190), 0.800, 0, "SUMMER_E_lg_up")
    Set shp = AddFigureText(sld, 846.875, 146.875, 16.250, 16.250, "UP" & vbLf & "↑", 5.500, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_lg_up_text")
    Set shp = AddFigureText(sld, 873.750, 148.750, 76.250, 15.000, "上采样 （× 2）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_lg_up")
    Set shp = AddFigureShape(sld, msoShapeOval, 846.875, 174.375, 16.250, 16.250, RGB(244, 172, 96), -1, 0.800, 0, "SUMMER_E_lg_down")
    Set shp = AddFigureText(sld, 846.875, 174.375, 16.250, 16.250, "Down" & vbLf & "↓", 4.000, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_lg_down_text")
    Set shp = AddFigureText(sld, 873.750, 175.000, 76.250, 15.000, "下采样 （× 2）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_lg_down")
    Set shp = AddFigureShape(sld, msoShapeOval, 846.875, 200.625, 16.250, 16.250, RGB(158, 226, 235), RGB(25, 150, 170), 0.800, 0, "SUMMER_E_lg_add")
    Set shp = AddFigureText(sld, 846.875, 200.625, 16.250, 16.250, "+", 7.000, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_lg_add_text")
    Set shp = AddFigureText(sld, 873.750, 201.250, 76.250, 15.000, "拼接 （Concat）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_lg_add")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 847.500, 227.500, 23.750, 11.250, RGB(253, 240, 232), RGB(226, 110, 50), 0.800, 0, "SUMMER_E_lg_losc")
    Set shp = AddFigureText(sld, 847.500, 228.125, 23.750, 10.000, "LOSC", 5.500, RGB(206, 90, 30), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_lg_losc_tag")
    Set shp = AddFigureText(sld, 875.000, 225.000, 75.000, 15.000, "方向性下采样建模", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_lg_losc")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 846.250, 252.500, 26.250, 11.250, RGB(235, 242, 255), RGB(37, 80, 216), 0.800, 0, "SUMMER_E_lg_cga")
    Set shp = AddFigureText(sld, 846.250, 253.125, 26.250, 10.000, "CGAFusion", 3.000, RGB(37, 80, 216), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_lg_cga_tag")
    Set shp = AddFigureText(sld, 875.000, 250.000, 75.000, 15.000, "高低频自适应融合", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_lg_cga")
    Set shp = AddFigureLine(sld, 848.750, 293.750, 868.750, 293.750, RGB(40, 44, 60), 1.400, false, false, true, "SUMMER_LD_lg_solid")
    Set shp = AddFigureText(sld, 875.000, 286.875, 75.000, 15.000, "数据流向", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_lg_solid")
    Set shp = AddFigureLine(sld, 848.750, 316.250, 868.750, 316.250, RGB(40, 44, 60), 1.400, true, false, true, "SUMMER_LD_lg_dash")
    Set shp = AddFigureText(sld, 875.000, 309.375, 75.000, 15.000, "跨层连接", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_lg_dash")
    Set shp = AddFigureLine(sld, 131.250, 185.000, 141.250, 185.000, RGB(37, 80, 216), 2.500, false, false, true, "SUMMER_L_input_backbone")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 8.125, 348.750, 296.875, 173.750, RGB(253, 253, 253), RGB(192, 203, 224), 0.800, 0, "SUMMER_E_panel_m1")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 310.625, 348.750, 256.250, 173.750, RGB(253, 253, 253), RGB(255, 105, 55), 0.800, 0, "SUMMER_E_panel_m2")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 572.500, 348.750, 378.750, 173.750, RGB(253, 253, 253), RGB(192, 203, 224), 0.800, 0, "SUMMER_E_panel_m3")
    Set shp = AddFigureText(sld, 7.500, 353.750, 296.250, 12.500, "模块一：CSP-MEEM（多尺度细粒度特征提取）", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_m1_title")
    Set shp = AddFigureText(sld, 59.375, 371.250, 56.250, 12.500, "输入图像", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m1_in")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 68.750, 385.625, 37.500, 25.625, RGB(228, 238, 253), RGB(37, 80, 216), 0.800, 0, "SUMMER_E_m1_csp")
    Set shp = AddFigureText(sld, 68.750, 388.750, 37.500, 18.750, "CSP" & vbLf & "分割", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m1_csp")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 25.000, 423.125, 28.125, 15.625, RGB(232, 245, 225), RGB(28, 148, 84), 0.800, 0, "SUMMER_E_m1_b0")
    Set shp = AddFigureText(sld, 25.000, 425.000, 28.125, 11.875, "分支1", 7.500, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m1_b0")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 72.500, 423.125, 28.125, 15.625, RGB(232, 245, 225), RGB(28, 148, 84), 0.800, 0, "SUMMER_E_m1_b1")
    Set shp = AddFigureText(sld, 72.500, 425.000, 28.125, 11.875, "分支2", 7.500, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m1_b1")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 117.500, 423.125, 28.125, 15.625, RGB(232, 245, 225), RGB(28, 148, 84), 0.800, 0, "SUMMER_E_m1_b2")
    Set shp = AddFigureText(sld, 117.500, 425.000, 28.125, 11.875, "分支n", 7.500, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m1_b2")
    Set shp = AddFigureText(sld, 100.625, 423.750, 16.875, 13.125, "···", 10.000, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m1_bdots")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 52.500, 448.750, 70.000, 15.625, RGB(226, 237, 252), RGB(37, 80, 216), 0.800, 0, "SUMMER_E_m1_fuse")
    Set shp = AddFigureText(sld, 52.500, 450.000, 70.000, 12.500, "融合", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m1_fuse")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 52.500, 473.750, 70.000, 16.250, RGB(255, 237, 215), RGB(230, 125, 35), 0.800, 0, "SUMMER_E_m1_ema")
    Set shp = AddFigureText(sld, 52.500, 475.625, 70.000, 12.500, "EMA 注意力", 8.000, RGB(200, 100, 30), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m1_ema")
    Set shp = AddFigureText(sld, 59.375, 496.875, 56.250, 12.500, "输出特征", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m1_out")
    Set shp = AddFigureLine(sld, 87.500, 383.750, 87.500, 385.625, RGB(40, 44, 60), 1.000, false, false, true, "SUMMER_L_m1_f1")
    Set shp = AddFigureLine(sld, 38.750, 413.750, 131.875, 413.750, RGB(40, 44, 60), 0.800, false, false, false, "SUMMER_L_m1_branch_bus")
    Set shp = AddFigureLine(sld, 38.750, 413.750, 38.750, 423.125, RGB(40, 44, 60), 0.800, false, false, true, "SUMMER_L_m1_branch_drop_0")
    Set shp = AddFigureLine(sld, 87.500, 413.750, 87.500, 423.125, RGB(40, 44, 60), 0.800, false, false, true, "SUMMER_L_m1_branch_drop_1")
    Set shp = AddFigureLine(sld, 131.875, 413.750, 131.875, 423.125, RGB(40, 44, 60), 0.800, false, false, true, "SUMMER_L_m1_branch_drop_2")
    Set shp = AddFigureLine(sld, 38.750, 438.750, 38.750, 456.250, RGB(40, 44, 60), 0.800, false, false, false, "SUMMER_L_m1_fanin_left_v")
    Set shp = AddFigureLine(sld, 38.750, 456.250, 52.500, 456.250, RGB(40, 44, 60), 0.800, false, false, true, "SUMMER_L_m1_fanin_left_h")
    Set shp = AddFigureLine(sld, 131.875, 438.750, 131.875, 456.250, RGB(40, 44, 60), 0.800, false, false, false, "SUMMER_L_m1_fanin_right_v")
    Set shp = AddFigureLine(sld, 131.875, 456.250, 122.500, 456.250, RGB(40, 44, 60), 0.800, false, false, true, "SUMMER_L_m1_fanin_right_h")
    Set shp = AddFigureLine(sld, 87.500, 411.250, 87.500, 413.750, RGB(40, 44, 60), 1.000, false, false, true, "SUMMER_L_m1_f2")
    Set shp = AddFigureLine(sld, 87.500, 438.750, 87.500, 448.750, RGB(40, 44, 60), 1.000, false, false, true, "SUMMER_L_m1_f3")
    Set shp = AddFigureLine(sld, 87.500, 464.375, 87.500, 473.750, RGB(40, 44, 60), 1.000, false, false, true, "SUMMER_L_m1_f4")
    Set shp = AddFigureLine(sld, 87.500, 490.000, 87.500, 497.500, RGB(40, 44, 60), 1.000, false, false, true, "SUMMER_L_m1_f5")
    Set shp = AddFigureText(sld, 162.500, 378.750, 153.750, 35.000, "•  多尺度并行感知，捕获不同" & vbLf & "大小的缺陷信息", 9.000, RGB(12, 12, 12), false, "left", "top", 0.500, "PingFang SC", "SUMMER_T_m1_p0")
    Set shp = AddFigureText(sld, 162.500, 420.000, 153.750, 35.000, "•  增强细节特征表达，保留" & vbLf & "边缘和纹理", 9.000, RGB(12, 12, 12), false, "left", "top", 0.500, "PingFang SC", "SUMMER_T_m1_p1")
    Set shp = AddFigureText(sld, 162.500, 461.250, 153.750, 35.000, "•  EMA 注意力增强关键特征，" & vbLf & "抑制冗余信息", 9.000, RGB(12, 12, 12), false, "left", "top", 0.500, "PingFang SC", "SUMMER_T_m1_p2")
    Set shp = AddFigureText(sld, 311.250, 353.750, 255.000, 12.500, "模块二：LOSC（方向性下采样建模）", 10.000, RGB(224, 58, 38), true, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_m2_title")
    Set shp = AddFigureText(sld, 325.000, 371.250, 75.000, 11.250, "输入特征", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m2_in")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 338.125, 385.625, 46.250, 30.000, RGB(214, 228, 252), RGB(120, 156, 228), 0.750, 0, "SUMMER_G_m2_in_bg")
    Set shp = AddFigureLine(sld, 349.688, 385.625, 349.688, 415.625, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m2_in_v1")
    Set shp = AddFigureLine(sld, 361.250, 385.625, 361.250, 415.625, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m2_in_v2")
    Set shp = AddFigureLine(sld, 372.812, 385.625, 372.812, 415.625, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m2_in_v3")
    Set shp = AddFigureLine(sld, 338.125, 395.625, 384.375, 395.625, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m2_in_h1")
    Set shp = AddFigureLine(sld, 338.125, 405.625, 384.375, 405.625, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m2_in_h2")
    Set shp = AddFigureShape(sld, msoShapeOval, 347.500, 433.750, 27.500, 27.500, RGB(244, 172, 96), -1, 0.800, 0, "SUMMER_E_m2_down")
    Set shp = AddFigureText(sld, 347.500, 433.750, 27.500, 27.500, "Down" & vbLf & "↓", 6.500, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_m2_down_text")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 343.750, 481.250, 36.250, 21.250, RGB(214, 228, 252), RGB(120, 156, 228), 0.750, 0, "SUMMER_G_m2_out_bg")
    Set shp = AddFigureLine(sld, 355.833, 481.250, 355.833, 502.500, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m2_out_v1")
    Set shp = AddFigureLine(sld, 367.917, 481.250, 367.917, 502.500, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m2_out_v2")
    Set shp = AddFigureLine(sld, 343.750, 491.875, 380.000, 491.875, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m2_out_h1")
    Set shp = AddFigureText(sld, 325.000, 503.750, 75.000, 11.250, "输出特征", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m2_out")
    Set shp = AddFigureLine(sld, 361.250, 415.625, 361.250, 433.750, RGB(40, 44, 60), 1.000, false, false, true, "SUMMER_L_m2_f1")
    Set shp = AddFigureLine(sld, 361.250, 461.250, 361.250, 481.250, RGB(40, 44, 60), 1.000, false, false, true, "SUMMER_L_m2_f2")
    Set shp = AddFigureText(sld, 421.875, 390.000, 147.500, 30.000, "•  引入方向性卷积核", 9.200, RGB(12, 12, 12), false, "left", "top", 0.500, "PingFang SC", "SUMMER_T_m2_p0")
    Set shp = AddFigureText(sld, 421.875, 415.000, 147.500, 30.000, "•  沿关键方向提取特征", 9.200, RGB(12, 12, 12), false, "left", "top", 0.500, "PingFang SC", "SUMMER_T_m2_p1")
    Set shp = AddFigureText(sld, 421.875, 440.000, 147.500, 30.000, "•  抑制无关干扰，保留方向信息", 9.200, RGB(12, 12, 12), false, "left", "top", 0.500, "PingFang SC", "SUMMER_T_m2_p2")
    Set shp = AddFigureText(sld, 421.875, 465.000, 147.500, 30.000, "•  增强对工业表面纹理与缺陷" & vbLf & "形态的建模能力", 9.200, RGB(12, 12, 12), false, "left", "top", 0.500, "PingFang SC", "SUMMER_T_m2_p3")
    Set shp = AddFigureText(sld, 572.500, 353.750, 381.250, 12.500, "模块三：CGAFusion（高低频自适应融合）", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_m3_title")
    Set shp = AddFigureText(sld, 581.250, 373.750, 81.250, 11.250, "输入特征 Fh（高频）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_m3_fh")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 591.250, 386.250, 35.000, 28.750, RGB(214, 228, 252), RGB(120, 156, 228), 0.750, 0, "SUMMER_G_m3_fh_bg")
    Set shp = AddFigureLine(sld, 600.000, 386.250, 600.000, 415.000, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m3_fh_v1")
    Set shp = AddFigureLine(sld, 608.750, 386.250, 608.750, 415.000, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m3_fh_v2")
    Set shp = AddFigureLine(sld, 617.500, 386.250, 617.500, 415.000, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m3_fh_v3")
    Set shp = AddFigureLine(sld, 591.250, 395.833, 626.250, 395.833, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m3_fh_h1")
    Set shp = AddFigureLine(sld, 591.250, 405.417, 626.250, 405.417, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m3_fh_h2")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 663.750, 387.500, 58.750, 25.000, RGB(235, 242, 255), RGB(37, 80, 216), 0.900, 0, "SUMMER_E_m3_high")
    Set shp = AddFigureLine(sld, 626.250, 400.000, 663.750, 400.000, RGB(40, 44, 60), 0.800, false, false, true, "SUMMER_L_m3_input_high")
    Set shp = AddFigureText(sld, 663.750, 390.000, 58.750, 20.000, "高频分支" & vbLf & "（细节信息）", 7.500, RGB(37, 80, 216), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m3_high")
    Set shp = AddFigureText(sld, 581.250, 426.250, 81.250, 11.250, "输入特征 Fl（低频）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_m3_fl")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 591.250, 438.750, 35.000, 28.750, RGB(224, 242, 216), RGB(112, 176, 92), 0.750, 0, "SUMMER_G_m3_fl_bg")
    Set shp = AddFigureLine(sld, 600.000, 438.750, 600.000, 467.500, RGB(112, 176, 92), 0.500, false, false, false, "SUMMER_G_m3_fl_v1")
    Set shp = AddFigureLine(sld, 608.750, 438.750, 608.750, 467.500, RGB(112, 176, 92), 0.500, false, false, false, "SUMMER_G_m3_fl_v2")
    Set shp = AddFigureLine(sld, 617.500, 438.750, 617.500, 467.500, RGB(112, 176, 92), 0.500, false, false, false, "SUMMER_G_m3_fl_v3")
    Set shp = AddFigureLine(sld, 591.250, 448.333, 626.250, 448.333, RGB(112, 176, 92), 0.500, false, false, false, "SUMMER_G_m3_fl_h1")
    Set shp = AddFigureLine(sld, 591.250, 457.917, 626.250, 457.917, RGB(112, 176, 92), 0.500, false, false, false, "SUMMER_G_m3_fl_h2")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 663.750, 440.000, 58.750, 25.000, RGB(235, 248, 238), RGB(28, 148, 84), 0.900, 0, "SUMMER_E_m3_low")
    Set shp = AddFigureLine(sld, 626.250, 452.500, 663.750, 452.500, RGB(40, 44, 60), 0.800, false, false, true, "SUMMER_L_m3_input_low")
    Set shp = AddFigureText(sld, 663.750, 442.500, 58.750, 20.000, "低频分支" & vbLf & "（结构信息）", 7.500, RGB(28, 148, 84), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m3_low")
    Set shp = AddFigureShape(sld, msoShapeOval, 746.250, 440.000, 25.000, 25.000, RGB(158, 226, 235), RGB(25, 150, 170), 0.800, 0, "SUMMER_E_m3_add")
    Set shp = AddFigureText(sld, 746.250, 440.000, 25.000, 25.000, "+", 6.500, RGB(0, 0, 0), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_m3_add_text")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 742.500, 481.250, 32.500, 25.000, RGB(214, 228, 252), RGB(120, 156, 228), 0.750, 0, "SUMMER_G_m3_out_bg")
    Set shp = AddFigureLine(sld, 753.333, 481.250, 753.333, 506.250, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m3_out_v1")
    Set shp = AddFigureLine(sld, 764.167, 481.250, 764.167, 506.250, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m3_out_v2")
    Set shp = AddFigureLine(sld, 742.500, 493.750, 775.000, 493.750, RGB(120, 156, 228), 0.500, false, false, false, "SUMMER_G_m3_out_h1")
    Set shp = AddFigureText(sld, 727.500, 507.500, 62.500, 11.250, "输出特征", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_m3_out")
    Set shp = AddFigureLine(sld, 722.500, 400.000, 758.750, 400.000, RGB(55, 58, 68), 1.100, false, false, false, "SUMMER_L_m3_f1_seg1")
    Set shp = AddFigureLine(sld, 758.750, 400.000, 758.750, 438.125, RGB(55, 58, 68), 1.100, false, false, true, "SUMMER_L_m3_f1_seg2")
    Set shp = AddFigureLine(sld, 722.500, 452.500, 744.375, 452.500, RGB(55, 58, 68), 1.100, false, false, true, "SUMMER_L_m3_f2_seg1")
    Set shp = AddFigureLine(sld, 758.750, 465.000, 758.750, 481.250, RGB(40, 44, 60), 1.000, false, false, true, "SUMMER_L_m3_f3")
    Set shp = AddFigureText(sld, 798.750, 388.750, 156.250, 27.500, "•  分离高频细节与低频结构信息", 9.000, RGB(12, 12, 12), false, "left", "top", 0.500, "PingFang SC", "SUMMER_T_m3_p0")
    Set shp = AddFigureText(sld, 798.750, 420.000, 156.250, 27.500, "•  空间联合注意力自适应融合", 9.000, RGB(12, 12, 12), false, "left", "top", 0.500, "PingFang SC", "SUMMER_T_m3_p1")
    Set shp = AddFigureText(sld, 798.750, 451.250, 156.250, 27.500, "•  突出关键细节，抑制噪声干扰", 9.000, RGB(12, 12, 12), false, "left", "top", 0.500, "PingFang SC", "SUMMER_T_m3_p2")
    Set shp = AddFigureText(sld, 798.750, 482.500, 156.250, 27.500, "•  提升特征表达与泛化能力", 9.000, RGB(12, 12, 12), false, "left", "top", 0.500, "PingFang SC", "SUMMER_T_m3_p3")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 7.500, 528.750, 191.250, 104.375, RGB(253, 253, 253), RGB(192, 203, 224), 0.800, 0, "SUMMER_E_panel_b1")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 218.750, 528.750, 330.625, 104.375, RGB(253, 253, 253), RGB(192, 203, 224), 0.800, 0, "SUMMER_E_panel_b2")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 567.500, 528.750, 141.250, 104.375, RGB(253, 253, 253), RGB(192, 203, 224), 0.800, 0, "SUMMER_E_panel_b3")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 718.750, 528.750, 232.500, 104.375, RGB(253, 253, 253), RGB(192, 203, 224), 0.800, 0, "SUMMER_E_panel_b4")
    Set shp = AddFigureText(sld, 7.500, 537.500, 201.250, 12.500, "输出多尺度特征金字塔", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_b1_title")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 29.375, 563.750)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 34.375, 558.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 55.625, 558.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 50.625, 563.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 29.375, 563.750
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_out_p3_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 150, 94): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 29.375, 563.750)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 50.625, 563.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 50.625, 583.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 29.375, 583.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 29.375, 563.750
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_out_p3_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 118, 62): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(255, 142, 86)
    shp.Fill.BackColor.RGB = RGB(237, 100, 44)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 50.625, 563.750)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 55.625, 558.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 55.625, 578.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 50.625, 583.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 50.625, 563.750
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_out_p3_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(207, 70, 14): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 21.250, 590.000, 46.250, 31.250, "P3" & vbLf & "（80×C）", 8.000, RGB(224, 58, 38), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_out_p3_label")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 86.250, 567.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 91.250, 562.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 111.250, 562.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 106.250, 567.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 86.250, 567.500
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_out_p4_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(177, 234, 164): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 86.250, 567.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 106.250, 567.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 106.250, 584.375
    ff.AddNodes msoSegmentLine, msoEditingAuto, 86.250, 584.375
    ff.AddNodes msoSegmentLine, msoEditingAuto, 86.250, 567.500
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_out_p4_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(145, 202, 132): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(169, 226, 156)
    shp.Fill.BackColor.RGB = RGB(127, 184, 114)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 106.250, 567.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 111.250, 562.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 111.250, 579.375
    ff.AddNodes msoSegmentLine, msoEditingAuto, 106.250, 584.375
    ff.AddNodes msoSegmentLine, msoEditingAuto, 106.250, 567.500
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_out_p4_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(97, 154, 84): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 77.500, 592.500, 43.750, 30.000, "P4" & vbLf & "（40×C）", 8.000, RGB(28, 148, 84), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_out_p4_label")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 142.500, 570.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 146.875, 566.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 164.375, 566.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 160.000, 570.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 142.500, 570.625
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_out_p5_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 142.500, 570.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 160.000, 570.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 160.000, 585.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 142.500, 585.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 142.500, 570.625
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_out_p5_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 160.000, 570.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 164.375, 566.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 164.375, 580.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 160.000, 585.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 160.000, 570.625
    Set shp = ff.ConvertToShape
    shp.Name = "SUMMER_E_out_p5_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 127.500, 595.000, 55.000, 27.500, "P5" & vbLf & "（20×C）", 8.000, RGB(37, 80, 216), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_out_p5_label")
    Set shp = AddFigureText(sld, 218.750, 535.000, 330.625, 15.000, "输入到 RT-DETR Head", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_b2_title")
    Set shp = AddFigureText(sld, 238.750, 552.500, 88.750, 26.250, "IoU-Aware" & vbLf & "Query Selection", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_b2_iou")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 240.000, 583.750, 10.000, 10.000, RGB(198, 178, 239), -1, 0.800, 0, "SUMMER_E_b2_sq1_0")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 254.375, 583.750, 10.000, 10.000, RGB(198, 178, 239), -1, 0.800, 0, "SUMMER_E_b2_sq1_1")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 268.750, 583.750, 10.000, 10.000, RGB(198, 178, 239), -1, 0.800, 0, "SUMMER_E_b2_sq1_2")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 283.125, 583.750, 10.000, 10.000, RGB(198, 178, 239), -1, 0.800, 0, "SUMMER_E_b2_sq1_3")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 297.500, 583.750, 10.000, 10.000, RGB(198, 178, 239), -1, 0.800, 0, "SUMMER_E_b2_sq1_4")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 311.875, 583.750, 10.000, 10.000, RGB(198, 178, 239), -1, 0.800, 0, "SUMMER_E_b2_sq1_5")
    Set shp = AddFigureText(sld, 327.500, 582.500, 18.750, 12.500, "····", 8.000, RGB(120, 126, 140), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_E_b2_sq1_dots")
    Set shp = AddFigureText(sld, 368.750, 546.875, 168.750, 13.750, "Transformer Decoder × 6 层", 8.500, RGB(37, 80, 216), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_b2_dec")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 371.250, 573.750, 13.750, 20.000, RGB(140, 110, 214), RGB(112, 62, 200), 0.800, 0, "SUMMER_E_b2_sq2_0")
    Set shp = AddFigureLine(sld, 378.125, 565.000, 378.125, 573.750, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_b2_stem_top_0")
    Set shp = AddFigureLine(sld, 378.125, 593.750, 378.125, 602.500, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_b2_stem_bottom_0")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 408.750, 573.750, 13.750, 20.000, RGB(140, 110, 214), RGB(112, 62, 200), 0.800, 0, "SUMMER_E_b2_sq2_1")
    Set shp = AddFigureLine(sld, 415.625, 565.000, 415.625, 573.750, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_b2_stem_top_1")
    Set shp = AddFigureLine(sld, 415.625, 593.750, 415.625, 602.500, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_b2_stem_bottom_1")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 446.250, 573.750, 13.750, 20.000, RGB(140, 110, 214), RGB(112, 62, 200), 0.800, 0, "SUMMER_E_b2_sq2_2")
    Set shp = AddFigureLine(sld, 453.125, 565.000, 453.125, 573.750, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_b2_stem_top_2")
    Set shp = AddFigureLine(sld, 453.125, 593.750, 453.125, 602.500, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_b2_stem_bottom_2")
    Set shp = AddFigureLine(sld, 385.000, 583.750, 408.750, 583.750, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_b2_chain_0")
    Set shp = AddFigureLine(sld, 422.500, 583.750, 446.250, 583.750, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_b2_chain_1")
    Set shp = AddFigureLine(sld, 460.000, 583.750, 467.500, 583.750, RGB(112, 62, 200), 0.700, false, false, false, "SUMMER_L_b2_chain_2")
    Set shp = AddFigureLine(sld, 500.000, 583.750, 515.000, 583.750, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_b2_chain_3")
    Set shp = AddFigureText(sld, 470.000, 576.875, 30.000, 15.000, "····", 9.000, RGB(120, 126, 140), true, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_b2_dots2")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 515.000, 573.750, 13.750, 20.000, RGB(140, 110, 214), -1, 1.000, 0, "SUMMER_E_b2_last")
    Set shp = AddFigureLine(sld, 521.875, 565.000, 521.875, 573.750, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_b2_last_top")
    Set shp = AddFigureLine(sld, 521.875, 593.750, 521.875, 602.500, RGB(112, 62, 200), 0.700, false, false, true, "SUMMER_L_b2_last_bottom")
    Set shp = AddFigureLine(sld, 328.750, 588.750, 368.750, 588.750, RGB(40, 44, 60), 1.200, false, false, true, "SUMMER_L_b2_flow")
    Set shp = AddFigureLine(sld, 198.750, 583.750, 218.750, 583.750, RGB(37, 80, 216), 2.500, false, false, true, "SUMMER_L_b1_b2")
    Set shp = AddFigureLine(sld, 549.375, 583.750, 567.500, 583.750, RGB(37, 80, 216), 2.500, false, false, true, "SUMMER_L_b2_b3")
    Set shp = AddFigureText(sld, 567.500, 537.500, 141.250, 12.500, "检测结果", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_b3_title")
    Set shp = AddFigureLine(sld, 587.500, 595.000, 620.000, 595.000, RGB(37, 80, 216), 1.000, false, false, false, "SUMMER_E_b3_bars_axis_x")
    Set shp = AddFigureLine(sld, 587.500, 595.000, 587.500, 561.250, RGB(37, 80, 216), 1.000, false, false, false, "SUMMER_E_b3_bars_axis_y")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 590.000, 579.375, 3.750, 15.625, RGB(70, 95, 205), -1, 1.000, 0, "SUMMER_E_b3_bars_0")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 595.625, 563.750, 3.750, 31.250, RGB(70, 95, 205), -1, 1.000, 0, "SUMMER_E_b3_bars_1")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 601.250, 583.750, 3.750, 11.250, RGB(70, 95, 205), -1, 1.000, 0, "SUMMER_E_b3_bars_2")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 606.875, 572.500, 3.750, 22.500, RGB(70, 95, 205), -1, 1.000, 0, "SUMMER_E_b3_bars_3")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 612.500, 587.500, 3.750, 7.500, RGB(70, 95, 205), -1, 1.000, 0, "SUMMER_E_b3_bars_4")
    Set shp = AddFigureText(sld, 572.500, 598.750, 62.500, 11.250, "类别概率", 7.500, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_b3_bars")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 648.750, 558.750, 23.750, 28.750, -1, RGB(224, 58, 38), 1.200, 0, "SUMMER_E_b3_box1")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 660.000, 568.750, 23.750, 26.250, -1, RGB(28, 148, 84), 1.200, 0, "SUMMER_E_b3_box2")
    Set shp = AddFigureText(sld, 630.000, 598.750, 77.500, 11.250, "归一化边界框坐标", 7.500, RGB(40, 44, 60), false, "center", "middle", 0.500, "PingFang SC", "SUMMER_T_b3_box")
    Set shp = AddFigureText(sld, 718.750, 531.875, 232.500, 17.500, "三模块协同优势", 11.000, RGB(0, 55, 210), true, "center", "middle", 0.000, "PingFang SC", "SUMMER_T_b4_title")
    Set shp = AddFigurePicture(sld, ResolveAssetPath(assetBase, "02_adv_precision.png", "outputs/reference-figures-regression/mphf-net-architecture/assets/02_adv_precision.png"), 725.000, 546.250, 26.250, 26.250, "SUMMER_E_b4_icon0")
    Set shp = AddFigureText(sld, 765.625, 552.500, 45.000, 17.500, "更精细：", 9.800, RGB(0, 55, 210), true, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_b4_head0")
    Set shp = AddFigureText(sld, 807.500, 552.500, 138.750, 17.500, "保留小目标细粒度特征", 10.500, RGB(18, 18, 18), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_b4_tail0")
    Set shp = AddFigurePicture(sld, ResolveAssetPath(assetBase, "03_adv_robust.png", "outputs/reference-figures-regression/mphf-net-architecture/assets/03_adv_robust.png"), 725.000, 571.250, 26.250, 26.250, "SUMMER_E_b4_icon1")
    Set shp = AddFigureText(sld, 765.625, 577.500, 45.000, 17.500, "更鲁棒：", 9.800, RGB(0, 125, 48), true, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_b4_head1")
    Set shp = AddFigureText(sld, 807.500, 577.500, 138.750, 17.500, "增强方向性缺陷感知", 10.500, RGB(18, 18, 18), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_b4_tail1")
    Set shp = AddFigurePicture(sld, ResolveAssetPath(assetBase, "04_adv_efficient.png", "outputs/reference-figures-regression/mphf-net-architecture/assets/04_adv_efficient.png"), 725.000, 596.250, 26.250, 26.250, "SUMMER_E_b4_icon2")
    Set shp = AddFigureText(sld, 765.625, 602.500, 45.000, 17.500, "更高效：", 9.800, RGB(255, 55, 30), true, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_b4_head2")
    Set shp = AddFigureText(sld, 807.500, 602.500, 138.750, 17.500, "自适应融合多尺度特征", 10.500, RGB(18, 18, 18), false, "left", "middle", 0.500, "PingFang SC", "SUMMER_T_b4_tail2")
    Set sld = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)
    Set shp = AddFigurePicture(sld, ResolveAssetPath(assetBase, "mphf-net-architecture.png", "assets/reference-figures/mphf-net-architecture.png"), 0.000, 0.000, 960.000, 640.000, "SUMMER_R_reference_full")
End Sub

Private Function AddFigureShape(sld As Slide, kind As MsoAutoShapeType, x As Single, y As Single, w As Single, h As Single, fillRgb As Long, lineRgb As Long, lineWeight As Single, fillTransparency As Integer, shapeName As String) As Shape
    Dim shp As Shape
    Set shp = sld.Shapes.AddShape(kind, x, y, w, h)
    shp.Name = shapeName
    shp.Shadow.Visible = msoFalse
    If kind = msoShapeRoundedRectangle Then shp.Adjustments(1) = 0.08
    If fillRgb < 0 Then shp.Fill.Visible = msoFalse Else shp.Fill.ForeColor.RGB = fillRgb: shp.Fill.Transparency = fillTransparency / 100#
    If lineRgb < 0 Then shp.Line.Visible = msoFalse Else shp.Line.ForeColor.RGB = lineRgb: shp.Line.Weight = lineWeight
    Set AddFigureShape = shp
End Function

Private Function AddFigureText(sld As Slide, x As Single, y As Single, w As Single, h As Single, txt As String, fontSize As Single, fontRgb As Long, isBold As Boolean, hAlign As String, vAlign As String, margin As Single, fontName As String, shapeName As String) As Shape
    Dim shp As Shape
    Set shp = sld.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    shp.Name = shapeName
    shp.Shadow.Visible = msoFalse
    With shp.TextFrame2
        .MarginLeft = margin: .MarginRight = margin: .MarginTop = margin: .MarginBottom = margin
        .WordWrap = msoTrue
        If vAlign = "top" Then .VerticalAnchor = msoAnchorTop Else If vAlign = "bottom" Then .VerticalAnchor = msoAnchorBottom Else .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = txt
        .TextRange.Font.Name = fontName
        .TextRange.Font.Size = fontSize
        .TextRange.Font.Bold = isBold
        .TextRange.Font.Fill.ForeColor.RGB = fontRgb
        If hAlign = "left" Then .TextRange.ParagraphFormat.Alignment = msoAlignLeft Else If hAlign = "right" Then .TextRange.ParagraphFormat.Alignment = msoAlignRight Else .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    Set AddFigureText = shp
End Function

Private Function AddFigureLine(sld As Slide, x1 As Single, y1 As Single, x2 As Single, y2 As Single, lineRgb As Long, lineWeight As Single, isDashed As Boolean, arrowBegin As Boolean, arrowEnd As Boolean, shapeName As String) As Shape
    Dim shp As Shape
    Set shp = sld.Shapes.AddLine(x1, y1, x2, y2)
    shp.Name = shapeName
    shp.Line.ForeColor.RGB = lineRgb: shp.Line.Weight = lineWeight
    If isDashed Then shp.Line.DashStyle = msoLineDash
    If arrowBegin Then shp.Line.BeginArrowheadStyle = msoArrowheadTriangle
    If arrowEnd Then shp.Line.EndArrowheadStyle = msoArrowheadTriangle
    Set AddFigureLine = shp
End Function

Private Function AddFigurePicture(sld As Slide, assetPath As String, x As Single, y As Single, w As Single, h As Single, shapeName As String) As Shape
    Dim shp As Shape
    Set shp = sld.Shapes.AddPicture(assetPath, msoFalse, msoTrue, x, y, w, h)
    shp.Name = shapeName: shp.LockAspectRatio = msoTrue
    Set AddFigurePicture = shp
End Function

Private Function ResolveAssetPath(baseFolder As String, assetFile As String, fallbackPath As String) As String
    Dim sep As String
    Dim candidate As String
    sep = Application.PathSeparator
    If InStr(baseFolder, "/") > 0 Then sep = "/"
    If Len(baseFolder) > 0 Then
        candidate = baseFolder & sep & "assets" & sep & assetFile
        If Len(Dir$(candidate)) > 0 Then ResolveAssetPath = candidate: Exit Function
    End If
    ResolveAssetPath = fallbackPath
End Function
