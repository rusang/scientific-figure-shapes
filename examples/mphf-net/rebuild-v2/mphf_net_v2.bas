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
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 6.250, 5.000, 947.500, 33.750, RGB(24, 33, 74), -1, 0.800, 0, "FIG_E_banner")
    Set shp = AddFigureText(sld, 6.250, 7.500, 947.500, 28.750, "MPHF-Net 工业表面缺陷检测网络结构设计图", 17.000, RGB(255, 255, 255), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_banner")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 7.500, 47.500, 125.000, 287.500, RGB(252, 252, 254), RGB(203, 210, 228), 1.200, 0, "FIG_E_panel_input")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 142.500, 47.500, 161.250, 287.500, RGB(252, 252, 254), RGB(203, 210, 228), 1.200, 0, "FIG_E_panel_backbone")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 311.250, 47.500, 358.750, 287.500, RGB(252, 252, 254), RGB(203, 210, 228), 1.200, 0, "FIG_E_panel_neck")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 678.750, 47.500, 150.000, 287.500, RGB(252, 252, 254), RGB(203, 210, 228), 1.200, 0, "FIG_E_panel_head")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 837.500, 47.500, 116.250, 287.500, RGB(252, 252, 254), RGB(203, 210, 228), 1.200, 0, "FIG_E_panel_legend")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 7.500, 347.500, 296.250, 177.500, RGB(252, 252, 254), RGB(203, 210, 228), 1.200, 0, "FIG_E_panel_m1")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 311.250, 347.500, 255.000, 177.500, RGB(252, 252, 254), RGB(203, 210, 228), 1.200, 0, "FIG_E_panel_m2")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 572.500, 347.500, 381.250, 177.500, RGB(252, 252, 254), RGB(203, 210, 228), 1.200, 0, "FIG_E_panel_m3")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 7.500, 532.500, 201.250, 102.500, RGB(252, 252, 254), RGB(203, 210, 228), 1.200, 0, "FIG_E_panel_b1")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 217.500, 532.500, 338.750, 102.500, RGB(252, 252, 254), RGB(203, 210, 228), 1.200, 0, "FIG_E_panel_b2")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 567.500, 532.500, 141.250, 102.500, RGB(252, 252, 254), RGB(203, 210, 228), 1.200, 0, "FIG_E_panel_b3")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 717.500, 532.500, 236.250, 102.500, RGB(252, 252, 254), RGB(203, 210, 228), 1.200, 0, "FIG_E_panel_b4")
    Set shp = AddFigureText(sld, 7.500, 53.750, 125.000, 13.750, "输入图像", 11.000, RGB(40, 44, 60), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_input_title")
    Set shp = AddFigurePicture(sld, ResolveAssetPath(assetBase, "01_R1_pcb.png", "/Users/shengshiliao/.skillshub/scientific-figure-shapes/examples/mphf-net/rebuild-v2/assets/01_R1_pcb.png"), 15.625, 101.250, 106.875, 152.500, "FIG_R_pcb")
    Set shp = AddFigureText(sld, 7.500, 268.750, 125.000, 15.000, "640 × 640 × 3", 11.000, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_input_size")
    Set shp = AddFigureText(sld, 142.500, 53.750, 161.250, 12.500, "Backbone：CSP-MEEM", 11.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_backbone_title")
    Set shp = AddFigureText(sld, 142.500, 67.500, 161.250, 10.000, "多尺度细粒度特征提取", 8.000, RGB(120, 126, 140), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_backbone_sub")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 187.500, 100.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 193.750, 93.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 215.000, 93.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 208.750, 100.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 187.500, 100.000
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_bb1_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 187.500, 100.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 208.750, 100.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 208.750, 118.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 187.500, 118.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 187.500, 100.000
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_bb1_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 208.750, 100.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 215.000, 93.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 215.000, 112.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 208.750, 118.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 208.750, 100.000
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_bb1_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 177.500, 147.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 185.000, 140.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 213.750, 140.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 206.250, 147.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 177.500, 147.500
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_bb2_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 177.500, 147.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 206.250, 147.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 206.250, 172.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 177.500, 172.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 177.500, 147.500
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_bb2_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 206.250, 147.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 213.750, 140.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 213.750, 165.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 206.250, 172.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 206.250, 147.500
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_bb2_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 166.250, 197.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 176.250, 187.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 215.000, 187.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 205.000, 197.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 166.250, 197.500
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_bb3_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 166.250, 197.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 205.000, 197.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 205.000, 228.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 166.250, 228.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 166.250, 197.500
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_bb3_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 205.000, 197.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 215.000, 187.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 215.000, 218.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 205.000, 228.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 205.000, 197.500
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_bb3_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 150.000, 256.250)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 163.750, 242.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 216.250, 242.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 202.500, 256.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 150.000, 256.250
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_bb4_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 150.000, 256.250)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 202.500, 256.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 202.500, 296.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 150.000, 296.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 150.000, 256.250
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_bb4_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 202.500, 256.250)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 216.250, 242.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 216.250, 282.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 202.500, 296.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 202.500, 256.250
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_bb4_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set shp = AddFigureLine(sld, 198.125, 118.750, 198.125, 140.000, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_bb_chain1")
    Set shp = AddFigureLine(sld, 191.875, 172.500, 191.875, 187.500, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_bb_chain2")
    Set shp = AddFigureLine(sld, 185.625, 228.750, 185.625, 242.500, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_bb_chain3")
    Set shp = AddFigureLine(sld, 215.000, 109.375, 232.500, 109.375, RGB(40, 44, 60), 1.200, false, false, false, "FIG_L_bb_bus_in1")
    Set shp = AddFigureLine(sld, 232.500, 109.375, 232.500, 258.750, RGB(40, 44, 60), 1.200, false, false, false, "FIG_L_bb_bus")
    Set shp = AddFigureLine(sld, 215.000, 213.125, 232.500, 213.125, RGB(40, 44, 60), 1.200, false, false, false, "FIG_L_bb_bus_in3")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 247.500, 128.125, 36.250, 27.500, RGB(245, 248, 255), RGB(37, 80, 216), 1.400, 0, "FIG_E_bb_c5")
    Set shp = AddFigureText(sld, 247.500, 130.625, 36.250, 12.500, "C5", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_bb_c5")
    Set shp = AddFigureText(sld, 247.500, 141.875, 36.250, 11.250, "20×C", 8.000, RGB(37, 80, 216), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_bb_c5_dim")
    Set shp = AddFigureLine(sld, 232.500, 141.875, 247.500, 141.875, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_bb_to_c5")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 247.500, 186.875, 36.250, 27.500, RGB(245, 248, 255), RGB(37, 80, 216), 1.400, 0, "FIG_E_bb_c4")
    Set shp = AddFigureText(sld, 247.500, 189.375, 36.250, 12.500, "C4", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_bb_c4")
    Set shp = AddFigureText(sld, 247.500, 200.625, 36.250, 11.250, "40×C", 8.000, RGB(37, 80, 216), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_bb_c4_dim")
    Set shp = AddFigureLine(sld, 232.500, 200.625, 247.500, 200.625, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_bb_to_c4")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 247.500, 245.000, 36.250, 27.500, RGB(245, 248, 255), RGB(224, 58, 38), 1.400, 0, "FIG_E_bb_c3")
    Set shp = AddFigureText(sld, 247.500, 247.500, 36.250, 12.500, "C3", 10.000, RGB(224, 58, 38), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_bb_c3")
    Set shp = AddFigureText(sld, 247.500, 258.750, 36.250, 11.250, "80×C", 8.000, RGB(224, 58, 38), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_bb_c3_dim")
    Set shp = AddFigureLine(sld, 216.250, 258.750, 247.500, 258.750, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_bb_to_c3")
    Set shp = AddFigureText(sld, 168.750, 296.250, 25.000, 10.000, "⋮", 11.000, RGB(120, 126, 140), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_backbone_dots")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 150.000, 311.250, 135.000, 18.750, RGB(235, 241, 255), RGB(203, 210, 228), 0.800, 0, "FIG_E_bb_outbar")
    Set shp = AddFigureText(sld, 150.000, 313.750, 135.000, 13.750, "输出多尺度特征", 9.000, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_bb_outbar")
    Set shp = AddFigureText(sld, 311.250, 53.750, 358.750, 12.500, "Neck：FPN-PAN （LOSC + CGAFusion）", 11.000, RGB(28, 148, 84), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_title")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 337.500, 72.500, 81.250, 16.250, RGB(226, 244, 230), RGB(28, 148, 84), 0.900, 0, "FIG_E_neck_td")
    Set shp = AddFigureText(sld, 337.500, 74.375, 81.250, 12.500, "Top-Down 路径", 8.500, RGB(28, 148, 84), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_td")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 517.500, 72.500, 81.250, 16.250, RGB(252, 236, 226), RGB(226, 120, 60), 0.900, 0, "FIG_E_neck_bu")
    Set shp = AddFigureText(sld, 517.500, 74.375, 81.250, 12.500, "Bottom-Up 路径", 8.500, RGB(206, 100, 40), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_bu")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 345.000, 95.000, 25.000, 17.500, RGB(245, 248, 255), RGB(37, 80, 216), 1.200, 0, "FIG_E_neck_c5")
    Set shp = AddFigureText(sld, 345.000, 97.500, 25.000, 12.500, "C5", 9.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_c5")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 317.500, 166.250, 25.000, 17.500, RGB(245, 248, 255), RGB(37, 80, 216), 1.200, 0, "FIG_E_neck_c4")
    Set shp = AddFigureText(sld, 317.500, 168.750, 25.000, 12.500, "C4", 9.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_c4")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 317.500, 255.000, 25.000, 17.500, RGB(245, 248, 255), RGB(37, 80, 216), 1.200, 0, "FIG_E_neck_c3")
    Set shp = AddFigureText(sld, 317.500, 257.500, 25.000, 12.500, "C3", 9.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_c3")
    Set shp = AddFigureShape(sld, msoShapeOval, 367.500, 130.000, 17.500, 17.500, RGB(222, 205, 248), RGB(140, 96, 216), 1.000, 0, "FIG_E_neck_up1")
    Set shp = AddFigureText(sld, 367.500, 130.000, 17.500, 17.500, "UP" & vbLf & "↑", 5.000, RGB(112, 62, 200), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_neck_up1_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 367.500, 206.250, 17.500, 17.500, RGB(222, 205, 248), RGB(140, 96, 216), 1.000, 0, "FIG_E_neck_up2")
    Set shp = AddFigureText(sld, 367.500, 206.250, 17.500, 17.500, "UP" & vbLf & "↑", 5.000, RGB(112, 62, 200), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_neck_up2_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 368.125, 166.875, 16.250, 16.250, RGB(52, 58, 78), RGB(30, 34, 48), 1.000, 0, "FIG_E_neck_add_td1")
    Set shp = AddFigureText(sld, 368.125, 166.875, 16.250, 16.250, "+", 9.000, RGB(255, 255, 255), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_neck_add_td1_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 368.125, 254.375, 16.250, 16.250, RGB(52, 58, 78), RGB(30, 34, 48), 1.000, 0, "FIG_E_neck_add_td2")
    Set shp = AddFigureText(sld, 368.125, 254.375, 16.250, 16.250, "+", 9.000, RGB(255, 255, 255), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_neck_add_td2_text")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 412.500, 165.625, 57.500, 18.750, RGB(240, 245, 255), RGB(37, 80, 216), 1.200, 0, "FIG_E_neck_fusion1")
    Set shp = AddFigureText(sld, 412.500, 168.750, 57.500, 12.500, "CGAFusion", 8.500, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_fusion1")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 412.500, 253.125, 57.500, 18.750, RGB(240, 245, 255), RGB(37, 80, 216), 1.200, 0, "FIG_E_neck_fusion2")
    Set shp = AddFigureText(sld, 412.500, 256.250, 57.500, 12.500, "CGAFusion", 8.500, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_fusion2")
    Set shp = AddFigureShape(sld, msoShapeOval, 549.375, 105.625, 16.250, 16.250, RGB(52, 58, 78), RGB(30, 34, 48), 1.000, 0, "FIG_E_neck_add_p5")
    Set shp = AddFigureText(sld, 549.375, 105.625, 16.250, 16.250, "+", 9.000, RGB(255, 255, 255), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_neck_add_p5_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 549.375, 205.625, 16.250, 16.250, RGB(52, 58, 78), RGB(30, 34, 48), 1.000, 0, "FIG_E_neck_add_p4")
    Set shp = AddFigureText(sld, 549.375, 205.625, 16.250, 16.250, "+", 9.000, RGB(255, 255, 255), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_neck_add_p4_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 549.375, 284.375, 16.250, 16.250, RGB(52, 58, 78), RGB(30, 34, 48), 1.000, 0, "FIG_E_neck_add_p3")
    Set shp = AddFigureText(sld, 549.375, 284.375, 16.250, 16.250, "+", 9.000, RGB(255, 255, 255), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_neck_add_p3_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 548.750, 142.500, 17.500, 17.500, RGB(250, 214, 170), RGB(206, 120, 40), 1.000, 0, "FIG_E_neck_down1")
    Set shp = AddFigureText(sld, 548.750, 142.500, 17.500, 17.500, "Down" & vbLf & "↓", 4.500, RGB(150, 74, 12), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_neck_down1_text")
    Set shp = AddFigureShape(sld, msoShapeOval, 548.750, 243.750, 17.500, 17.500, RGB(250, 214, 170), RGB(206, 120, 40), 1.000, 0, "FIG_E_neck_down2")
    Set shp = AddFigureText(sld, 548.750, 243.750, 17.500, 17.500, "Down" & vbLf & "↓", 4.500, RGB(150, 74, 12), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_neck_down2_text")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 535.000, 180.000, 45.000, 16.250, RGB(253, 240, 232), RGB(226, 110, 50), 1.200, 0, "FIG_E_neck_losc")
    Set shp = AddFigureText(sld, 535.000, 181.875, 45.000, 12.500, "LOSC", 9.000, RGB(206, 90, 30), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_losc")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 602.500, 103.750)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 607.500, 98.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 626.250, 98.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 103.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 602.500, 103.750
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_neck_p5_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 602.500, 103.750)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 103.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 120.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 602.500, 120.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 602.500, 103.750
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_neck_p5_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 621.250, 103.750)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 626.250, 98.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 626.250, 115.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 120.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 103.750
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_neck_p5_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 633.750, 96.250, 32.500, 12.500, "P5", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_p5")
    Set shp = AddFigureText(sld, 633.750, 107.500, 32.500, 10.000, "20×C", 8.000, RGB(37, 80, 216), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_p5_dim")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 602.500, 203.750)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 607.500, 198.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 626.250, 198.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 203.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 602.500, 203.750
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_neck_p4_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(177, 234, 164): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 602.500, 203.750)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 203.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 220.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 602.500, 220.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 602.500, 203.750
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_neck_p4_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(145, 202, 132): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(169, 226, 156)
    shp.Fill.BackColor.RGB = RGB(127, 184, 114)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 621.250, 203.750)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 626.250, 198.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 626.250, 215.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 220.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 203.750
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_neck_p4_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(97, 154, 84): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 633.750, 196.250, 32.500, 12.500, "P4", 10.000, RGB(28, 148, 84), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_p4")
    Set shp = AddFigureText(sld, 633.750, 207.500, 32.500, 10.000, "40×C", 8.000, RGB(28, 148, 84), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_p4_dim")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 602.500, 282.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 607.500, 277.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 626.250, 277.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 282.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 602.500, 282.500
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_neck_p3_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 150, 94): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 602.500, 282.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 282.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 298.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 602.500, 298.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 602.500, 282.500
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_neck_p3_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 118, 62): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(255, 142, 86)
    shp.Fill.BackColor.RGB = RGB(237, 100, 44)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 621.250, 282.500)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 626.250, 277.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 626.250, 293.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 298.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 621.250, 282.500
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_neck_p3_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(207, 70, 14): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 633.750, 275.000, 32.500, 12.500, "P3", 10.000, RGB(224, 58, 38), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_p3")
    Set shp = AddFigureText(sld, 633.750, 286.250, 32.500, 10.000, "80×C", 8.000, RGB(224, 58, 38), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_p3_dim")
    Set shp = AddFigureLine(sld, 357.500, 112.500, 373.125, 129.375, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_c5_up")
    Set shp = AddFigureLine(sld, 376.250, 147.500, 376.250, 166.875, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_up1_add")
    Set shp = AddFigureLine(sld, 342.500, 175.000, 368.125, 175.000, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_c4_add")
    Set shp = AddFigureLine(sld, 384.375, 175.000, 412.500, 175.000, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_add1_fusion")
    Set shp = AddFigureLine(sld, 376.250, 223.750, 376.250, 254.375, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_up2_add")
    Set shp = AddFigureLine(sld, 342.500, 263.125, 368.125, 263.125, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_c3_add")
    Set shp = AddFigureLine(sld, 384.375, 262.500, 412.500, 262.500, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_add2_fusion")
    Set shp = AddFigureLine(sld, 565.625, 113.750, 602.500, 113.750, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_addp5_cube")
    Set shp = AddFigureLine(sld, 565.625, 213.750, 602.500, 213.750, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_addp4_cube")
    Set shp = AddFigureLine(sld, 565.625, 292.500, 602.500, 292.500, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_addp3_cube")
    Set shp = AddFigureLine(sld, 557.500, 121.875, 557.500, 142.500, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_addp5_down")
    Set shp = AddFigureLine(sld, 557.500, 160.000, 557.500, 180.000, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_down_losc")
    Set shp = AddFigureLine(sld, 557.500, 196.250, 557.500, 205.625, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_losc_addp4")
    Set shp = AddFigureLine(sld, 557.500, 221.875, 557.500, 243.750, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_addp4_down2")
    Set shp = AddFigureLine(sld, 557.500, 261.250, 557.500, 284.375, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_neck_down2_addp3")
    Set shp = AddFigureLine(sld, 470.000, 175.000, 493.750, 175.000, RGB(120, 126, 140), 1.000, true, false, false, "FIG_LD_fusion1_addp5_1")
    Set shp = AddFigureLine(sld, 493.750, 175.000, 493.750, 113.750, RGB(120, 126, 140), 1.000, true, false, false, "FIG_LD_fusion1_addp5_2")
    Set shp = AddFigureLine(sld, 493.750, 113.750, 549.375, 113.750, RGB(120, 126, 140), 1.000, true, false, true, "FIG_LD_fusion1_addp5_3")
    Set shp = AddFigureLine(sld, 470.000, 262.500, 493.750, 262.500, RGB(120, 126, 140), 1.000, true, false, false, "FIG_LD_fusion2_addp3_1")
    Set shp = AddFigureLine(sld, 493.750, 262.500, 493.750, 292.500, RGB(120, 126, 140), 1.000, true, false, false, "FIG_LD_fusion2_addp3_2")
    Set shp = AddFigureLine(sld, 493.750, 292.500, 549.375, 292.500, RGB(120, 126, 140), 1.000, true, false, true, "FIG_LD_fusion2_addp3_3")
    Set shp = AddFigureLine(sld, 441.250, 184.375, 441.250, 253.125, RGB(120, 126, 140), 1.000, true, false, true, "FIG_LD_fusion1_fusion2_1")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 316.250, 317.500, 156.250, 17.500, RGB(235, 242, 255), RGB(37, 80, 216), 0.800, 0, "FIG_E_neck_note_cga")
    Set shp = AddFigureText(sld, 316.250, 320.000, 156.250, 12.500, "CGAFusion（高低频自适应融合）", 8.500, RGB(37, 80, 216), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_note_cga")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 483.750, 317.500, 156.250, 17.500, RGB(253, 238, 232), RGB(224, 58, 38), 0.800, 0, "FIG_E_neck_note_losc")
    Set shp = AddFigureText(sld, 483.750, 320.000, 156.250, 12.500, "LOSC（方向性下采样建模）", 8.500, RGB(224, 58, 38), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_neck_note_losc")
    Set shp = AddFigureText(sld, 678.750, 53.750, 150.000, 12.500, "Head：RT-DETR 原始检测头", 9.500, RGB(112, 62, 200), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_head_title")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 686.250, 88.750, 135.000, 60.000, RGB(255, 255, 255), RGB(203, 210, 228), 1.000, 0, "FIG_E_head_iou")
    Set shp = AddFigureText(sld, 686.250, 95.000, 135.000, 22.500, "IoU-Aware" & vbLf & "Query Selection", 9.000, RGB(40, 44, 60), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_head_iou")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 697.500, 126.250, 7.500, 7.500, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_head_iou_sq_0")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 708.125, 126.250, 7.500, 7.500, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_head_iou_sq_1")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 718.750, 126.250, 7.500, 7.500, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_head_iou_sq_2")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 729.375, 126.250, 7.500, 7.500, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_head_iou_sq_3")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 740.000, 126.250, 7.500, 7.500, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_head_iou_sq_4")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 750.625, 126.250, 7.500, 7.500, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_head_iou_sq_5")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 761.250, 126.250, 7.500, 7.500, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_head_iou_sq_6")
    Set shp = AddFigureText(sld, 773.125, 125.000, 18.750, 10.000, "····", 8.000, RGB(120, 126, 140), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_head_iou_sq_dots")
    Set shp = AddFigureLine(sld, 753.750, 148.750, 753.750, 182.500, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_head_iou_dec")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 686.250, 182.500, 135.000, 67.500, RGB(255, 255, 255), RGB(203, 210, 228), 1.000, 0, "FIG_E_head_dec")
    Set shp = AddFigureText(sld, 686.250, 187.500, 135.000, 13.750, "Transformer Decoder × 6 层", 8.500, RGB(40, 44, 60), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_head_dec")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 696.250, 212.500, 8.750, 8.750, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_head_dec_sq_0")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 709.375, 212.500, 8.750, 8.750, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_head_dec_sq_1")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 722.500, 212.500, 8.750, 8.750, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_head_dec_sq_2")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 735.625, 212.500, 8.750, 8.750, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_head_dec_sq_3")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 748.750, 212.500, 8.750, 8.750, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_head_dec_sq_4")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 761.875, 212.500, 8.750, 8.750, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_head_dec_sq_5")
    Set shp = AddFigureText(sld, 776.250, 211.250, 18.750, 11.250, "····", 8.000, RGB(120, 126, 140), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_head_dec_sq_dots")
    Set shp = AddFigureLine(sld, 753.750, 250.000, 753.750, 270.000, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_head_dec_out")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 686.250, 270.000, 135.000, 57.500, RGB(255, 255, 255), RGB(203, 210, 228), 1.000, 0, "FIG_E_head_out")
    Set shp = AddFigureText(sld, 686.250, 273.750, 135.000, 21.250, "输出：" & vbLf & "类别概率 + 归一化边界框坐标", 8.500, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_head_out")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 710.000, 307.500, 4.375, 11.250, RGB(90, 118, 216), -1, 0.800, 0, "FIG_E_head_bars_0")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 716.250, 300.000, 4.375, 18.750, RGB(90, 118, 216), -1, 0.800, 0, "FIG_E_head_bars_1")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 722.500, 311.250, 4.375, 7.500, RGB(90, 118, 216), -1, 0.800, 0, "FIG_E_head_bars_2")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 728.750, 303.750, 4.375, 15.000, RGB(90, 118, 216), -1, 0.800, 0, "FIG_E_head_bars_3")
    Set shp = AddFigureLine(sld, 708.125, 319.375, 736.250, 319.375, RGB(40, 44, 60), 0.800, false, false, false, "FIG_E_head_bars_axis")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 760.000, 298.750, 25.000, 20.000, -1, RGB(37, 80, 216), 1.200, 0, "FIG_E_head_bbox")
    Set shp = AddFigureText(sld, 837.500, 53.750, 116.250, 12.500, "图例说明", 10.500, RGB(40, 44, 60), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_legend_title")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 846.250, 85.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 849.375, 82.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.625, 82.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 85.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 846.250, 85.625
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_lg_p5_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 846.250, 85.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 85.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 93.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 846.250, 93.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 846.250, 85.625
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_lg_p5_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 857.500, 85.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.625, 82.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.625, 90.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 93.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 85.625
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_lg_p5_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 867.500, 81.250, 85.000, 13.750, "P5 / C5 （20×C）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_lg_p5")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 846.250, 108.125)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 849.375, 105.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.625, 105.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 108.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 846.250, 108.125
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_lg_p4_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(177, 234, 164): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 846.250, 108.125)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 108.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 116.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 846.250, 116.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 846.250, 108.125
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_lg_p4_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(145, 202, 132): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(169, 226, 156)
    shp.Fill.BackColor.RGB = RGB(127, 184, 114)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 857.500, 108.125)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.625, 105.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.625, 113.125
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 116.250
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 108.125
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_lg_p4_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(97, 154, 84): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 867.500, 103.750, 85.000, 13.750, "P4 / C4 （40×C）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_lg_p4")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 846.250, 130.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 849.375, 127.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.625, 127.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 130.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 846.250, 130.625
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_lg_p3_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 150, 94): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 846.250, 130.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 130.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 138.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 846.250, 138.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 846.250, 130.625
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_lg_p3_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 118, 62): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(255, 142, 86)
    shp.Fill.BackColor.RGB = RGB(237, 100, 44)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 857.500, 130.625)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.625, 127.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 860.625, 135.625
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 138.750
    ff.AddNodes msoSegmentLine, msoEditingAuto, 857.500, 130.625
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_lg_p3_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(207, 70, 14): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 867.500, 126.250, 85.000, 13.750, "P3 / C3 （80×C）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_lg_p3")
    Set shp = AddFigureShape(sld, msoShapeOval, 846.250, 150.000, 12.500, 12.500, RGB(222, 205, 248), RGB(140, 96, 216), 1.000, 0, "FIG_E_lg_up")
    Set shp = AddFigureText(sld, 846.250, 150.000, 12.500, 12.500, "UP", 5.000, RGB(112, 62, 200), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_lg_up_text")
    Set shp = AddFigureText(sld, 867.500, 148.750, 85.000, 13.750, "上采样 （× 2）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_lg_up")
    Set shp = AddFigureShape(sld, msoShapeOval, 846.250, 172.500, 12.500, 12.500, RGB(250, 214, 170), RGB(206, 120, 40), 1.000, 0, "FIG_E_lg_down")
    Set shp = AddFigureText(sld, 846.250, 172.500, 12.500, 12.500, "Down", 4.000, RGB(150, 74, 12), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_lg_down_text")
    Set shp = AddFigureText(sld, 867.500, 171.250, 85.000, 13.750, "下采样 （× 2）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_lg_down")
    Set shp = AddFigureShape(sld, msoShapeOval, 846.250, 195.000, 12.500, 12.500, RGB(52, 58, 78), RGB(30, 34, 48), 1.000, 0, "FIG_E_lg_add")
    Set shp = AddFigureText(sld, 846.250, 195.000, 12.500, 12.500, "+", 7.000, RGB(255, 255, 255), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_lg_add_text")
    Set shp = AddFigureText(sld, 867.500, 193.750, 85.000, 13.750, "拼接 （Concat）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_lg_add")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 845.000, 218.750, 20.000, 10.000, RGB(253, 240, 232), RGB(226, 110, 50), 0.800, 0, "FIG_E_lg_losc")
    Set shp = AddFigureText(sld, 845.000, 219.375, 20.000, 8.750, "LOSC", 5.500, RGB(206, 90, 30), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_lg_losc_tag")
    Set shp = AddFigureText(sld, 867.500, 216.875, 85.000, 13.750, "方向性下采样建模", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_lg_losc")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 843.750, 240.000, 27.500, 10.000, RGB(235, 242, 255), RGB(37, 80, 216), 0.800, 0, "FIG_E_lg_cga")
    Set shp = AddFigureText(sld, 843.750, 240.625, 27.500, 8.750, "CGAFusion", 4.500, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_lg_cga_tag")
    Set shp = AddFigureText(sld, 867.500, 238.125, 85.000, 13.750, "高低频自适应融合", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_lg_cga")
    Set shp = AddFigureLine(sld, 846.250, 267.500, 865.000, 267.500, RGB(40, 44, 60), 1.200, false, false, true, "FIG_LD_lg_solid")
    Set shp = AddFigureText(sld, 867.500, 260.625, 85.000, 13.750, "数据流向", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_lg_solid")
    Set shp = AddFigureLine(sld, 846.250, 288.750, 865.000, 288.750, RGB(40, 44, 60), 1.200, true, false, true, "FIG_LD_lg_dash")
    Set shp = AddFigureText(sld, 867.500, 281.875, 85.000, 13.750, "跨层连接", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_lg_dash")
    Set shp = AddFigureLine(sld, 132.500, 185.000, 142.500, 185.000, RGB(37, 80, 216), 4.000, false, false, true, "FIG_L_input_backbone")
    Set shp = AddFigureLine(sld, 303.750, 185.000, 311.250, 185.000, RGB(37, 80, 216), 4.000, false, false, true, "FIG_L_backbone_neck")
    Set shp = AddFigureLine(sld, 670.000, 185.000, 678.750, 185.000, RGB(37, 80, 216), 4.000, false, false, true, "FIG_L_neck_head")
    Set shp = AddFigureText(sld, 7.500, 353.750, 296.250, 12.500, "模块一：CSP-MEEM（多尺度细粒度特征提取）", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_m1_title")
    Set shp = AddFigureText(sld, 37.500, 372.500, 75.000, 12.500, "输入图像", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m1_in")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 52.500, 388.750, 45.000, 18.750, RGB(235, 242, 255), RGB(37, 80, 216), 0.900, 0, "FIG_E_m1_csp")
    Set shp = AddFigureText(sld, 52.500, 391.250, 45.000, 13.750, "CSP 分割", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m1_csp")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 17.500, 421.250, 32.500, 16.250, RGB(240, 246, 255), RGB(37, 80, 216), 0.800, 0, "FIG_E_m1_b0")
    Set shp = AddFigureText(sld, 17.500, 423.125, 32.500, 12.500, "分支1", 7.500, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m1_b0")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 58.750, 421.250, 32.500, 16.250, RGB(240, 246, 255), RGB(37, 80, 216), 0.800, 0, "FIG_E_m1_b1")
    Set shp = AddFigureText(sld, 58.750, 423.125, 32.500, 12.500, "分支2", 7.500, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m1_b1")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 100.000, 421.250, 32.500, 16.250, RGB(240, 246, 255), RGB(37, 80, 216), 0.800, 0, "FIG_E_m1_b2")
    Set shp = AddFigureText(sld, 100.000, 423.125, 32.500, 12.500, "分支n", 7.500, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m1_b2")
    Set shp = AddFigureText(sld, 97.500, 423.125, 18.750, 12.500, "···", 8.000, RGB(120, 126, 140), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m1_bdots")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 43.750, 447.500, 62.500, 16.250, RGB(235, 242, 255), RGB(37, 80, 216), 0.900, 0, "FIG_E_m1_fuse")
    Set shp = AddFigureText(sld, 43.750, 449.375, 62.500, 12.500, "融合", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m1_fuse")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 38.750, 471.250, 72.500, 16.250, RGB(255, 240, 226), RGB(226, 130, 50), 0.900, 0, "FIG_E_m1_ema")
    Set shp = AddFigureText(sld, 38.750, 473.125, 72.500, 12.500, "EMA 注意力", 8.000, RGB(200, 100, 30), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m1_ema")
    Set shp = AddFigureText(sld, 37.500, 493.750, 75.000, 12.500, "输出特征", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m1_out")
    Set shp = AddFigureLine(sld, 75.000, 385.000, 75.000, 388.750, RGB(40, 44, 60), 1.000, false, false, true, "FIG_L_m1_f1")
    Set shp = AddFigureLine(sld, 75.000, 407.500, 75.000, 421.250, RGB(40, 44, 60), 1.000, false, false, true, "FIG_L_m1_f2")
    Set shp = AddFigureLine(sld, 75.000, 437.500, 75.000, 447.500, RGB(40, 44, 60), 1.000, false, false, true, "FIG_L_m1_f3")
    Set shp = AddFigureLine(sld, 75.000, 463.750, 75.000, 471.250, RGB(40, 44, 60), 1.000, false, false, true, "FIG_L_m1_f4")
    Set shp = AddFigureLine(sld, 75.000, 487.500, 75.000, 493.750, RGB(40, 44, 60), 1.000, false, false, true, "FIG_L_m1_f5")
    Set shp = AddFigureText(sld, 141.250, 378.750, 153.750, 35.000, "•  多尺度并行感知，捕获不同大小的缺陷信息", 8.500, RGB(40, 44, 60), false, "left", "top", 0.000, "PingFang SC", "FIG_T_m1_p0")
    Set shp = AddFigureText(sld, 141.250, 420.000, 153.750, 35.000, "•  增强细节特征表达，保留边缘和纹理", 8.500, RGB(40, 44, 60), false, "left", "top", 0.000, "PingFang SC", "FIG_T_m1_p1")
    Set shp = AddFigureText(sld, 141.250, 461.250, 153.750, 35.000, "•  EMA 注意力增强关键特征，抑制冗余信息", 8.500, RGB(40, 44, 60), false, "left", "top", 0.000, "PingFang SC", "FIG_T_m1_p2")
    Set shp = AddFigureText(sld, 311.250, 353.750, 255.000, 12.500, "模块二：LOSC（方向性下采样建模）", 10.000, RGB(224, 58, 38), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_m2_title")
    Set shp = AddFigureText(sld, 325.000, 371.250, 75.000, 11.250, "输入特征", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m2_in")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 337.500, 383.750, 47.500, 35.000, RGB(214, 228, 252), RGB(120, 156, 228), 0.750, 0, "FIG_G_m2_in_bg")
    Set shp = AddFigureLine(sld, 349.375, 383.750, 349.375, 418.750, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m2_in_v1")
    Set shp = AddFigureLine(sld, 361.250, 383.750, 361.250, 418.750, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m2_in_v2")
    Set shp = AddFigureLine(sld, 373.125, 383.750, 373.125, 418.750, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m2_in_v3")
    Set shp = AddFigureLine(sld, 337.500, 395.417, 385.000, 395.417, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m2_in_h1")
    Set shp = AddFigureLine(sld, 337.500, 407.083, 385.000, 407.083, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m2_in_h2")
    Set shp = AddFigureShape(sld, msoShapeOval, 351.875, 431.875, 18.750, 18.750, RGB(250, 214, 170), RGB(206, 120, 40), 1.000, 0, "FIG_E_m2_down")
    Set shp = AddFigureText(sld, 351.875, 431.875, 18.750, 18.750, "Down" & vbLf & "↓", 4.500, RGB(150, 74, 12), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_m2_down_text")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 345.000, 467.500, 32.500, 25.000, RGB(214, 228, 252), RGB(120, 156, 228), 0.750, 0, "FIG_G_m2_out_bg")
    Set shp = AddFigureLine(sld, 355.833, 467.500, 355.833, 492.500, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m2_out_v1")
    Set shp = AddFigureLine(sld, 366.667, 467.500, 366.667, 492.500, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m2_out_v2")
    Set shp = AddFigureLine(sld, 345.000, 480.000, 377.500, 480.000, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m2_out_h1")
    Set shp = AddFigureText(sld, 325.000, 496.250, 75.000, 11.250, "输出特征", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m2_out")
    Set shp = AddFigureLine(sld, 361.250, 418.750, 361.250, 431.250, RGB(40, 44, 60), 1.000, false, false, true, "FIG_L_m2_f1")
    Set shp = AddFigureLine(sld, 361.250, 451.250, 361.250, 467.500, RGB(40, 44, 60), 1.000, false, false, true, "FIG_L_m2_f2")
    Set shp = AddFigureText(sld, 412.500, 375.000, 147.500, 30.000, "•  引入方向性卷积核", 8.500, RGB(40, 44, 60), false, "left", "top", 0.000, "PingFang SC", "FIG_T_m2_p0")
    Set shp = AddFigureText(sld, 412.500, 407.500, 147.500, 30.000, "•  沿关键方向提取特征", 8.500, RGB(40, 44, 60), false, "left", "top", 0.000, "PingFang SC", "FIG_T_m2_p1")
    Set shp = AddFigureText(sld, 412.500, 440.000, 147.500, 30.000, "•  抑制无关干扰，保留方向信息", 8.500, RGB(40, 44, 60), false, "left", "top", 0.000, "PingFang SC", "FIG_T_m2_p2")
    Set shp = AddFigureText(sld, 412.500, 472.500, 147.500, 30.000, "•  增强对工业表面纹理与缺陷形态的建模能力", 8.500, RGB(40, 44, 60), false, "left", "top", 0.000, "PingFang SC", "FIG_T_m2_p3")
    Set shp = AddFigureText(sld, 572.500, 353.750, 381.250, 12.500, "模块三：CGAFusion（高低频自适应融合）", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_m3_title")
    Set shp = AddFigureText(sld, 581.250, 373.750, 81.250, 11.250, "输入特征 Fh（高频）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_m3_fh")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 591.250, 386.250, 40.000, 28.750, RGB(214, 228, 252), RGB(120, 156, 228), 0.750, 0, "FIG_G_m3_fh_bg")
    Set shp = AddFigureLine(sld, 601.250, 386.250, 601.250, 415.000, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m3_fh_v1")
    Set shp = AddFigureLine(sld, 611.250, 386.250, 611.250, 415.000, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m3_fh_v2")
    Set shp = AddFigureLine(sld, 621.250, 386.250, 621.250, 415.000, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m3_fh_v3")
    Set shp = AddFigureLine(sld, 591.250, 395.833, 631.250, 395.833, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m3_fh_h1")
    Set shp = AddFigureLine(sld, 591.250, 405.417, 631.250, 405.417, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m3_fh_h2")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 653.750, 387.500, 60.000, 25.000, RGB(235, 242, 255), RGB(37, 80, 216), 0.900, 0, "FIG_E_m3_high")
    Set shp = AddFigureText(sld, 653.750, 390.000, 60.000, 20.000, "高频分支" & vbLf & "（细节信息）", 7.500, RGB(37, 80, 216), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m3_high")
    Set shp = AddFigureText(sld, 581.250, 426.250, 81.250, 11.250, "输入特征 Fl（低频）", 7.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_m3_fl")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 591.250, 438.750, 40.000, 28.750, RGB(224, 242, 216), RGB(112, 176, 92), 0.750, 0, "FIG_G_m3_fl_bg")
    Set shp = AddFigureLine(sld, 601.250, 438.750, 601.250, 467.500, RGB(112, 176, 92), 0.500, false, false, false, "FIG_G_m3_fl_v1")
    Set shp = AddFigureLine(sld, 611.250, 438.750, 611.250, 467.500, RGB(112, 176, 92), 0.500, false, false, false, "FIG_G_m3_fl_v2")
    Set shp = AddFigureLine(sld, 621.250, 438.750, 621.250, 467.500, RGB(112, 176, 92), 0.500, false, false, false, "FIG_G_m3_fl_v3")
    Set shp = AddFigureLine(sld, 591.250, 448.333, 631.250, 448.333, RGB(112, 176, 92), 0.500, false, false, false, "FIG_G_m3_fl_h1")
    Set shp = AddFigureLine(sld, 591.250, 457.917, 631.250, 457.917, RGB(112, 176, 92), 0.500, false, false, false, "FIG_G_m3_fl_h2")
    Set shp = AddFigureShape(sld, msoShapeRoundedRectangle, 653.750, 440.000, 60.000, 25.000, RGB(235, 248, 238), RGB(28, 148, 84), 0.900, 0, "FIG_E_m3_low")
    Set shp = AddFigureText(sld, 653.750, 442.500, 60.000, 20.000, "低频分支" & vbLf & "（结构信息）", 7.500, RGB(28, 148, 84), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m3_low")
    Set shp = AddFigureShape(sld, msoShapeOval, 735.000, 417.500, 17.500, 17.500, RGB(52, 58, 78), RGB(30, 34, 48), 1.000, 0, "FIG_E_m3_add")
    Set shp = AddFigureText(sld, 735.000, 417.500, 17.500, 17.500, "+", 9.000, RGB(255, 255, 255), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_m3_add_text")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 728.750, 456.250, 32.500, 25.000, RGB(214, 228, 252), RGB(120, 156, 228), 0.750, 0, "FIG_G_m3_out_bg")
    Set shp = AddFigureLine(sld, 739.583, 456.250, 739.583, 481.250, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m3_out_v1")
    Set shp = AddFigureLine(sld, 750.417, 456.250, 750.417, 481.250, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m3_out_v2")
    Set shp = AddFigureLine(sld, 728.750, 468.750, 761.250, 468.750, RGB(120, 156, 228), 0.500, false, false, false, "FIG_G_m3_out_h1")
    Set shp = AddFigureText(sld, 712.500, 485.000, 68.750, 11.250, "输出特征", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_m3_out")
    Set shp = AddFigureLine(sld, 713.750, 400.000, 735.000, 400.000, RGB(40, 44, 60), 1.000, false, false, false, "FIG_L_m3_f1_1")
    Set shp = AddFigureLine(sld, 735.000, 400.000, 735.000, 420.000, RGB(40, 44, 60), 1.000, false, false, true, "FIG_L_m3_f1_2")
    Set shp = AddFigureLine(sld, 713.750, 452.500, 735.000, 431.250, RGB(40, 44, 60), 1.000, false, false, true, "FIG_L_m3_f2")
    Set shp = AddFigureLine(sld, 743.750, 435.000, 743.750, 456.250, RGB(40, 44, 60), 1.000, false, false, true, "FIG_L_m3_f3")
    Set shp = AddFigureText(sld, 792.500, 375.000, 156.250, 27.500, "•  分离高频细节与低频结构信息", 8.500, RGB(40, 44, 60), false, "left", "top", 0.000, "PingFang SC", "FIG_T_m3_p0")
    Set shp = AddFigureText(sld, 792.500, 407.500, 156.250, 27.500, "•  空间联合注意力自适应融合", 8.500, RGB(40, 44, 60), false, "left", "top", 0.000, "PingFang SC", "FIG_T_m3_p1")
    Set shp = AddFigureText(sld, 792.500, 440.000, 156.250, 27.500, "•  突出关键细节，抑制噪声干扰", 8.500, RGB(40, 44, 60), false, "left", "top", 0.000, "PingFang SC", "FIG_T_m3_p2")
    Set shp = AddFigureText(sld, 792.500, 472.500, 156.250, 27.500, "•  提升特征表达与泛化能力", 8.500, RGB(40, 44, 60), false, "left", "top", 0.000, "PingFang SC", "FIG_T_m3_p3")
    Set shp = AddFigureText(sld, 7.500, 537.500, 201.250, 12.500, "输出多尺度特征金字塔", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_b1_title")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 28.750, 565.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 33.750, 560.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 68.750, 560.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 63.750, 565.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 28.750, 565.000
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_out_p3_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 150, 94): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 28.750, 565.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 63.750, 565.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 63.750, 592.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 28.750, 592.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 28.750, 565.000
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_out_p3_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(255, 118, 62): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(255, 142, 86)
    shp.Fill.BackColor.RGB = RGB(237, 100, 44)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 63.750, 565.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 68.750, 560.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 68.750, 587.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 63.750, 592.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 63.750, 565.000
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_out_p3_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(207, 70, 14): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(185, 48, 0): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 21.250, 601.250, 56.250, 20.000, "P3" & vbLf & "（80×C）", 8.000, RGB(224, 58, 38), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_out_p3_label")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 88.750, 570.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 93.750, 565.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 122.500, 565.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 117.500, 570.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 88.750, 570.000
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_out_p4_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(177, 234, 164): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 88.750, 570.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 117.500, 570.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 117.500, 592.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 88.750, 592.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 88.750, 570.000
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_out_p4_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(145, 202, 132): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(169, 226, 156)
    shp.Fill.BackColor.RGB = RGB(127, 184, 114)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 117.500, 570.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 122.500, 565.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 122.500, 587.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 117.500, 592.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 117.500, 570.000
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_out_p4_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(97, 154, 84): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(75, 132, 62): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 81.250, 601.250, 50.000, 20.000, "P4" & vbLf & "（40×C）", 8.000, RGB(28, 148, 84), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_out_p4_label")
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 141.250, 575.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 146.250, 570.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 170.000, 570.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 165.000, 575.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 141.250, 575.000
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_out_p5_top"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(174, 212, 255): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 141.250, 575.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 165.000, 575.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 165.000, 592.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 141.250, 592.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 141.250, 575.000
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_out_p5_front"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(142, 180, 238): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    shp.Fill.TwoColorGradient msoGradientVertical, 1
    shp.Fill.ForeColor.RGB = RGB(166, 204, 255)
    shp.Fill.BackColor.RGB = RGB(124, 162, 220)
    Set ff = sld.Shapes.BuildFreeform(msoEditingAuto, 165.000, 575.000)
    ff.AddNodes msoSegmentLine, msoEditingAuto, 170.000, 570.000
    ff.AddNodes msoSegmentLine, msoEditingAuto, 170.000, 587.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 165.000, 592.500
    ff.AddNodes msoSegmentLine, msoEditingAuto, 165.000, 575.000
    Set shp = ff.ConvertToShape
    shp.Name = "FIG_E_out_p5_right"
    shp.Shadow.Visible = msoFalse
    shp.Fill.ForeColor.RGB = RGB(94, 132, 190): shp.Fill.Transparency = 0 / 100#
    shp.Line.ForeColor.RGB = RGB(72, 110, 168): shp.Line.Weight = 0.800
    Set shp = AddFigureText(sld, 133.750, 601.250, 45.000, 20.000, "P5" & vbLf & "（20×C）", 8.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_out_p5_label")
    Set shp = AddFigureText(sld, 217.500, 537.500, 338.750, 12.500, "输入到 RT-DETR Head", 10.000, RGB(40, 44, 60), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_b2_title")
    Set shp = AddFigureText(sld, 237.500, 557.500, 81.250, 21.250, "IoU-Aware" & vbLf & "Query Selection", 8.000, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_b2_iou")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 241.250, 583.750, 7.500, 7.500, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_b2_sq1_0")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 251.875, 583.750, 7.500, 7.500, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_b2_sq1_1")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 262.500, 583.750, 7.500, 7.500, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_b2_sq1_2")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 273.125, 583.750, 7.500, 7.500, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_b2_sq1_3")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 283.750, 583.750, 7.500, 7.500, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_b2_sq1_4")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 294.375, 583.750, 7.500, 7.500, RGB(150, 120, 220), RGB(110, 82, 178), 0.600, 0, "FIG_E_b2_sq1_5")
    Set shp = AddFigureText(sld, 306.250, 582.500, 18.750, 10.000, "····", 8.000, RGB(120, 126, 140), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_b2_sq1_dots")
    Set shp = AddFigureText(sld, 350.000, 550.000, 150.000, 11.250, "Transformer Decoder × 6 层", 8.500, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_b2_dec")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 365.000, 576.250, 10.000, 10.000, RGB(140, 110, 214), RGB(110, 82, 178), 0.600, 0, "FIG_E_b2_sq2_0")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 380.625, 576.250, 10.000, 10.000, RGB(140, 110, 214), RGB(110, 82, 178), 0.600, 0, "FIG_E_b2_sq2_1")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 396.250, 576.250, 10.000, 10.000, RGB(140, 110, 214), RGB(110, 82, 178), 0.600, 0, "FIG_E_b2_sq2_2")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 411.875, 576.250, 10.000, 10.000, RGB(140, 110, 214), RGB(110, 82, 178), 0.600, 0, "FIG_E_b2_sq2_3")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 427.500, 576.250, 10.000, 10.000, RGB(140, 110, 214), RGB(110, 82, 178), 0.600, 0, "FIG_E_b2_sq2_4")
    Set shp = AddFigureText(sld, 444.375, 575.000, 18.750, 12.500, "····", 8.000, RGB(120, 126, 140), true, "center", "middle", 0.000, "PingFang SC", "FIG_E_b2_sq2_dots")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 495.000, 576.250, 10.000, 10.000, RGB(140, 110, 214), RGB(110, 82, 178), 0.600, 0, "FIG_E_b2_last")
    Set shp = AddFigureLine(sld, 322.500, 581.250, 365.000, 581.250, RGB(40, 44, 60), 1.200, false, false, true, "FIG_L_b2_flow")
    Set shp = AddFigureLine(sld, 208.750, 583.750, 217.500, 583.750, RGB(37, 80, 216), 4.000, false, false, true, "FIG_L_b1_b2")
    Set shp = AddFigureLine(sld, 556.250, 583.750, 567.500, 583.750, RGB(37, 80, 216), 4.000, false, false, true, "FIG_L_b2_b3")
    Set shp = AddFigureText(sld, 567.500, 537.500, 141.250, 12.500, "检测结果", 10.000, RGB(40, 44, 60), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_b3_title")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 585.000, 568.750, 4.375, 11.250, RGB(90, 118, 216), -1, 0.800, 0, "FIG_E_b3_bars_0")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 591.250, 561.250, 4.375, 18.750, RGB(90, 118, 216), -1, 0.800, 0, "FIG_E_b3_bars_1")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 597.500, 572.500, 4.375, 7.500, RGB(90, 118, 216), -1, 0.800, 0, "FIG_E_b3_bars_2")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 603.750, 565.000, 4.375, 15.000, RGB(90, 118, 216), -1, 0.800, 0, "FIG_E_b3_bars_3")
    Set shp = AddFigureLine(sld, 583.125, 580.625, 611.250, 580.625, RGB(40, 44, 60), 0.800, false, false, false, "FIG_E_b3_bars_axis")
    Set shp = AddFigureText(sld, 572.500, 591.250, 62.500, 11.250, "类别概率", 7.500, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_b3_bars")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 643.750, 557.500, 28.750, 22.500, -1, RGB(224, 58, 38), 1.200, 0, "FIG_E_b3_box1")
    Set shp = AddFigureShape(sld, msoShapeRectangle, 655.000, 567.500, 28.750, 22.500, -1, RGB(28, 148, 84), 1.200, 0, "FIG_E_b3_box2")
    Set shp = AddFigureText(sld, 631.250, 591.250, 75.000, 18.750, "归一化边界" & vbLf & "框坐标", 7.500, RGB(40, 44, 60), false, "center", "middle", 0.000, "PingFang SC", "FIG_T_b3_box")
    Set shp = AddFigureText(sld, 717.500, 537.500, 236.250, 12.500, "三模块协同优势", 10.000, RGB(37, 80, 216), true, "center", "middle", 0.000, "PingFang SC", "FIG_T_b4_title")
    Set shp = AddFigureShape(sld, msoShapeOval, 731.250, 555.000, 12.500, 12.500, RGB(37, 80, 216), -1, 0.800, 0, "FIG_E_b4_icon0")
    Set shp = AddFigureText(sld, 750.000, 555.000, 43.750, 13.750, "更精细：", 8.500, RGB(37, 80, 216), true, "left", "middle", 0.000, "PingFang SC", "FIG_T_b4_head0")
    Set shp = AddFigureText(sld, 792.500, 555.000, 156.250, 13.750, "保留小目标细粒度特征", 8.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_b4_tail0")
    Set shp = AddFigureShape(sld, msoShapeOval, 731.250, 580.000, 12.500, 12.500, RGB(28, 148, 84), -1, 0.800, 0, "FIG_E_b4_icon1")
    Set shp = AddFigureText(sld, 750.000, 580.000, 43.750, 13.750, "更鲁棒：", 8.500, RGB(28, 148, 84), true, "left", "middle", 0.000, "PingFang SC", "FIG_T_b4_head1")
    Set shp = AddFigureText(sld, 792.500, 580.000, 156.250, 13.750, "增强方向性缺陷感知", 8.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_b4_tail1")
    Set shp = AddFigureShape(sld, msoShapeOval, 731.250, 605.000, 12.500, 12.500, RGB(224, 58, 38), -1, 0.800, 0, "FIG_E_b4_icon2")
    Set shp = AddFigureText(sld, 750.000, 605.000, 43.750, 13.750, "更高效：", 8.500, RGB(224, 58, 38), true, "left", "middle", 0.000, "PingFang SC", "FIG_T_b4_head2")
    Set shp = AddFigureText(sld, 792.500, 605.000, 156.250, 13.750, "自适应融合多尺度特征", 8.500, RGB(40, 44, 60), false, "left", "middle", 0.000, "PingFang SC", "FIG_T_b4_tail2")
    Set sld = pres.Slides(1)
    Set shp = sld.Shapes.Range(Array("FIG_E_banner", "FIG_T_banner")).Group
    shp.Name = "FIG_G_banner"
    Set shp = sld.Shapes.Range(Array("FIG_E_panel_input", "FIG_E_panel_backbone", "FIG_E_panel_neck", "FIG_E_panel_head", "FIG_E_panel_legend", "FIG_E_panel_m1", "FIG_E_panel_m2", "FIG_E_panel_m3", "FIG_E_panel_b1", "FIG_E_panel_b2", "FIG_E_panel_b3", "FIG_E_panel_b4")).Group
    shp.Name = "FIG_G_panels"
    Set shp = sld.Shapes.Range(Array("FIG_T_input_title", "FIG_R_pcb", "FIG_T_input_size")).Group
    shp.Name = "FIG_G_input"
    Set shp = sld.Shapes.Range(Array("FIG_T_backbone_title", "FIG_T_backbone_sub", "FIG_E_bb1_top", "FIG_E_bb1_front", "FIG_E_bb1_right", "FIG_E_bb2_top", "FIG_E_bb2_front", "FIG_E_bb2_right", "FIG_E_bb3_top", "FIG_E_bb3_front", "FIG_E_bb3_right", "FIG_E_bb4_top", "FIG_E_bb4_front", "FIG_E_bb4_right", "FIG_L_bb_chain1", "FIG_L_bb_chain2", "FIG_L_bb_chain3", "FIG_L_bb_bus_in1", "FIG_L_bb_bus", "FIG_L_bb_bus_in3", "FIG_E_bb_c5", "FIG_T_bb_c5", "FIG_T_bb_c5_dim", "FIG_L_bb_to_c5", "FIG_E_bb_c4", "FIG_T_bb_c4", "FIG_T_bb_c4_dim", "FIG_L_bb_to_c4", "FIG_E_bb_c3", "FIG_T_bb_c3", "FIG_T_bb_c3_dim", "FIG_L_bb_to_c3", "FIG_T_backbone_dots", "FIG_E_bb_outbar", "FIG_T_bb_outbar")).Group
    shp.Name = "FIG_G_backbone"
    Set shp = sld.Shapes.Range(Array("FIG_T_neck_title", "FIG_E_neck_td", "FIG_T_neck_td", "FIG_E_neck_bu", "FIG_T_neck_bu", "FIG_E_neck_c5", "FIG_T_neck_c5", "FIG_E_neck_c4", "FIG_T_neck_c4", "FIG_E_neck_c3", "FIG_T_neck_c3", "FIG_E_neck_up1", "FIG_E_neck_up1_text", "FIG_E_neck_up2", "FIG_E_neck_up2_text", "FIG_E_neck_add_td1", "FIG_E_neck_add_td1_text", "FIG_E_neck_add_td2", "FIG_E_neck_add_td2_text", "FIG_E_neck_fusion1", "FIG_T_neck_fusion1", "FIG_E_neck_fusion2", "FIG_T_neck_fusion2", "FIG_E_neck_add_p5", "FIG_E_neck_add_p5_text", "FIG_E_neck_add_p4", "FIG_E_neck_add_p4_text", "FIG_E_neck_add_p3", "FIG_E_neck_add_p3_text", "FIG_E_neck_down1", "FIG_E_neck_down1_text", "FIG_E_neck_down2", "FIG_E_neck_down2_text", "FIG_E_neck_losc", "FIG_T_neck_losc", "FIG_E_neck_p5_top", "FIG_E_neck_p5_front", "FIG_E_neck_p5_right", "FIG_T_neck_p5", "FIG_T_neck_p5_dim", "FIG_E_neck_p4_top", "FIG_E_neck_p4_front", "FIG_E_neck_p4_right", "FIG_T_neck_p4", "FIG_T_neck_p4_dim", "FIG_E_neck_p3_top", "FIG_E_neck_p3_front", "FIG_E_neck_p3_right", "FIG_T_neck_p3", "FIG_T_neck_p3_dim", "FIG_L_neck_c5_up", "FIG_L_neck_up1_add", "FIG_L_neck_c4_add", "FIG_L_neck_add1_fusion", "FIG_L_neck_up2_add", "FIG_L_neck_c3_add", "FIG_L_neck_add2_fusion", "FIG_L_neck_addp5_cube", "FIG_L_neck_addp4_cube", "FIG_L_neck_addp3_cube", "FIG_L_neck_addp5_down", "FIG_L_neck_down_losc", "FIG_L_neck_losc_addp4", "FIG_L_neck_addp4_down2", "FIG_L_neck_down2_addp3", "FIG_LD_fusion1_addp5_1", "FIG_LD_fusion1_addp5_2", "FIG_LD_fusion1_addp5_3", "FIG_LD_fusion2_addp3_1", "FIG_LD_fusion2_addp3_2", "FIG_LD_fusion2_addp3_3", "FIG_LD_fusion1_fusion2_1", "FIG_E_neck_note_cga", "FIG_T_neck_note_cga", "FIG_E_neck_note_losc", "FIG_T_neck_note_losc")).Group
    shp.Name = "FIG_G_neck"
    Set shp = sld.Shapes.Range(Array("FIG_T_head_title", "FIG_E_head_iou", "FIG_T_head_iou", "FIG_E_head_iou_sq_0", "FIG_E_head_iou_sq_1", "FIG_E_head_iou_sq_2", "FIG_E_head_iou_sq_3", "FIG_E_head_iou_sq_4", "FIG_E_head_iou_sq_5", "FIG_E_head_iou_sq_6", "FIG_E_head_iou_sq_dots", "FIG_L_head_iou_dec", "FIG_E_head_dec", "FIG_T_head_dec", "FIG_E_head_dec_sq_0", "FIG_E_head_dec_sq_1", "FIG_E_head_dec_sq_2", "FIG_E_head_dec_sq_3", "FIG_E_head_dec_sq_4", "FIG_E_head_dec_sq_5", "FIG_E_head_dec_sq_dots", "FIG_L_head_dec_out", "FIG_E_head_out", "FIG_T_head_out", "FIG_E_head_bars_0", "FIG_E_head_bars_1", "FIG_E_head_bars_2", "FIG_E_head_bars_3", "FIG_E_head_bars_axis", "FIG_E_head_bbox")).Group
    shp.Name = "FIG_G_head"
    Set shp = sld.Shapes.Range(Array("FIG_T_legend_title", "FIG_E_lg_p5_top", "FIG_E_lg_p5_front", "FIG_E_lg_p5_right", "FIG_T_lg_p5", "FIG_E_lg_p4_top", "FIG_E_lg_p4_front", "FIG_E_lg_p4_right", "FIG_T_lg_p4", "FIG_E_lg_p3_top", "FIG_E_lg_p3_front", "FIG_E_lg_p3_right", "FIG_T_lg_p3", "FIG_E_lg_up", "FIG_E_lg_up_text", "FIG_T_lg_up", "FIG_E_lg_down", "FIG_E_lg_down_text", "FIG_T_lg_down", "FIG_E_lg_add", "FIG_E_lg_add_text", "FIG_T_lg_add", "FIG_E_lg_losc", "FIG_T_lg_losc_tag", "FIG_T_lg_losc", "FIG_E_lg_cga", "FIG_T_lg_cga_tag", "FIG_T_lg_cga", "FIG_LD_lg_solid", "FIG_T_lg_solid", "FIG_LD_lg_dash", "FIG_T_lg_dash")).Group
    shp.Name = "FIG_G_legend"
    Set shp = sld.Shapes.Range(Array("FIG_L_input_backbone", "FIG_L_backbone_neck", "FIG_L_neck_head")).Group
    shp.Name = "FIG_G_flow"
    Set shp = sld.Shapes.Range(Array("FIG_T_m1_title", "FIG_T_m1_in", "FIG_E_m1_csp", "FIG_T_m1_csp", "FIG_E_m1_b0", "FIG_T_m1_b0", "FIG_E_m1_b1", "FIG_T_m1_b1", "FIG_E_m1_b2", "FIG_T_m1_b2", "FIG_T_m1_bdots", "FIG_E_m1_fuse", "FIG_T_m1_fuse", "FIG_E_m1_ema", "FIG_T_m1_ema", "FIG_T_m1_out", "FIG_L_m1_f1", "FIG_L_m1_f2", "FIG_L_m1_f3", "FIG_L_m1_f4", "FIG_L_m1_f5", "FIG_T_m1_p0", "FIG_T_m1_p1", "FIG_T_m1_p2")).Group
    shp.Name = "FIG_G_m1"
    Set shp = sld.Shapes.Range(Array("FIG_T_m2_title", "FIG_T_m2_in", "FIG_G_m2_in_bg", "FIG_G_m2_in_v1", "FIG_G_m2_in_v2", "FIG_G_m2_in_v3", "FIG_G_m2_in_h1", "FIG_G_m2_in_h2", "FIG_E_m2_down", "FIG_E_m2_down_text", "FIG_G_m2_out_bg", "FIG_G_m2_out_v1", "FIG_G_m2_out_v2", "FIG_G_m2_out_h1", "FIG_T_m2_out", "FIG_L_m2_f1", "FIG_L_m2_f2", "FIG_T_m2_p0", "FIG_T_m2_p1", "FIG_T_m2_p2", "FIG_T_m2_p3")).Group
    shp.Name = "FIG_G_m2"
    Set shp = sld.Shapes.Range(Array("FIG_T_m3_title", "FIG_T_m3_fh", "FIG_G_m3_fh_bg", "FIG_G_m3_fh_v1", "FIG_G_m3_fh_v2", "FIG_G_m3_fh_v3", "FIG_G_m3_fh_h1", "FIG_G_m3_fh_h2", "FIG_E_m3_high", "FIG_T_m3_high", "FIG_T_m3_fl", "FIG_G_m3_fl_bg", "FIG_G_m3_fl_v1", "FIG_G_m3_fl_v2", "FIG_G_m3_fl_v3", "FIG_G_m3_fl_h1", "FIG_G_m3_fl_h2", "FIG_E_m3_low", "FIG_T_m3_low", "FIG_E_m3_add", "FIG_E_m3_add_text", "FIG_G_m3_out_bg", "FIG_G_m3_out_v1", "FIG_G_m3_out_v2", "FIG_G_m3_out_h1", "FIG_T_m3_out", "FIG_L_m3_f1_1", "FIG_L_m3_f1_2", "FIG_L_m3_f2", "FIG_L_m3_f3", "FIG_T_m3_p0", "FIG_T_m3_p1", "FIG_T_m3_p2", "FIG_T_m3_p3")).Group
    shp.Name = "FIG_G_m3"
    Set shp = sld.Shapes.Range(Array("FIG_T_b1_title", "FIG_E_out_p3_top", "FIG_E_out_p3_front", "FIG_E_out_p3_right", "FIG_E_out_p3_label", "FIG_E_out_p4_top", "FIG_E_out_p4_front", "FIG_E_out_p4_right", "FIG_E_out_p4_label", "FIG_E_out_p5_top", "FIG_E_out_p5_front", "FIG_E_out_p5_right", "FIG_E_out_p5_label", "FIG_T_b2_title", "FIG_T_b2_iou", "FIG_E_b2_sq1_0", "FIG_E_b2_sq1_1", "FIG_E_b2_sq1_2", "FIG_E_b2_sq1_3", "FIG_E_b2_sq1_4", "FIG_E_b2_sq1_5", "FIG_E_b2_sq1_dots", "FIG_T_b2_dec", "FIG_E_b2_sq2_0", "FIG_E_b2_sq2_1", "FIG_E_b2_sq2_2", "FIG_E_b2_sq2_3", "FIG_E_b2_sq2_4", "FIG_E_b2_sq2_dots", "FIG_E_b2_last", "FIG_L_b2_flow", "FIG_L_b1_b2", "FIG_L_b2_b3", "FIG_T_b3_title", "FIG_E_b3_bars_0", "FIG_E_b3_bars_1", "FIG_E_b3_bars_2", "FIG_E_b3_bars_3", "FIG_E_b3_bars_axis", "FIG_T_b3_bars", "FIG_E_b3_box1", "FIG_E_b3_box2", "FIG_T_b3_box", "FIG_T_b4_title", "FIG_E_b4_icon0", "FIG_T_b4_head0", "FIG_T_b4_tail0", "FIG_E_b4_icon1", "FIG_T_b4_head1", "FIG_T_b4_tail1", "FIG_E_b4_icon2", "FIG_T_b4_head2", "FIG_T_b4_tail2")).Group
    shp.Name = "FIG_G_bottom"
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
