# MPHF-Net 从零重建 v2 报告(2026-08-23)

按 skill 当前版工作流(Fidelity Escalation)从零重跑:ShapeCanvas 官方
运行时 + FIG_ 前缀命名契约;场景坐标为自绘(源图 px,uniform 0.625 pt/px)。

## 相对 v1 的样式升级

- 立方体改**等轴测三面**(top/right 斜面 freeform + front 90° 垂直渐变
  ±24/-18,edge=front-70 描边),不再是平顶矩形拼接;
- UP/Down 徽章改浅底描边 + 符号,+ 徽章深底白字贴原图;
- CGAFusion 跨层连接改**正交虚线折线**(ortho polyline),贴原图路由;
- 主流程箭头 4pt 粗线 + 三角头;检测框/输出框虚线样式;
- 柱状图补底部轴线(出头设计,登记 ignore_dangling)。

## 门禁与验证(命令输出)

- pptx_connector_audit + routing_manifest(routes 13 条,其中 head 两条
  带 source/source_edge;badge 字标 ignore_text_shapes;轴线/图例样例
  ignore_dangling)→ passed true。
- pptx_layering_audit + layering_manifest(13 组等轴测三面块;backbone
  金字塔 size/z/overlap 链)→ passed true。
- pptx_text_audit + text_manifest(6 条标题 max_lines=1 门禁;曾抓出
  Head 标题与 Decoder 行折行,已修)→ passed true。
- pptx_editability_audit → passed true(唯一 raster 为 FIG_R_pcb crop)。
- render_delta_probe:mad 0.0824,SSIM 0.455(v1 0.0873/0.425;
  codex v6 基准 0.0644/0.601——差距主要在图标细节与字体渲染,如实报告)。

## 交付物

build_v2.py(可复跑)、mphf_net_v2.pptx(可编辑,等轴测样式)、
mphf_net_v2.bas(VBA)、scene-manifest.json、routing/layering/text 三份
审计 manifest、assets/01_R1_pcb.png + crops.json、preview.png(诊断预览)。
