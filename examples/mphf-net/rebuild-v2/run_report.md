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

## 箭头覆盖补齐(2026-08-23 第二轮)

用户对照发现多面板缺箭头。根因:审计缺覆盖率维度(该画没画不可见),
已在 skill 层新增 min_connector_count 清点门禁(commit 46a2814)。本轮补:

- Neck 跨层虚线改为原图 4 条拓扑:fusion1→+P5、fusion1→UP2、
  fusion2→+P4、fusion2→+P3(删自造的 fusion1→fusion2 直线),并改
  FIG_L_ 前缀纳入锚定/清点审计;
- 模块一:CSP→三分支扇出 3 条 + 三分支→融合扇入 3 条(原为单轴线);
- 模块三:Fh/Fl grid→高低频分支 2 条;
- Decoder 方块序列连线:Head 5 段 + b2 4 段。

验收:connector audit passed(68 条被审计,min_connector_count=68 随
构建落盘防缩水);layering/text/editability passed;mad 0.0818 /
SSIM 0.4656(逐轮 0.425→0.455→0.464→0.466)。
