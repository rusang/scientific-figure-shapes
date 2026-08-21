# MPHF-Net 整图重建 + 像素级验证(a2)

目标:把 `assets/reference-figures/mphf-net-architecture.png`(1536×1024)完整重建为可编辑 PPTX,渲染 preview,与原图跑 `render_delta_probe.py`,并通过 connector / layering 双审计。
范围:python-pptx 直建 editable deck(SKILL.md 第 11 条 fallback 路径);PCB 电路板图按 Crop Contract 保留为图像。非目标:VBA 宏路径、逐像素 1:1(字体渲染差异不可消除,按 SKILL "report whatever comparison data is available")。
Grounding:SKILL.md Fast Path/Fidelity Escalation、references/fidelity-review-gates.md 8 门禁、修复后的两个审计脚本(develop `6d4edc1`)。

- [x] 1. 裁剪 R1(PCB 板图)crop → `examples/mphf-net/assets/`(bbox 25,162,171,244;crops.json 落盘)
- [x] 2. 编写 `examples/mphf-net/build_mphf_deck.py`(约 200 shape,脚本跑通产出 pptx)
- [x] 3. routing + layering manifest(backbone 4 cuboid 含 expected_face_bounds_pt;两审计 passed / exit 0)
- [x] 4. soffice headless 渲染 preview.png(1280×853,比例一致)
- [x] 5. render_delta_probe:mad 0.0864 / SSIM 0.425;Render Gate 目测通过
- [x] 6. 迭代两轮:金字塔层叠坐标修正 + 12 处 badge 碰撞修复(端点退让,非豁免)
- [x] 7. 交付物入库 examples/mphf-net/ 并提交 push(见 progress)

风险:soffice 渲染中文字体与源图字体不同 → delta 有底噪,以结构指标 + 目测门禁为准;回滚:产物独立目录 examples/mphf-net/,可整体撤销。
