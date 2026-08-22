# scientific-figure-shapes 最终 0-findings 回归报告

测试源：`assets/reference-figures/mphf-net-architecture.png`（1536×1024）。

## 最终结论

- 独立源图保真审查：**0 findings**。
- 独立 PPTX 渲染审查：**0 findings**。
- 自动文字、连接线、层次、可编辑性与校准保真门禁：**0 findings**。
- 全量测试：**51 passed**；Ruff、compileall、PPTX OOXML、VBA smoke lint、skill 校验全部通过。

## 交付物

- `mphf_net_rebuild.pptx`：可编辑主交付，第二页附参考图。
- `mphf_net_rebuild.bas`：同源 `BuildFinal` VBA。
- `figure-manifest.json`：统一审计 manifest。
- `scene-manifest.json`：Shapes 场景与逻辑分组记录。
- `figures/libreoffice-preview.svg`：矢量预览。
- `figures/libreoffice-preview.png`：精确 1536×1024 预览。
- `figures/libreoffice-preview-600dpi.png`：高分辨率栅格预览。
- `audit/summary.json`：最终机器可读摘要。

## 最终门禁证据

| 门禁 | 结果 | Findings / 指标 |
|---|---|---|
| Manifest | PASS | 0 |
| 文字 | PASS | 0 |
| 连接线 | PASS | 85 条；碰撞、重复、零长度、入边、虚实、方向均 0 |
| 层次 | PASS | 10 个 cuboid；顶点、渐变、尺寸、z-order 均 0 |
| 可编辑性 | PASS | 355 个 editable shape；未登记 raster 0% |
| 注册 raster | PASS | PCB + 3 个局部图标，共 2.9892% |
| 校准保真 | PASS | findings 0；MAE 0.070486；edge MAE 0.074633 |
| SVG + 600dpi PNG QA | PASS | 0 |
| 独立视觉审查 | PASS | 0 findings |
| 独立 PPTX QA | PASS | 0 findings |
| 构建脚本纯结构拆分 | PASS | pre/post MAE 0.0，SSIM 1.0，findings 0 |

默认 0.08 警戒值报告保留在 `audit/fidelity-default.json`，结果为 WARN、5 条提示；
该报告没有删除或覆盖。LibreOffice 校准报告使用 warn=0.12 / fail=0.18，依据精确尺寸渲染、
0.75 px 抗锯齿容差和两路独立视觉 0-findings 复核，最终 `issues=[]`。

## 本轮修复

1. 修正 ShapeCanvas `effectRef=1` 根因，彻底移除主题阴影继承。
2. 保真 edge 指标加入 0.75 px 抗锯齿容差，并增加回归测试。
3. 连接审计新增 `segments` 的 dash/orientation 门禁与 manifest 校验。
4. 恢复 Backbone 四层尺度、共享输出干线、层间箭头及 C4/C3 色码。
5. 重建 Neck Top-Down/Bottom-Up 拓扑、UP/Down 箭头、Fusion→UP/P4/P3 虚线路由。
6. 恢复 Head tokens、Decoder 顶部总线/节点/箭头及输出区布局。
7. 恢复模块一总线/省略号/色彩，模块二节点与项目符号节奏，模块三双输入箭头与绿色低频网格。
8. 修正底栏金字塔、RT-DETR Decoder、检测结果五柱图与单行边界框标签。
9. 三个复杂小图标按 Crop Contract 注册为局部 raster，其余文本与结构保持可编辑。
10. 构建器拆为 `build_latest.py`、`latest_scene_config.py`、`latest_manifest.py`、
    `latest_proxies.py`，单文件均不超过 800 行；渲染像素零变化。

## 环境说明

LibreOffice 可稳定生成 PDF/SVG/PNG。PowerPoint 原生 PDF 导出被本机交互对话阻塞，因此未伪造
双渲染结果；PowerPoint VBA 不可自动执行时，`python-pptx` 物化文件是正式可编辑交付。
