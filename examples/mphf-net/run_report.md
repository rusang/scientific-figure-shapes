# MPHF-Net 整图重建验证报告

日期:2026-08-22。源图:`assets/reference-figures/mphf-net-architecture.png`(1536×1024 px)。
映射:uniform 0.625 pt/px → slide 960×640 pt。构建:`build_mphf_deck.py`(python-pptx,editable fallback 路径)。

## 交付物

| 文件 | 说明 |
|------|------|
| `build_mphf_deck.py` | 可复跑构建脚本(同时生成两份审计 manifest) |
| `mphf_net_rebuild.pptx` | 可编辑重建稿(全部 `SUMMER_` 命名,约 200 shape) |
| `assets/01_R1_pcb.png` + `crops.json` | 保留 crop(PCB 板图,bbox 25,162,171,244 px) |
| `layering_manifest.json` / `routing_manifest.json` | 审计清单(backbone 金字塔含 expected_face_bounds_pt) |
| `preview.png` | 诊断预览(LibreOffice headless 渲染,非交付稿) |

## 验证结果(命令 + 实际输出)

- `pptx_layering_audit.py … --manifest layering_manifest.json` → `passed: true`,exit 0。
  覆盖:backbone 4 层 cuboid(三面/参考几何 ±1pt/尺寸链/z 序/相邻重叠)+ Neck 与底行 P3/P4/P5 共 10 组三面块。
- `pptx_connector_audit.py … --manifest routing_manifest.json` → `passed: true`,exit 0。
  覆盖:9 条命名路由 target-edge 校验;首轮 12 处 badge 文字碰撞已通过收窄 badge 文本框 + 端点退让 3px 修复(非豁免)。
- `render_delta_probe.py 源图 preview.png` → `mean_abs_delta 0.0864`,`SSIM 0.425`,尺寸比例一致。
  按 SKILL 约定不设 SSIM 硬阈值,如实报告;主要残差来自字体渲染差异与微布局偏移。
- Render Gate 目测:14 个面板全部就位,顺序与源图一致,无意外空白,crop 无拉伸。

## 未编辑区域(诚实清单)

- PCB 电路板图为图像 crop(纹理密集,按 Crop Contract 保留)。
- 其余全部为可编辑 shape/text/connector。
