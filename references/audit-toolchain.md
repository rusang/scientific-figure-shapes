# Audit Toolchain

本页给出高保真重建的最小可复现校验链。所有审计都读取同一份
`figure-manifest.json`；不要为每个脚本维护互相漂移的独立配置。

## 1. 先校验 manifest

```bash
python scripts/manifest_validate.py figure-manifest.json --pretty
```

结构错误必须先修复。`simplified` 或 `omitted` 若未标记 `approved: true`，会产生警告；
高保真任务仍应视为人工复核项。

## 2. 生成同源可编辑输出

优先使用 `office_shape_canvas.py` 在一次场景定义中同时生成 `.pptx` 与 `.bas`。宏入口统一为
`BuildFinal`。macOS 上先运行能力探测；若 `do Visual Basic` 不可用，直接交付
`python-pptx` 物化结果，不重复尝试 AppleScript。

简单图标可先描摹：

```bash
python scripts/vector_trace.py icon.png \
  --threshold 245 --min-area 20 \
  --svg-out assets/icon.svg --pptx-out assets/icon-editable.pptx --pretty
```

## 3. 结构审计

```bash
python scripts/pptx_text_audit.py final.pptx \
  --manifest figure-manifest.json --json-out audit/text.json --pretty

python scripts/pptx_connector_audit.py final.pptx \
  --manifest figure-manifest.json --json-out audit/connectors.json --pretty

python scripts/pptx_layering_audit.py final.pptx \
  --manifest figure-manifest.json --json-out audit/layering.json --pretty

python scripts/pptx_editability_audit.py final.pptx \
  --manifest figure-manifest.json --json-out audit/editability.json --pretty
```

箭头碰撞、重复线段、零长度线段、错误入边、缺失立体面、错误顶点、渐变色/方向不符、
未登记的大面积栅格图都属于失败。

## 4. 参考图保真审计

```bash
python scripts/fidelity_audit.py reference.png preview.png \
  --manifest figure-manifest.json \
  --rows 4 --cols 4 --warn-mae 0.08 --fail-mae 0.16 \
  --heatmap audit/fidelity-heatmap.png \
  --json-out audit/fidelity.json --pretty
```

阈值不是通用真理。先用默认值定位大问题，再针对文字密集图或渐变背景调整。显著元素 ROI
优先级高于全图均值；全图差异低不能抵消小图标、箭头或标签缺失。

## 5. 双渲染兼容性

仅在 PowerPoint 与 LibreOffice 两份 PNG 都可用时运行：

```bash
python scripts/dual_renderer_probe.py \
  render/powerpoint.png render/libreoffice.png \
  --max-mean-delta 0.015 \
  --heatmap audit/renderer-delta.png \
  --json-out audit/renderer-delta.json --pretty
```

双渲染差分用于发现字体、渐变、透明度和连接线兼容性问题；它不替代参考图保真审计。
