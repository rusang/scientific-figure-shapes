# Toolchain Demo

这个示例展示同一份 Shapes 场景如何生成可编辑 `.pptx`、`BuildFinal` VBA、逻辑分组场景清单，
以及可供审计脚本共享的 `figure-manifest.json`。

```bash
python examples/toolchain-demo/build_demo.py --outdir /tmp/figure-shapes-demo

python scripts/manifest_validate.py \
  /tmp/figure-shapes-demo/figure-manifest.json --pretty

python scripts/pptx_text_audit.py \
  /tmp/figure-shapes-demo/editable-demo.pptx \
  --manifest /tmp/figure-shapes-demo/figure-manifest.json --pretty

python scripts/pptx_connector_audit.py \
  /tmp/figure-shapes-demo/editable-demo.pptx \
  --manifest /tmp/figure-shapes-demo/figure-manifest.json --pretty

python scripts/pptx_editability_audit.py \
  /tmp/figure-shapes-demo/editable-demo.pptx \
  --manifest /tmp/figure-shapes-demo/figure-manifest.json --pretty
```
