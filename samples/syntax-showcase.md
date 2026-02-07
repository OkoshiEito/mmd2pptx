# syntax-showcase.merged.pptx

複数記法を1ファイルに集約したサンプルです（1ファイル=1スライド）。

## 入力 `.mmd`

1. `samples/flowchart_syntax_coverage/03-02-markdown-strings.mmd`
2. `samples/flowchart_syntax_coverage/05-01-expanded-node-shapes.mmd`
3. `samples/flowchart_syntax_coverage/07-01-links.mmd`
4. `samples/class_syntax_coverage.mmd`
5. `samples/sequence-all-expr.mmd`
6. `samples/architecture-syntax.mmd`

## 再生成コマンド

```bash
node dist/cli.js \
  samples/flowchart_syntax_coverage/03-02-markdown-strings.mmd \
  samples/flowchart_syntax_coverage/05-01-expanded-node-shapes.mmd \
  samples/flowchart_syntax_coverage/07-01-links.mmd \
  samples/class_syntax_coverage.mmd \
  samples/sequence-all-expr.mmd \
  samples/architecture-syntax.mmd \
  -o samples/syntax-showcase.merged.pptx --renderer python
```
