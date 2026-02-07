# mmd2pptx

`.mmd` (Mermaid) を直接 `.pptx` に変換するためのコンパイラ型 CLI です。

## 特徴

- `Parse -> Normalize -> Measure -> Layout -> Render -> Embed` のフェーズ分割
- 中間表現 `IR` を JSON で出力可能
- flowchart/graph/classDiagram を自前解析
- sequenceDiagram を python レンダラで直接変換
- patch YAML でノード位置/サイズ・レイアウトを上書き
- 生成した pptx に `元mmd + patch` を埋め込み

## インストール

```bash
# Node.js 16+
npm install
npm run build
```

`python` レンダラを使う場合は `uv` が必要です。
`--renderer auto`（既定）は `uv` が見つかれば `python`、なければ `js` を使います。

## 使い方

```bash
# 基本（buildは省略可）
node dist/cli.js samples/flow.mmd

# buildサブコマンド明示も可
node dist/cli.js build samples/flow.mmd -o samples/flow.pptx

# 複数mmdを1つのpptxに集約（1ファイル=1スライド）
node dist/cli.js samples/flow.mmd samples/self-loop-inline.mmd -o samples/combined.pptx --renderer python

# patch適用
node dist/cli.js samples/flow.mmd --patch samples/flow.patch.yml -o samples/flow.pptx

# JSレンダラを使う場合
node dist/cli.js samples/flow.mmd --renderer js -o samples/flow.pptx

# sequenceDiagram (python renderer)
node dist/cli.js build samples/sequence-all-expr.mmd --renderer python -o samples/sequence-all-expr.pptx

# classDiagram (python renderer)
node dist/cli.js build samples/class_syntax_coverage.mmd --renderer python -o samples/class_syntax_coverage.pptx

# IRを出力
node dist/cli.js build samples/flow.mmd --ir-out samples/flow.ir.json -o samples/flow.pptx

# 依存関係チェック
node dist/cli.js doctor
```

複数入力時の補足:

- `--renderer python` のみ対応
- `--patch` は未対応
- `--ir-out` は未対応
- `-o` 省略時は `<先頭入力名>.merged.pptx` を出力

開発時は `tsx` で直接実行できます。

```bash
npm run dev -- samples/flow.mmd -o samples/flow.pptx
```

## 配布（npm CLI）

```bash
# グローバルインストール
npm i -g mmd2pptx

# 単発実行
npx mmd2pptx samples/flow.mmd
```

公開手順（メンテナ向け）:

```bash
npm run lint
npm run build
npm version patch
npm publish --access public
```

## patch YAML 例

```yaml
layout:
  ranksep: 100
  nodesep: 70

nodes:
  Start:
    dx: -20
  Decide:
    w: 220
    h: 110

subgraphs:
  Core:
    padding: 36
```

## フェーズ設計

- Parse: Mermaid テキストを AST 相当オブジェクトへ
- Normalize: ノード/エッジ/サブグラフを正規化して IR 化
- Measure: 文字幅推定でノードサイズを決定
- Layout: Dagre で座標/経路を計算
- Render: python-pptx (既定) / PptxGenJS (任意) で図形生成
- Embed: 元 mmd と patch をノートへ埋め込み

## 現在の対応範囲 (MVP)

- `flowchart` / `graph`
- `classDiagram`（関係マーカー・多重度ラベル・namespace・note・classDef/class/style/cssClass）
- `sequenceDiagram`（python renderer）
- ノード形状: 四角/角丸/円/菱形/平行四辺形
- エッジ: 実線/点線/太線、矢印、ラベル (`|label|`)
- subgraph

## 既知の制約

- `sequenceDiagram` は IR 非経由で直接レンダリング（`--ir-out` はダミー情報のみ）
- `stateDiagram` など未対応図種は順次対応予定
- エッジラベルは `-->|label|` 形式を優先サポート
- `python` レンダラではコネクタ接続、`js` レンダラでは線分描画

## 将来拡張

- `extract` コマンド (pptx から patch 抽出)
- ELK レイアウト切替
- `extract` 実装（pptx2patch）
