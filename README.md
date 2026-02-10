# mmd2pptx

`.mmd` (Mermaid) を直接 `.pptx` に変換するためのコンパイラ型 CLI です。

## 特徴

- `Parse -> Normalize -> Measure -> Layout -> Render -> Embed` のフェーズ分割
- 中間表現 `IR` を JSON で出力可能
- flowchart/graph/classDiagram/architecture-beta を自前解析
- sequenceDiagram を python レンダラで直接変換
- architecture service の icon 指定を PNG 化して貼り込み（Iconify経由・ローカルキャッシュ）
- patch YAML でノード位置/サイズ・レイアウトを上書き
- 生成した pptx に `元mmd + patch` を埋め込み

## インストール

```bash
# Node.js 16+
npm install
npm run build
```

レンダラは `python` に統一されています（`--renderer` は不要）。
`python3`（または `python`）実行環境と `python-pptx` が必要です。

## 使い方

```bash
# 基本（buildは省略可）
node dist/cli.js samples/flow.mmd

# buildサブコマンド明示も可
node dist/cli.js build samples/flow.mmd -o samples/flow.pptx

# 複数mmdを1つのpptxに集約（1ファイル=1スライド）
node dist/cli.js samples/flow.mmd samples/self-loop-inline.mmd -o samples/combined.pptx

# ディレクトリ配下の .mmd を再帰収集して1つのpptxへ集約
# 各スライド左上に元ファイル名を表示
node dist/cli.js build samples -o samples/samples.merged.pptx

# patch適用
node dist/cli.js samples/flow.mmd --patch samples/flow.patch.yml -o samples/flow.pptx

# sequenceDiagram
node dist/cli.js build samples/sequence-all-expr.mmd -o samples/sequence-all-expr.pptx

# classDiagram
node dist/cli.js build samples/class_syntax_coverage.mmd -o samples/class_syntax_coverage.pptx

# IRを出力
node dist/cli.js build samples/flow.mmd --ir-out samples/flow.ir.json -o samples/flow.pptx

# 依存関係チェック
node dist/cli.js doctor
```

複数入力時の補足:

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
- Render: python-pptx で図形生成
- Embed: 元 mmd と patch をノートへ埋め込み

## 現在の対応範囲 (MVP)

- `flowchart` / `graph`
- `classDiagram`（関係マーカー・多重度ラベル・namespace・note・classDef/class/style/cssClass）
- `architecture-beta`（group/service/junction、`id:Side -- Side:id`、`{group}` 接続、service icon）
- `sequenceDiagram`
- ノード形状: 四角/角丸/円/菱形/平行四辺形
- エッジ: 実線/点線/太線、矢印、ラベル (`|label|`)
- subgraph

## 既知の制約

- `sequenceDiagram` は IR 非経由で直接レンダリング（`--ir-out` はダミー情報のみ）
- `stateDiagram` など未対応図種は順次対応予定
- エッジラベルは `-->|label|` 形式を優先サポート
- python レンダラ前提の実装

## 将来拡張

- `extract` コマンド (pptx から patch 抽出)
- ELK レイアウト切替
- `extract` 実装（pptx2patch）
