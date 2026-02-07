# mermaid2pptx 動作確認用: Flowchart 記法網羅（.mmd）

このファイルは公式ドキュメント（Flowchart Syntax）に載っている主要な記法を一通り通すためのテスト用です。  
https://mermaid.js.org/syntax/flowchart.html

---

## 1) graph / flowchart 宣言 + Direction（TB/TD/LR/RL/BT）

```mermaid
graph TB
  A[graph TB] --> B[Top to Bottom]
```

```mermaid
flowchart TD
  A[flowchart TD] --> B[Top-Down]
```

```mermaid
flowchart LR
  A[flowchart LR] --> B[Left to Right]
```

```mermaid
flowchart RL
  A[flowchart RL] --> B[Right to Left]
```

```mermaid
flowchart BT
  A[flowchart BT] --> B[Bottom to Top]
```

---

## 2) ノードID / テキスト / Unicode / Markdown（従来）

```mermaid
flowchart TD
  idOnly
  idWithText["表示テキスト（idと別）"]
  unicode["😄 Unicode は \" \" で囲む"]
  md["`**太字** と *斜体*`"]
  idOnly --> idWithText --> unicode --> md
```

---

## 3) Markdown Strings（自動折り返し + 改行 / 無効化）

```mermaid
flowchart LR
  A["`**Bold**
*Italic*
New line (newline char)`"]
  B["`長い文章は自動で折り返されます（markdownAutoWrap=true のとき）。`"]
  A -- "`**EdgeLabel**`" --> B
```

```mermaid
---
config:
  markdownAutoWrap: false
---
flowchart LR
  A["`AutoWrap OFF: とてもとてもとてもとてもとてもとても長い文章を折り返さずに表示（しようとする）`"] --> B["`OK`"]
```

---

## 4) 旧来ノード形状（Node shapes）

```mermaid
flowchart LR
  A[Rectangle]
  B(Rounded)
  C([Stadium])
  D[[Subroutine]]
  E[(Cylinder/Database)]
  F((Circle))
  G>Asymmetric]
  H{Rhombus}
  I{{Hexagon}}
  J[/Parallelogram/]
  K[\Parallelogram alt\]
  L[/Trapezoid\]
  M[\Trapezoid alt/]
  N(((Double circle)))
  A --> B --> C --> D --> E --> F --> G --> H --> I --> J --> K --> L --> M --> N
```

---

## 5) Expanded Node Shapes（v11.3.0+）: `@{ shape: ... }`

> 形状名（short name）を広く通すため、1つの図にまとめています。

```mermaid
flowchart LR
  bang@{ shape: bang, label: "bang" } --- notch@{ shape: notch-rect, label: "notch-rect" } --- cloud@{ shape: cloud, label: "cloud" } --- hour@{ shape: hourglass, label: "hourglass" } --- bolt@{ shape: bolt, label: "bolt" }

  brace@{ shape: brace, label: "brace" } --- bracer@{ shape: brace-r, label: "brace-r" } --- braces@{ shape: braces, label: "braces" }

  leanR@{ shape: lean-r, label: "lean-r" } --- leanL@{ shape: lean-l, label: "lean-l" } --- cyl@{ shape: cyl, label: "cyl" } --- diam@{ shape: diam, label: "diam" } --- delay@{ shape: delay, label: "delay" }

  hCyl@{ shape: h-cyl, label: "h-cyl" } --- linCyl@{ shape: lin-cyl, label: "lin-cyl" } --- curvTrap@{ shape: curv-trap, label: "curv-trap" } --- divRect@{ shape: div-rect, label: "div-rect" } --- doc@{ shape: doc, label: "doc" }

  rounded@{ shape: rounded, label: "rounded" } --- tri@{ shape: tri, label: "tri" } --- fork@{ shape: fork, label: "fork" } --- win@{ shape: win-pane, label: "win-pane" } --- fcirc@{ shape: f-circ, label: "f-circ" }

  linDoc@{ shape: lin-doc, label: "lin-doc" } --- linRect@{ shape: lin-rect, label: "lin-rect" } --- notchPent@{ shape: notch-pent, label: "notch-pent" } --- flipTri@{ shape: flip-tri, label: "flip-tri" } --- slRect@{ shape: sl-rect, label: "sl-rect" }

  trapT@{ shape: trap-t, label: "trap-t" } --- docs@{ shape: docs, label: "docs" } --- stRect@{ shape: st-rect, label: "st-rect" } --- odd@{ shape: odd, label: "odd" } --- flag@{ shape: flag, label: "flag" }

  hex@{ shape: hex, label: "hex" } --- trapB@{ shape: trap-b, label: "trap-b" } --- rect@{ shape: rect, label: "rect" } --- circle@{ shape: circle, label: "circle" } --- smcirc@{ shape: sm-circ, label: "sm-circ" }

  dblcirc@{ shape: dbl-circ, label: "dbl-circ" } --- frcirc@{ shape: fr-circ, label: "fr-circ" } --- bowRect@{ shape: bow-rect, label: "bow-rect" } --- frRect@{ shape: fr-rect, label: "fr-rect" } --- crossCirc@{ shape: cross-circ, label: "cross-circ" }

  tagDoc@{ shape: tag-doc, label: "tag-doc" } --- tagRect@{ shape: tag-rect, label: "tag-rect" } --- stadium@{ shape: stadium, label: "stadium" } --- text@{ shape: text, label: "text" }
```

---

## 6) Special shapes（icon / image）

```mermaid
flowchart LR
  %% icon shape（要: icon pack 登録）
  icon1@{ icon: "fa:fa-user", form: "circle", label: "Icon (circle)", pos: "b", h: 48 }

  %% image shape（data URI 1x1 PNGで外部アクセス不要）
  img1@{ img: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+Xv0sAAAAASUVORK5CYII=", label: "Image (data URI)", pos: "b", w: 60, h: 60, constraint: "on" }

  icon1 --> img1
```

---

## 7) Links（矢印 / オープン / ラベル / 点線 / 太線 / 透明）

```mermaid
flowchart LR
  A --> B
  B --- C
  C -->|pipe label| D
  D -- inline label --> E
  E -. dotted .-> F
  F == thick ==> G
  G == thick label ==> H
  H ~~~ I
```

---

## 8) Chaining / Multiple node links（表現力の高い1行記法）

```mermaid
graph LR
  A -- text --> B -- text2 --> C
  a --> b & c --> d
```

```mermaid
graph TB
  A & B --> C & D
```

---

## 9) Edge IDs / Animation / classDef（edgeへclass適用）/ Edge curve override

```mermaid
---
config:
  flowchart:
    curve: stepBefore
---
flowchart LR
  A e1@--> B
  B e2@==> C
  A e3@--> D

  %% edge metadata: animation / curve
  e1@{ animate: true }
  e2@{ animation: fast }
  e3@{ curve: linear }

  %% classDef で edge をアニメーション（stroke-dasharray のカンマは \, でエスケープ）
  classDef animate stroke-dasharray: 5\,5,animation: dash 2s linear infinite;
  class e1 animate
```

---

## 10) New arrow types（circle / cross） & Multi directional arrows

```mermaid
flowchart LR
  A---oB
  B---xC
```

```mermaid
flowchart LR
  A o--o B
  B <--> C
  C x--x D
```

---

## 11) Minimum length of a link（ダッシュ追加）

```mermaid
flowchart LR
  A --> B --> C --> D
  B ----> E
  C --> F
```

```mermaid
flowchart LR
  A --> B --> C --> D
  B -- label ---> E
  C --> F
```

---

## 12) Special characters / Entity codes

```mermaid
graph LR
  id1["This is the (text) in the box"] --> id2["Text with / or \ and {braces}"]
```

```mermaid
graph LR
  A["A double quote:#quot;"] --> B["A dec char:#9829;"]
```

---

## 13) Subgraphs（基本 / explicit id / subgraphと外部のリンク / subgraph内direction / limitation）

```mermaid
flowchart TB
  c1 --> a2
  subgraph one
    a1 --> a2
  end
```

```mermaid
flowchart TB
  c1 --> a2
  subgraph SG1[one (explicit id)]
    a1 --> a2
  end
```

```mermaid
flowchart TB
  subgraph SG
    a --> b
  end
  c --> SG
  SG --> d
```

```mermaid
flowchart TB
  subgraph outer
    direction LR
    x1 --> x2
    x2 --> x3
  end
```

```mermaid
flowchart TB
  %% Limitation: subgraph 内のノードが外部とリンクすると direction が親を継承
  subgraph outer
    direction LR
    y1 --> y2 --> y3
  end
  y2 --> outside
```

---

## 14) Comments（%%） / Styling（style, classDef, class, :::, linkStyle, default class）

```mermaid
flowchart LR
  %% コメントは %% で始まる「単独行」
  A[Normal] --> B[Styled]
  B --> C[Classed]
  C --> D[Default class]

  style B fill:#f9f,stroke:#333,stroke-width:4px

  classDef hot fill:#ff9,stroke:#333,stroke-width:4px
  class C hot

  classDef default fill:#eef,stroke:#333,stroke-width:2px

  %% linkStyle（0-based index）+ default
  linkStyle 1 stroke:#f00,stroke-width:4px,color:red
  linkStyle default stroke:#999
```

```mermaid
flowchart LR
  %% ::: 省略記法
  A:::hot --> B:::hot --> C
  classDef hot fill:#f96,stroke:#333,stroke-width:2px
```

---

## 15) Interaction（click）

```mermaid
flowchart LR
  A --> B --> C --> D
  click A callback "Tooltip"
  click B "https://www.github.com" "This is a link"
  click C call callback() "Tooltip"
  click D href "https://www.github.com" "This is a link" "_blank"
```

---

## 16) Graph declarations（スペース許容 / セミコロン省略）

```mermaid
graph LR;
  A[No space around label]-->B[OK];
  A --> C[Spaces between vertices and link]
  C --> D[No semicolon at end is OK]
```

---

## 17) Configuration（Renderer: elk）

```mermaid
---
config:
  flowchart:
    defaultRenderer: "elk"
---
flowchart LR
  A --> B --> C
```
