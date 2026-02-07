export type DiagramDirection = "TD" | "TB" | "BT" | "LR" | "RL";

export type NodeShape =
  | "rect"
  | "roundRect"
  | "circle"
  | "diamond"
  | "parallelogram"
  | "parallelogramAlt"
  | "trapezoid"
  | "trapezoidAlt"
  | "cylinder"
  | "subroutine"
  | "hexagon"
  | "cloud"
  | "lightningBolt"
  | "explosion"
  | "notchedRightArrow"
  | "hourglass"
  | "hCylinder"
  | "curvedTrapezoid"
  | "forkBar"
  | "windowPane"
  | "filledCircle"
  | "smallCircle"
  | "framedCircle"
  | "linedDocument"
  | "linedRect"
  | "wave"
  | "stackedRect"
  | "framedRect"
  | "braceLeft"
  | "braceRight"
  | "bracePair"
  | "card"
  | "delay"
  | "internalStorage"
  | "document"
  | "multiDocument"
  | "triangle"
  | "rightTriangle"
  | "chevron"
  | "plaqueTabs"
  | "pentagon"
  | "decagon"
  | "foldedCorner"
  | "donut"
  | "summingJunction";

export type EdgeLineStyle = "solid" | "dotted" | "thick" | "invisible";

export type ArrowType = "none" | "start" | "end" | "both";
export type EdgeMarker = "none" | "arrow" | "triangle" | "diamond" | "openDiamond" | "circle";
export type EdgeSide = "T" | "B" | "L" | "R";

export interface Point {
  x: number;
  y: number;
}

export interface ParsedNode {
  id: string;
  label?: string;
  shape?: NodeShape;
  icon?: string;
  isJunction?: boolean;
  subgraphId?: string;
  inlineClasses?: string[];
  line: number;
  raw: string;
}

export interface ParsedEdge {
  from: string;
  to: string;
  label?: string;
  startLabel?: string;
  endLabel?: string;
  startMarker?: EdgeMarker;
  endMarker?: EdgeMarker;
  startSide?: EdgeSide;
  endSide?: EdgeSide;
  startViaGroup?: boolean;
  endViaGroup?: boolean;
  style: EdgeLineStyle;
  arrow: ArrowType;
  line: number;
  raw: string;
}

export interface ParsedSubgraph {
  id: string;
  title: string;
  parentId?: string;
  line: number;
}

export interface DiagramAst {
  source: string;
  direction: DiagramDirection;
  title?: string;
  nodes: ParsedNode[];
  edges: ParsedEdge[];
  subgraphs: ParsedSubgraph[];
  classDefs: Record<string, Record<string, string>>;
  classAssignments: Record<string, string[]>;
  styleOverrides: Record<string, Record<string, string>>;
  layoutHints: {
    nodeSpacing?: number;
    rankSpacing?: number;
  };
}

export interface NodeStyle {
  fill: string;
  stroke: string;
  text: string;
  strokeWidth: number;
  radius?: number;
  fontSize: number;
  bold?: boolean;
}

export interface EdgeStyle {
  color: string;
  width: number;
  lineStyle: EdgeLineStyle;
  arrow: ArrowType;
  startMarker: EdgeMarker;
  endMarker: EdgeMarker;
  startSide?: EdgeSide;
  endSide?: EdgeSide;
  startViaGroup?: boolean;
  endViaGroup?: boolean;
  fontSize: number;
}

export interface SubgraphStyle {
  stroke: string;
  fill: string;
  text: string;
  dash: "solid" | "dash";
  strokeWidth: number;
  padding: number;
}

export interface IrNode {
  id: string;
  label: string;
  shape: NodeShape;
  icon?: string;
  isJunction?: boolean;
  subgraphId?: string;
  width: number;
  height: number;
  x: number;
  y: number;
  style: NodeStyle;
}

export interface IrEdge {
  id: string;
  from: string;
  to: string;
  label?: string;
  startLabel?: string;
  endLabel?: string;
  points: Point[];
  labelPosition?: Point;
  style: EdgeStyle;
}

export interface IrSubgraph {
  id: string;
  title: string;
  parentId?: string;
  nodeIds: string[];
  x: number;
  y: number;
  width: number;
  height: number;
  style: SubgraphStyle;
}

export interface LayoutConfig {
  ranksep: number;
  nodesep: number;
  edgesep: number;
  marginx: number;
  marginy: number;
}

export interface DiagramIr {
  meta: {
    direction: DiagramDirection;
    source: string;
    title?: string;
  };
  config: {
    layout: LayoutConfig;
    fontFamily: string;
    lang: string;
  };
  nodes: IrNode[];
  edges: IrEdge[];
  subgraphs: IrSubgraph[];
  bounds: {
    minX: number;
    minY: number;
    maxX: number;
    maxY: number;
    width: number;
    height: number;
  };
}

export interface NodePatch {
  x?: number;
  y?: number;
  dx?: number;
  dy?: number;
  w?: number;
  h?: number;
  minW?: number;
  minH?: number;
  maxW?: number;
  maxH?: number;
}

export interface EdgePatch {
  points?: Point[];
  labelDx?: number;
  labelDy?: number;
}

export interface SubgraphPatch {
  padding?: number;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
}

export interface DiagramPatch {
  layout?: Partial<LayoutConfig>;
  nodes?: Record<string, NodePatch>;
  edges?: Record<string, EdgePatch>;
  subgraphs?: Record<string, SubgraphPatch>;
  renderer?: {
    fontFamily?: string;
    lang?: string;
  };
}

export interface BuildOptions {
  patch?: DiagramPatch;
  irOutPath?: string;
  fontFamily?: string;
  lang?: string;
  targetAspectRatio?: number;
}
