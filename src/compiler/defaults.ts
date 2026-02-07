import type { DiagramIr, LayoutConfig, NodeStyle, EdgeStyle, SubgraphStyle } from "../types.js";

export const DEFAULT_LAYOUT_CONFIG: LayoutConfig = {
  ranksep: 90,
  nodesep: 60,
  edgesep: 30,
  marginx: 30,
  marginy: 30,
};

export const DEFAULT_NODE_STYLE: NodeStyle = {
  fill: "F8FAFC",
  stroke: "0F172A",
  text: "0F172A",
  strokeWidth: 1,
  radius: 0.08,
  fontSize: 14,
  bold: false,
};

export const DEFAULT_EDGE_STYLE: EdgeStyle = {
  color: "1E293B",
  width: 1.3,
  lineStyle: "solid",
  arrow: "end",
  startMarker: "none",
  endMarker: "arrow",
  fontSize: 11,
};

export const DEFAULT_SUBGRAPH_STYLE: SubgraphStyle = {
  stroke: "64748B",
  fill: "F8FAFC",
  text: "334155",
  dash: "dash",
  strokeWidth: 1,
  padding: 24,
};

export function emptyIr(source: string): DiagramIr {
  return {
    meta: {
      direction: "TD",
      source,
    },
    config: {
      layout: { ...DEFAULT_LAYOUT_CONFIG },
      fontFamily: "Yu Gothic UI",
      lang: "ja-JP",
    },
    nodes: [],
    edges: [],
    subgraphs: [],
    bounds: {
      minX: 0,
      minY: 0,
      maxX: 0,
      maxY: 0,
      width: 0,
      height: 0,
    },
  };
}
