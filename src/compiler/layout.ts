import dagre from "dagre";
import type { DiagramDirection, DiagramIr, Point } from "../types.js";
import { recomputeBounds, recomputeSubgraphBounds } from "./geometry.js";

interface LayoutOptions {
  targetAspectRatio?: number;
}

interface LayoutState {
  nodePositions: Map<string, { x: number; y: number }>;
  edgeRoutes: Map<string, { points: Point[]; labelPosition?: Point }>;
  subgraphBounds: Map<string, { x: number; y: number; width: number; height: number }>;
  bounds: DiagramIr["bounds"];
}

const MIN_LAYOUT_ASPECT = 4 / 3;
const MAX_LAYOUT_ASPECT = 16 / 9;

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function clampAspect(value: number | undefined): number {
  if (!Number.isFinite(value)) {
    return MAX_LAYOUT_ASPECT;
  }
  return clamp(value as number, MIN_LAYOUT_ASPECT, MAX_LAYOUT_ASPECT);
}

function toRankdir(direction: DiagramDirection): "TB" | "BT" | "LR" | "RL" {
  if (direction === "BT") {
    return "BT";
  }
  if (direction === "LR") {
    return "LR";
  }
  if (direction === "RL") {
    return "RL";
  }
  return "TB";
}

function fallbackLabelPosition(points: Point[]): Point | undefined {
  if (points.length < 2) {
    return undefined;
  }

  const lengths: number[] = [];
  let total = 0;

  for (let i = 0; i < points.length - 1; i += 1) {
    const dx = points[i + 1].x - points[i].x;
    const dy = points[i + 1].y - points[i].y;
    const length = Math.hypot(dx, dy);
    lengths.push(length);
    total += length;
  }

  if (total <= 0) {
    const p0 = points[0];
    const p1 = points[points.length - 1];
    return {
      x: (p0.x + p1.x) / 2,
      y: (p0.y + p1.y) / 2,
    };
  }

  const target = total / 2;
  let walked = 0;

  for (let i = 0; i < lengths.length; i += 1) {
    const segment = lengths[i];
    if (segment <= 0) {
      continue;
    }

    if (walked + segment < target) {
      walked += segment;
      continue;
    }

    const p0 = points[i];
    const p1 = points[i + 1];
    const t = (target - walked) / segment;
    const x = p0.x + (p1.x - p0.x) * t;
    const y = p0.y + (p1.y - p0.y) * t;

    const nx = -(p1.y - p0.y) / segment;
    const ny = (p1.x - p0.x) / segment;
    return {
      x: x + nx * 14,
      y: y + ny * 14,
    };
  }

  return {
    x: points[0].x,
    y: points[0].y,
  };
}

function applyLayout(ir: DiagramIr): void {
  const graph = new dagre.graphlib.Graph({ multigraph: true, compound: true });
  const subgraphIds = new Set(ir.subgraphs.map((subgraph) => subgraph.id));
  graph.setGraph({
    rankdir: toRankdir(ir.meta.direction),
    ranksep: ir.config.layout.ranksep,
    nodesep: ir.config.layout.nodesep,
    edgesep: ir.config.layout.edgesep,
    marginx: ir.config.layout.marginx,
    marginy: ir.config.layout.marginy,
  });
  graph.setDefaultEdgeLabel(() => ({}));

  for (const subgraph of ir.subgraphs) {
    graph.setNode(subgraph.id, {
      width: 1,
      height: 1,
      label: subgraph.title,
    });
  }

  for (const node of ir.nodes) {
    if (subgraphIds.has(node.id)) {
      continue;
    }

    graph.setNode(node.id, {
      width: node.width,
      height: node.height,
    });

    if (node.subgraphId && graph.hasNode(node.subgraphId)) {
      graph.setParent(node.id, node.subgraphId);
    }
  }

  const subgraphAnchorNode = new Map<string, string>();
  for (const node of ir.nodes) {
    if (!node.subgraphId) {
      continue;
    }
    if (subgraphIds.has(node.id)) {
      continue;
    }
    if (!subgraphAnchorNode.has(node.subgraphId)) {
      subgraphAnchorNode.set(node.subgraphId, node.id);
    }
  }

  for (const edge of ir.edges) {
    let fromId = edge.from;
    let toId = edge.to;

    if (subgraphIds.has(fromId)) {
      fromId = subgraphAnchorNode.get(fromId) ?? fromId;
    }
    if (subgraphIds.has(toId)) {
      toId = subgraphAnchorNode.get(toId) ?? toId;
    }

    if (!graph.hasNode(fromId) || !graph.hasNode(toId)) {
      continue;
    }

    graph.setEdge(
      { v: fromId, w: toId, name: edge.id },
      {
        id: edge.id,
        label: edge.label,
      },
    );
  }

  dagre.layout(graph);

  for (const node of ir.nodes) {
    const layoutNode = graph.node(node.id) as { x: number; y: number; width: number; height: number } | undefined;
    if (!layoutNode) {
      continue;
    }

    node.x = layoutNode.x - node.width / 2;
    node.y = layoutNode.y - node.height / 2;
  }

  for (const edge of ir.edges) {
    const layoutEdge = graph.edge({ v: edge.from, w: edge.to, name: edge.id }) as
      | { points?: Array<{ x: number; y: number }>; x?: number; y?: number }
      | undefined;

    if (!layoutEdge) {
      continue;
    }

    edge.points = Array.isArray(layoutEdge.points)
      ? layoutEdge.points.map((point) => ({
          x: point.x,
          y: point.y,
        }))
      : [];

    if (typeof layoutEdge.x === "number" && typeof layoutEdge.y === "number") {
      edge.labelPosition = {
        x: layoutEdge.x,
        y: layoutEdge.y,
      };
    } else if (edge.label) {
      edge.labelPosition = fallbackLabelPosition(edge.points);
    }
  }

  recomputeSubgraphBounds(ir);
  recomputeBounds(ir);
}

function diagramAspect(ir: DiagramIr): number {
  return ir.bounds.width / Math.max(1, ir.bounds.height);
}

function captureLayoutState(ir: DiagramIr): LayoutState {
  return {
    nodePositions: new Map(ir.nodes.map((node) => [node.id, { x: node.x, y: node.y }])),
    edgeRoutes: new Map(
      ir.edges.map((edge) => [
        edge.id,
        {
          points: edge.points.map((point) => ({ x: point.x, y: point.y })),
          labelPosition: edge.labelPosition ? { ...edge.labelPosition } : undefined,
        },
      ]),
    ),
    subgraphBounds: new Map(
      ir.subgraphs.map((subgraph) => [
        subgraph.id,
        {
          x: subgraph.x,
          y: subgraph.y,
          width: subgraph.width,
          height: subgraph.height,
        },
      ]),
    ),
    bounds: { ...ir.bounds },
  };
}

function restoreLayoutState(ir: DiagramIr, state: LayoutState): void {
  for (const node of ir.nodes) {
    const pos = state.nodePositions.get(node.id);
    if (!pos) {
      continue;
    }
    node.x = pos.x;
    node.y = pos.y;
  }

  for (const edge of ir.edges) {
    const route = state.edgeRoutes.get(edge.id);
    if (!route) {
      continue;
    }
    edge.points = route.points.map((point) => ({ ...point }));
    edge.labelPosition = route.labelPosition ? { ...route.labelPosition } : undefined;
  }

  for (const subgraph of ir.subgraphs) {
    const bounds = state.subgraphBounds.get(subgraph.id);
    if (!bounds) {
      continue;
    }
    subgraph.x = bounds.x;
    subgraph.y = bounds.y;
    subgraph.width = bounds.width;
    subgraph.height = bounds.height;
  }

  ir.bounds = { ...state.bounds };
}

function tuneLayoutForAspect(ir: DiagramIr, targetAspectRatio: number): void {
  const target = clampAspect(targetAspectRatio);
  const rankdir = toRankdir(ir.meta.direction);
  const isVerticalFlow = rankdir === "TB" || rankdir === "BT";

  const baseNodesep = ir.config.layout.nodesep;
  const baseRanksep = ir.config.layout.ranksep;

  applyLayout(ir);

  let bestState = captureLayoutState(ir);
  let bestNodesep = ir.config.layout.nodesep;
  let bestRanksep = ir.config.layout.ranksep;
  let bestScore = Math.abs(Math.log(diagramAspect(ir) / target));

  let candidateNodesep = baseNodesep;
  let candidateRanksep = baseRanksep;

  for (let step = 0; step < 4; step += 1) {
    const aspect = diagramAspect(ir);
    const ratio = aspect / target;

    if (ratio > 1.06) {
      if (isVerticalFlow) {
        candidateNodesep = clamp(candidateNodesep * 0.86, 22, 260);
        candidateRanksep = clamp(candidateRanksep * 1.20, 40, 380);
      } else {
        candidateNodesep = clamp(candidateNodesep * 1.20, 22, 260);
        candidateRanksep = clamp(candidateRanksep * 0.86, 40, 380);
      }
    } else if (ratio < 0.94) {
      if (isVerticalFlow) {
        candidateNodesep = clamp(candidateNodesep * 1.16, 22, 260);
        candidateRanksep = clamp(candidateRanksep * 0.88, 40, 380);
      } else {
        candidateNodesep = clamp(candidateNodesep * 0.88, 22, 260);
        candidateRanksep = clamp(candidateRanksep * 1.16, 40, 380);
      }
    } else {
      break;
    }

    ir.config.layout.nodesep = candidateNodesep;
    ir.config.layout.ranksep = candidateRanksep;
    applyLayout(ir);

    const score = Math.abs(Math.log(diagramAspect(ir) / target));
    if (score < bestScore) {
      bestScore = score;
      bestState = captureLayoutState(ir);
      bestNodesep = candidateNodesep;
      bestRanksep = candidateRanksep;
    }
  }

  restoreLayoutState(ir, bestState);
  ir.config.layout.nodesep = bestNodesep;
  ir.config.layout.ranksep = bestRanksep;
}

export function layoutDiagram(ir: DiagramIr, options: LayoutOptions = {}): void {
  if (options.targetAspectRatio) {
    tuneLayoutForAspect(ir, options.targetAspectRatio);
    return;
  }

  applyLayout(ir);
}
