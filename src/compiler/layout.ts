import dagre from "dagre";
import type { DiagramDirection, DiagramIr, EdgeSide, Point } from "../types.js";
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

type JunctionRelation = "left" | "right" | "up" | "down";

function relationFromSides(nodeSide?: EdgeSide, junctionSide?: EdgeSide): JunctionRelation | undefined {
  if (nodeSide === "R" && junctionSide === "L") {
    return "left";
  }
  if (nodeSide === "L" && junctionSide === "R") {
    return "right";
  }
  if (nodeSide === "B" && junctionSide === "T") {
    return "up";
  }
  if (nodeSide === "T" && junctionSide === "B") {
    return "down";
  }

  if (junctionSide === "L") {
    return "left";
  }
  if (junctionSide === "R") {
    return "right";
  }
  if (junctionSide === "T") {
    return "up";
  }
  if (junctionSide === "B") {
    return "down";
  }

  return undefined;
}

function enforceJunctionSidePlacement(ir: DiagramIr): Set<string> {
  const moved = new Set<string>();
  const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));
  const junctions = ir.nodes.filter((node) => node.isJunction);

  for (const junction of junctions) {
    const junctionCx = junction.x + junction.width / 2;
    const junctionCy = junction.y + junction.height / 2;

    for (const edge of ir.edges) {
      let otherId: string | undefined;
      let otherSide: EdgeSide | undefined;
      let junctionSide: EdgeSide | undefined;

      if (edge.from === junction.id) {
        otherId = edge.to;
        otherSide = edge.style.endSide;
        junctionSide = edge.style.startSide;
      } else if (edge.to === junction.id) {
        otherId = edge.from;
        otherSide = edge.style.startSide;
        junctionSide = edge.style.endSide;
      }

      if (!otherId) {
        continue;
      }

      const other = nodeById.get(otherId);
      if (!other || other.isJunction) {
        continue;
      }

      const relation = relationFromSides(otherSide, junctionSide);
      if (!relation) {
        continue;
      }

      const gap = Math.max(42, Math.min(150, Math.max(other.width, other.height) * 0.72));

      if (relation === "left") {
        other.x = junction.x - gap - other.width;
        other.y = junctionCy - other.height / 2;
        moved.add(other.id);
        continue;
      }
      if (relation === "right") {
        other.x = junction.x + junction.width + gap;
        other.y = junctionCy - other.height / 2;
        moved.add(other.id);
        continue;
      }
      if (relation === "up") {
        other.x = junctionCx - other.width / 2;
        other.y = junction.y - gap - other.height;
        moved.add(other.id);
        continue;
      }

      other.x = junctionCx - other.width / 2;
      other.y = junction.y + junction.height + gap;
      moved.add(other.id);
    }
  }

  return moved;
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

  const subgraphChildren = new Map<string, string[]>();
  for (const subgraph of ir.subgraphs) {
    const parentId = subgraph.parentId;
    if (!parentId || parentId === subgraph.id || !graph.hasNode(parentId) || !graph.hasNode(subgraph.id)) {
      continue;
    }

    graph.setParent(subgraph.id, parentId);
    const children = subgraphChildren.get(parentId) ?? [];
    children.push(subgraph.id);
    subgraphChildren.set(parentId, children);
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

  const resolveSubgraphAnchor = (subgraphId: string, visited: Set<string> = new Set()): string | undefined => {
    if (subgraphAnchorNode.has(subgraphId)) {
      return subgraphAnchorNode.get(subgraphId);
    }

    if (visited.has(subgraphId)) {
      return undefined;
    }
    visited.add(subgraphId);

    const children = subgraphChildren.get(subgraphId) ?? [];
    for (const childId of children) {
      const anchor = resolveSubgraphAnchor(childId, visited);
      if (anchor) {
        subgraphAnchorNode.set(subgraphId, anchor);
        return anchor;
      }
    }

    return undefined;
  };

  for (const subgraph of ir.subgraphs) {
    resolveSubgraphAnchor(subgraph.id);
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

  const movedByJunction = enforceJunctionSidePlacement(ir);

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

  if (movedByJunction.size > 0) {
    for (const edge of ir.edges) {
      if (movedByJunction.has(edge.from) || movedByJunction.has(edge.to)) {
        edge.points = [];
        edge.labelPosition = undefined;
      }
    }
  }

  recomputeSubgraphBounds(ir);
  recomputeBounds(ir);
}

function diagramAspect(ir: DiagramIr): number {
  return ir.bounds.width / Math.max(1, ir.bounds.height);
}

function normalizedFitScale(ir: DiagramIr, targetAspectRatio: number): number {
  const target = clampAspect(targetAspectRatio);
  const width = Math.max(1, ir.bounds.width);
  const height = Math.max(1, ir.bounds.height);
  // Compare layouts by the uniform scaling factor required to fit into a
  // targetAspectRatio rectangle (height=1, width=target). Larger is better.
  return Math.min(target / width, 1 / height);
}

function swapLayoutDirection(direction: DiagramDirection): DiagramDirection {
  const rankdir = toRankdir(direction);
  if (rankdir === "LR" || rankdir === "RL") {
    return "TD";
  }
  return "LR";
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

interface LayoutQuality {
  fit: number;
  gapPx: number;
  gapScaled: number;
  coverage: number;
  edgeLenScaled: number;
  boundsWidth: number;
  boundsHeight: number;
  primary: number;
}

interface TopLevelBlock {
  id: string;
  kind: "subgraph" | "ungrouped";
  x: number;
  y: number;
  width: number;
  height: number;
  nodeIds: string[];
}

function coverageForAspect(aspect: number, targetAspectRatio: number): number {
  const target = clampAspect(targetAspectRatio);
  if (!Number.isFinite(aspect) || aspect <= 0) {
    return 0;
  }

  const ratio = aspect / target;
  if (!Number.isFinite(ratio) || ratio <= 0) {
    return 0;
  }

  return Math.min(ratio, 1 / ratio);
}

function nodeAnchorPoint(node: { x: number; y: number; width: number; height: number }, side?: EdgeSide): Point {
  const margin = Math.min(12, Math.max(4, Math.min(node.width, node.height) * 0.22));

  if (side === "T") {
    return {
      x: clamp(node.x + node.width / 2, node.x + margin, node.x + node.width - margin),
      y: node.y,
    };
  }
  if (side === "B") {
    return {
      x: clamp(node.x + node.width / 2, node.x + margin, node.x + node.width - margin),
      y: node.y + node.height,
    };
  }
  if (side === "L") {
    return {
      x: node.x,
      y: clamp(node.y + node.height / 2, node.y + margin, node.y + node.height - margin),
    };
  }
  if (side === "R") {
    return {
      x: node.x + node.width,
      y: clamp(node.y + node.height / 2, node.y + margin, node.y + node.height - margin),
    };
  }

  return {
    x: node.x + node.width / 2,
    y: node.y + node.height / 2,
  };
}

function buildSubgraphChildren(ir: DiagramIr): Map<string, string[]> {
  const children = new Map<string, string[]>();
  for (const subgraph of ir.subgraphs) {
    if (!subgraph.parentId) {
      continue;
    }

    const list = children.get(subgraph.parentId) ?? [];
    list.push(subgraph.id);
    children.set(subgraph.parentId, list);
  }
  return children;
}

function buildSubgraphAnchorNodes(ir: DiagramIr, childrenByParent: Map<string, string[]>): Map<string, string> {
  const subgraphIds = new Set(ir.subgraphs.map((subgraph) => subgraph.id));
  const anchorBySubgraph = new Map<string, string>();

  for (const node of ir.nodes) {
    if (!node.subgraphId) {
      continue;
    }
    if (subgraphIds.has(node.id)) {
      continue;
    }
    if (!anchorBySubgraph.has(node.subgraphId)) {
      anchorBySubgraph.set(node.subgraphId, node.id);
    }
  }

  const resolveAnchor = (subgraphId: string, visited: Set<string> = new Set()): string | undefined => {
    if (anchorBySubgraph.has(subgraphId)) {
      return anchorBySubgraph.get(subgraphId);
    }

    if (visited.has(subgraphId)) {
      return undefined;
    }
    visited.add(subgraphId);

    const children = childrenByParent.get(subgraphId) ?? [];
    for (const childId of children) {
      const anchor = resolveAnchor(childId, visited);
      if (anchor) {
        anchorBySubgraph.set(subgraphId, anchor);
        return anchor;
      }
    }

    return undefined;
  };

  for (const subgraph of ir.subgraphs) {
    resolveAnchor(subgraph.id);
  }

  return anchorBySubgraph;
}

function rerouteEdgesStraight(ir: DiagramIr): void {
  const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));
  const subgraphIds = new Set(ir.subgraphs.map((subgraph) => subgraph.id));
  const childrenByParent = buildSubgraphChildren(ir);
  const anchorBySubgraph = buildSubgraphAnchorNodes(ir, childrenByParent);

  for (const edge of ir.edges) {
    let fromId = edge.from;
    let toId = edge.to;

    if (subgraphIds.has(fromId)) {
      fromId = anchorBySubgraph.get(fromId) ?? fromId;
    }
    if (subgraphIds.has(toId)) {
      toId = anchorBySubgraph.get(toId) ?? toId;
    }

    const fromNode = nodeById.get(fromId);
    const toNode = nodeById.get(toId);
    if (!fromNode || !toNode) {
      edge.points = [];
      edge.labelPosition = undefined;
      continue;
    }

    const p0 = nodeAnchorPoint(fromNode, edge.style.startSide);
    const p1 = nodeAnchorPoint(toNode, edge.style.endSide);
    edge.points = [p0, p1];
    edge.labelPosition = edge.label ? fallbackLabelPosition(edge.points) : undefined;
  }

  recomputeBounds(ir);
}

function rectDistance(
  a: { x: number; y: number; width: number; height: number },
  b: { x: number; y: number; width: number; height: number },
): number {
  const ax0 = a.x;
  const ay0 = a.y;
  const ax1 = a.x + a.width;
  const ay1 = a.y + a.height;

  const bx0 = b.x;
  const by0 = b.y;
  const bx1 = b.x + b.width;
  const by1 = b.y + b.height;

  let dx = 0;
  if (ax1 < bx0) {
    dx = bx0 - ax1;
  } else if (bx1 < ax0) {
    dx = ax0 - bx1;
  }

  let dy = 0;
  if (ay1 < by0) {
    dy = by0 - ay1;
  } else if (by1 < ay0) {
    dy = ay0 - by1;
  }

  return Math.hypot(dx, dy);
}

function minNodeGapPx(ir: DiagramIr): number {
  const nodes = ir.nodes;
  if (nodes.length < 2) {
    return 0;
  }

  let best = Number.POSITIVE_INFINITY;
  for (let i = 0; i < nodes.length; i += 1) {
    const a = nodes[i];
    for (let j = i + 1; j < nodes.length; j += 1) {
      const b = nodes[j];
      const d = rectDistance(a, b);
      if (d < best) {
        best = d;
        if (best <= 0) {
          return 0;
        }
      }
    }
  }

  return Number.isFinite(best) ? best : 0;
}

function totalEdgeLengthPx(ir: DiagramIr): number {
  let total = 0;
  for (const edge of ir.edges) {
    if (edge.points.length >= 2) {
      const p0 = edge.points[0];
      const p1 = edge.points[edge.points.length - 1];
      total += Math.hypot(p1.x - p0.x, p1.y - p0.y);
    }
  }
  return total;
}

function computeLayoutQuality(ir: DiagramIr, targetAspectRatio: number): LayoutQuality {
  const aspect = diagramAspect(ir);
  const fit = normalizedFitScale(ir, targetAspectRatio);
  const gapPx = minNodeGapPx(ir);
  const gapScaled = gapPx * fit;
  const coverage = coverageForAspect(aspect, targetAspectRatio);
  const edgeLenScaled = totalEdgeLengthPx(ir) * fit;

  // Priority (user request):
  //  1) make objects big AND keep them apart (fit + spacing)
  //  2) reduce whitespace (coverage)
  //  3) minimize line length (edgeLenScaled) [weak tie-breaker only]
  const desiredGapScaled = 0.016;
  const gapRatioRaw = desiredGapScaled > 0 ? gapScaled / desiredGapScaled : 0;
  const gapRatio = clamp(gapRatioRaw, 0.3, 1.6);
  // Penalize tight layouts, but keep the "make objects bigger" goal dominant.
  const gapPenalty = gapRatio < 1 ? gapRatio ** 1.5 : 1 + (gapRatio - 1) * 0.1;
  const primary = fit * gapPenalty;

  return {
    fit,
    gapPx,
    gapScaled,
    coverage,
    edgeLenScaled,
    boundsWidth: Math.max(1, ir.bounds.width),
    boundsHeight: Math.max(1, ir.bounds.height),
    primary,
  };
}

function isBetterQuality(candidate: LayoutQuality, best: LayoutQuality): boolean {
  if (candidate.primary > best.primary * 1.01) {
    return true;
  }
  if (best.primary > candidate.primary * 1.01) {
    return false;
  }

  if (candidate.coverage > best.coverage * 1.01) {
    return true;
  }
  if (best.coverage > candidate.coverage * 1.01) {
    return false;
  }

  // Prefer bigger spacing when primary and coverage are similar.
  if (candidate.gapScaled > best.gapScaled * 1.05) {
    return true;
  }
  if (best.gapScaled > candidate.gapScaled * 1.05) {
    return false;
  }

  // Keep line-shortening as a very weak preference, and never trade spacing away for it.
  if (candidate.gapScaled >= best.gapScaled * 0.98) {
    if (candidate.edgeLenScaled < best.edgeLenScaled * 0.85) {
      return true;
    }
    if (best.edgeLenScaled < candidate.edgeLenScaled * 0.85) {
      return false;
    }
  }

  if (candidate.fit > best.fit) {
    return true;
  }

  const areaCandidate = candidate.boundsWidth * candidate.boundsHeight;
  const areaBest = best.boundsWidth * best.boundsHeight;
  return areaCandidate < areaBest;
}

function collectTopLevelBlocks(ir: DiagramIr): TopLevelBlock[] {
  const blocks: TopLevelBlock[] = [];
  const childrenByParent = buildSubgraphChildren(ir);

  const rootSubgraphs = ir.subgraphs.filter((subgraph) => !subgraph.parentId);
  for (const root of rootSubgraphs) {
    const descendant = new Set<string>();
    const stack: string[] = [root.id];
    while (stack.length > 0) {
      const id = stack.pop();
      if (!id || descendant.has(id)) {
        continue;
      }
      descendant.add(id);
      const children = childrenByParent.get(id) ?? [];
      for (const childId of children) {
        stack.push(childId);
      }
    }

    const nodeIds = ir.nodes.filter((node) => node.subgraphId && descendant.has(node.subgraphId)).map((node) => node.id);
    blocks.push({
      id: root.id,
      kind: "subgraph",
      x: root.x,
      y: root.y,
      width: root.width,
      height: root.height,
      nodeIds,
    });
  }

  const ungrouped = ir.nodes.filter((node) => !node.subgraphId);
  if (ungrouped.length > 0) {
    let minX = Number.POSITIVE_INFINITY;
    let minY = Number.POSITIVE_INFINITY;
    let maxX = Number.NEGATIVE_INFINITY;
    let maxY = Number.NEGATIVE_INFINITY;

    for (const node of ungrouped) {
      minX = Math.min(minX, node.x);
      minY = Math.min(minY, node.y);
      maxX = Math.max(maxX, node.x + node.width);
      maxY = Math.max(maxY, node.y + node.height);
    }

    blocks.push({
      id: "__ungrouped__",
      kind: "ungrouped",
      x: minX,
      y: minY,
      width: maxX - minX,
      height: maxY - minY,
      nodeIds: ungrouped.map((node) => node.id),
    });
  }

  const rankdir = toRankdir(ir.meta.direction);
  const axis = rankdir === "LR" || rankdir === "RL" ? "x" : "y";
  blocks.sort((a, b) => (axis === "x" ? a.x - b.x : a.y - b.y));
  if (rankdir === "RL" || rankdir === "BT") {
    blocks.reverse();
  }

  return blocks;
}

function applyPackedRows(ir: DiagramIr, rows: TopLevelBlock[][], gap: number): void {
  const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));
  let cursorY = 0;

  for (const row of rows) {
    let cursorX = 0;
    const rowHeight = Math.max(...row.map((block) => block.height));

    for (const block of row) {
      const dx = cursorX - block.x;
      const dy = cursorY - block.y;
      for (const nodeId of block.nodeIds) {
        const node = nodeById.get(nodeId);
        if (!node) {
          continue;
        }
        node.x += dx;
        node.y += dy;
      }

      cursorX += block.width + gap;
    }

    cursorY += rowHeight + gap;
  }

  recomputeSubgraphBounds(ir);
  rerouteEdgesStraight(ir);
}

function wrapTopLevelBlocks(ir: DiagramIr, targetAspectRatio: number): void {
  const blocks = collectTopLevelBlocks(ir);
  if (blocks.length <= 1) {
    return;
  }

  // Gap between top-level blocks (subgraphs / ungrouped section).
  const gap = clamp((ir.config.layout.nodesep + ir.config.layout.ranksep) / 2, 40, 160);

  rerouteEdgesStraight(ir);
  const baseState = captureLayoutState(ir);
  let bestState = baseState;
  let bestQuality = computeLayoutQuality(ir, targetAspectRatio);

  const n = blocks.length;
  const maxRows = Math.min(5, n);

  const tryRows = (rows: TopLevelBlock[][]): void => {
    restoreLayoutState(ir, baseState);
    applyPackedRows(ir, rows, gap);
    const q = computeLayoutQuality(ir, targetAspectRatio);
    if (isBetterQuality(q, bestQuality)) {
      bestQuality = q;
      bestState = captureLayoutState(ir);
    }
  };

  if (n <= 10) {
    const breaks = n - 1;
    const maxMask = 1 << breaks;
    for (let mask = 0; mask < maxMask; mask += 1) {
      const rows: TopLevelBlock[][] = [];
      let current: TopLevelBlock[] = [blocks[0]];
      for (let i = 0; i < breaks; i += 1) {
        if (mask & (1 << i)) {
          rows.push(current);
          current = [];
        }
        current.push(blocks[i + 1]);
      }
      if (current.length > 0) {
        rows.push(current);
      }

      if (rows.length > maxRows) {
        continue;
      }

      tryRows(rows);
    }
  } else {
    // Greedy fallback for very large diagrams: wrap into 3-5 rows.
    const target = clampAspect(targetAspectRatio);
    const approxRows = clamp(Math.round(Math.sqrt((n * 1.2) / Math.max(0.5, target))), 3, maxRows);
    const perRow = Math.max(1, Math.ceil(n / approxRows));
    const rows: TopLevelBlock[][] = [];
    for (let i = 0; i < n; i += perRow) {
      rows.push(blocks.slice(i, i + perRow));
    }
    tryRows(rows);
  }

  restoreLayoutState(ir, bestState);
}

function uniqueNumbers(values: number[]): number[] {
  return [...new Set(values.map((value) => Math.round(value * 1000) / 1000))].sort((a, b) => a - b);
}

function spacingCandidates(base: number, extras: number[], min: number, max: number): number[] {
  const multipliers = [0.85, 1.0, 1.25, 1.6, 2.0];
  const out: number[] = [];
  for (const mult of multipliers) {
    out.push(clamp(base * mult, min, max));
  }
  for (const extra of extras) {
    out.push(clamp(extra, min, max));
  }
  return uniqueNumbers(out);
}

function optimizeLayoutForSlide(ir: DiagramIr, targetAspectRatio: number): void {
  const baseDirection = ir.meta.direction;
  const baseNodesep = ir.config.layout.nodesep;
  const baseRanksep = ir.config.layout.ranksep;

  const swappedDirection = swapLayoutDirection(baseDirection);
  const directionCandidates: DiagramDirection[] =
    swappedDirection !== baseDirection ? [baseDirection, swappedDirection] : [baseDirection];

  const nodesepCandidates = spacingCandidates(baseNodesep, [60, 90, 130, 180, 240], 22, 360);
  const ranksepCandidates = spacingCandidates(baseRanksep, [25, 40, 60, 90, 140], 18, 380);

  let bestState: LayoutState | undefined;
  let bestQuality: LayoutQuality | undefined;
  let bestDirection = baseDirection;
  let bestNodesep = baseNodesep;
  let bestRanksep = baseRanksep;

  for (const direction of directionCandidates) {
    ir.meta.direction = direction;

    for (const nodesep of nodesepCandidates) {
      for (const ranksep of ranksepCandidates) {
        ir.config.layout.nodesep = nodesep;
        ir.config.layout.ranksep = ranksep;

        applyLayout(ir);
        rerouteEdgesStraight(ir);
        wrapTopLevelBlocks(ir, targetAspectRatio);
        const quality = computeLayoutQuality(ir, targetAspectRatio);

        if (!bestQuality || isBetterQuality(quality, bestQuality)) {
          bestQuality = quality;
          bestState = captureLayoutState(ir);
          bestDirection = direction;
          bestNodesep = nodesep;
          bestRanksep = ranksep;
        }
      }
    }
  }

  if (bestState && bestQuality) {
    restoreLayoutState(ir, bestState);
    ir.meta.direction = bestDirection;
    ir.config.layout.nodesep = bestNodesep;
    ir.config.layout.ranksep = bestRanksep;
    return;
  }

  ir.meta.direction = baseDirection;
  ir.config.layout.nodesep = baseNodesep;
  ir.config.layout.ranksep = baseRanksep;
  applyLayout(ir);
}

export function layoutDiagram(ir: DiagramIr, options: LayoutOptions = {}): void {
  if (options.targetAspectRatio) {
    optimizeLayoutForSlide(ir, options.targetAspectRatio);
    return;
  }

  applyLayout(ir);
}
