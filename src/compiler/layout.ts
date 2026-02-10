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

interface HardConstraintStats {
  textOverflow: number;
  nodeOverlap: number;
  nodeGap: number;
  labelNodeOverlap: number;
  labelLabelOverlap: number;
  nodeOutsideSubgraph: number;
  subgraphOverlap: number;
  edgeNodeGap: number;
  totalViolations: number;
  penalty: number;
  feasible: boolean;
}

interface ReadabilityStats {
  edgeCrossings: number;
  edgeThroughNode: number;
  parallelOverlapScore: number;
}

interface LayoutQuality {
  hard: HardConstraintStats;
  readability: ReadabilityStats;
  fit: number;
  objectSizeScore: number;
  spacingScore: number;
  nodeAreaRatio: number;
  coverage: number;
  aspectError: number;
  minGapPx: number;
  minGapScaled: number;
  avgGapPx: number;
  avgGapScaled: number;
  densityPenalty: number;
  edgeLenScaled: number;
  bends: number;
  boundsWidth: number;
  boundsHeight: number;
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

const HARD_MIN_FONT_PT = 8;
const HARD_MIN_NODE_GAP_PX = 14;
const HARD_MIN_LABEL_GAP_PX = 6;
const HARD_MIN_EDGE_NODE_GAP_PX = 7;
const HARD_MIN_SUBGRAPH_GAP_PX = 10;

interface Rect {
  x: number;
  y: number;
  width: number;
  height: number;
}

interface EdgeSegment {
  edgeId: string;
  from: string;
  to: string;
  p0: Point;
  p1: Point;
}

interface EdgeLabelBox extends Rect {
  edgeId: string;
}

function effectiveTextUnits(text: string): number {
  let units = 0;
  for (const ch of text) {
    if (/\s/u.test(ch)) {
      units += 0.35;
      continue;
    }
    if (/[\u3000-\u9FFF\uF900-\uFAFF]/u.test(ch)) {
      units += 1.75;
      continue;
    }
    units += 1.0;
  }
  return units;
}

function estimateNodeTextMinSize(node: DiagramIr["nodes"][number]): { minWidth: number; minHeight: number } {
  if (node.isJunction) {
    return {
      minWidth: 1,
      minHeight: 1,
    };
  }

  const fontSize = Math.max(8, node.style.fontSize || 14);
  const lines = (node.label || node.id)
    .replace(/\r\n/g, "\n")
    .split("\n")
    .filter((line) => line.length > 0);
  const safeLines = lines.length > 0 ? lines : [node.id];
  const longest = Math.max(...safeLines.map((line) => effectiveTextUnits(line)), effectiveTextUnits(node.id));
  const minWidth = (longest * fontSize * 0.70 + 48) * 1.08;
  const iconHeadroom = node.icon ? Math.max(20, fontSize * 1.3) : 0;
  const minHeight = (safeLines.length * (fontSize * 1.30) + 34 + iconHeadroom) * 1.08;
  return {
    minWidth,
    minHeight,
  };
}

function rectArea(rect: Rect): number {
  return Math.max(0, rect.width) * Math.max(0, rect.height);
}

function intersectsRect(a: Rect, b: Rect): boolean {
  return !(a.x + a.width <= b.x || b.x + b.width <= a.x || a.y + a.height <= b.y || b.y + b.height <= a.y);
}

function rectIntersectionArea(a: Rect, b: Rect): number {
  const left = Math.max(a.x, b.x);
  const top = Math.max(a.y, b.y);
  const right = Math.min(a.x + a.width, b.x + b.width);
  const bottom = Math.min(a.y + a.height, b.y + b.height);
  if (right <= left || bottom <= top) {
    return 0;
  }
  return (right - left) * (bottom - top);
}

function expandedRect(rect: Rect, margin: number): Rect {
  return {
    x: rect.x - margin,
    y: rect.y - margin,
    width: rect.width + margin * 2,
    height: rect.height + margin * 2,
  };
}

function rectContainsRect(outer: Rect, inner: Rect, margin = 0): boolean {
  return (
    inner.x >= outer.x + margin &&
    inner.y >= outer.y + margin &&
    inner.x + inner.width <= outer.x + outer.width - margin &&
    inner.y + inner.height <= outer.y + outer.height - margin
  );
}

function isInsideRect(point: Point, rect: Rect): boolean {
  return point.x > rect.x && point.x < rect.x + rect.width && point.y > rect.y && point.y < rect.y + rect.height;
}

function orientation(a: Point, b: Point, c: Point): number {
  return (b.x - a.x) * (c.y - a.y) - (b.y - a.y) * (c.x - a.x);
}

function onSegment(a: Point, b: Point, p: Point): boolean {
  return p.x >= Math.min(a.x, b.x) - 1e-6 && p.x <= Math.max(a.x, b.x) + 1e-6 && p.y >= Math.min(a.y, b.y) - 1e-6 && p.y <= Math.max(a.y, b.y) + 1e-6;
}

function segmentIntersectionStrict(a0: Point, a1: Point, b0: Point, b1: Point): Point | undefined {
  const x1 = a0.x;
  const y1 = a0.y;
  const x2 = a1.x;
  const y2 = a1.y;
  const x3 = b0.x;
  const y3 = b0.y;
  const x4 = b1.x;
  const y4 = b1.y;
  const denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4);
  if (Math.abs(denom) < 1e-9) {
    return undefined;
  }

  const t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / denom;
  const u = ((x1 - x3) * (y1 - y2) - (y1 - y3) * (x1 - x2)) / denom;
  if (!(t > 0.02 && t < 0.98 && u > 0.02 && u < 0.98)) {
    return undefined;
  }

  return {
    x: x1 + t * (x2 - x1),
    y: y1 + t * (y2 - y1),
  };
}

function segmentLength(a: Point, b: Point): number {
  return Math.hypot(b.x - a.x, b.y - a.y);
}

function dot(ax: number, ay: number, bx: number, by: number): number {
  return ax * bx + ay * by;
}

function distancePointToSegment(point: Point, a: Point, b: Point): number {
  const vx = b.x - a.x;
  const vy = b.y - a.y;
  const len2 = vx * vx + vy * vy;
  if (len2 <= 1e-9) {
    return Math.hypot(point.x - a.x, point.y - a.y);
  }
  const t = clamp(((point.x - a.x) * vx + (point.y - a.y) * vy) / len2, 0, 1);
  const px = a.x + vx * t;
  const py = a.y + vy * t;
  return Math.hypot(point.x - px, point.y - py);
}

function distancePointToRect(point: Point, rect: Rect): number {
  const dx = point.x < rect.x ? rect.x - point.x : point.x > rect.x + rect.width ? point.x - (rect.x + rect.width) : 0;
  const dy = point.y < rect.y ? rect.y - point.y : point.y > rect.y + rect.height ? point.y - (rect.y + rect.height) : 0;
  return Math.hypot(dx, dy);
}

function distanceSegmentToSegment(a0: Point, a1: Point, b0: Point, b1: Point): number {
  const strict = segmentIntersectionStrict(a0, a1, b0, b1);
  if (strict) {
    return 0;
  }
  const distances = [
    distancePointToSegment(a0, b0, b1),
    distancePointToSegment(a1, b0, b1),
    distancePointToSegment(b0, a0, a1),
    distancePointToSegment(b1, a0, a1),
  ];
  return Math.min(...distances);
}

function distanceSegmentToRect(a: Point, b: Point, rect: Rect): number {
  if (isInsideRect(a, rect) || isInsideRect(b, rect)) {
    return 0;
  }

  const r0: Point = { x: rect.x, y: rect.y };
  const r1: Point = { x: rect.x + rect.width, y: rect.y };
  const r2: Point = { x: rect.x + rect.width, y: rect.y + rect.height };
  const r3: Point = { x: rect.x, y: rect.y + rect.height };

  if (
    segmentIntersectionStrict(a, b, r0, r1) ||
    segmentIntersectionStrict(a, b, r1, r2) ||
    segmentIntersectionStrict(a, b, r2, r3) ||
    segmentIntersectionStrict(a, b, r3, r0)
  ) {
    return 0;
  }

  const cornerDistances = [distancePointToSegment(r0, a, b), distancePointToSegment(r1, a, b), distancePointToSegment(r2, a, b), distancePointToSegment(r3, a, b)];
  const endpointDistances = [distancePointToRect(a, rect), distancePointToRect(b, rect)];
  return Math.min(...cornerDistances, ...endpointDistances);
}

function segmentPassesThroughRect(a: Point, b: Point, rect: Rect): boolean {
  const inner = expandedRect(rect, -2);
  if (inner.width <= 0 || inner.height <= 0) {
    return false;
  }
  return distanceSegmentToRect(a, b, inner) <= 0.001;
}

function estimateEdgeLabelRect(edge: DiagramIr["edges"][number]): EdgeLabelBox | undefined {
  if (!edge.label || edge.label.trim().length === 0) {
    return undefined;
  }

  const center = edge.labelPosition ?? fallbackLabelPosition(edge.points);
  if (!center) {
    return undefined;
  }

  const fontSize = Math.max(8, edge.style.fontSize || 11);
  const text = edge.label.replace(/\r\n/g, "\n");
  const lines = text.split("\n");
  const longestUnits = Math.max(...lines.map((line) => effectiveTextUnits(line)), 1);
  const width = clamp(longestUnits * fontSize * 0.72 + 20, 52, 320);
  const height = clamp(lines.length * fontSize * 1.28 + 8, 18, 140);

  return {
    edgeId: edge.id,
    x: center.x - width / 2,
    y: center.y - height / 2,
    width,
    height,
  };
}

function areRelatedSubgraphs(aId: string, bId: string, parentById: Map<string, string | undefined>): boolean {
  const isAncestor = (candidate: string, child: string): boolean => {
    let current: string | undefined = child;
    const visited = new Set<string>();
    while (current) {
      if (current === candidate) {
        return true;
      }
      if (visited.has(current)) {
        break;
      }
      visited.add(current);
      current = parentById.get(current);
    }
    return false;
  };

  return isAncestor(aId, bId) || isAncestor(bId, aId);
}

function nodeAnchorPoint(
  node: { x: number; y: number; width: number; height: number },
  side?: EdgeSide,
  alongOffset = 0,
): Point {
  const margin = Math.min(12, Math.max(4, Math.min(node.width, node.height) * 0.22));

  if (side === "T") {
    return {
      x: clamp(node.x + node.width / 2 + alongOffset, node.x + margin, node.x + node.width - margin),
      y: node.y,
    };
  }
  if (side === "B") {
    return {
      x: clamp(node.x + node.width / 2 + alongOffset, node.x + margin, node.x + node.width - margin),
      y: node.y + node.height,
    };
  }
  if (side === "L") {
    return {
      x: node.x,
      y: clamp(node.y + node.height / 2 + alongOffset, node.y + margin, node.y + node.height - margin),
    };
  }
  if (side === "R") {
    return {
      x: node.x + node.width,
      y: clamp(node.y + node.height / 2 + alongOffset, node.y + margin, node.y + node.height - margin),
    };
  }

  return {
    x: node.x + node.width / 2,
    y: node.y + node.height / 2,
  };
}

function sideNormal(side: EdgeSide): Point {
  if (side === "T") {
    return { x: 0, y: -1 };
  }
  if (side === "B") {
    return { x: 0, y: 1 };
  }
  if (side === "L") {
    return { x: -1, y: 0 };
  }
  return { x: 1, y: 0 };
}

function connectionCost(src: Rect, dst: Rect, srcSide: EdgeSide, dstSide: EdgeSide): number {
  const s = nodeAnchorPoint(src, srcSide);
  const t = nodeAnchorPoint(dst, dstSide);
  const vx = t.x - s.x;
  const vy = t.y - s.y;
  const dist = Math.hypot(vx, vy);
  if (dist < 1e-6) {
    return 0;
  }

  const srcN = sideNormal(srcSide);
  const dstN = sideNormal(dstSide);
  const srcDot = dot(srcN.x, srcN.y, vx, vy);
  const dstDot = dot(dstN.x, dstN.y, -vx, -vy);

  let penalty = 0;
  if (srcDot <= 0) {
    penalty += 42 + dist * 0.42;
  }
  if (dstDot <= 0) {
    penalty += 42 + dist * 0.42;
  }

  const srcVertical = srcSide === "T" || srcSide === "B";
  const dstVertical = dstSide === "T" || dstSide === "B";
  if (srcVertical !== dstVertical) {
    penalty += 5;
  }

  return dist + penalty;
}

function chooseConnectionSides(src: Rect, dst: Rect, avoidExact?: [EdgeSide, EdgeSide]): [EdgeSide, EdgeSide] {
  const sides: EdgeSide[] = ["T", "L", "B", "R"];
  let best: [EdgeSide, EdgeSide] = ["R", "L"];
  let bestCost = Number.POSITIVE_INFINITY;

  for (const s of sides) {
    for (const t of sides) {
      let cost = connectionCost(src, dst, s, t);
      if (avoidExact && avoidExact[0] === s && avoidExact[1] === t) {
        cost += 2200;
      }
      if (cost < bestCost) {
        bestCost = cost;
        best = [s, t];
      }
    }
  }

  return best;
}

function chooseConnectionSidesWithHints(
  src: Rect,
  dst: Rect,
  hintedSrcSide: EdgeSide | undefined,
  hintedDstSide: EdgeSide | undefined,
  strictHints: boolean,
): [EdgeSide, EdgeSide] {
  const sides: EdgeSide[] = ["T", "L", "B", "R"];
  if (strictHints) {
    if (hintedSrcSide && hintedDstSide) {
      return [hintedSrcSide, hintedDstSide];
    }
    if (hintedSrcSide) {
      const bestDst = [...sides].sort((a, b) => connectionCost(src, dst, hintedSrcSide, a) - connectionCost(src, dst, hintedSrcSide, b))[0];
      return [hintedSrcSide, bestDst];
    }
    if (hintedDstSide) {
      const bestSrc = [...sides].sort((a, b) => connectionCost(src, dst, a, hintedDstSide) - connectionCost(src, dst, b, hintedDstSide))[0];
      return [bestSrc, hintedDstSide];
    }
  }

  const [autoSrc, autoDst] = chooseConnectionSides(src, dst);
  const autoCost = connectionCost(src, dst, autoSrc, autoDst);
  const tolerance = Math.max(10, autoCost * 0.2);

  if (hintedSrcSide && hintedDstSide) {
    const hintedCost = connectionCost(src, dst, hintedSrcSide, hintedDstSide);
    if (hintedCost <= autoCost + tolerance) {
      return [hintedSrcSide, hintedDstSide];
    }
    return [autoSrc, autoDst];
  }

  if (hintedSrcSide) {
    let bestDst = sides[0];
    let bestCost = Number.POSITIVE_INFINITY;
    for (const s of sides) {
      const cost = connectionCost(src, dst, hintedSrcSide, s);
      if (cost < bestCost) {
        bestCost = cost;
        bestDst = s;
      }
    }
    if (bestCost <= autoCost + tolerance) {
      return [hintedSrcSide, bestDst];
    }
    return [autoSrc, autoDst];
  }

  if (hintedDstSide) {
    let bestSrc = sides[0];
    let bestCost = Number.POSITIVE_INFINITY;
    for (const s of sides) {
      const cost = connectionCost(src, dst, s, hintedDstSide);
      if (cost < bestCost) {
        bestCost = cost;
        bestSrc = s;
      }
    }
    if (bestCost <= autoCost + tolerance) {
      return [bestSrc, hintedDstSide];
    }
    return [autoSrc, autoDst];
  }

  return [autoSrc, autoDst];
}

function sideSortAxis(side: EdgeSide, fromNode: Rect, toNode: Rect, sourceSide: boolean): number {
  const src = sourceSide ? fromNode : toNode;
  const dst = sourceSide ? toNode : fromNode;
  const dstCx = dst.x + dst.width / 2;
  const dstCy = dst.y + dst.height / 2;
  const srcCx = src.x + src.width / 2;
  const srcCy = src.y + src.height / 2;

  if (side === "T" || side === "B") {
    return dstCx + (dstCy - srcCy) * 0.03;
  }
  return dstCy + (dstCx - srcCx) * 0.03;
}

function laneOffsets(count: number, spanPx: number): number[] {
  if (count <= 1) {
    return [0];
  }

  const usable = Math.max(0, spanPx - 24);
  const step = usable <= 1e-6 ? 0 : clamp(usable / Math.max(1, count), 6, 15);
  const center = (count - 1) / 2;
  const out: number[] = [];
  for (let i = 0; i < count; i += 1) {
    out.push((i - center) * step);
  }
  return out;
}

function chooseSelfLoopPoints(
  node: Rect,
  viewport: { minX: number; minY: number; maxX: number; maxY: number },
): { points: Point[]; labelPosition: Point } {
  const leftSpace = node.x - viewport.minX;
  const rightSpace = viewport.maxX - (node.x + node.width);
  const topSpace = node.y - viewport.minY;
  const bottomSpace = viewport.maxY - (node.y + node.height);
  const candidates: Array<["tr" | "br" | "bl" | "tl", number]> = [
    ["tr", rightSpace + topSpace],
    ["br", rightSpace + bottomSpace],
    ["bl", leftSpace + bottomSpace],
    ["tl", leftSpace + topSpace],
  ];
  const mode = candidates.sort((a, b) => b[1] - a[1])[0][0];
  const loopX = clamp(Math.max(node.width * 0.72, 26), 26, 120);
  const loopY = clamp(Math.max(node.height * 1.05, 26), 26, 120);

  if (mode === "tr") {
    return {
      points: [
        { x: node.x + node.width, y: node.y + node.height * 0.58 },
        { x: node.x + node.width + loopX, y: node.y + node.height * 0.58 },
        { x: node.x + node.width + loopX, y: node.y - loopY },
        { x: node.x + node.width * 0.58, y: node.y - loopY },
        { x: node.x + node.width * 0.58, y: node.y },
      ],
      labelPosition: { x: node.x + node.width + loopX * 0.52, y: node.y - loopY * 0.56 },
    };
  }

  if (mode === "br") {
    return {
      points: [
        { x: node.x + node.width, y: node.y + node.height * 0.42 },
        { x: node.x + node.width + loopX, y: node.y + node.height * 0.42 },
        { x: node.x + node.width + loopX, y: node.y + node.height + loopY },
        { x: node.x + node.width * 0.58, y: node.y + node.height + loopY },
        { x: node.x + node.width * 0.58, y: node.y + node.height },
      ],
      labelPosition: { x: node.x + node.width + loopX * 0.52, y: node.y + node.height + loopY * 0.56 },
    };
  }

  if (mode === "bl") {
    return {
      points: [
        { x: node.x, y: node.y + node.height * 0.42 },
        { x: node.x - loopX, y: node.y + node.height * 0.42 },
        { x: node.x - loopX, y: node.y + node.height + loopY },
        { x: node.x + node.width * 0.42, y: node.y + node.height + loopY },
        { x: node.x + node.width * 0.42, y: node.y + node.height },
      ],
      labelPosition: { x: node.x - loopX * 0.52, y: node.y + node.height + loopY * 0.56 },
    };
  }

  return {
    points: [
      { x: node.x, y: node.y + node.height * 0.58 },
      { x: node.x - loopX, y: node.y + node.height * 0.58 },
      { x: node.x - loopX, y: node.y - loopY },
      { x: node.x + node.width * 0.42, y: node.y - loopY },
      { x: node.x + node.width * 0.42, y: node.y },
    ],
    labelPosition: { x: node.x - loopX * 0.52, y: node.y - loopY * 0.56 },
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
  const subgraphById = new Map(ir.subgraphs.map((subgraph) => [subgraph.id, subgraph]));
  const subgraphIds = new Set(ir.subgraphs.map((subgraph) => subgraph.id));
  const childrenByParent = buildSubgraphChildren(ir);
  const anchorBySubgraph = buildSubgraphAnchorNodes(ir, childrenByParent);
  const viewport = {
    minX: Number.isFinite(ir.bounds.minX) ? ir.bounds.minX - 220 : -220,
    minY: Number.isFinite(ir.bounds.minY) ? ir.bounds.minY - 220 : -220,
    maxX: Number.isFinite(ir.bounds.maxX) ? ir.bounds.maxX + 220 : 220,
    maxY: Number.isFinite(ir.bounds.maxY) ? ir.bounds.maxY + 220 : 220,
  };

  interface PendingRoute {
    edge: DiagramIr["edges"][number];
    srcAnchorId: string;
    dstAnchorId: string;
    srcAnchor: Rect;
    dstAnchor: Rect;
    srcSide: EdgeSide;
    dstSide: EdgeSide;
    srcOffset: number;
    dstOffset: number;
    selfLoop: boolean;
    selfLoopNode?: Rect;
  }

  const pendingRoutes: PendingRoute[] = [];

  const resolveNodeId = (rawId: string): string => {
    if (!subgraphIds.has(rawId)) {
      return rawId;
    }
    return anchorBySubgraph.get(rawId) ?? rawId;
  };

  for (const edge of ir.edges) {
    const fromNodeId = resolveNodeId(edge.from);
    const toNodeId = resolveNodeId(edge.to);
    const fromNode = nodeById.get(fromNodeId);
    const toNode = nodeById.get(toNodeId);
    if (!fromNode || !toNode) {
      edge.points = [];
      edge.labelPosition = undefined;
      continue;
    }

    if (fromNodeId === toNodeId && !edge.style.startViaGroup && !edge.style.endViaGroup) {
      pendingRoutes.push({
        edge,
        srcAnchorId: fromNodeId,
        dstAnchorId: toNodeId,
        srcAnchor: fromNode,
        dstAnchor: toNode,
        srcSide: "R",
        dstSide: "T",
        srcOffset: 0,
        dstOffset: 0,
        selfLoop: true,
        selfLoopNode: fromNode,
      });
      continue;
    }

    let srcAnchorId = fromNodeId;
    let dstAnchorId = toNodeId;
    let srcAnchor: Rect = fromNode;
    let dstAnchor: Rect = toNode;

    if (edge.style.startViaGroup && fromNode.subgraphId) {
      const group = subgraphById.get(fromNode.subgraphId);
      if (group) {
        srcAnchorId = group.id;
        srcAnchor = group;
      }
    }

    if (edge.style.endViaGroup && toNode.subgraphId) {
      const group = subgraphById.get(toNode.subgraphId);
      if (group) {
        dstAnchorId = group.id;
        dstAnchor = group;
      }
    }

    const strictHints = Boolean(nodeById.get(srcAnchorId)?.isJunction || nodeById.get(dstAnchorId)?.isJunction);
    const [srcSide, dstSide] = chooseConnectionSidesWithHints(
      srcAnchor,
      dstAnchor,
      edge.style.startSide,
      edge.style.endSide,
      strictHints,
    );

    pendingRoutes.push({
      edge,
      srcAnchorId,
      dstAnchorId,
      srcAnchor,
      dstAnchor,
      srcSide,
      dstSide,
      srcOffset: 0,
      dstOffset: 0,
      selfLoop: false,
    });
  }

  const sideGroups = new Map<string, Array<{ routeIndex: number; sourceSide: boolean }>>();
  for (let i = 0; i < pendingRoutes.length; i += 1) {
    const route = pendingRoutes[i];
    if (route.selfLoop) {
      continue;
    }

    const srcKey = `${route.srcAnchorId}:${route.srcSide}:src`;
    const dstKey = `${route.dstAnchorId}:${route.dstSide}:dst`;
    const srcMembers = sideGroups.get(srcKey) ?? [];
    srcMembers.push({ routeIndex: i, sourceSide: true });
    sideGroups.set(srcKey, srcMembers);

    const dstMembers = sideGroups.get(dstKey) ?? [];
    dstMembers.push({ routeIndex: i, sourceSide: false });
    sideGroups.set(dstKey, dstMembers);
  }

  for (const [key, members] of sideGroups) {
    if (members.length <= 1) {
      continue;
    }

    const [anchorId, sideToken] = key.split(":");
    const side = sideToken as EdgeSide;
    const anchorNode = nodeById.get(anchorId);
    if (anchorNode?.isJunction) {
      continue;
    }

    const anchor = anchorNode ?? subgraphById.get(anchorId);
    if (!anchor) {
      continue;
    }

    const span = side === "T" || side === "B" ? anchor.width : anchor.height;
    const sorted = [...members].sort((a, b) => {
      const routeA = pendingRoutes[a.routeIndex];
      const routeB = pendingRoutes[b.routeIndex];
      const axisA = sideSortAxis(side, routeA.srcAnchor, routeA.dstAnchor, a.sourceSide);
      const axisB = sideSortAxis(side, routeB.srcAnchor, routeB.dstAnchor, b.sourceSide);
      return axisA - axisB;
    });

    const offsets = laneOffsets(sorted.length, span);
    for (let i = 0; i < sorted.length; i += 1) {
      const member = sorted[i];
      if (member.sourceSide) {
        pendingRoutes[member.routeIndex].srcOffset = offsets[i];
      } else {
        pendingRoutes[member.routeIndex].dstOffset = offsets[i];
      }
    }
  }

  for (const route of pendingRoutes) {
    if (route.selfLoop && route.selfLoopNode) {
      const loop = chooseSelfLoopPoints(route.selfLoopNode, viewport);
      route.edge.points = loop.points.map((point) => ({ ...point }));
      route.edge.labelPosition = route.edge.label ? { ...loop.labelPosition } : undefined;
      continue;
    }

    const p0 = nodeAnchorPoint(route.srcAnchor, route.srcSide, route.srcOffset);
    const p1 = nodeAnchorPoint(route.dstAnchor, route.dstSide, route.dstOffset);
    route.edge.points = [p0, p1];
    route.edge.labelPosition = route.edge.label ? fallbackLabelPosition(route.edge.points) : undefined;
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

function averageNearestNodeGapPx(ir: DiagramIr): number {
  const nodes = ir.nodes;
  if (nodes.length <= 1) {
    return 0;
  }

  let total = 0;
  for (let i = 0; i < nodes.length; i += 1) {
    const a = nodes[i];
    let best = Number.POSITIVE_INFINITY;
    for (let j = 0; j < nodes.length; j += 1) {
      if (i === j) {
        continue;
      }
      const d = rectDistance(a, nodes[j]);
      if (d < best) {
        best = d;
      }
    }
    if (Number.isFinite(best)) {
      total += best;
    }
  }

  return total / Math.max(1, nodes.length);
}

function buildEdgeSegments(ir: DiagramIr): EdgeSegment[] {
  const out: EdgeSegment[] = [];
  const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));

  for (const edge of ir.edges) {
    if (edge.points.length >= 2) {
      for (let i = 0; i < edge.points.length - 1; i += 1) {
        out.push({
          edgeId: edge.id,
          from: edge.from,
          to: edge.to,
          p0: edge.points[i],
          p1: edge.points[i + 1],
        });
      }
      continue;
    }

    const from = nodeById.get(edge.from);
    const to = nodeById.get(edge.to);
    if (!from || !to) {
      continue;
    }
    out.push({
      edgeId: edge.id,
      from: edge.from,
      to: edge.to,
      p0: { x: from.x + from.width / 2, y: from.y + from.height / 2 },
      p1: { x: to.x + to.width / 2, y: to.y + to.height / 2 },
    });
  }

  return out;
}

function totalEdgeLengthPx(ir: DiagramIr): number {
  let total = 0;
  for (const edge of ir.edges) {
    if (edge.points.length < 2) {
      continue;
    }
    for (let i = 0; i < edge.points.length - 1; i += 1) {
      total += segmentLength(edge.points[i], edge.points[i + 1]);
    }
  }
  return total;
}

function totalBends(ir: DiagramIr): number {
  let bends = 0;
  for (const edge of ir.edges) {
    bends += Math.max(0, edge.points.length - 2);
  }
  return bends;
}

function computeHardConstraints(ir: DiagramIr, segments: EdgeSegment[], labelBoxes: EdgeLabelBox[]): HardConstraintStats {
  let textOverflow = 0;
  let nodeOverlap = 0;
  let nodeGap = 0;
  let labelNodeOverlap = 0;
  let labelLabelOverlap = 0;
  let nodeOutsideSubgraph = 0;
  let subgraphOverlap = 0;
  let edgeNodeGap = 0;
  let penalty = 0;

  for (const node of ir.nodes) {
    const estimated = estimateNodeTextMinSize(node);
    if ((node.style.fontSize || 14) < HARD_MIN_FONT_PT) {
      textOverflow += 1;
      penalty += 180;
    }
    if (node.width + 0.5 < estimated.minWidth || node.height + 0.5 < estimated.minHeight) {
      textOverflow += 1;
      penalty += 220;
    }
  }

  for (let i = 0; i < ir.nodes.length; i += 1) {
    const a = ir.nodes[i];
    for (let j = i + 1; j < ir.nodes.length; j += 1) {
      const b = ir.nodes[j];
      const overlapArea = rectIntersectionArea(a, b);
      if (overlapArea > 0.01) {
        nodeOverlap += 1;
        penalty += overlapArea * 10;
      }
      const distance = rectDistance(a, b);
      if (distance < HARD_MIN_NODE_GAP_PX) {
        nodeGap += 1;
        penalty += (HARD_MIN_NODE_GAP_PX - distance) * 16;
      }
    }
  }

  const subgraphById = new Map(ir.subgraphs.map((subgraph) => [subgraph.id, subgraph]));
  for (const node of ir.nodes) {
    if (!node.subgraphId) {
      continue;
    }
    const parent = subgraphById.get(node.subgraphId);
    if (!parent) {
      continue;
    }
    if (!rectContainsRect(parent, node, 1)) {
      nodeOutsideSubgraph += 1;
      penalty += 180;
    }
  }

  const parentById = new Map(ir.subgraphs.map((subgraph) => [subgraph.id, subgraph.parentId]));
  for (let i = 0; i < ir.subgraphs.length; i += 1) {
    const a = ir.subgraphs[i];
    for (let j = i + 1; j < ir.subgraphs.length; j += 1) {
      const b = ir.subgraphs[j];
      if (areRelatedSubgraphs(a.id, b.id, parentById)) {
        continue;
      }
      const overlap = rectIntersectionArea(a, b);
      const gap = rectDistance(a, b);
      if (overlap > 0.01 || gap < HARD_MIN_SUBGRAPH_GAP_PX) {
        subgraphOverlap += 1;
        penalty += overlap * 8 + Math.max(0, HARD_MIN_SUBGRAPH_GAP_PX - gap) * 18;
      }
    }
  }

  const expandedNodeRects = ir.nodes.map((node) => ({
    id: node.id,
    rect: expandedRect(node, 2),
  }));
  for (const labelBox of labelBoxes) {
    for (const node of expandedNodeRects) {
      const overlap = rectIntersectionArea(labelBox, node.rect);
      if (overlap <= 0.01) {
        continue;
      }
      labelNodeOverlap += 1;
      penalty += overlap * 20;
    }
  }
  for (let i = 0; i < labelBoxes.length; i += 1) {
    const a = labelBoxes[i];
    for (let j = i + 1; j < labelBoxes.length; j += 1) {
      const b = labelBoxes[j];
      const overlap = rectIntersectionArea(a, b);
      const gap = rectDistance(a, b);
      if (overlap > 0.01 || gap < HARD_MIN_LABEL_GAP_PX) {
        labelLabelOverlap += 1;
        penalty += overlap * 28 + Math.max(0, HARD_MIN_LABEL_GAP_PX - gap) * 22;
      }
    }
  }

  const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));
  for (const seg of segments) {
    for (const node of ir.nodes) {
      if (node.id === seg.from || node.id === seg.to) {
        continue;
      }
      const distance = distanceSegmentToRect(seg.p0, seg.p1, node);
      if (distance < HARD_MIN_EDGE_NODE_GAP_PX) {
        edgeNodeGap += 1;
        penalty += (HARD_MIN_EDGE_NODE_GAP_PX - distance) * 12;
      }
    }

    // Edges that point to missing nodes are also hard-invalid.
    if (!nodeById.has(seg.from) || !nodeById.has(seg.to)) {
      edgeNodeGap += 1;
      penalty += 120;
    }
  }

  const totalViolations =
    textOverflow +
    nodeOverlap +
    nodeGap +
    labelNodeOverlap +
    labelLabelOverlap +
    nodeOutsideSubgraph +
    subgraphOverlap +
    edgeNodeGap;

  return {
    textOverflow,
    nodeOverlap,
    nodeGap,
    labelNodeOverlap,
    labelLabelOverlap,
    nodeOutsideSubgraph,
    subgraphOverlap,
    edgeNodeGap,
    totalViolations,
    penalty,
    feasible: totalViolations === 0,
  };
}

function computeReadability(ir: DiagramIr, segments: EdgeSegment[]): ReadabilityStats {
  let edgeCrossings = 0;
  let edgeThroughNode = 0;
  let parallelOverlapScore = 0;

  const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));
  for (const seg of segments) {
    for (const node of ir.nodes) {
      if (node.id === seg.from || node.id === seg.to) {
        continue;
      }
      if (segmentPassesThroughRect(seg.p0, seg.p1, node)) {
        edgeThroughNode += 1;
      }
    }

    if (!nodeById.has(seg.from) || !nodeById.has(seg.to)) {
      edgeThroughNode += 1;
    }
  }

  for (let i = 0; i < segments.length; i += 1) {
    const a = segments[i];
    const avx = a.p1.x - a.p0.x;
    const avy = a.p1.y - a.p0.y;
    const alen = Math.hypot(avx, avy);
    if (alen < 1e-6) {
      continue;
    }

    for (let j = i + 1; j < segments.length; j += 1) {
      const b = segments[j];
      if (a.edgeId === b.edgeId) {
        continue;
      }

      const sharesEndpointNode =
        a.from === b.from || a.from === b.to || a.to === b.from || a.to === b.to;
      const crossing = segmentIntersectionStrict(a.p0, a.p1, b.p0, b.p1);
      if (crossing && !sharesEndpointNode) {
        edgeCrossings += 1;
      }

      const bvx = b.p1.x - b.p0.x;
      const bvy = b.p1.y - b.p0.y;
      const blen = Math.hypot(bvx, bvy);
      if (blen < 1e-6) {
        continue;
      }
      const cos = Math.abs(dot(avx, avy, bvx, bvy) / (alen * blen));
      if (cos < 0.965) {
        continue;
      }

      const distance = distanceSegmentToSegment(a.p0, a.p1, b.p0, b.p1);
      if (distance > 14) {
        continue;
      }

      const useA = alen >= blen;
      const axisX = useA ? avx / alen : bvx / blen;
      const axisY = useA ? avy / alen : bvy / blen;
      const aMin = Math.min(dot(a.p0.x, a.p0.y, axisX, axisY), dot(a.p1.x, a.p1.y, axisX, axisY));
      const aMax = Math.max(dot(a.p0.x, a.p0.y, axisX, axisY), dot(a.p1.x, a.p1.y, axisX, axisY));
      const bMin = Math.min(dot(b.p0.x, b.p0.y, axisX, axisY), dot(b.p1.x, b.p1.y, axisX, axisY));
      const bMax = Math.max(dot(b.p0.x, b.p0.y, axisX, axisY), dot(b.p1.x, b.p1.y, axisX, axisY));
      const overlap = Math.max(0, Math.min(aMax, bMax) - Math.max(aMin, bMin));
      if (overlap <= 16) {
        continue;
      }
      parallelOverlapScore += overlap / Math.max(2, distance + 1);
    }
  }

  return {
    edgeCrossings,
    edgeThroughNode,
    parallelOverlapScore,
  };
}

function computeLayoutQuality(ir: DiagramIr, targetAspectRatio: number): LayoutQuality {
  const aspect = diagramAspect(ir);
  const target = clampAspect(targetAspectRatio);
  const fit = normalizedFitScale(ir, target);
  const minGapPx = minNodeGapPx(ir);
  const avgGapPx = averageNearestNodeGapPx(ir);
  const minGapScaled = minGapPx * fit;
  const avgGapScaled = avgGapPx * fit;
  const coverage = coverageForAspect(aspect, target);
  const aspectError = Math.abs(Math.log(Math.max(1e-6, aspect / target)));
  const edgeLenScaled = totalEdgeLengthPx(ir) * fit;
  const bends = totalBends(ir);
  const boundsArea = Math.max(1, ir.bounds.width * ir.bounds.height);
  const totalNodeArea = ir.nodes.reduce((sum, node) => sum + Math.max(0, node.width * node.height), 0);
  const nodeAreaRatio = totalNodeArea / boundsArea;

  const labelBoxes = ir.edges.map((edge) => estimateEdgeLabelRect(edge)).filter((label): label is EdgeLabelBox => Boolean(label));
  const segments = buildEdgeSegments(ir);
  const hard = computeHardConstraints(ir, segments, labelBoxes);
  const readability = computeReadability(ir, segments);

  const spacingCenter = 0.020;
  const spacingHalfRange = 0.012;
  const spacingAvgScore = clamp(1 - Math.abs(avgGapScaled - spacingCenter) / spacingHalfRange, 0, 1);
  const spacingMinScore = clamp(minGapScaled / 0.012, 0, 1);
  const spacingScore = spacingAvgScore * 0.68 + spacingMinScore * 0.32;
  const nodeAreaNorm = clamp(nodeAreaRatio / 0.26, 0, 2.4);
  const objectSizeScore = fit * (0.30 + nodeAreaNorm * 0.70);

  const desiredMinGapScaled = 0.022;
  const desiredMaxGapScaled = 0.040;
  let densityPenalty = 0;
  if (avgGapScaled < desiredMinGapScaled) {
    densityPenalty += (desiredMinGapScaled - avgGapScaled) * 9;
  } else if (avgGapScaled > desiredMaxGapScaled) {
    densityPenalty += (avgGapScaled - desiredMaxGapScaled) * 10.0;
  }
  if (minGapScaled < 0.010) {
    densityPenalty += (0.010 - minGapScaled) * 12;
  }
  if (nodeAreaRatio < 0.20) {
    densityPenalty += (0.20 - nodeAreaRatio) * 88;
  }

  return {
    hard,
    readability,
    fit,
    objectSizeScore,
    spacingScore,
    nodeAreaRatio,
    coverage,
    aspectError,
    minGapPx,
    minGapScaled,
    avgGapPx,
    avgGapScaled,
    densityPenalty,
    edgeLenScaled,
    bends,
    boundsWidth: Math.max(1, ir.bounds.width),
    boundsHeight: Math.max(1, ir.bounds.height),
  };
}

function compareLower(a: number, b: number, tolerance = 1e-6): number {
  if (a < b - tolerance) {
    return -1;
  }
  if (a > b + tolerance) {
    return 1;
  }
  return 0;
}

function compareHigher(a: number, b: number, tolerance = 1e-6): number {
  return compareLower(b, a, tolerance);
}

function isBetterQuality(candidate: LayoutQuality, best: LayoutQuality): boolean {
  // Stage 0: hard constraints (must satisfy first).
  const textOverflowCmp = compareLower(candidate.hard.textOverflow, best.hard.textOverflow);
  if (textOverflowCmp !== 0) {
    return textOverflowCmp < 0;
  }

  if (candidate.hard.feasible !== best.hard.feasible) {
    return candidate.hard.feasible;
  }
  if (!candidate.hard.feasible || !best.hard.feasible) {
    const totalCmp = compareLower(candidate.hard.totalViolations, best.hard.totalViolations);
    if (totalCmp !== 0) {
      return totalCmp < 0;
    }
    const penaltyCmp = compareLower(candidate.hard.penalty, best.hard.penalty, 0.5);
    if (penaltyCmp !== 0) {
      return penaltyCmp < 0;
    }
  }

  // Stage 1: object legibility and presence.
  const objectSizeCmp = compareHigher(candidate.objectSizeScore, best.objectSizeScore, 6e-5);
  if (objectSizeCmp !== 0) {
    return objectSizeCmp < 0;
  }
  const areaRatioCmp = compareHigher(candidate.nodeAreaRatio, best.nodeAreaRatio, 2e-4);
  if (areaRatioCmp !== 0) {
    return areaRatioCmp < 0;
  }
  const fitCmp = compareHigher(candidate.fit, best.fit, 5e-5);
  if (fitCmp !== 0) {
    return fitCmp < 0;
  }
  const spacingCmp = compareHigher(candidate.spacingScore, best.spacingScore, 4e-4);
  if (spacingCmp !== 0) {
    return spacingCmp < 0;
  }
  const densityCmp = compareLower(candidate.densityPenalty, best.densityPenalty, 0.0018);
  if (densityCmp !== 0) {
    return densityCmp < 0;
  }

  // Stage 2: page usage.
  const aspectCmp = compareLower(candidate.aspectError, best.aspectError, 0.003);
  if (aspectCmp !== 0) {
    return aspectCmp < 0;
  }
  const coverageCmp = compareHigher(candidate.coverage, best.coverage, 0.004);
  if (coverageCmp !== 0) {
    return coverageCmp < 0;
  }
  const avgGapCmp = compareHigher(candidate.avgGapScaled, best.avgGapScaled, 4e-4);
  if (avgGapCmp !== 0) {
    return avgGapCmp < 0;
  }

  // Stage 3: edge readability.
  const crossingCmp = compareLower(candidate.readability.edgeCrossings, best.readability.edgeCrossings);
  if (crossingCmp !== 0) {
    return crossingCmp < 0;
  }
  const throughCmp = compareLower(candidate.readability.edgeThroughNode, best.readability.edgeThroughNode);
  if (throughCmp !== 0) {
    return throughCmp < 0;
  }
  const parallelCmp = compareLower(candidate.readability.parallelOverlapScore, best.readability.parallelOverlapScore, 0.02);
  if (parallelCmp !== 0) {
    return parallelCmp < 0;
  }

  // Stage 4: secondary aesthetics and compactness.
  const bendCmp = compareLower(candidate.bends, best.bends);
  if (bendCmp !== 0) {
    return bendCmp < 0;
  }
  const edgeLenCmp = compareLower(candidate.edgeLenScaled, best.edgeLenScaled, 0.01);
  if (edgeLenCmp !== 0) {
    return edgeLenCmp < 0;
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

function applyNodeRepulsion(ir: DiagramIr, desiredGapPx: number): void {
  if (ir.nodes.length <= 1) {
    return;
  }

  const iterations = 16;
  const subgraphById = new Map(ir.subgraphs.map((subgraph) => [subgraph.id, subgraph]));

  for (let iter = 0; iter < iterations; iter += 1) {
    const forces = new Map<string, { x: number; y: number }>();
    for (const node of ir.nodes) {
      forces.set(node.id, { x: 0, y: 0 });
    }

    for (let i = 0; i < ir.nodes.length; i += 1) {
      const a = ir.nodes[i];
      for (let j = i + 1; j < ir.nodes.length; j += 1) {
        const b = ir.nodes[j];
        const ra = expandedRect(a, desiredGapPx / 2);
        const rb = expandedRect(b, desiredGapPx / 2);
        if (!intersectsRect(ra, rb)) {
          continue;
        }

        const overlapX = Math.min(ra.x + ra.width, rb.x + rb.width) - Math.max(ra.x, rb.x);
        const overlapY = Math.min(ra.y + ra.height, rb.y + rb.height) - Math.max(ra.y, rb.y);
        if (overlapX <= 0 || overlapY <= 0) {
          continue;
        }

        const acx = a.x + a.width / 2;
        const bcx = b.x + b.width / 2;
        const acy = a.y + a.height / 2;
        const bcy = b.y + b.height / 2;
        const forceA = forces.get(a.id);
        const forceB = forces.get(b.id);
        if (!forceA || !forceB) {
          continue;
        }

        if (overlapX < overlapY) {
          const dir = acx === bcx ? (a.id < b.id ? -1 : 1) : acx < bcx ? -1 : 1;
          const push = overlapX / 2 + 0.6;
          forceA.x += dir * push;
          forceB.x -= dir * push;
        } else {
          const dir = acy === bcy ? (a.id < b.id ? -1 : 1) : acy < bcy ? -1 : 1;
          const push = overlapY / 2 + 0.6;
          forceA.y += dir * push;
          forceB.y -= dir * push;
        }
      }
    }

    let globalCx = 0;
    let globalCy = 0;
    for (const node of ir.nodes) {
      globalCx += node.x + node.width / 2;
      globalCy += node.y + node.height / 2;
    }
    globalCx /= ir.nodes.length;
    globalCy /= ir.nodes.length;

    const maxStep = clamp(desiredGapPx * (0.32 - iter * 0.010), 2.0, 22);
    for (const node of ir.nodes) {
      if (node.isJunction) {
        continue;
      }

      const force = forces.get(node.id);
      if (!force) {
        continue;
      }

      const cx = node.x + node.width / 2;
      const cy = node.y + node.height / 2;
      force.x += (globalCx - cx) * 0.022;
      force.y += (globalCy - cy) * 0.022;

      if (node.subgraphId) {
        const sub = subgraphById.get(node.subgraphId);
        if (sub) {
          const scx = sub.x + sub.width / 2;
          const scy = sub.y + sub.height / 2;
          force.x += (scx - cx) * 0.032;
          force.y += (scy - cy) * 0.032;
        }
      }

      const mag = Math.hypot(force.x, force.y);
      if (mag > maxStep && mag > 1e-6) {
        const scale = maxStep / mag;
        force.x *= scale;
        force.y *= scale;
      }
      node.x += force.x;
      node.y += force.y;
    }
  }

  recomputeSubgraphBounds(ir);
  recomputeBounds(ir);
}

function spreadTopLevelBlocks(ir: DiagramIr, desiredGapPx: number): void {
  if (ir.subgraphs.length === 0 && ir.nodes.length <= 1) {
    return;
  }

  const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));
  const iterations = 12;

  for (let iter = 0; iter < iterations; iter += 1) {
    const blocks = collectTopLevelBlocks(ir);
    if (blocks.length <= 1) {
      return;
    }

    const moveByBlock = new Map<string, { dx: number; dy: number }>();
    for (const block of blocks) {
      moveByBlock.set(block.id, { dx: 0, dy: 0 });
    }

    let moved = false;
    for (let i = 0; i < blocks.length; i += 1) {
      const a = blocks[i];
      for (let j = i + 1; j < blocks.length; j += 1) {
        const b = blocks[j];
        const ra = expandedRect(a, Math.max(0, desiredGapPx / 4));
        const rb = expandedRect(b, Math.max(0, desiredGapPx / 4));
        if (!intersectsRect(ra, rb)) {
          continue;
        }

        const overlapX = Math.min(ra.x + ra.width, rb.x + rb.width) - Math.max(ra.x, rb.x);
        const overlapY = Math.min(ra.y + ra.height, rb.y + rb.height) - Math.max(ra.y, rb.y);
        if (overlapX <= 0 || overlapY <= 0) {
          continue;
        }

        const acx = a.x + a.width / 2;
        const bcx = b.x + b.width / 2;
        const acy = a.y + a.height / 2;
        const bcy = b.y + b.height / 2;
        const ma = moveByBlock.get(a.id);
        const mb = moveByBlock.get(b.id);
        if (!ma || !mb) {
          continue;
        }

        moved = true;
        if (overlapX < overlapY) {
          const dir = acx === bcx ? (a.id < b.id ? -1 : 1) : acx < bcx ? -1 : 1;
          const push = overlapX / 2 + 0.5;
          ma.dx += dir * push;
          mb.dx -= dir * push;
        } else {
          const dir = acy === bcy ? (a.id < b.id ? -1 : 1) : acy < bcy ? -1 : 1;
          const push = overlapY / 2 + 0.5;
          ma.dy += dir * push;
          mb.dy -= dir * push;
        }
      }
    }

    if (!moved) {
      break;
    }

    for (const block of blocks) {
      const move = moveByBlock.get(block.id);
      if (!move) {
        continue;
      }
      const mag = Math.hypot(move.dx, move.dy);
      const maxStep = clamp(desiredGapPx * 0.32, 1.4, 26);
      if (mag > maxStep && mag > 1e-6) {
        const scale = maxStep / mag;
        move.dx *= scale;
        move.dy *= scale;
      }
      for (const nodeId of block.nodeIds) {
        const node = nodeById.get(nodeId);
        if (!node) {
          continue;
        }
        node.x += move.dx;
        node.y += move.dy;
      }
    }

    recomputeSubgraphBounds(ir);
    recomputeBounds(ir);
  }

  rerouteEdgesStraight(ir);
}

function scaleNodePositionsAroundCenter(ir: DiagramIr, scaleX: number, scaleY: number): void {
  recomputeBounds(ir);
  const cx = (ir.bounds.minX + ir.bounds.maxX) / 2;
  const cy = (ir.bounds.minY + ir.bounds.maxY) / 2;
  for (const node of ir.nodes) {
    node.x = cx + (node.x - cx) * scaleX;
    node.y = cy + (node.y - cy) * scaleY;
  }
  recomputeSubgraphBounds(ir);
  recomputeBounds(ir);
}

function compactLayoutByGapTargets(ir: DiagramIr, targetMinGapPx: number, targetAvgGapPx: number): void {
  const maxIter = 18;
  for (let iter = 0; iter < maxIter; iter += 1) {
    const minGap = minNodeGapPx(ir);
    const avgGap = averageNearestNodeGapPx(ir);
    if (minGap <= targetMinGapPx * 1.10 && avgGap <= targetAvgGapPx * 1.10) {
      break;
    }

    const minRatio = targetMinGapPx / Math.max(1, minGap);
    const avgRatio = targetAvgGapPx / Math.max(1, avgGap);
    const factor = clamp(Math.max(minRatio, avgRatio), 0.80, 0.965);
    if (factor >= 0.999) {
      break;
    }

    const before = captureLayoutState(ir);
    scaleNodePositionsAroundCenter(ir, factor, factor);
    rerouteEdgesStraight(ir);

    const minAfter = minNodeGapPx(ir);
    if (minAfter < targetMinGapPx * 0.90) {
      restoreLayoutState(ir, before);
      break;
    }
  }
}

function refineLayoutForReadability(ir: DiagramIr): void {
  const desiredGap = clamp((ir.config.layout.nodesep + ir.config.layout.ranksep) / 4.4, 8, 24);
  applyNodeRepulsion(ir, desiredGap);
  const blocks = collectTopLevelBlocks(ir);
  let hasBlockOverlap = false;
  for (let i = 0; i < blocks.length; i += 1) {
    for (let j = i + 1; j < blocks.length; j += 1) {
      if (rectIntersectionArea(blocks[i], blocks[j]) > 0.1) {
        hasBlockOverlap = true;
        break;
      }
    }
    if (hasBlockOverlap) {
      break;
    }
  }
  if (hasBlockOverlap) {
    spreadTopLevelBlocks(ir, desiredGap * 0.45);
  }
  rerouteEdgesStraight(ir);
  const compactMinGap = clamp(desiredGap * 1.25, 16, 44);
  const compactAvgGap = clamp(compactMinGap * 2.2, 38, 96);
  compactLayoutByGapTargets(ir, compactMinGap, compactAvgGap);
  rerouteEdgesStraight(ir);
}

function rebalanceAspectBySpreading(ir: DiagramIr, targetAspectRatio: number): void {
  void ir;
  void targetAspectRatio;
}

function wrapTopLevelBlocks(ir: DiagramIr, targetAspectRatio: number): void {
  const blocks = collectTopLevelBlocks(ir);
  if (blocks.length <= 1) {
    return;
  }

  // Gap between top-level blocks (subgraphs / ungrouped section).
  const gap = clamp((ir.config.layout.nodesep + ir.config.layout.ranksep) / 2.6, 10, 72);

  rerouteEdgesStraight(ir);
  refineLayoutForReadability(ir);
  rebalanceAspectBySpreading(ir, targetAspectRatio);
  const baseState = captureLayoutState(ir);
  let bestState = baseState;
  let bestQuality = computeLayoutQuality(ir, targetAspectRatio);

  const n = blocks.length;
  const maxRows = Math.min(5, n);

  const tryRows = (rows: TopLevelBlock[][]): void => {
    restoreLayoutState(ir, baseState);
    applyPackedRows(ir, rows, gap);
    refineLayoutForReadability(ir);
    rebalanceAspectBySpreading(ir, targetAspectRatio);
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
  const multipliers = [0.9, 1.0, 1.2, 1.45, 1.8, 2.2];
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

  const nodesepCandidates = spacingCandidates(baseNodesep, [28, 36, 48, 64, 84, 108, 136], 14, 180);
  const ranksepCandidates = spacingCandidates(baseRanksep, [22, 30, 42, 56, 74, 96, 126], 12, 180);

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
        refineLayoutForReadability(ir);
        rebalanceAspectBySpreading(ir, targetAspectRatio);
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
  rerouteEdgesStraight(ir);
  refineLayoutForReadability(ir);
  rebalanceAspectBySpreading(ir, targetAspectRatio);
}

export function layoutDiagram(ir: DiagramIr, options: LayoutOptions = {}): void {
  if (options.targetAspectRatio) {
    optimizeLayoutForSlide(ir, options.targetAspectRatio);
    return;
  }

  applyLayout(ir);
}
