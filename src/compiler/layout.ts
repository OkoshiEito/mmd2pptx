import dagre from "dagre";
import type { DiagramDirection, DiagramIr, EdgeSide, Point } from "../types.js";
import { recomputeBounds, recomputeSubgraphBounds } from "./geometry.js";

interface LayoutOptions {
  targetAspectRatio?: number;
}

interface LayoutState {
  nodePositions: Map<string, { x: number; y: number }>;
  edgeRoutes: Map<string, { points: Point[]; labelPosition?: Point }>;
  edgeSides: Map<string, { startSide?: EdgeSide; endSide?: EdgeSide }>;
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

type Rankdir = "TB" | "BT" | "LR" | "RL";

interface EdgeGeom {
  fromId: string;
  toId: string;
}

interface IndexedEdgeGeom extends EdgeGeom {
  fromIndex: number;
  toIndex: number;
}

interface Segment {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

const EDGE_SIDES: EdgeSide[] = ["T", "L", "B", "R"];

function nodeCenter(node: DiagramIr["nodes"][number], override?: { x?: number; y?: number }): Point {
  return {
    x: (override?.x ?? node.x) + node.width / 2,
    y: (override?.y ?? node.y) + node.height / 2,
  };
}

function sideNormal(side: EdgeSide): Point {
  if (side === "T") {
    return { x: 0, y: -1 };
  }
  if (side === "L") {
    return { x: -1, y: 0 };
  }
  if (side === "B") {
    return { x: 0, y: 1 };
  }
  return { x: 1, y: 0 };
}

function sideAnchor(node: DiagramIr["nodes"][number], side: EdgeSide): Point {
  const cx = node.x + node.width / 2;
  const cy = node.y + node.height / 2;
  if (side === "T") {
    return { x: cx, y: node.y };
  }
  if (side === "L") {
    return { x: node.x, y: cy };
  }
  if (side === "B") {
    return { x: cx, y: node.y + node.height };
  }
  return { x: node.x + node.width, y: cy };
}

function sideOutwardAnchor(node: DiagramIr["nodes"][number], side: EdgeSide, distance: number): Point {
  const anchor = sideAnchor(node, side);
  const normal = sideNormal(side);
  return {
    x: anchor.x + normal.x * distance,
    y: anchor.y + normal.y * distance,
  };
}

function segmentForEdge(
  edge: EdgeGeom,
  nodeById: Map<string, DiagramIr["nodes"][number]>,
  override?: { nodeId: string; x: number; y: number },
): Segment | undefined {
  const from = nodeById.get(edge.fromId);
  const to = nodeById.get(edge.toId);
  if (!from || !to || from.id === to.id) {
    return undefined;
  }

  const fromCenter = nodeCenter(
    from,
    override && override.nodeId === from.id
      ? {
          x: override.x,
          y: override.y,
        }
      : undefined,
  );
  const toCenter = nodeCenter(
    to,
    override && override.nodeId === to.id
      ? {
          x: override.x,
          y: override.y,
        }
      : undefined,
  );

  return {
    x1: fromCenter.x,
    y1: fromCenter.y,
    x2: toCenter.x,
    y2: toCenter.y,
  };
}

function orientation(ax: number, ay: number, bx: number, by: number, cx: number, cy: number): number {
  return (bx - ax) * (cy - ay) - (by - ay) * (cx - ax);
}

function segmentsCross(a: Segment, b: Segment): boolean {
  const o1 = orientation(a.x1, a.y1, a.x2, a.y2, b.x1, b.y1);
  const o2 = orientation(a.x1, a.y1, a.x2, a.y2, b.x2, b.y2);
  const o3 = orientation(b.x1, b.y1, b.x2, b.y2, a.x1, a.y1);
  const o4 = orientation(b.x1, b.y1, b.x2, b.y2, a.x2, a.y2);

  const eps = 1e-6;
  if (Math.abs(o1) < eps || Math.abs(o2) < eps || Math.abs(o3) < eps || Math.abs(o4) < eps) {
    return false;
  }

  return (o1 > 0) !== (o2 > 0) && (o3 > 0) !== (o4 > 0);
}

function directionalPenalty(from: Point, to: Point, rankdir: Rankdir): number {
  const minForward = 22;

  if (rankdir === "TB") {
    const forward = to.y - from.y;
    return forward >= minForward ? 0 : (minForward - forward) * 8.8;
  }
  if (rankdir === "BT") {
    const forward = from.y - to.y;
    return forward >= minForward ? 0 : (minForward - forward) * 8.8;
  }
  if (rankdir === "LR") {
    const forward = to.x - from.x;
    return forward >= minForward ? 0 : (minForward - forward) * 8.8;
  }

  const forward = from.x - to.x;
  return forward >= minForward ? 0 : (minForward - forward) * 8.8;
}

function expandedOverlapArea(
  a: DiagramIr["nodes"][number],
  b: DiagramIr["nodes"][number],
  gap: number,
  overrideA?: { x: number; y: number },
): number {
  const ax0 = (overrideA?.x ?? a.x) - gap;
  const ay0 = (overrideA?.y ?? a.y) - gap;
  const ax1 = ax0 + a.width + gap * 2;
  const ay1 = ay0 + a.height + gap * 2;

  const bx0 = b.x - gap;
  const by0 = b.y - gap;
  const bx1 = bx0 + b.width + gap * 2;
  const by1 = by0 + b.height + gap * 2;

  const ox = Math.max(0, Math.min(ax1, bx1) - Math.max(ax0, bx0));
  const oy = Math.max(0, Math.min(ay1, by1) - Math.max(ay0, by0));
  return ox * oy;
}

function resolveNodeCollisions(nodes: DiagramIr["nodes"], gap: number, passes: number): void {
  if (nodes.length < 2) {
    return;
  }

  for (let pass = 0; pass < passes; pass += 1) {
    for (let i = 0; i < nodes.length; i += 1) {
      const a = nodes[i];
      const acx = a.x + a.width / 2;
      const acy = a.y + a.height / 2;
      for (let j = i + 1; j < nodes.length; j += 1) {
        const b = nodes[j];
        const bcx = b.x + b.width / 2;
        const bcy = b.y + b.height / 2;

        const needX = (a.width + b.width) / 2 + gap;
        const needY = (a.height + b.height) / 2 + gap;
        const dx = bcx - acx;
        const dy = bcy - acy;
        const overlapX = needX - Math.abs(dx);
        const overlapY = needY - Math.abs(dy);

        if (overlapX <= 0 || overlapY <= 0) {
          continue;
        }

        if (overlapX < overlapY) {
          const push = overlapX / 2 + 0.5;
          const sign = dx >= 0 ? 1 : -1;
          a.x -= push * sign;
          b.x += push * sign;
        } else {
          const push = overlapY / 2 + 0.5;
          const sign = dy >= 0 ? 1 : -1;
          a.y -= push * sign;
          b.y += push * sign;
        }
      }
    }
  }
}

function cellKey(ix: number, iy: number): string {
  return `${ix},${iy}`;
}

function buildPointSpatialIndex(cx: number[], cy: number[], cellSize: number): Map<string, number[]> {
  const buckets = new Map<string, number[]>();
  for (let i = 0; i < cx.length; i += 1) {
    const key = cellKey(Math.floor(cx[i] / cellSize), Math.floor(cy[i] / cellSize));
    const list = buckets.get(key) ?? [];
    list.push(i);
    buckets.set(key, list);
  }
  return buckets;
}

function queryPointSpatialIndex(
  buckets: Map<string, number[]>,
  cellSize: number,
  x: number,
  y: number,
  radius: number,
): number[] {
  const minX = Math.floor((x - radius) / cellSize);
  const maxX = Math.floor((x + radius) / cellSize);
  const minY = Math.floor((y - radius) / cellSize);
  const maxY = Math.floor((y + radius) / cellSize);

  const out: number[] = [];
  for (let ix = minX; ix <= maxX; ix += 1) {
    for (let iy = minY; iy <= maxY; iy += 1) {
      const list = buckets.get(cellKey(ix, iy));
      if (!list) {
        continue;
      }
      out.push(...list);
    }
  }
  return out;
}

function applyCrossingNudges(
  edges: IndexedEdgeGeom[],
  centersX: number[],
  centersY: number[],
  movableNodeIndices: Set<number>,
  fx: number[],
  fy: number[],
  strength: number,
): void {
  if (edges.length < 2 || strength <= 0) {
    return;
  }

  const segments: Segment[] = [];
  const midX: number[] = [];
  const midY: number[] = [];
  const cellSize = 140;
  const buckets = new Map<string, number[]>();

  for (let i = 0; i < edges.length; i += 1) {
    const edge = edges[i];
    const segment: Segment = {
      x1: centersX[edge.fromIndex],
      y1: centersY[edge.fromIndex],
      x2: centersX[edge.toIndex],
      y2: centersY[edge.toIndex],
    };
    segments.push(segment);
    const mx = (segment.x1 + segment.x2) / 2;
    const my = (segment.y1 + segment.y2) / 2;
    midX.push(mx);
    midY.push(my);
    const key = cellKey(Math.floor(mx / cellSize), Math.floor(my / cellSize));
    const list = buckets.get(key) ?? [];
    list.push(i);
    buckets.set(key, list);
  }

  for (let i = 0; i < edges.length; i += 1) {
    const edgeA = edges[i];
    const segmentA = segments[i];
    const ax = segmentA.x2 - segmentA.x1;
    const ay = segmentA.y2 - segmentA.y1;
    const lenA = Math.hypot(ax, ay);
    if (lenA < 1e-6) {
      continue;
    }

    const bucketX = Math.floor(midX[i] / cellSize);
    const bucketY = Math.floor(midY[i] / cellSize);
    for (let bx = bucketX - 1; bx <= bucketX + 1; bx += 1) {
      for (let by = bucketY - 1; by <= bucketY + 1; by += 1) {
        const list = buckets.get(cellKey(bx, by));
        if (!list) {
          continue;
        }

        for (const j of list) {
          if (j <= i) {
            continue;
          }
          const edgeB = edges[j];
          if (
            edgeA.fromIndex === edgeB.fromIndex ||
            edgeA.fromIndex === edgeB.toIndex ||
            edgeA.toIndex === edgeB.fromIndex ||
            edgeA.toIndex === edgeB.toIndex
          ) {
            continue;
          }

          const segmentB = segments[j];
          if (!segmentsCross(segmentA, segmentB)) {
            continue;
          }

          const bxv = segmentB.x2 - segmentB.x1;
          const byv = segmentB.y2 - segmentB.y1;
          const lenB = Math.hypot(bxv, byv);
          if (lenB < 1e-6) {
            continue;
          }

          const aNx = -ay / lenA;
          const aNy = ax / lenA;
          const bNx = -byv / lenB;
          const bNy = bxv / lenB;
          const phase = ((i + j) & 1) === 0 ? 1 : -1;
          const push = strength;

          if (movableNodeIndices.has(edgeA.fromIndex)) {
            fx[edgeA.fromIndex] += aNx * push * phase;
            fy[edgeA.fromIndex] += aNy * push * phase;
          }
          if (movableNodeIndices.has(edgeA.toIndex)) {
            fx[edgeA.toIndex] += aNx * push * phase;
            fy[edgeA.toIndex] += aNy * push * phase;
          }
          if (movableNodeIndices.has(edgeB.fromIndex)) {
            fx[edgeB.fromIndex] -= bNx * push * phase;
            fy[edgeB.fromIndex] -= bNy * push * phase;
          }
          if (movableNodeIndices.has(edgeB.toIndex)) {
            fx[edgeB.toIndex] -= bNx * push * phase;
            fy[edgeB.toIndex] -= bNy * push * phase;
          }
        }
      }
    }
  }
}

function optimizeNodePlacement(ir: DiagramIr, rankdir: Rankdir, pinnedNodeIds: Set<string>): Set<string> {
  const moved = new Set<string>();
  const allNodes = ir.nodes.filter((node) => !node.isJunction);
  if (allNodes.length <= 1) {
    return moved;
  }

  const nodeIndexById = new Map(allNodes.map((node, index) => [node.id, index]));
  const nodeById = new Map(allNodes.map((node) => [node.id, node]));
  const movableNodes = allNodes.filter((node) => !pinnedNodeIds.has(node.id));
  if (movableNodes.length <= 1) {
    return moved;
  }

  const indexedEdges: IndexedEdgeGeom[] = ir.edges
    .map((edge) => {
      const fromIndex = nodeIndexById.get(edge.from);
      const toIndex = nodeIndexById.get(edge.to);
      if (fromIndex === undefined || toIndex === undefined || fromIndex === toIndex) {
        return undefined;
      }
      return {
        fromId: edge.from,
        toId: edge.to,
        fromIndex,
        toIndex,
      } satisfies IndexedEdgeGeom;
    })
    .filter((edge): edge is IndexedEdgeGeom => Boolean(edge));

  const incidentEdges = new Map<string, EdgeGeom[]>();
  for (const edge of indexedEdges) {
    const fromList = incidentEdges.get(edge.fromId) ?? [];
    fromList.push(edge);
    incidentEdges.set(edge.fromId, fromList);
    const toList = incidentEdges.get(edge.toId) ?? [];
    toList.push(edge);
    incidentEdges.set(edge.toId, toList);
  }

  const incidentEdgeIndices: number[][] = Array.from({ length: allNodes.length }, () => []);
  for (let i = 0; i < indexedEdges.length; i += 1) {
    const edge = indexedEdges[i];
    incidentEdgeIndices[edge.fromIndex].push(i);
    incidentEdgeIndices[edge.toIndex].push(i);
  }

  const movableNodeIndices = new Set<number>(movableNodes.map((node) => nodeIndexById.get(node.id) as number));
  const anchors = new Map(movableNodes.map((node) => [node.id, { x: node.x, y: node.y }]));
  const optimizationOrder = [...movableNodes].sort(
    (a, b) => (incidentEdges.get(b.id)?.length ?? 0) - (incidentEdges.get(a.id)?.length ?? 0),
  );

  const minGap = Math.max(8, Math.min(26, ir.config.layout.nodesep * 0.2));
  const centersX = new Array<number>(allNodes.length);
  const centersY = new Array<number>(allNodes.length);
  const reverseEdgeSet = new Set<string>(indexedEdges.map((edge) => `${edge.toIndex}:${edge.fromIndex}`));

  const refreshCenters = (): void => {
    for (let i = 0; i < allNodes.length; i += 1) {
      const node = allNodes[i];
      centersX[i] = node.x + node.width / 2;
      centersY[i] = node.y + node.height / 2;
    }
  };

  refreshCenters();

  const idealEdgeLength = clamp((ir.config.layout.nodesep + ir.config.layout.ranksep) * 0.55, 58, 260);
  const repulsionRadius = clamp(idealEdgeLength * 1.75, 120, 300);
  const iterations = clamp(Math.round(26 + Math.sqrt(allNodes.length) * 9), 28, 88);
  let temperature = clamp(Math.max(14, idealEdgeLength * 0.24), 12, 70);

  for (let iter = 0; iter < iterations; iter += 1) {
    refreshCenters();

    const fx = new Array<number>(allNodes.length).fill(0);
    const fy = new Array<number>(allNodes.length).fill(0);
    const nodeBuckets = buildPointSpatialIndex(centersX, centersY, 120);

    for (const node of movableNodes) {
      const i = nodeIndexById.get(node.id);
      if (i === undefined) {
        continue;
      }

      const x = centersX[i];
      const y = centersY[i];
      const near = queryPointSpatialIndex(nodeBuckets, 120, x, y, repulsionRadius);

      for (const j of near) {
        if (j === i) {
          continue;
        }
        const other = allNodes[j];
        const dx = x - centersX[j];
        const dy = y - centersY[j];
        const dist = Math.hypot(dx, dy);
        const safeDist = Math.max(1e-3, dist);

        const overlapX = (node.width + other.width) / 2 + minGap - Math.abs(dx);
        const overlapY = (node.height + other.height) / 2 + minGap - Math.abs(dy);
        if (overlapX > 0 && overlapY > 0) {
          if (overlapX < overlapY) {
            fx[i] += (dx >= 0 ? 1 : -1) * (overlapX * 1.48 + 4.2);
          } else {
            fy[i] += (dy >= 0 ? 1 : -1) * (overlapY * 1.48 + 4.2);
          }
          continue;
        }

        if (safeDist <= repulsionRadius) {
          const r = (repulsionRadius - safeDist) / repulsionRadius;
          const force = r * r * 22;
          fx[i] += (dx / safeDist) * force;
          fy[i] += (dy / safeDist) * force;
        }
      }
    }

    for (const edge of indexedEdges) {
      const fromIdx = edge.fromIndex;
      const toIdx = edge.toIndex;
      const dx = centersX[toIdx] - centersX[fromIdx];
      const dy = centersY[toIdx] - centersY[fromIdx];
      const dist = Math.max(1e-3, Math.hypot(dx, dy));
      const ux = dx / dist;
      const uy = dy / dist;

      const fromNode = allNodes[fromIdx];
      const toNode = allNodes[toIdx];
      const localTarget = idealEdgeLength + (Math.max(fromNode.width, fromNode.height) + Math.max(toNode.width, toNode.height)) * 0.1;
      const stretch = dist - localTarget;
      const spring = stretch * 0.045;

      if (movableNodeIndices.has(fromIdx)) {
        fx[fromIdx] += ux * spring;
        fy[fromIdx] += uy * spring;
      }
      if (movableNodeIndices.has(toIdx)) {
        fx[toIdx] -= ux * spring;
        fy[toIdx] -= uy * spring;
      }

      const fromPoint = { x: centersX[fromIdx], y: centersY[fromIdx] };
      const toPoint = { x: centersX[toIdx], y: centersY[toIdx] };
      const dirPenalty = directionalPenalty(fromPoint, toPoint, rankdir);
      if (dirPenalty > 0) {
        const force = Math.min(18, dirPenalty * 0.12);
        if (rankdir === "TB") {
          if (movableNodeIndices.has(fromIdx)) {
            fy[fromIdx] -= force;
          }
          if (movableNodeIndices.has(toIdx)) {
            fy[toIdx] += force;
          }
        } else if (rankdir === "BT") {
          if (movableNodeIndices.has(fromIdx)) {
            fy[fromIdx] += force;
          }
          if (movableNodeIndices.has(toIdx)) {
            fy[toIdx] -= force;
          }
        } else if (rankdir === "LR") {
          if (movableNodeIndices.has(fromIdx)) {
            fx[fromIdx] -= force;
          }
          if (movableNodeIndices.has(toIdx)) {
            fx[toIdx] += force;
          }
        } else {
          if (movableNodeIndices.has(fromIdx)) {
            fx[fromIdx] += force;
          }
          if (movableNodeIndices.has(toIdx)) {
            fx[toIdx] -= force;
          }
        }
      }

      if (reverseEdgeSet.has(`${fromIdx}:${toIdx}`) && reverseEdgeSet.has(`${toIdx}:${fromIdx}`)) {
        const nx = -uy;
        const ny = ux;
        const phase = ((fromIdx * 73856093) ^ (toIdx * 19349663)) & 1 ? 1 : -1;
        const sep = 6.2 * phase;
        if (movableNodeIndices.has(fromIdx)) {
          fx[fromIdx] += nx * sep;
          fy[fromIdx] += ny * sep;
        }
        if (movableNodeIndices.has(toIdx)) {
          fx[toIdx] += nx * sep;
          fy[toIdx] += ny * sep;
        }
      }
    }

    if ((iter + 1) % 3 === 0) {
      applyCrossingNudges(indexedEdges, centersX, centersY, movableNodeIndices, fx, fy, Math.max(1.5, temperature * 0.14));
    }

    for (const node of movableNodes) {
      const idx = nodeIndexById.get(node.id);
      if (idx === undefined) {
        continue;
      }

      const anchor = anchors.get(node.id);
      if (anchor) {
        const anchorX = anchor.x + node.width / 2;
        const anchorY = anchor.y + node.height / 2;
        fx[idx] += (anchorX - centersX[idx]) * 0.038;
        fy[idx] += (anchorY - centersY[idx]) * 0.038;
      }

      const incident = incidentEdgeIndices[idx];
      if (incident.length > 0) {
        let sumX = 0;
        let sumY = 0;
        for (const edgeIdx of incident) {
          const edge = indexedEdges[edgeIdx];
          const otherIdx = edge.fromIndex === idx ? edge.toIndex : edge.fromIndex;
          sumX += centersX[otherIdx];
          sumY += centersY[otherIdx];
        }
        const avgX = sumX / incident.length;
        const avgY = sumY / incident.length;
        fx[idx] += (avgX - centersX[idx]) * 0.018;
        fy[idx] += (avgY - centersY[idx]) * 0.018;
      }
    }

    const maxStep = Math.max(2.2, temperature);
    for (const node of movableNodes) {
      const idx = nodeIndexById.get(node.id);
      if (idx === undefined) {
        continue;
      }

      let dx = fx[idx];
      let dy = fy[idx];
      const mag = Math.hypot(dx, dy);
      if (mag > maxStep) {
        const scale = maxStep / mag;
        dx *= scale;
        dy *= scale;
      }

      if (Math.abs(dx) < 0.01 && Math.abs(dy) < 0.01) {
        continue;
      }

      node.x += dx;
      node.y += dy;
      moved.add(node.id);
    }

    if ((iter + 1) % 2 === 0) {
      resolveNodeCollisions(movableNodes, Math.max(4, minGap * 0.85), 1);
    }

    temperature = Math.max(2.4, temperature * 0.91);
  }

  const localScore = (node: DiagramIr["nodes"][number], candX: number, candY: number): number => {
    let score = 0;

    for (const other of allNodes) {
      if (other.id === node.id) {
        continue;
      }

      const overlap = expandedOverlapArea(node, other, minGap, { x: candX, y: candY });
      if (overlap > 0) {
        score += overlap * 5.8;
      }
    }

    const center = nodeCenter(node, { x: candX, y: candY });
    const incident = incidentEdges.get(node.id) ?? [];
    for (const edge of incident) {
      const otherId = edge.fromId === node.id ? edge.toId : edge.fromId;
      const other = nodeById.get(otherId);
      if (!other) {
        continue;
      }

      const otherCenter = nodeCenter(other);
      const dist = Math.hypot(center.x - otherCenter.x, center.y - otherCenter.y);
      score += dist * 0.56;
      if (dist < 26) {
        score += (26 - dist) * 34;
      }

      const fromCenter = edge.fromId === node.id ? center : otherCenter;
      const toCenter = edge.toId === node.id ? center : otherCenter;
      score += directionalPenalty(fromCenter, toCenter, rankdir);
    }

    for (const edge of incident) {
      const segmentA = segmentForEdge(edge, nodeById, { nodeId: node.id, x: candX, y: candY });
      if (!segmentA) {
        continue;
      }

      for (const otherEdge of indexedEdges) {
        if (otherEdge === edge) {
          continue;
        }

        if (
          otherEdge.fromId === edge.fromId ||
          otherEdge.fromId === edge.toId ||
          otherEdge.toId === edge.fromId ||
          otherEdge.toId === edge.toId
        ) {
          continue;
        }

        const segmentB = segmentForEdge(otherEdge, nodeById);
        if (!segmentB) {
          continue;
        }

        if (segmentsCross(segmentA, segmentB)) {
          score += 760;
        }
      }
    }

    const anchor = anchors.get(node.id);
    if (anchor) {
      const drift = Math.hypot(candX - anchor.x, candY - anchor.y);
      score += drift * 0.38;
    }

    return score;
  };

  let step = clamp(Math.max(12, ir.config.layout.nodesep * 0.32), 10, 56);
  const candidateOffsets = (s: number): Array<{ dx: number; dy: number }> => [
    { dx: 0, dy: 0 },
    { dx: s, dy: 0 },
    { dx: -s, dy: 0 },
    { dx: 0, dy: s },
    { dx: 0, dy: -s },
    { dx: s * 0.7, dy: s * 0.7 },
    { dx: s * 0.7, dy: -s * 0.7 },
    { dx: -s * 0.7, dy: s * 0.7 },
    { dx: -s * 0.7, dy: -s * 0.7 },
  ];

  for (let round = 0; round < 5; round += 1) {
    for (const node of optimizationOrder) {
      const baseX = node.x;
      const baseY = node.y;
      let bestX = baseX;
      let bestY = baseY;
      let bestScore = localScore(node, baseX, baseY);

      for (const offset of candidateOffsets(step)) {
        const candX = baseX + offset.dx;
        const candY = baseY + offset.dy;
        const score = localScore(node, candX, candY);
        if (score + 1e-6 < bestScore) {
          bestScore = score;
          bestX = candX;
          bestY = candY;
        }
      }

      if (Math.abs(bestX - baseX) > 0.01 || Math.abs(bestY - baseY) > 0.01) {
        node.x = bestX;
        node.y = bestY;
        moved.add(node.id);
      }
    }

    resolveNodeCollisions(movableNodes, Math.max(4, minGap * 0.75), 1);
    step = Math.max(4, step * 0.7);
  }

  return moved;
}

function chooseEdgeSides(
  fromNode: DiagramIr["nodes"][number],
  toNode: DiagramIr["nodes"][number],
  rankdir: Rankdir,
  fixedStartSide: EdgeSide | undefined,
  fixedEndSide: EdgeSide | undefined,
  sideLoad: Map<string, number>,
  pairLoad: Map<string, number>,
): { startSide: EdgeSide; endSide: EdgeSide } {
  const startCandidates = fixedStartSide ? [fixedStartSide] : EDGE_SIDES;
  const endCandidates = fixedEndSide ? [fixedEndSide] : EDGE_SIDES;

  let bestStart: EdgeSide = startCandidates[0];
  let bestEnd: EdgeSide = endCandidates[0];
  let bestScore = Number.POSITIVE_INFINITY;

  for (const startSide of startCandidates) {
    for (const endSide of endCandidates) {
      const p0 = sideAnchor(fromNode, startSide);
      const p1 = sideAnchor(toNode, endSide);
      const dx = p1.x - p0.x;
      const dy = p1.y - p0.y;
      const dist = Math.max(1e-3, Math.hypot(dx, dy));

      let score = dist + directionalPenalty(p0, p1, rankdir) * 0.84;
      const startNormal = sideNormal(startSide);
      const endNormal = sideNormal(endSide);
      const startDot = startNormal.x * dx + startNormal.y * dy;
      const endDot = endNormal.x * (-dx) + endNormal.y * (-dy);
      if (startDot <= 0) {
        score += 48 + dist * 0.42;
      }
      if (endDot <= 0) {
        score += 48 + dist * 0.42;
      }

      if ((startSide === "T" || startSide === "B") !== (endSide === "T" || endSide === "B")) {
        score += 5;
      }

      score += (sideLoad.get(`${fromNode.id}:${startSide}`) ?? 0) * 18;
      score += (sideLoad.get(`${toNode.id}:${endSide}`) ?? 0) * 18;

      const pairKey = `${fromNode.id}->${toNode.id}:${startSide}${endSide}`;
      const reverseKey = `${toNode.id}->${fromNode.id}:${endSide}${startSide}`;
      score += (pairLoad.get(pairKey) ?? 0) * 86;
      score += (pairLoad.get(reverseKey) ?? 0) * 134;

      if (score < bestScore) {
        bestScore = score;
        bestStart = startSide;
        bestEnd = endSide;
      }
    }
  }

  return {
    startSide: bestStart,
    endSide: bestEnd,
  };
}

function inferEdgeSideHints(ir: DiagramIr, rankdir: Rankdir): void {
  const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));
  const sideLoad = new Map<string, number>();
  const pairLoad = new Map<string, number>();

  const sortableEdges = [...ir.edges].sort((a, b) => {
    const aFrom = nodeById.get(a.from);
    const aTo = nodeById.get(a.to);
    const bFrom = nodeById.get(b.from);
    const bTo = nodeById.get(b.to);

    const aDist = aFrom && aTo ? Math.hypot(aTo.x - aFrom.x, aTo.y - aFrom.y) : 0;
    const bDist = bFrom && bTo ? Math.hypot(bTo.x - bFrom.x, bTo.y - bFrom.y) : 0;
    if (Math.abs(aDist - bDist) > 1e-6) {
      return bDist - aDist;
    }
    return a.id.localeCompare(b.id);
  });

  for (const edge of sortableEdges) {
    const from = nodeById.get(edge.from);
    const to = nodeById.get(edge.to);
    if (!from || !to) {
      continue;
    }

    if (from.id === to.id) {
      edge.style.startSide = edge.style.startSide ?? "R";
      edge.style.endSide = edge.style.endSide ?? "T";
      continue;
    }

    if (from.isJunction || to.isJunction) {
      continue;
    }

    const sides = chooseEdgeSides(
      from,
      to,
      rankdir,
      edge.style.startSide,
      edge.style.endSide,
      sideLoad,
      pairLoad,
    );
    edge.style.startSide = sides.startSide;
    edge.style.endSide = sides.endSide;

    const startKey = `${from.id}:${sides.startSide}`;
    const endKey = `${to.id}:${sides.endSide}`;
    sideLoad.set(startKey, (sideLoad.get(startKey) ?? 0) + 1);
    sideLoad.set(endKey, (sideLoad.get(endKey) ?? 0) + 1);

    const pairKey = `${from.id}->${to.id}:${sides.startSide}${sides.endSide}`;
    pairLoad.set(pairKey, (pairLoad.get(pairKey) ?? 0) + 1);
  }
}

function simplifyPolyline(points: Point[]): Point[] {
  if (points.length <= 2) {
    return points;
  }

  const deduped: Point[] = [];
  for (const point of points) {
    const prev = deduped[deduped.length - 1];
    if (!prev || Math.hypot(prev.x - point.x, prev.y - point.y) > 1e-3) {
      deduped.push(point);
    }
  }

  if (deduped.length <= 2) {
    return deduped;
  }

  const simplified: Point[] = [deduped[0]];
  for (let i = 1; i < deduped.length - 1; i += 1) {
    const a = simplified[simplified.length - 1];
    const b = deduped[i];
    const c = deduped[i + 1];
    const area = Math.abs((b.x - a.x) * (c.y - a.y) - (b.y - a.y) * (c.x - a.x));
    if (area > 1e-3) {
      simplified.push(b);
    }
  }
  simplified.push(deduped[deduped.length - 1]);
  return simplified;
}

function rebuildEdgeRoutesFromSideAnchors(ir: DiagramIr): void {
  const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));

  for (const edge of ir.edges) {
    const from = nodeById.get(edge.from);
    const to = nodeById.get(edge.to);
    if (!from || !to) {
      edge.points = [];
      edge.labelPosition = undefined;
      continue;
    }

    const startSide = edge.style.startSide ?? "R";
    const endSide = edge.style.endSide ?? "L";

    if (from.id === to.id) {
      const loopExtent = Math.max(30, Math.min(160, Math.max(from.width, from.height) * 0.95));
      const start = sideAnchor(from, startSide);
      const end = sideAnchor(to, endSide);
      const startOut = sideOutwardAnchor(from, startSide, loopExtent);
      const endOut = sideOutwardAnchor(to, endSide, loopExtent);
      const bridge =
        startSide === "L" || startSide === "R"
          ? { x: startOut.x, y: endOut.y }
          : { x: endOut.x, y: startOut.y };
      const points = simplifyPolyline([start, startOut, bridge, endOut, end]);
      edge.points = points;
      edge.labelPosition = edge.label ? fallbackLabelPosition(points) : undefined;
      continue;
    }

    const start = sideAnchor(from, startSide);
    const end = sideAnchor(to, endSide);
    const startOut = sideOutwardAnchor(from, startSide, Math.max(12, Math.min(56, Math.min(from.width, from.height) * 0.32)));
    const endOut = sideOutwardAnchor(to, endSide, Math.max(12, Math.min(56, Math.min(to.width, to.height) * 0.32)));

    const startIsHorizontal = startSide === "L" || startSide === "R";
    const endIsHorizontal = endSide === "L" || endSide === "R";

    let middlePoints: Point[] = [];
    if (startIsHorizontal === endIsHorizontal) {
      if (startIsHorizontal) {
        const midX = (startOut.x + endOut.x) / 2;
        middlePoints = [
          { x: midX, y: startOut.y },
          { x: midX, y: endOut.y },
        ];
      } else {
        const midY = (startOut.y + endOut.y) / 2;
        middlePoints = [
          { x: startOut.x, y: midY },
          { x: endOut.x, y: midY },
        ];
      }
    } else if (startIsHorizontal) {
      middlePoints = [{ x: endOut.x, y: startOut.y }];
    } else {
      middlePoints = [{ x: startOut.x, y: endOut.y }];
    }

    const points = simplifyPolyline([start, startOut, ...middlePoints, endOut, end]);
    edge.points = points;
    edge.labelPosition = edge.label ? fallbackLabelPosition(points) : undefined;
  }
}

interface EdgeSegmentRef {
  edgeId: string;
  fromId: string;
  toId: string;
  segment: Segment;
}

function polylineLength(points: Point[]): number {
  if (points.length < 2) {
    return 0;
  }

  let total = 0;
  for (let i = 0; i < points.length - 1; i += 1) {
    total += Math.hypot(points[i + 1].x - points[i].x, points[i + 1].y - points[i].y);
  }
  return total;
}

function collectEdgeSegments(ir: DiagramIr): EdgeSegmentRef[] {
  const out: EdgeSegmentRef[] = [];
  for (const edge of ir.edges) {
    if (!edge.points || edge.points.length < 2) {
      continue;
    }

    for (let i = 0; i < edge.points.length - 1; i += 1) {
      const p0 = edge.points[i];
      const p1 = edge.points[i + 1];
      if (Math.hypot(p1.x - p0.x, p1.y - p0.y) < 1e-3) {
        continue;
      }
      out.push({
        edgeId: edge.id,
        fromId: edge.from,
        toId: edge.to,
        segment: {
          x1: p0.x,
          y1: p0.y,
          x2: p1.x,
          y2: p1.y,
        },
      });
    }
  }
  return out;
}

function pointInExpandedNode(point: Point, node: DiagramIr["nodes"][number], padding: number): boolean {
  return (
    point.x >= node.x - padding &&
    point.x <= node.x + node.width + padding &&
    point.y >= node.y - padding &&
    point.y <= node.y + node.height + padding
  );
}

function overlapSize(
  a: { x: number; y: number; width: number; height: number },
  b: { x: number; y: number; width: number; height: number },
): { overlapX: number; overlapY: number; area: number } {
  const overlapX = Math.max(0, Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x));
  const overlapY = Math.max(0, Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y));
  return {
    overlapX,
    overlapY,
    area: overlapX * overlapY,
  };
}

function buildTopLevelSubgraphMembers(ir: DiagramIr): Map<string, Set<string>> {
  const subgraphById = new Map(ir.subgraphs.map((subgraph) => [subgraph.id, subgraph]));
  const parentById = new Map(ir.subgraphs.map((subgraph) => [subgraph.id, subgraph.parentId]));
  const topLevelIds = new Set(
    ir.subgraphs
      .filter((subgraph) => !subgraph.parentId || !subgraphById.has(subgraph.parentId) || subgraph.parentId === subgraph.id)
      .map((subgraph) => subgraph.id),
  );

  const members = new Map<string, Set<string>>();
  for (const topId of topLevelIds) {
    members.set(topId, new Set<string>());
  }

  const findTopLevel = (subgraphId: string | undefined): string | undefined => {
    if (!subgraphId) {
      return undefined;
    }
    let current = subgraphId;
    const visited = new Set<string>();
    while (current && !visited.has(current)) {
      visited.add(current);
      if (topLevelIds.has(current)) {
        return current;
      }
      const next = parentById.get(current);
      if (!next || next === current) {
        return undefined;
      }
      current = next;
    }
    return undefined;
  };

  for (const node of ir.nodes) {
    const topId = findTopLevel(node.subgraphId);
    if (!topId) {
      continue;
    }
    const set = members.get(topId);
    if (set) {
      set.add(node.id);
    }
  }

  return members;
}

function translateNodeSet(ir: DiagramIr, nodeIds: Set<string>, dx: number, dy: number): boolean {
  if (Math.abs(dx) < 1e-6 && Math.abs(dy) < 1e-6) {
    return false;
  }

  let moved = false;
  for (const node of ir.nodes) {
    if (!nodeIds.has(node.id)) {
      continue;
    }
    node.x += dx;
    node.y += dy;
    moved = true;
  }
  return moved;
}

function enforceTopLevelSubgraphSeparation(ir: DiagramIr, passes: number = 8): boolean {
  if (ir.subgraphs.length < 2) {
    return false;
  }

  const subgraphById = new Map(ir.subgraphs.map((subgraph) => [subgraph.id, subgraph]));
  const topLevelSubgraphs = ir.subgraphs.filter(
    (subgraph) => !subgraph.parentId || !subgraphById.has(subgraph.parentId) || subgraph.parentId === subgraph.id,
  );
  if (topLevelSubgraphs.length < 2) {
    return false;
  }

  const memberSets = buildTopLevelSubgraphMembers(ir);
  recomputeSubgraphBounds(ir);
  recomputeBounds(ir);

  const targetGap = clamp(Math.min(ir.config.layout.nodesep, ir.config.layout.ranksep) * 0.22, 8, 22);
  let moved = false;

  for (let pass = 0; pass < passes; pass += 1) {
    let movedThisPass = false;

    for (let i = 0; i < topLevelSubgraphs.length; i += 1) {
      for (let j = i + 1; j < topLevelSubgraphs.length; j += 1) {
        const a = topLevelSubgraphs[i];
        const b = topLevelSubgraphs[j];
        const overlap = overlapSize(a, b);
        if (overlap.overlapX <= 0 || overlap.overlapY <= 0) {
          continue;
        }

        const aMembers = memberSets.get(a.id);
        const bMembers = memberSets.get(b.id);
        if (!aMembers || !bMembers || aMembers.size === 0 || bMembers.size === 0) {
          continue;
        }

        const aCx = a.x + a.width / 2;
        const aCy = a.y + a.height / 2;
        const bCx = b.x + b.width / 2;
        const bCy = b.y + b.height / 2;

        if (overlap.overlapX < overlap.overlapY) {
          const push = overlap.overlapX / 2 + targetGap / 2;
          const sign = aCx <= bCx ? -1 : 1;
          const movedA = translateNodeSet(ir, aMembers, sign * push, 0);
          const movedB = translateNodeSet(ir, bMembers, -sign * push, 0);
          movedThisPass = movedThisPass || movedA || movedB;
        } else {
          const push = overlap.overlapY / 2 + targetGap / 2;
          const sign = aCy <= bCy ? -1 : 1;
          const movedA = translateNodeSet(ir, aMembers, 0, sign * push);
          const movedB = translateNodeSet(ir, bMembers, 0, -sign * push);
          movedThisPass = movedThisPass || movedA || movedB;
        }
      }
    }

    if (!movedThisPass) {
      break;
    }

    moved = true;
    recomputeSubgraphBounds(ir);
    recomputeBounds(ir);
  }

  return moved;
}

function scoreLayoutQuality(ir: DiagramIr, rankdir: Rankdir): number {
  const nonJunctionNodes = ir.nodes.filter((node) => !node.isJunction);
  let nodeOverlapPenalty = 0;
  const overlapGap = Math.max(6, Math.min(20, ir.config.layout.nodesep * 0.18));
  for (let i = 0; i < nonJunctionNodes.length; i += 1) {
    for (let j = i + 1; j < nonJunctionNodes.length; j += 1) {
      nodeOverlapPenalty += expandedOverlapArea(nonJunctionNodes[i], nonJunctionNodes[j], overlapGap);
    }
  }

  const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));
  let edgeLengthPenalty = 0;
  let bendPenalty = 0;
  let backwardPenalty = 0;
  let edgeNodePenalty = 0;
  let subgraphOverlapPenalty = 0;

  for (const edge of ir.edges) {
    if (!edge.points || edge.points.length < 2) {
      continue;
    }
    edgeLengthPenalty += polylineLength(edge.points);
    bendPenalty += Math.max(0, edge.points.length - 2);

    const fromNode = nodeById.get(edge.from);
    const toNode = nodeById.get(edge.to);
    if (fromNode && toNode && fromNode.id !== toNode.id) {
      backwardPenalty += directionalPenalty(nodeCenter(fromNode), nodeCenter(toNode), rankdir);
    }

    for (let i = 0; i < edge.points.length - 1; i += 1) {
      const p0 = edge.points[i];
      const p1 = edge.points[i + 1];
      const mid: Point = { x: (p0.x + p1.x) / 2, y: (p0.y + p1.y) / 2 };
      for (const node of nonJunctionNodes) {
        if (node.id === edge.from || node.id === edge.to) {
          continue;
        }
        if (pointInExpandedNode(mid, node, 4)) {
          edgeNodePenalty += 1;
          break;
        }
      }
    }
  }

  for (let i = 0; i < ir.subgraphs.length; i += 1) {
    for (let j = i + 1; j < ir.subgraphs.length; j += 1) {
      const overlap = overlapSize(ir.subgraphs[i], ir.subgraphs[j]);
      subgraphOverlapPenalty += overlap.area;
    }
  }

  const segments = collectEdgeSegments(ir);
  let crossingPenalty = 0;
  for (let i = 0; i < segments.length; i += 1) {
    for (let j = i + 1; j < segments.length; j += 1) {
      const a = segments[i];
      const b = segments[j];
      if (
        a.fromId === b.fromId ||
        a.fromId === b.toId ||
        a.toId === b.fromId ||
        a.toId === b.toId ||
        a.edgeId === b.edgeId
      ) {
        continue;
      }
      if (segmentsCross(a.segment, b.segment)) {
        crossingPenalty += 1;
      }
    }
  }

  const spreadPenalty = ir.bounds.width * 0.024 + ir.bounds.height * 0.024;

  return (
    nodeOverlapPenalty * 9.2 +
    subgraphOverlapPenalty * 22 +
    crossingPenalty * 1280 +
    edgeNodePenalty * 620 +
    backwardPenalty * 4.4 +
    edgeLengthPenalty * 0.11 +
    bendPenalty * 22 +
    spreadPenalty
  );
}

function hashString32(text: string): number {
  let hash = 2166136261;
  for (let i = 0; i < text.length; i += 1) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return hash >>> 0;
}

function deterministicNoise(nodeId: string, seed: number, salt: number): number {
  const hash = hashString32(`${seed}:${salt}:${nodeId}`);
  return ((hash % 20001) / 10000) - 1;
}

function applyDeterministicJitter(ir: DiagramIr, pinnedNodeIds: Set<string>, rankdir: Rankdir, seed: number): void {
  const movable = ir.nodes.filter((node) => !node.isJunction && !pinnedNodeIds.has(node.id));
  if (movable.length === 0) {
    return;
  }

  const baseAmplitude = clamp(ir.config.layout.nodesep * 0.24, 7, 36);
  const minorAmplitude = clamp(baseAmplitude * 0.58, 4, 20);
  for (const node of movable) {
    const nMain = deterministicNoise(node.id, seed, 17);
    const nMinor = deterministicNoise(node.id, seed, 53);

    if (rankdir === "TB" || rankdir === "BT") {
      node.x += nMain * baseAmplitude;
      node.y += nMinor * minorAmplitude;
    } else {
      node.x += nMinor * minorAmplitude;
      node.y += nMain * baseAmplitude;
    }
  }
}

function optimizeLayoutWithRestarts(ir: DiagramIr, rankdir: Rankdir, pinnedNodeIds: Set<string>): void {
  const baseState = captureLayoutState(ir);
  const candidateCount = clamp(Math.round(3 + Math.sqrt(Math.max(1, ir.nodes.length)) * 0.8), 3, 10);

  let bestState = captureLayoutState(ir);
  let bestScore = Number.POSITIVE_INFINITY;

  for (let candidate = 0; candidate < candidateCount; candidate += 1) {
    restoreLayoutState(ir, baseState);

    if (candidate > 0) {
      applyDeterministicJitter(ir, pinnedNodeIds, rankdir, candidate);
    }

    optimizeNodePlacement(ir, rankdir, pinnedNodeIds);
    inferEdgeSideHints(ir, rankdir);
    rebuildEdgeRoutesFromSideAnchors(ir);
    recomputeSubgraphBounds(ir);
    recomputeBounds(ir);
    const movedBySubgraphConstraint = enforceTopLevelSubgraphSeparation(ir, 8);
    if (movedBySubgraphConstraint) {
      inferEdgeSideHints(ir, rankdir);
      rebuildEdgeRoutesFromSideAnchors(ir);
      recomputeSubgraphBounds(ir);
      recomputeBounds(ir);
    }

    const score = scoreLayoutQuality(ir, rankdir);
    if (score + 1e-6 < bestScore) {
      bestScore = score;
      bestState = captureLayoutState(ir);
    }
  }

  restoreLayoutState(ir, bestState);
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

  const rankdir = toRankdir(ir.meta.direction);
  const movedByJunction = enforceJunctionSidePlacement(ir);
  optimizeLayoutWithRestarts(ir, rankdir, movedByJunction);
  const movedBySubgraphConstraint = enforceTopLevelSubgraphSeparation(ir, 10);
  if (movedBySubgraphConstraint) {
    inferEdgeSideHints(ir, rankdir);
    rebuildEdgeRoutesFromSideAnchors(ir);
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
    edgeSides: new Map(
      ir.edges.map((edge) => [
        edge.id,
        {
          startSide: edge.style.startSide,
          endSide: edge.style.endSide,
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
    const sides = state.edgeSides.get(edge.id);
    if (route) {
      edge.points = route.points.map((point) => ({ ...point }));
      edge.labelPosition = route.labelPosition ? { ...route.labelPosition } : undefined;
    }
    if (sides) {
      edge.style.startSide = sides.startSide;
      edge.style.endSide = sides.endSide;
    }
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
