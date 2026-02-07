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

interface EdgeGeom {
  fromId: string;
  toId: string;
}

interface Segment {
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

function nodeCenter(node: DiagramIr["nodes"][number], override?: { x?: number; y?: number }): Point {
  return {
    x: (override?.x ?? node.x) + node.width / 2,
    y: (override?.y ?? node.y) + node.height / 2,
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

function directionalPenalty(from: Point, to: Point, rankdir: "TB" | "BT" | "LR" | "RL"): number {
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

function optimizeNodePlacement(ir: DiagramIr, rankdir: "TB" | "BT" | "LR" | "RL", pinnedNodeIds: Set<string>): Set<string> {
  const moved = new Set<string>();
  const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));
  const movableNodes = ir.nodes.filter((node) => !node.isJunction && !pinnedNodeIds.has(node.id));
  if (movableNodes.length <= 1) {
    return moved;
  }

  const edges: EdgeGeom[] = ir.edges
    .map((edge) => ({ fromId: edge.from, toId: edge.to }))
    .filter((edge) => nodeById.has(edge.fromId) && nodeById.has(edge.toId) && edge.fromId !== edge.toId);

  const incidentEdges = new Map<string, EdgeGeom[]>();
  for (const edge of edges) {
    const fromList = incidentEdges.get(edge.fromId) ?? [];
    fromList.push(edge);
    incidentEdges.set(edge.fromId, fromList);
    const toList = incidentEdges.get(edge.toId) ?? [];
    toList.push(edge);
    incidentEdges.set(edge.toId, toList);
  }

  const anchors = new Map(movableNodes.map((node) => [node.id, { x: node.x, y: node.y }]));
  const optimizationOrder = [...movableNodes].sort(
    (a, b) => (incidentEdges.get(b.id)?.length ?? 0) - (incidentEdges.get(a.id)?.length ?? 0),
  );

  const allNonJunction = ir.nodes.filter((node) => !node.isJunction);
  const minGap = Math.max(8, Math.min(24, ir.config.layout.nodesep * 0.18));

  const localScore = (node: DiagramIr["nodes"][number], candX: number, candY: number): number => {
    let score = 0;

    for (const other of allNonJunction) {
      if (other.id === node.id) {
        continue;
      }

      const overlap = expandedOverlapArea(node, other, minGap, { x: candX, y: candY });
      if (overlap > 0) {
        score += overlap * 5.2;
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
      score += dist * 0.52;
      if (dist < 28) {
        score += (28 - dist) * 26;
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

      for (const otherEdge of edges) {
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
          score += 680;
        }
      }
    }

    const anchor = anchors.get(node.id);
    if (anchor) {
      const drift = Math.hypot(candX - anchor.x, candY - anchor.y);
      score += drift * 0.36;
    }

    return score;
  };

  let step = clamp(Math.max(18, ir.config.layout.nodesep * 0.44), 18, 84);
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

  for (let round = 0; round < 7; round += 1) {
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

    resolveNodeCollisions(movableNodes, minGap, 2);
    step = Math.max(6, step * 0.72);
  }

  return moved;
}

function rebuildEdgeRoutesFromNodeCenters(ir: DiagramIr): void {
  const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));

  for (const edge of ir.edges) {
    const from = nodeById.get(edge.from);
    const to = nodeById.get(edge.to);
    if (!from || !to || from.id === to.id) {
      edge.points = [];
      edge.labelPosition = undefined;
      continue;
    }

    const p0 = nodeCenter(from);
    const p1 = nodeCenter(to);
    edge.points = [p0, p1];
    edge.labelPosition = edge.label
      ? {
          x: (p0.x + p1.x) / 2,
          y: (p0.y + p1.y) / 2,
        }
      : undefined;
  }
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
  const movedByOptimization = optimizeNodePlacement(ir, toRankdir(ir.meta.direction), movedByJunction);
  const movedNodeIds = new Set<string>([...movedByJunction, ...movedByOptimization]);

  if (movedNodeIds.size > 0) {
    rebuildEdgeRoutesFromNodeCenters(ir);
  } else {
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
