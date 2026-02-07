import type { DiagramDirection, DiagramIr, Point } from "../types.js";

type Rankdir = "TB" | "BT" | "LR" | "RL";

interface SegmentRef {
  edgeId: string;
  fromId: string;
  toId: string;
  x1: number;
  y1: number;
  x2: number;
  y2: number;
}

export interface ReadabilityMetrics {
  nodeOverlapArea: number;
  subgraphOverlapArea: number;
  nodeOutOfSubgraphCount: number;
  edgeCrossings: number;
  edgeThroughNodeCount: number;
  edgeLabelNodeOverlapCount: number;
  edgeLabelLabelOverlapCount: number;
  edgeLabelDistanceMean: number;
  lowContrastCount: number;
  textOverflowRiskCount: number;
  totalEdgeBends: number;
  directionalBackflowPenalty: number;
  occupancyRatio: number;
  occupancyPenalty: number;
}

export interface ReadabilityEvaluation {
  penalty: number;
  score: number;
  metrics: ReadabilityMetrics;
}

interface Rect {
  x: number;
  y: number;
  w: number;
  h: number;
}

function toRankdir(direction: DiagramDirection): Rankdir {
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

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function orientation(ax: number, ay: number, bx: number, by: number, cx: number, cy: number): number {
  return (bx - ax) * (cy - ay) - (by - ay) * (cx - ax);
}

function segmentsCross(a: SegmentRef, b: SegmentRef): boolean {
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
    return forward >= minForward ? 0 : minForward - forward;
  }
  if (rankdir === "BT") {
    const forward = from.y - to.y;
    return forward >= minForward ? 0 : minForward - forward;
  }
  if (rankdir === "LR") {
    const forward = to.x - from.x;
    return forward >= minForward ? 0 : minForward - forward;
  }
  const forward = from.x - to.x;
  return forward >= minForward ? 0 : minForward - forward;
}

function expandedOverlapArea(a: Rect, b: Rect, gap: number): number {
  const ax0 = a.x - gap;
  const ay0 = a.y - gap;
  const ax1 = a.x + a.w + gap;
  const ay1 = a.y + a.h + gap;
  const bx0 = b.x - gap;
  const by0 = b.y - gap;
  const bx1 = b.x + b.w + gap;
  const by1 = b.y + b.h + gap;
  const ox = Math.max(0, Math.min(ax1, bx1) - Math.max(ax0, bx0));
  const oy = Math.max(0, Math.min(ay1, by1) - Math.max(ay0, by0));
  return ox * oy;
}

function segmentToPointDistance(seg: SegmentRef, p: Point): number {
  const vx = seg.x2 - seg.x1;
  const vy = seg.y2 - seg.y1;
  const wx = p.x - seg.x1;
  const wy = p.y - seg.y1;
  const c1 = vx * wx + vy * wy;
  if (c1 <= 0) {
    return Math.hypot(p.x - seg.x1, p.y - seg.y1);
  }
  const c2 = vx * vx + vy * vy;
  if (c2 <= 1e-6) {
    return Math.hypot(p.x - seg.x1, p.y - seg.y1);
  }
  const t = Math.min(1, Math.max(0, c1 / c2));
  const px = seg.x1 + vx * t;
  const py = seg.y1 + vy * t;
  return Math.hypot(p.x - px, p.y - py);
}

function rectOverlapArea(a: Rect, b: Rect): number {
  const ox = Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x));
  const oy = Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y));
  return ox * oy;
}

function parseHexColor(input: string | undefined): { r: number; g: number; b: number } | undefined {
  const raw = String(input ?? "").trim().replace(/^#/u, "");
  if (!/^[0-9a-fA-F]{6}$/u.test(raw)) {
    return undefined;
  }
  return {
    r: Number.parseInt(raw.slice(0, 2), 16),
    g: Number.parseInt(raw.slice(2, 4), 16),
    b: Number.parseInt(raw.slice(4, 6), 16),
  };
}

function linearizeSrgb(v: number): number {
  const x = v / 255;
  if (x <= 0.03928) {
    return x / 12.92;
  }
  return ((x + 0.055) / 1.055) ** 2.4;
}

function luminance(rgb: { r: number; g: number; b: number }): number {
  return 0.2126 * linearizeSrgb(rgb.r) + 0.7152 * linearizeSrgb(rgb.g) + 0.0722 * linearizeSrgb(rgb.b);
}

function contrastRatio(fg: string | undefined, bg: string | undefined): number {
  const fgRgb = parseHexColor(fg);
  const bgRgb = parseHexColor(bg);
  if (!fgRgb || !bgRgb) {
    return 4.5;
  }
  const l1 = luminance(fgRgb);
  const l2 = luminance(bgRgb);
  const bright = Math.max(l1, l2);
  const dark = Math.min(l1, l2);
  return (bright + 0.05) / (dark + 0.05);
}

function estimateTextBoxPx(text: string, fontSize: number, maxWidth: number): { w: number; h: number } {
  const chars = Math.max(1, Array.from(text).length);
  const charW = Math.max(4, fontSize * 0.56);
  const rawW = chars * charW + 10;
  const w = clamp(rawW, 32, Math.max(32, maxWidth));
  const cols = Math.max(1, Math.floor(w / charW));
  const rows = Math.max(1, Math.ceil(chars / cols));
  const h = rows * fontSize * 1.28 + 8;
  return { w, h };
}

function collectSegments(ir: DiagramIr): SegmentRef[] {
  const segments: SegmentRef[] = [];
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
      segments.push({
        edgeId: edge.id,
        fromId: edge.from,
        toId: edge.to,
        x1: p0.x,
        y1: p0.y,
        x2: p1.x,
        y2: p1.y,
      });
    }
  }
  return segments;
}

export function evaluateReadability(ir: DiagramIr): ReadabilityEvaluation {
  const rankdir = toRankdir(ir.meta.direction);
  const nodes = ir.nodes.filter((node) => !node.isJunction);
  const nodeRects = nodes.map((node) => ({
    id: node.id,
    x: node.x,
    y: node.y,
    w: node.width,
    h: node.height,
  }));

  let nodeOverlapArea = 0;
  for (let i = 0; i < nodeRects.length; i += 1) {
    for (let j = i + 1; j < nodeRects.length; j += 1) {
      nodeOverlapArea += expandedOverlapArea(nodeRects[i], nodeRects[j], 2);
    }
  }

  let subgraphOverlapArea = 0;
  const subgraphs = ir.subgraphs || [];
  for (let i = 0; i < subgraphs.length; i += 1) {
    for (let j = i + 1; j < subgraphs.length; j += 1) {
      subgraphOverlapArea += rectOverlapArea(
        { x: subgraphs[i].x, y: subgraphs[i].y, w: subgraphs[i].width, h: subgraphs[i].height },
        { x: subgraphs[j].x, y: subgraphs[j].y, w: subgraphs[j].width, h: subgraphs[j].height },
      );
    }
  }

  const subgraphById = new Map(subgraphs.map((subgraph) => [subgraph.id, subgraph]));
  let nodeOutOfSubgraphCount = 0;
  for (const node of nodes) {
    const subgraphId = node.subgraphId;
    if (!subgraphId) {
      continue;
    }
    const subgraph = subgraphById.get(subgraphId);
    if (!subgraph) {
      continue;
    }
    const margin = 2;
    const inside =
      node.x >= subgraph.x - margin &&
      node.y >= subgraph.y - margin &&
      node.x + node.width <= subgraph.x + subgraph.width + margin &&
      node.y + node.height <= subgraph.y + subgraph.height + margin;
    if (!inside) {
      nodeOutOfSubgraphCount += 1;
    }
  }

  const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));
  const segments = collectSegments(ir);

  let edgeCrossings = 0;
  for (let i = 0; i < segments.length; i += 1) {
    for (let j = i + 1; j < segments.length; j += 1) {
      const a = segments[i];
      const b = segments[j];
      if (
        a.edgeId === b.edgeId ||
        a.fromId === b.fromId ||
        a.fromId === b.toId ||
        a.toId === b.fromId ||
        a.toId === b.toId
      ) {
        continue;
      }
      if (segmentsCross(a, b)) {
        edgeCrossings += 1;
      }
    }
  }

  let edgeThroughNodeCount = 0;
  for (const segment of segments) {
    const mid = {
      x: (segment.x1 + segment.x2) / 2,
      y: (segment.y1 + segment.y2) / 2,
    };
    for (const node of nodes) {
      if (node.id === segment.fromId || node.id === segment.toId) {
        continue;
      }
      if (mid.x >= node.x - 3 && mid.x <= node.x + node.width + 3 && mid.y >= node.y - 3 && mid.y <= node.y + node.height + 3) {
        edgeThroughNodeCount += 1;
        break;
      }
    }
  }

  const labelRects: Array<{ edgeId: string; rect: Rect }> = [];
  let edgeLabelNodeOverlapCount = 0;
  let edgeLabelDistanceMean = 0;
  let labelDistCount = 0;

  for (const edge of ir.edges) {
    if (!edge.label || !edge.labelPosition) {
      continue;
    }

    const font = clamp(edge.style.fontSize || 11, 8, 24);
    const box = estimateTextBoxPx(edge.label, font, 220);
    const rect: Rect = {
      x: edge.labelPosition.x - box.w / 2,
      y: edge.labelPosition.y - box.h / 2,
      w: box.w,
      h: box.h,
    };

    labelRects.push({ edgeId: edge.id, rect });

    for (const node of nodes) {
      if (node.id === edge.from || node.id === edge.to) {
        continue;
      }
      if (rectOverlapArea(rect, { x: node.x, y: node.y, w: node.width, h: node.height }) > 0) {
        edgeLabelNodeOverlapCount += 1;
        break;
      }
    }

    const ownSegments = segments.filter((segment) => segment.edgeId === edge.id);
    if (ownSegments.length > 0) {
      let minDist = Number.POSITIVE_INFINITY;
      for (const seg of ownSegments) {
        minDist = Math.min(minDist, segmentToPointDistance(seg, edge.labelPosition));
      }
      if (Number.isFinite(minDist)) {
        edgeLabelDistanceMean += minDist;
        labelDistCount += 1;
      }
    }
  }

  let edgeLabelLabelOverlapCount = 0;
  for (let i = 0; i < labelRects.length; i += 1) {
    for (let j = i + 1; j < labelRects.length; j += 1) {
      if (rectOverlapArea(labelRects[i].rect, labelRects[j].rect) > 0) {
        edgeLabelLabelOverlapCount += 1;
      }
    }
  }

  let lowContrastCount = 0;
  for (const node of nodes) {
    if (contrastRatio(node.style.text, node.style.fill) < 4.5) {
      lowContrastCount += 1;
    }
  }
  for (const subgraph of ir.subgraphs) {
    if (contrastRatio(subgraph.style.text, subgraph.style.fill) < 4.5) {
      lowContrastCount += 1;
    }
  }

  let textOverflowRiskCount = 0;
  for (const node of nodes) {
    if (!node.label) {
      continue;
    }
    const est = estimateTextBoxPx(node.label, clamp(node.style.fontSize, 8, 24), Math.max(48, node.width - 10));
    if (est.h > node.height * 0.98 || est.w > node.width * 1.02) {
      textOverflowRiskCount += 1;
    }
  }

  let totalEdgeBends = 0;
  let directionalBackflowPenalty = 0;
  for (const edge of ir.edges) {
    totalEdgeBends += Math.max(0, (edge.points?.length ?? 0) - 2);
    const from = nodeById.get(edge.from);
    const to = nodeById.get(edge.to);
    if (!from || !to) {
      continue;
    }
    directionalBackflowPenalty += directionalPenalty(
      { x: from.x + from.width / 2, y: from.y + from.height / 2 },
      { x: to.x + to.width / 2, y: to.y + to.height / 2 },
      rankdir,
    );
  }

  const totalNodeArea = nodes.reduce((sum, node) => sum + node.width * node.height, 0);
  const boundsArea = Math.max(1, ir.bounds.width * ir.bounds.height);
  const occupancyRatio = totalNodeArea / boundsArea;
  const occupancyPenalty =
    occupancyRatio > 0.16 ? (occupancyRatio - 0.16) : (0.16 - occupancyRatio) * 0.18;

  const metrics: ReadabilityMetrics = {
    nodeOverlapArea,
    subgraphOverlapArea,
    nodeOutOfSubgraphCount,
    edgeCrossings,
    edgeThroughNodeCount,
    edgeLabelNodeOverlapCount,
    edgeLabelLabelOverlapCount,
    edgeLabelDistanceMean: labelDistCount > 0 ? edgeLabelDistanceMean / labelDistCount : 0,
    lowContrastCount,
    textOverflowRiskCount,
    totalEdgeBends,
    directionalBackflowPenalty,
    occupancyRatio,
    occupancyPenalty,
  };

  const penalty =
    metrics.nodeOverlapArea * 8.8 +
    metrics.subgraphOverlapArea * 26 +
    metrics.nodeOutOfSubgraphCount * 780 +
    metrics.edgeCrossings * 1300 +
    metrics.edgeThroughNodeCount * 520 +
    metrics.edgeLabelNodeOverlapCount * 460 +
    metrics.edgeLabelLabelOverlapCount * 560 +
    Math.max(0, metrics.edgeLabelDistanceMean - 20) * 12 +
    metrics.lowContrastCount * 210 +
    metrics.textOverflowRiskCount * 300 +
    metrics.totalEdgeBends * 20 +
    metrics.directionalBackflowPenalty * 4.5 +
    metrics.occupancyPenalty * 2400;

  const score = clamp(100 - Math.log10(1 + penalty) * 18, 0, 100);
  return {
    penalty,
    score,
    metrics,
  };
}
