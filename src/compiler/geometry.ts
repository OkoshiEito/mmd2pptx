import type { DiagramIr, IrNode, IrSubgraph } from "../types.js";

function titleTextUnits(line: string): number {
  let units = 0;
  for (const ch of line) {
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

function titleBandHeight(title: string): number {
  const lines = title
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean).length;

  if (lines <= 0) {
    return 0;
  }

  const wrappedLines = title
    .replace(/\r\n/g, "\n")
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean)
    .reduce((sum, line) => sum + Math.max(1, Math.ceil(titleTextUnits(line) / 34)), 0);

  // Reserve generous vertical space above subgraph members so long group descriptions
  // remain visible and do not collide with nodes.
  return Math.max(72, 16 + wrappedLines * 22);
}

export function recomputeSubgraphBounds(ir: DiagramIr): void {
  const nodeMap = new Map<string, IrNode>();
  for (const node of ir.nodes) {
    nodeMap.set(node.id, node);
  }

  for (const subgraph of ir.subgraphs) {
    const members = subgraph.nodeIds
      .map((id) => nodeMap.get(id))
      .filter((node): node is IrNode => Boolean(node));

    if (members.length === 0) {
      continue;
    }

    const minX = Math.min(...members.map((node) => node.x));
    const minY = Math.min(...members.map((node) => node.y));
    const maxX = Math.max(...members.map((node) => node.x + node.width));
    const maxY = Math.max(...members.map((node) => node.y + node.height));

    const padding = subgraph.style.padding;
    const titleBand = titleBandHeight(subgraph.title);
    subgraph.x = minX - padding;
    subgraph.y = minY - padding - titleBand;
    subgraph.width = maxX - minX + padding * 2;
    subgraph.height = maxY - minY + padding * 2 + titleBand;
  }
}

export function recomputeBounds(ir: DiagramIr): void {
  let minX = Number.POSITIVE_INFINITY;
  let minY = Number.POSITIVE_INFINITY;
  let maxX = Number.NEGATIVE_INFINITY;
  let maxY = Number.NEGATIVE_INFINITY;

  for (const node of ir.nodes) {
    minX = Math.min(minX, node.x);
    minY = Math.min(minY, node.y);
    maxX = Math.max(maxX, node.x + node.width);
    maxY = Math.max(maxY, node.y + node.height);
  }

  for (const subgraph of ir.subgraphs) {
    minX = Math.min(minX, subgraph.x);
    minY = Math.min(minY, subgraph.y);
    maxX = Math.max(maxX, subgraph.x + subgraph.width);
    maxY = Math.max(maxY, subgraph.y + subgraph.height);
  }

  for (const edge of ir.edges) {
    for (const point of edge.points) {
      minX = Math.min(minX, point.x);
      minY = Math.min(minY, point.y);
      maxX = Math.max(maxX, point.x);
      maxY = Math.max(maxY, point.y);
    }
  }

  if (!Number.isFinite(minX) || !Number.isFinite(minY) || !Number.isFinite(maxX) || !Number.isFinite(maxY)) {
    ir.bounds = {
      minX: 0,
      minY: 0,
      maxX: 0,
      maxY: 0,
      width: 0,
      height: 0,
    };
    return;
  }

  ir.bounds = {
    minX,
    minY,
    maxX,
    maxY,
    width: maxX - minX,
    height: maxY - minY,
  };
}

export function sortSubgraphsByArea(subgraphs: IrSubgraph[]): IrSubgraph[] {
  return [...subgraphs].sort((a, b) => a.width * a.height - b.width * b.height);
}
