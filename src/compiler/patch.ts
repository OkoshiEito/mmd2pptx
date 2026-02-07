import yaml from "js-yaml";
import type { DiagramIr, DiagramPatch, EdgePatch, NodePatch, SubgraphPatch } from "../types.js";
import { recomputeBounds, recomputeSubgraphBounds } from "./geometry.js";

function asNumber(input: unknown): number | undefined {
  return typeof input === "number" && Number.isFinite(input) ? input : undefined;
}

function clamp(value: number, min?: number, max?: number): number {
  let next = value;
  if (min !== undefined) {
    next = Math.max(min, next);
  }
  if (max !== undefined) {
    next = Math.min(max, next);
  }
  return next;
}

function applyNodeSizePatch(node: DiagramIr["nodes"][number], patch: NodePatch): void {
  const minW = asNumber(patch.minW);
  const minH = asNumber(patch.minH);
  const maxW = asNumber(patch.maxW);
  const maxH = asNumber(patch.maxH);

  if (asNumber(patch.w) !== undefined) {
    node.width = clamp(asNumber(patch.w) as number, minW, maxW);
  } else {
    node.width = clamp(node.width, minW, maxW);
  }

  if (asNumber(patch.h) !== undefined) {
    node.height = clamp(asNumber(patch.h) as number, minH, maxH);
  } else {
    node.height = clamp(node.height, minH, maxH);
  }
}

function applyNodePositionPatch(node: DiagramIr["nodes"][number], patch: NodePatch): void {
  const x = asNumber(patch.x);
  const y = asNumber(patch.y);
  const dx = asNumber(patch.dx) ?? 0;
  const dy = asNumber(patch.dy) ?? 0;

  if (x !== undefined) {
    node.x = x;
  }

  if (y !== undefined) {
    node.y = y;
  }

  if (dx !== 0) {
    node.x += dx;
  }

  if (dy !== 0) {
    node.y += dy;
  }
}

function applyEdgePatch(edge: DiagramIr["edges"][number], patch: EdgePatch): void {
  if (Array.isArray(patch.points) && patch.points.length > 1) {
    edge.points = patch.points
      .map((point) => {
        const x = asNumber(point.x);
        const y = asNumber(point.y);
        if (x === undefined || y === undefined) {
          return null;
        }
        return { x, y };
      })
      .filter((point): point is { x: number; y: number } => Boolean(point));
  }

  if (!edge.labelPosition) {
    return;
  }

  const labelDx = asNumber(patch.labelDx) ?? 0;
  const labelDy = asNumber(patch.labelDy) ?? 0;
  edge.labelPosition.x += labelDx;
  edge.labelPosition.y += labelDy;
}

function applySubgraphPatch(subgraph: DiagramIr["subgraphs"][number], patch: SubgraphPatch): void {
  const padding = asNumber(patch.padding);
  if (padding !== undefined) {
    subgraph.style.padding = padding;
  }

  const x = asNumber(patch.x);
  const y = asNumber(patch.y);
  const w = asNumber(patch.w);
  const h = asNumber(patch.h);

  if (x !== undefined) {
    subgraph.x = x;
  }

  if (y !== undefined) {
    subgraph.y = y;
  }

  if (w !== undefined) {
    subgraph.width = w;
  }

  if (h !== undefined) {
    subgraph.height = h;
  }
}

export function parsePatchYaml(raw: string): DiagramPatch {
  const loaded = yaml.load(raw);
  if (!loaded || typeof loaded !== "object") {
    return {};
  }
  return loaded as DiagramPatch;
}

export function applyPatchPreLayout(ir: DiagramIr, patch?: DiagramPatch): void {
  if (!patch) {
    return;
  }

  if (patch.layout) {
    ir.config.layout = {
      ...ir.config.layout,
      ...patch.layout,
    };
  }

  if (patch.renderer?.fontFamily) {
    ir.config.fontFamily = patch.renderer.fontFamily;
  }

  if (patch.renderer?.lang) {
    ir.config.lang = patch.renderer.lang;
  }

  if (patch.nodes) {
    const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));
    for (const [nodeId, nodePatch] of Object.entries(patch.nodes)) {
      const node = nodeById.get(nodeId);
      if (!node) {
        continue;
      }
      applyNodeSizePatch(node, nodePatch);
    }
  }

  if (patch.subgraphs) {
    const subgraphById = new Map(ir.subgraphs.map((subgraph) => [subgraph.id, subgraph]));
    for (const [subgraphId, subgraphPatch] of Object.entries(patch.subgraphs)) {
      const subgraph = subgraphById.get(subgraphId);
      if (!subgraph) {
        continue;
      }
      const padding = asNumber(subgraphPatch.padding);
      if (padding !== undefined) {
        subgraph.style.padding = padding;
      }
    }
  }
}

export function applyPatchPostLayout(ir: DiagramIr, patch?: DiagramPatch): void {
  if (!patch) {
    recomputeSubgraphBounds(ir);
    recomputeBounds(ir);
    return;
  }

  if (patch.nodes) {
    const nodeById = new Map(ir.nodes.map((node) => [node.id, node]));
    for (const [nodeId, nodePatch] of Object.entries(patch.nodes)) {
      const node = nodeById.get(nodeId);
      if (!node) {
        continue;
      }
      applyNodePositionPatch(node, nodePatch);
    }
  }

  if (patch.edges) {
    const edgeById = new Map(ir.edges.map((edge) => [edge.id, edge]));
    for (const [edgeId, edgePatch] of Object.entries(patch.edges)) {
      const edge = edgeById.get(edgeId);
      if (!edge) {
        continue;
      }
      applyEdgePatch(edge, edgePatch);
    }
  }

  recomputeSubgraphBounds(ir);

  if (patch.subgraphs) {
    const subgraphById = new Map(ir.subgraphs.map((subgraph) => [subgraph.id, subgraph]));
    for (const [subgraphId, subgraphPatch] of Object.entries(patch.subgraphs)) {
      const subgraph = subgraphById.get(subgraphId);
      if (!subgraph) {
        continue;
      }
      applySubgraphPatch(subgraph, subgraphPatch);
    }
  }

  recomputeBounds(ir);
}
