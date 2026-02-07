import { DEFAULT_EDGE_STYLE, DEFAULT_LAYOUT_CONFIG, DEFAULT_NODE_STYLE, DEFAULT_SUBGRAPH_STYLE, emptyIr } from "./defaults.js";
import type { DiagramAst, DiagramIr, IrNode, ParsedEdge, ParsedNode } from "../types.js";

interface NormalizeOptions {
  fontFamily?: string;
  lang?: string;
}

function cloneNodeStyle() {
  return { ...DEFAULT_NODE_STYLE };
}

function cloneEdgeStyle() {
  return { ...DEFAULT_EDGE_STYLE };
}

function colorFrom(input?: string): string | undefined {
  if (!input) {
    return undefined;
  }

  const normalized = input.trim().replace(/^#/, "");
  if (/^[0-9a-f]{6}$/iu.test(normalized)) {
    return normalized.toUpperCase();
  }
  return undefined;
}

function numberFromCss(input?: string): number | undefined {
  if (!input) {
    return undefined;
  }

  const match = input.match(/-?[0-9]+(?:\.[0-9]+)?/u);
  if (!match) {
    return undefined;
  }
  const value = Number(match[0]);
  return Number.isFinite(value) ? value : undefined;
}

function applyStyleToNode(node: IrNode, styleMap?: Record<string, string>): void {
  if (!styleMap) {
    return;
  }

  const fill = colorFrom(styleMap.fill);
  if (fill) {
    node.style.fill = fill;
  }

  const stroke = colorFrom(styleMap.stroke);
  if (stroke) {
    node.style.stroke = stroke;
  }

  const text = colorFrom(styleMap.color);
  if (text) {
    node.style.text = text;
  }

  const strokeWidth = numberFromCss(styleMap["stroke-width"]);
  if (strokeWidth !== undefined) {
    node.style.strokeWidth = strokeWidth;
  }

  const fontSize = numberFromCss(styleMap["font-size"]);
  if (fontSize !== undefined) {
    node.style.fontSize = fontSize;
  }

  if (styleMap["font-weight"]) {
    const weight = styleMap["font-weight"].toLowerCase();
    node.style.bold = weight.includes("bold") || Number(weight) >= 600;
  }
}

function applyStyleToSubgraph(subgraph: DiagramIr["subgraphs"][number], styleMap?: Record<string, string>): void {
  if (!styleMap) {
    return;
  }

  const fill = colorFrom(styleMap.fill);
  if (fill) {
    subgraph.style.fill = fill;
  }

  const stroke = colorFrom(styleMap.stroke);
  if (stroke) {
    subgraph.style.stroke = stroke;
  }

  const text = colorFrom(styleMap.color);
  if (text) {
    subgraph.style.text = text;
  }

  const strokeWidth = numberFromCss(styleMap["stroke-width"]);
  if (strokeWidth !== undefined) {
    subgraph.style.strokeWidth = strokeWidth;
  }

  const dash = styleMap["stroke-dasharray"];
  if (dash) {
    const dashValue = numberFromCss(dash);
    if (dashValue !== undefined) {
      subgraph.style.dash = dashValue > 0 ? "dash" : "solid";
    }
  }
}

function slugify(value: string): string {
  return value
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .slice(0, 64);
}

function ensureNode(
  nodeById: Map<string, IrNode>,
  nodeOrder: string[],
  parsed: Pick<ParsedNode, "id" | "label" | "shape" | "subgraphId">,
): IrNode {
  const existing = nodeById.get(parsed.id);
  if (existing) {
    const nextLabel = parsed.label?.trim();
    if (nextLabel && (existing.label === existing.id || nextLabel.length >= existing.label.length)) {
      existing.label = nextLabel;
    }

    if (parsed.shape) {
      existing.shape = parsed.shape;
    }

    if (parsed.subgraphId) {
      existing.subgraphId = parsed.subgraphId;
    }

    return existing;
  }

  const node: IrNode = {
    id: parsed.id,
    label: parsed.label?.trim() || parsed.id,
    shape: parsed.shape ?? "rect",
    subgraphId: parsed.subgraphId,
    width: 0,
    height: 0,
    x: 0,
    y: 0,
    style: cloneNodeStyle(),
  };

  nodeById.set(node.id, node);
  nodeOrder.push(node.id);
  return node;
}

function edgeIdFrom(edge: ParsedEdge, index: number, counts: Map<string, number>): string {
  const base = slugify(`${edge.from}_${edge.to}_${edge.label ?? ""}_${edge.style}_${edge.arrow}`) || "edge";
  const next = (counts.get(base) ?? 0) + 1;
  counts.set(base, next);
  return `${base}_${String(index + 1).padStart(3, "0")}_${next}`;
}

export function normalizeDiagram(ast: DiagramAst, options: NormalizeOptions = {}): DiagramIr {
  const ir = emptyIr(ast.source);
  ir.meta.direction = ast.direction;
  if (ast.title?.trim()) {
    ir.meta.title = ast.title.trim();
  }
  ir.config.layout = { ...DEFAULT_LAYOUT_CONFIG };

  if (ast.layoutHints.rankSpacing !== undefined) {
    ir.config.layout.ranksep = ast.layoutHints.rankSpacing;
  }
  if (ast.layoutHints.nodeSpacing !== undefined) {
    ir.config.layout.nodesep = ast.layoutHints.nodeSpacing;
  }

  if (options.fontFamily) {
    ir.config.fontFamily = options.fontFamily;
  }

  if (options.lang) {
    ir.config.lang = options.lang;
  }

  const nodeById = new Map<string, IrNode>();
  const nodeOrder: string[] = [];

  for (const parsedNode of ast.nodes) {
    ensureNode(nodeById, nodeOrder, parsedNode);
  }

  const edgeIdCounter = new Map<string, number>();
  ir.edges = ast.edges.map((parsedEdge, index) => {
    ensureNode(nodeById, nodeOrder, { id: parsedEdge.from });
    ensureNode(nodeById, nodeOrder, { id: parsedEdge.to });

    return {
      id: edgeIdFrom(parsedEdge, index, edgeIdCounter),
      from: parsedEdge.from,
      to: parsedEdge.to,
      label: parsedEdge.label,
      startLabel: parsedEdge.startLabel,
      endLabel: parsedEdge.endLabel,
      points: [],
      style: {
        ...cloneEdgeStyle(),
        lineStyle: parsedEdge.style,
        arrow: parsedEdge.arrow,
        startMarker: parsedEdge.startMarker ?? "none",
        endMarker: parsedEdge.endMarker ?? (parsedEdge.arrow === "end" || parsedEdge.arrow === "both" ? "arrow" : "none"),
        startSide: parsedEdge.startSide,
        endSide: parsedEdge.endSide,
        startViaGroup: parsedEdge.startViaGroup,
        endViaGroup: parsedEdge.endViaGroup,
      },
    };
  });

  ir.nodes = nodeOrder.map((id) => nodeById.get(id)).filter((n): n is IrNode => Boolean(n));

  const subgraphById = new Map<string, { id: string; title: string; parentId?: string }>();
  for (const subgraph of ast.subgraphs) {
    const existing = subgraphById.get(subgraph.id);
    if (!existing) {
      subgraphById.set(subgraph.id, { id: subgraph.id, title: subgraph.title, parentId: subgraph.parentId });
      continue;
    }

    if (!existing.parentId && subgraph.parentId) {
      existing.parentId = subgraph.parentId;
    }
    if (subgraph.title && subgraph.title.length >= existing.title.length) {
      existing.title = subgraph.title;
    }
  }

  for (const node of ir.nodes) {
    if (node.subgraphId && !subgraphById.has(node.subgraphId)) {
      subgraphById.set(node.subgraphId, { id: node.subgraphId, title: node.subgraphId });
    }
  }

  const ensureParentSubgraphs = (): void => {
    let changed = false;
    for (const subgraph of [...subgraphById.values()]) {
      if (!subgraph.parentId || subgraphById.has(subgraph.parentId)) {
        continue;
      }
      subgraphById.set(subgraph.parentId, { id: subgraph.parentId, title: subgraph.parentId });
      changed = true;
    }

    if (changed) {
      ensureParentSubgraphs();
    }
  };
  ensureParentSubgraphs();

  ir.subgraphs = [...subgraphById.values()].map((subgraph) => {
    const nodeIds = ir.nodes.filter((node) => node.subgraphId === subgraph.id).map((node) => node.id);
    return {
      id: subgraph.id,
      title: subgraph.title,
      parentId: subgraph.parentId,
      nodeIds,
      x: 0,
      y: 0,
      width: 0,
      height: 0,
      style: { ...DEFAULT_SUBGRAPH_STYLE },
    };
  });

  const subgraphByIdMap = new Map(ir.subgraphs.map((subgraph) => [subgraph.id, subgraph]));
  const titleNodeIds = new Set<string>();
  const defaultClassStyle = ast.classDefs.default ?? ast.classDefs.DEFAULT;

  for (const node of ir.nodes) {
    const classes = ast.classAssignments[node.id] ?? [];
    const isGroupTitle = classes.some((name) => name.toLowerCase() === "group_title");
    if (!isGroupTitle) {
      continue;
    }

    if (node.subgraphId) {
      const subgraph = subgraphByIdMap.get(node.subgraphId);
      const label = node.label.trim();
      if (subgraph && label && label.length >= subgraph.title.trim().length) {
        subgraph.title = label;
      }
    }

    titleNodeIds.add(node.id);
  }

  if (titleNodeIds.size > 0) {
    ir.nodes = ir.nodes.filter((node) => !titleNodeIds.has(node.id));
    ir.edges = ir.edges.filter((edge) => !titleNodeIds.has(edge.from) && !titleNodeIds.has(edge.to));

    for (const subgraph of ir.subgraphs) {
      subgraph.nodeIds = subgraph.nodeIds.filter((id) => !titleNodeIds.has(id));
    }
  }

  for (const node of ir.nodes) {
    if (defaultClassStyle) {
      applyStyleToNode(node, defaultClassStyle);
    }

    const classes = ast.classAssignments[node.id] ?? [];
    for (const className of classes) {
      applyStyleToNode(node, ast.classDefs[className]);
    }
    applyStyleToNode(node, ast.styleOverrides[node.id]);
  }

  for (const subgraph of ir.subgraphs) {
    applyStyleToSubgraph(subgraph, ast.styleOverrides[subgraph.id]);
  }

  return ir;
}
