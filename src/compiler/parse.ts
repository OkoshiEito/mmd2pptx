import type {
  ArrowType,
  DiagramAst,
  DiagramDirection,
  EdgeMarker,
  EdgeLineStyle,
  NodeShape,
  ParsedEdge,
  ParsedNode,
  ParsedSubgraph,
} from "../types.js";

const HEADER_RE = /^(flowchart|graph|stateDiagram(?:-v2)?|classDiagram(?:-v2)?)\b(?:\s+([A-Za-z]{2}))?/i;
const CLASS_HEADER_RE = /^classDiagram(?:-v2)?\b/i;
const EDGE_TOKEN_RE = /(?:([A-Za-z0-9_:-]+)@)?([ox]?<?[-=.~]{2,}>?[ox]?)/giu;

interface NodeExpr {
  id: string;
  label?: string;
  shape?: NodeShape;
  classes?: string[];
}

function toDirection(input?: string): DiagramDirection {
  const upper = (input ?? "TD").toUpperCase();
  if (upper === "TB" || upper === "TD") {
    return "TD";
  }
  if (upper === "BT" || upper === "LR" || upper === "RL") {
    return upper;
  }
  return "TD";
}

function parseInitDirective(line: string): { nodeSpacing?: number; rankSpacing?: number } | null {
  const match = line.match(/%%\{init:\s*([\s\S]*?)\}%%/i);
  if (!match) {
    return null;
  }

  const payload = match[1];
  const nodeSpacingMatch = payload.match(/['"]?nodeSpacing['"]?\s*:\s*([0-9]+(?:\.[0-9]+)?)/i);
  const rankSpacingMatch = payload.match(/['"]?rankSpacing['"]?\s*:\s*([0-9]+(?:\.[0-9]+)?)/i);

  const nodeSpacing = nodeSpacingMatch ? Number(nodeSpacingMatch[1]) : undefined;
  const rankSpacing = rankSpacingMatch ? Number(rankSpacingMatch[1]) : undefined;

  if (nodeSpacing === undefined && rankSpacing === undefined) {
    return null;
  }

  return {
    nodeSpacing,
    rankSpacing,
  };
}

function stripComment(line: string): string {
  const idx = line.indexOf("%%");
  if (idx === -1) {
    return line;
  }
  return line.slice(0, idx);
}

function cleanLabel(raw: string): string {
  let next = raw.trim();
  const quoted = next.match(/^(["'`])([\s\S]*)\1$/u);
  next = quoted ? quoted[2].trim() : next;

  // Mermaid labels commonly use HTML break tags; map them to plain newlines for IR/rendering.
  next = next.replace(/<br\s*\/?>/gi, "\n").replace(/&nbsp;/gi, " ");

  // Some nested forms like [("text")] or ["text"] may leak wrappers into labels.
  for (let i = 0; i < 3; i += 1) {
    const candidates: RegExp[] = [/^\(([\s\S]*)\)$/u, /^\[([\s\S]*)\]$/u, /^\/([\s\S]*)\/$/u];
    let unwrapped = false;

    for (const re of candidates) {
      const match = next.match(re);
      if (!match) {
        continue;
      }

      const inner = match[1].trim();
      if (!inner) {
        continue;
      }

      const innerQuoted = inner.match(/^(["'`])([\s\S]*)\1$/u);
      next = (innerQuoted ? innerQuoted[2] : inner).trim();
      unwrapped = true;
      break;
    }

    if (!unwrapped) {
      break;
    }
  }

  return next;
}

function normalizeShapeLabel(shape: NodeShape, label: string): string {
  if (shape === "roundRect") {
    const nestedRect = label.match(/^\[(.*)\]$/u);
    if (nestedRect) {
      return cleanLabel(nestedRect[1]);
    }
  }

  if (shape === "circle") {
    const nestedRound = label.match(/^\((.*)\)$/u);
    if (nestedRound) {
      return cleanLabel(nestedRound[1]);
    }
  }

  return label;
}

function extractInlineClasses(raw: string): { core: string; classes: string[] } {
  let core = raw.trim();
  const classes: string[] = [];
  const suffixRe = /:::([A-Za-z0-9_:-]+)\s*$/u;

  while (true) {
    const matched = core.match(suffixRe);
    if (!matched || matched.index === undefined) {
      break;
    }
    classes.unshift(matched[1]);
    core = core.slice(0, matched.index).trimEnd();
  }

  return {
    core,
    classes,
  };
}

function splitStatements(line: string): string[] {
  const out: string[] = [];
  let depthParen = 0;
  let depthBracket = 0;
  let depthBrace = 0;
  let current = "";

  for (const ch of line) {
    if (ch === "(") {
      depthParen += 1;
    } else if (ch === ")") {
      depthParen = Math.max(0, depthParen - 1);
    } else if (ch === "[") {
      depthBracket += 1;
    } else if (ch === "]") {
      depthBracket = Math.max(0, depthBracket - 1);
    } else if (ch === "{") {
      depthBrace += 1;
    } else if (ch === "}") {
      depthBrace = Math.max(0, depthBrace - 1);
    }

    if (ch === ";" && depthParen === 0 && depthBracket === 0 && depthBrace === 0) {
      if (current.trim()) {
        out.push(current.trim());
      }
      current = "";
      continue;
    }

    current += ch;
  }

  if (current.trim()) {
    out.push(current.trim());
  }

  return out;
}

function buildLogicalLines(lines: string[]): Array<{ text: string; lineNumber: number }> {
  const out: Array<{ text: string; lineNumber: number }> = [];

  let buffer = "";
  let startLine = 1;
  let depthParen = 0;
  let depthBracket = 0;
  let depthBrace = 0;
  let inSingle = false;
  let inDouble = false;
  let inBacktick = false;

  for (let i = 0; i < lines.length; i += 1) {
    const raw = lines[i] ?? "";
    if (!buffer) {
      buffer = raw;
      startLine = i + 1;
    } else {
      buffer += `\n${raw}`;
    }

    for (let p = 0; p < raw.length; p += 1) {
      const ch = raw[p];
      const next = p + 1 < raw.length ? raw[p + 1] : "";
      if (ch === "\\" && (inSingle || inDouble || inBacktick) && next) {
        p += 1;
        continue;
      }

      if (ch === "'" && !inDouble && !inBacktick) {
        inSingle = !inSingle;
        continue;
      }
      if (ch === '"' && !inSingle && !inBacktick) {
        inDouble = !inDouble;
        continue;
      }
      if (ch === "`" && !inSingle && !inDouble) {
        inBacktick = !inBacktick;
        continue;
      }

      if (inSingle || inDouble || inBacktick) {
        continue;
      }

      if (ch === "(") {
        depthParen += 1;
      } else if (ch === ")") {
        depthParen = Math.max(0, depthParen - 1);
      } else if (ch === "[") {
        depthBracket += 1;
      } else if (ch === "]") {
        depthBracket = Math.max(0, depthBracket - 1);
      } else if (ch === "{") {
        depthBrace += 1;
      } else if (ch === "}") {
        depthBrace = Math.max(0, depthBrace - 1);
      }
    }

    const closed = !inSingle && !inDouble && !inBacktick && depthParen === 0 && depthBracket === 0 && depthBrace === 0;
    if (closed) {
      out.push({ text: buffer, lineNumber: startLine });
      buffer = "";
    }
  }

  if (buffer) {
    out.push({ text: buffer, lineNumber: startLine });
  }

  return out;
}

function parseNodeExpr(raw: string): NodeExpr | null {
  const extracted = extractInlineClasses(raw);
  const token = extracted.core;
  if (!token) {
    return null;
  }

  const expanded = token.match(/^([A-Za-z0-9_:.\/-]+)\s*@\{\s*([\s\S]*?)\s*\}$/u);
  if (expanded) {
    const id = expanded[1].trim();
    const body = expanded[2];

    const fieldValue = (key: string): string | undefined => {
      const pattern = new RegExp(
        String.raw`(?:^|,)\s*${key}\s*:\s*(?:"([^"]*)"|'([^']*)'|` + "`([^`]*)`" + String.raw`|([^,}]+))`,
        "iu",
      );
      const matched = body.match(pattern);
      if (!matched) {
        return undefined;
      }
      return (matched[1] ?? matched[2] ?? matched[3] ?? matched[4] ?? "").trim();
    };

    const shapeAlias = (fieldValue("shape") ?? "").toLowerCase();
    const shapeMap: Record<string, NodeShape> = {
      rect: "rect",
      rounded: "roundRect",
      stadium: "roundRect",
      bow: "roundRect",
      "bow-rect": "roundRect",
      circle: "circle",
      "sm-circ": "smallCircle",
      "f-circ": "filledCircle",
      "fr-circ": "framedCircle",
      "dbl-circ": "donut",
      "cross-circ": "summingJunction",
      diam: "diamond",
      diamond: "diamond",
      hex: "hexagon",
      hexagon: "hexagon",
      cyl: "cylinder",
      can: "cylinder",
      "h-cyl": "hCylinder",
      "lin-cyl": "cylinder",
      "lean-r": "parallelogram",
      "lean-l": "parallelogramAlt",
      "sl-rect": "parallelogram",
      "trap-t": "trapezoid",
      "trap-b": "trapezoidAlt",
      "curv-trap": "curvedTrapezoid",
      tri: "triangle",
      triangle: "triangle",
      "flip-tri": "rightTriangle",
      fork: "forkBar",
      notch: "card",
      "notch-rect": "card",
      "notch-pent": "pentagon",
      cloud: "cloud",
      hourglass: "hourglass",
      bolt: "lightningBolt",
      bang: "explosion",
      brace: "braceLeft",
      "brace-l": "braceLeft",
      "brace-r": "braceRight",
      braces: "bracePair",
      delay: "delay",
      doc: "document",
      "lin-doc": "linedDocument",
      "tag-doc": "document",
      docs: "multiDocument",
      "div-rect": "internalStorage",
      "st-rect": "stackedRect",
      "lin-rect": "linedRect",
      "fr-rect": "framedRect",
      "tag-rect": "foldedCorner",
      "win-pane": "windowPane",
      odd: "decagon",
      flag: "wave",
      text: "rect",
    };

    const shape = shapeMap[shapeAlias] ?? "rect";
    const suppressLabelAliases = new Set(["fork", "f-circ", "sm-circ", "fr-circ"]);
    const label = suppressLabelAliases.has(shapeAlias) ? "" : cleanLabel(fieldValue("label") ?? id);
    return {
      id,
      label,
      shape,
      classes: extracted.classes,
    };
  }

    const patterns: Array<[RegExp, NodeShape]> = [
    [/^([A-Za-z0-9_:.\/-]+)\[\(([\s\S]*)\)\]$/u, "cylinder"],
    [/^([A-Za-z0-9_:.\/-]+)\[\[([\s\S]*)\]\]$/u, "subroutine"],
    [/^([A-Za-z0-9_:.\/-]+)\{\{([\s\S]*)\}\}$/u, "hexagon"],
    [/^([A-Za-z0-9_:.\/-]+)\(\(([\s\S]*)\)\)$/u, "circle"],
    [/^([A-Za-z0-9_:.\/-]+)\(([\s\S]*)\)$/u, "roundRect"],
    [/^([A-Za-z0-9_:.\/-]+)\{([\s\S]*)\}$/u, "diamond"],
    [/^([A-Za-z0-9_:.\/-]+)\[\\([\s\S]*)\\\]$/u, "parallelogramAlt"],
    [/^([A-Za-z0-9_:.\/-]+)\[\/([\s\S]*)\\\]$/u, "trapezoid"],
    [/^([A-Za-z0-9_:.\/-]+)\[\\([\s\S]*)\/\]$/u, "trapezoidAlt"],
    [/^([A-Za-z0-9_:.\/-]+)\[\/([\s\S]*)\/\]$/u, "parallelogram"],
    [/^([A-Za-z0-9_:.\/-]+)\[([\s\S]*)\]$/u, "rect"],
  ];

  for (const [re, shape] of patterns) {
    const match = token.match(re);
    if (!match) {
      continue;
    }

    const id = match[1].trim();
    const label = normalizeShapeLabel(shape, cleanLabel(match[2]));
    return {
      id,
      label,
      shape,
      classes: extracted.classes,
    };
  }

  const plain = token.match(/^([A-Za-z0-9_:.\/-]+)$/u);
  if (!plain) {
    return null;
  }

  return {
    id: plain[1],
    label: plain[1],
    classes: extracted.classes,
  };
}

function parseSubgraphHeader(raw: string): { id: string; title: string } {
  const rest = raw.replace(/^subgraph\s+/i, "").trim();

  const withTitle = rest.match(/^([A-Za-z0-9_:.\/-]+)\[(.*)\]$/u);
  if (withTitle) {
    return {
      id: withTitle[1].trim(),
      title: cleanLabel(withTitle[2]),
    };
  }

  const withSpace = rest.match(/^([A-Za-z0-9_:.\/-]+)\s+(.+)$/u);
  if (withSpace) {
    return {
      id: withSpace[1].trim(),
      title: cleanLabel(withSpace[2]),
    };
  }

  return {
    id: rest,
    title: cleanLabel(rest),
  };
}

function parseStyleMap(raw: string): Record<string, string> {
  const out: Record<string, string> = {};
  const parts = raw
    .split(",")
    .map((part) => part.trim())
    .filter(Boolean);

  for (const part of parts) {
    const idx = part.indexOf(":");
    if (idx === -1) {
      continue;
    }

    const key = part.slice(0, idx).trim().toLowerCase();
    const value = part.slice(idx + 1).trim();
    if (!key || !value) {
      continue;
    }

    out[key] = value;
  }

  return out;
}

function mergeStyleMap(target: Record<string, string>, source: Record<string, string>): Record<string, string> {
  return {
    ...target,
    ...source,
  };
}

function parseClassDefStatement(raw: string): { names: string[]; styles: Record<string, string> } | null {
  const match = raw.match(/^classDef\s+([A-Za-z0-9_.,:-]+)\s+(.+)$/iu);
  if (!match) {
    return null;
  }

  const names = match[1]
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);
  if (names.length === 0) {
    return null;
  }

  const styles = parseStyleMap(match[2]);
  return {
    names,
    styles,
  };
}

function parseClassStatement(raw: string): { ids: string[]; classes: string[] } | null {
  const match = raw.match(/^class\s+([A-Za-z0-9_.,:\/-]+)\s+([A-Za-z0-9_.,:-]+)\s*$/iu);
  if (!match) {
    return null;
  }

  const ids = match[1]
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);

  const classes = match[2]
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);

  if (ids.length === 0 || classes.length === 0) {
    return null;
  }

  return { ids, classes };
}

function parseStyleStatement(raw: string): { id: string; styles: Record<string, string> } | null {
  const match = raw.match(/^style\s+([A-Za-z0-9_:.\/-]+)\s+(.+)$/iu);
  if (!match) {
    return null;
  }

  return {
    id: match[1].trim(),
    styles: parseStyleMap(match[2]),
  };
}

function mapEdgeOperator(op: string): { style: EdgeLineStyle; arrow: ArrowType } {
  const style: EdgeLineStyle = op.includes("~")
    ? "invisible"
    : op.includes(".")
      ? "dotted"
      : op.includes("=")
        ? "thick"
        : "solid";
  const hasStart = op.startsWith("<");
  const hasEnd = op.endsWith(">");
  const arrow: ArrowType = hasStart && hasEnd ? "both" : hasStart ? "start" : hasEnd ? "end" : "none";

  return {
    style,
    arrow,
  };
}

function parseInlineLabeledArrow(
  line: string,
  lineNumber: number,
): { nodes: ParsedNode[]; edges: ParsedEdge[] } | null {
  const match = line.match(/^(.+?)\s+([-=.~]{2,3})\s+(.+?)\s+([<>\-=.~]{2,4})\s+(.+)$/u);
  if (!match) {
    return null;
  }

  const leftExpr = match[1].trim();
  const leftOp = match[2].trim();
  const middleRaw = match[3].trim();
  const label = cleanLabel(middleRaw);
  const rightOp = match[4].trim();
  const rightExpr = match[5].trim();

  if (!/^[<>\-=.~]+$/u.test(leftOp) || !/^[<>\-=.~]+$/u.test(rightOp)) {
    return null;
  }

  // Prefer node-chaining interpretation when the middle token is a node-like expression.
  // Example: `A --- B --- C` should be two edges, not inline label "B".
  const isQuotedLabel = /^(["'`])[\s\S]*\1$/u.test(middleRaw);
  if (!isQuotedLabel && parseNodeExpr(label) !== null) {
    return null;
  }

  const leftNodeExpr = parseNodeExpr(leftExpr);
  const rightNodeExpr = parseNodeExpr(rightExpr);
  if (!leftNodeExpr || !rightNodeExpr) {
    return null;
  }

  const style: EdgeLineStyle = `${leftOp}${rightOp}`.includes("~")
    ? "invisible"
    : `${leftOp}${rightOp}`.includes(".")
      ? "dotted"
      : `${leftOp}${rightOp}`.includes("=")
        ? "thick"
        : "solid";
  const hasStart = rightOp.includes("<");
  const hasEnd = rightOp.includes(">");
  const arrow: ArrowType = hasStart && hasEnd ? "both" : hasStart ? "start" : hasEnd ? "end" : "none";

  const nodes: ParsedNode[] = [
    {
      id: leftNodeExpr.id,
      label: leftNodeExpr.label,
      shape: leftNodeExpr.shape,
      line: lineNumber,
      raw: leftExpr,
    },
    {
      id: rightNodeExpr.id,
      label: rightNodeExpr.label,
      shape: rightNodeExpr.shape,
      line: lineNumber,
      raw: rightExpr,
    },
  ];

  const edges: ParsedEdge[] = [
    {
      from: leftNodeExpr.id,
      to: rightNodeExpr.id,
      label,
      style,
      arrow,
      line: lineNumber,
      raw: line,
    },
  ];

  return { nodes, edges };
}

function splitNodeGroup(raw: string): string[] {
  const out: string[] = [];
  let current = "";
  let depthParen = 0;
  let depthBracket = 0;
  let depthBrace = 0;
  let inSingle = false;
  let inDouble = false;
  let inBacktick = false;

  for (let i = 0; i < raw.length; i += 1) {
    const ch = raw[i];
    const next = i + 1 < raw.length ? raw[i + 1] : "";
    if (ch === "\\" && (inSingle || inDouble || inBacktick) && next) {
      current += ch + next;
      i += 1;
      continue;
    }

    if (ch === "'" && !inDouble && !inBacktick) {
      inSingle = !inSingle;
      current += ch;
      continue;
    }
    if (ch === '"' && !inSingle && !inBacktick) {
      inDouble = !inDouble;
      current += ch;
      continue;
    }
    if (ch === "`" && !inSingle && !inDouble) {
      inBacktick = !inBacktick;
      current += ch;
      continue;
    }

    if (!inSingle && !inDouble && !inBacktick) {
      if (ch === "(") {
        depthParen += 1;
      } else if (ch === ")") {
        depthParen = Math.max(0, depthParen - 1);
      } else if (ch === "[") {
        depthBracket += 1;
      } else if (ch === "]") {
        depthBracket = Math.max(0, depthBracket - 1);
      } else if (ch === "{") {
        depthBrace += 1;
      } else if (ch === "}") {
        depthBrace = Math.max(0, depthBrace - 1);
      } else if (ch === "&" && depthParen === 0 && depthBracket === 0 && depthBrace === 0) {
        if (current.trim()) {
          out.push(current.trim());
        }
        current = "";
        continue;
      }
    }

    current += ch;
  }

  if (current.trim()) {
    out.push(current.trim());
  }
  return out;
}

function tokenizeEdgeLine(line: string): { nodes: string[]; ops: string[] } | null {
  const nodes: string[] = [];
  const ops: string[] = [];

  EDGE_TOKEN_RE.lastIndex = 0;
  let last = 0;
  let matched = false;

  while (true) {
    const m = EDGE_TOKEN_RE.exec(line);
    if (!m) {
      break;
    }

    const left = line.slice(last, m.index).trim();
    if (!matched) {
      if (!left) {
        return null;
      }
      nodes.push(left);
    } else {
      nodes.push(left);
    }

    ops.push(m[2]);
    last = m.index + m[0].length;
    matched = true;
  }

  if (!matched) {
    return null;
  }

  const tail = line.slice(last).trim();
  if (!tail) {
    return null;
  }
  nodes.push(tail);

  if (nodes.length !== ops.length + 1) {
    return null;
  }
  return { nodes, ops };
}

function parseNodeGroup(raw: string, lineNumber: number): ParsedNode[] | null {
  const parts = splitNodeGroup(raw);
  if (parts.length === 0) {
    return null;
  }

  const out: ParsedNode[] = [];
  for (const part of parts) {
    const parsed = parseNodeExpr(part);
    if (!parsed) {
      return null;
    }
    out.push({
      id: parsed.id,
      label: parsed.label,
      shape: parsed.shape,
      inlineClasses: parsed.classes,
      line: lineNumber,
      raw: part,
    });
  }
  return out;
}

function mergeOperators(left: string, right: string): string {
  const merged = `${left}${right}`;
  const normalized = merged.replace(/[ox]/giu, "");
  return normalized;
}

function operatorCore(op: string): string {
  return op.replace(/[ox<>]/giu, "");
}

function hasOperatorArrow(op: string): boolean {
  return /[<>]/u.test(op);
}

function shouldInterpretAsInlineLabel(currentOp: string, nextOp: string, middleRaw: string): boolean {
  if (middleRaw.includes("&")) {
    return false;
  }

  const curCore = operatorCore(currentOp);
  const nextCore = operatorCore(nextOp);
  if (!curCore || !nextCore) {
    return false;
  }

  // Inline label syntax is typically `-- text -->`, `== text ==>`, `-. text .->`.
  // Interpret when first leg has no arrow and second leg has an arrow.
  if (hasOperatorArrow(currentOp) || !hasOperatorArrow(nextOp)) {
    return false;
  }

  const curKind = curCore[0];
  const nextKind = nextCore[0];
  if (!["-", "=", ".", "~"].includes(curKind) || !["-", "=", ".", "~"].includes(nextKind)) {
    return false;
  }

  return true;
}

function isEdgeMetadataDirective(raw: string): boolean {
  const matched = raw.match(/^([A-Za-z0-9_:.\/-]+)\s*@\{\s*([\s\S]*?)\s*\}$/u);
  if (!matched) {
    return false;
  }
  const body = matched[2].toLowerCase();
  if (/(^|[,\s])(shape|label|icon|img|form|pos|w|h|constraint)\s*:/u.test(body)) {
    return false;
  }
  return /(^|[,\s])(animate|animation|curve|class|style)\s*:/u.test(body);
}

function parseEdgeStatement(line: string, lineNumber: number): { nodes: ParsedNode[]; edges: ParsedEdge[] } | null {
  const inlineLabeled = parseInlineLabeledArrow(line, lineNumber);
  if (inlineLabeled) {
    return inlineLabeled;
  }

  const tokenized = tokenizeEdgeLine(line);
  if (!tokenized) {
    return null;
  }

  const discoveredNodes: ParsedNode[] = [];
  const discoveredEdges: ParsedEdge[] = [];

  let leftGroup = parseNodeGroup(tokenized.nodes[0], lineNumber);
  if (!leftGroup) {
    return null;
  }
  discoveredNodes.push(...leftGroup);

  for (let i = 0; i < tokenized.ops.length; i += 1) {
    const op = tokenized.ops[i];
    const middleRaw = tokenized.nodes[i + 1].trim();

    let label: string | undefined;
    let rightGroup: ParsedNode[] | null = null;
    let effectiveOp = op;

    const piped = middleRaw.match(/^\|([\s\S]+)\|\s*(.+)$/u);
    if (piped) {
      label = cleanLabel(piped[1]);
      rightGroup = parseNodeGroup(piped[2].trim(), lineNumber);
    } else {
      if (i + 1 < tokenized.ops.length && shouldInterpretAsInlineLabel(op, tokenized.ops[i + 1], middleRaw)) {
        const nextNodeRaw = tokenized.nodes[i + 2].trim();
        const nextGroup = parseNodeGroup(nextNodeRaw, lineNumber);
        if (nextGroup) {
          label = cleanLabel(middleRaw);
          effectiveOp = mergeOperators(op, tokenized.ops[i + 1]);
          rightGroup = nextGroup;
          i += 1;
        }
      }

      if (!rightGroup) {
        rightGroup = parseNodeGroup(middleRaw, lineNumber);
      }
      if (!rightGroup && i + 1 < tokenized.ops.length) {
        const nextNodeRaw = tokenized.nodes[i + 2].trim();
        const nextGroup = parseNodeGroup(nextNodeRaw, lineNumber);
        if (nextGroup) {
          label = cleanLabel(middleRaw);
          effectiveOp = mergeOperators(op, tokenized.ops[i + 1]);
          rightGroup = nextGroup;
          i += 1;
        }
      }

      if (!label && rightGroup && rightGroup.length === 1) {
        const expr = rightGroup[0].raw;
        const stateLabel = expr.match(/^([A-Za-z0-9_./-]+)\s*:\s*(.+)$/u);
        if (stateLabel) {
          const maybeNode = parseNodeGroup(stateLabel[1].trim(), lineNumber);
          if (maybeNode) {
            rightGroup = maybeNode;
            label = cleanLabel(stateLabel[2]);
          }
        }
      }
    }

    if (!rightGroup) {
      return null;
    }

    discoveredNodes.push(...rightGroup);
    const mapped = mapEdgeOperator(effectiveOp);
    for (const leftNode of leftGroup) {
      for (const rightNode of rightGroup) {
        discoveredEdges.push({
          from: leftNode.id,
          to: rightNode.id,
          label,
          style: mapped.style,
          arrow: mapped.arrow,
          line: lineNumber,
          raw: line,
        });
      }
    }
    leftGroup = rightGroup;
  }

  return {
    nodes: discoveredNodes,
    edges: discoveredEdges,
  };
}

interface ClassNodeBuffer {
  id: string;
  label: string;
  members: string[];
  annotation?: string;
  subgraphId?: string;
}

function stripQuotesOrBackticks(input: string): string {
  const trimmed = input.trim();
  const matched = trimmed.match(/^["'`]([\s\S]*)["'`]$/u);
  return matched ? matched[1] : trimmed;
}

function normalizeClassId(input: string): string {
  const stripped = stripQuotesOrBackticks(input);
  const withoutGenerics = stripped.replace(/~[^~]+~/gu, "");
  return withoutGenerics.trim();
}

function classDisplayName(input: string, fallbackId: string): string {
  const stripped = stripQuotesOrBackticks(input);
  if (!stripped) {
    return fallbackId;
  }

  const withGenerics = stripped.replace(/~([^~]+)~/gu, "<$1>");
  return withGenerics.trim() || fallbackId;
}

function tokenizeQuoted(input: string): string[] {
  const tokens = input.match(/"([^"\\]|\\.)*"|'([^'\\]|\\.)*'|`[^`]*`|\S+/gu);
  return tokens ? [...tokens] : [];
}

function unquoteToken(input: string): string {
  const token = input.trim();
  if (
    (token.startsWith('"') && token.endsWith('"')) ||
    (token.startsWith("'") && token.endsWith("'")) ||
    (token.startsWith("`") && token.endsWith("`"))
  ) {
    return token.slice(1, -1);
  }
  return token;
}

function splitTrailingColonLabel(input: string): { core: string; label?: string } {
  let inSingle = false;
  let inDouble = false;
  let inBacktick = false;

  for (let i = 0; i < input.length; i += 1) {
    const ch = input[i];
    const next = i + 1 < input.length ? input[i + 1] : "";
    if (ch === "\\" && (inSingle || inDouble || inBacktick) && next) {
      i += 1;
      continue;
    }

    if (ch === "'" && !inDouble && !inBacktick) {
      inSingle = !inSingle;
      continue;
    }
    if (ch === '"' && !inSingle && !inBacktick) {
      inDouble = !inDouble;
      continue;
    }
    if (ch === "`" && !inSingle && !inDouble) {
      inBacktick = !inBacktick;
      continue;
    }

    if (ch === ":" && !inSingle && !inDouble && !inBacktick) {
      return {
        core: input.slice(0, i).trim(),
        label: cleanLabel(input.slice(i + 1)),
      };
    }
  }

  return {
    core: input.trim(),
  };
}

function markerFromClassRelationToken(token?: string): EdgeMarker {
  if (!token) {
    return "none";
  }
  if (token === "*" ) {
    return "diamond";
  }
  if (token === "o") {
    return "openDiamond";
  }
  if (token === "()") {
    return "circle";
  }
  if (token === "<|" || token === "|>") {
    return "triangle";
  }
  if (token === "<" || token === ">") {
    return "arrow";
  }
  return "none";
}

function arrowFromMarkers(startMarker: EdgeMarker, endMarker: EdgeMarker): ArrowType {
  const hasStart = startMarker !== "none";
  const hasEnd = endMarker !== "none";
  if (hasStart && hasEnd) {
    return "both";
  }
  if (hasStart) {
    return "start";
  }
  if (hasEnd) {
    return "end";
  }
  return "none";
}

function parseClassRelationOperator(
  op: string,
): { lineStyle: EdgeLineStyle; startMarker: EdgeMarker; endMarker: EdgeMarker; arrow: ArrowType } | null {
  const matched = op.match(/^(<\||\|>|\*|o|<|>|\(\))?(--|\.\.)(<\||\|>|\*|o|<|>|\(\))?$/u);
  if (!matched) {
    return null;
  }

  const startMarker = markerFromClassRelationToken(matched[1]);
  const endMarker = markerFromClassRelationToken(matched[3]);
  return {
    lineStyle: matched[2] === ".." ? "dotted" : "solid",
    startMarker,
    endMarker,
    arrow: arrowFromMarkers(startMarker, endMarker),
  };
}

function parseCssClassStatement(raw: string): { ids: string[]; classes: string[] } | null {
  const matched = raw.match(/^cssClass\s+(.+?)\s+([A-Za-z0-9_.,:-]+)\s*;?$/iu);
  if (!matched) {
    return null;
  }

  const idExpr = matched[1].trim();
  const classExpr = matched[2].trim();
  const ids = idExpr
    .replace(/^["']|["']$/gu, "")
    .split(",")
    .map((value) => normalizeClassId(value))
    .filter(Boolean);
  const classes = classExpr
    .split(",")
    .map((value) => value.trim())
    .filter(Boolean);

  if (ids.length === 0 || classes.length === 0) {
    return null;
  }
  return { ids, classes };
}

function parseClassDeclarationStatement(raw: string): {
  id: string;
  label: string;
  annotation?: string;
  classes: string[];
  openBlock: boolean;
} | null {
  let body = raw.replace(/^class\s+/iu, "").trim();
  if (!body) {
    return null;
  }

  if (
    !body.includes("[") &&
    !body.includes("{") &&
    !body.includes(":::") &&
    !body.includes("<<") &&
    /^[A-Za-z0-9_.,:\/-]+\s+[A-Za-z0-9_.,:-]+$/u.test(body)
  ) {
    return null;
  }

  let openBlock = false;
  if (body.endsWith("{")) {
    openBlock = true;
    body = body.slice(0, -1).trim();
  }

  const withClasses = extractInlineClasses(body);
  body = withClasses.core;

  let annotation: string | undefined;
  const annotationMatch = body.match(/<<\s*([^>]+)\s*>>/u);
  if (annotationMatch && annotationMatch.index !== undefined) {
    annotation = cleanLabel(annotationMatch[1]);
    body = `${body.slice(0, annotationMatch.index)}${body.slice(annotationMatch.index + annotationMatch[0].length)}`.trim();
  }

  let idToken = body;
  let label: string | undefined;
  let rest = "";
  let handledBacktickId = false;
  if (body.startsWith("`")) {
    const closing = body.indexOf("`", 1);
    if (closing !== -1) {
      idToken = body.slice(0, closing + 1).trim();
      rest = body.slice(closing + 1).trim();
      handledBacktickId = true;
    }
  }
  if (!handledBacktickId) {
    const matched = body.match(/^([^\s\[]+)([\s\S]*)$/u);
    if (matched) {
      idToken = matched[1].trim();
      rest = matched[2].trim();
    }
  }

  const labelMatch = rest.match(/^\[(.+)\]\s*$/u);
  if (labelMatch) {
    label = cleanLabel(labelMatch[1]);
  }

  const id = normalizeClassId(idToken);
  if (!id) {
    return null;
  }

  return {
    id,
    label: label || classDisplayName(idToken, id),
    annotation,
    classes: withClasses.classes,
    openBlock,
  };
}

function parseClassRelationStatement(
  raw: string,
  lineNumber: number,
): { nodes: ParsedNode[]; edge: ParsedEdge } | null {
  const split = splitTrailingColonLabel(raw);
  const tokens = tokenizeQuoted(split.core);
  if (tokens.length < 3) {
    return null;
  }

  let opIndex = -1;
  let opInfo: ReturnType<typeof parseClassRelationOperator> = null;
  for (let i = 0; i < tokens.length; i += 1) {
    const parsed = parseClassRelationOperator(tokens[i]);
    if (!parsed) {
      continue;
    }
    opIndex = i;
    opInfo = parsed;
    break;
  }

  if (opIndex === -1 || !opInfo) {
    return null;
  }

  const leftTokens = tokens.slice(0, opIndex);
  const rightTokens = tokens.slice(opIndex + 1);
  if (leftTokens.length === 0 || rightTokens.length === 0) {
    return null;
  }

  let startLabel: string | undefined;
  if (leftTokens.length >= 2 && /^".*"$|^'.*'$|^`.*`$/u.test(leftTokens[leftTokens.length - 1])) {
    startLabel = cleanLabel(unquoteToken(leftTokens.pop() ?? ""));
  }

  let endLabel: string | undefined;
  if (rightTokens.length >= 2 && /^".*"$|^'.*'$|^`.*`$/u.test(rightTokens[0])) {
    endLabel = cleanLabel(unquoteToken(rightTokens.shift() ?? ""));
  }

  const leftExpr = leftTokens.join(" ").trim();
  const rightExpr = rightTokens.join(" ").trim();
  const fromId = normalizeClassId(leftExpr);
  const toId = normalizeClassId(rightExpr);
  if (!fromId || !toId) {
    return null;
  }

  return {
    nodes: [
      {
        id: fromId,
        label: classDisplayName(leftExpr, fromId),
        shape: "rect",
        line: lineNumber,
        raw: leftExpr,
      },
      {
        id: toId,
        label: classDisplayName(rightExpr, toId),
        shape: "rect",
        line: lineNumber,
        raw: rightExpr,
      },
    ],
    edge: {
      from: fromId,
      to: toId,
      label: split.label,
      startLabel,
      endLabel,
      startMarker: opInfo.startMarker,
      endMarker: opInfo.endMarker,
      style: opInfo.lineStyle,
      arrow: opInfo.arrow,
      line: lineNumber,
      raw,
    },
  };
}

function classNodeLabel(node: ClassNodeBuffer): string {
  const header = node.annotation ? `«${node.annotation}»\n${node.label}` : node.label;
  if (node.members.length === 0) {
    return header;
  }

  const attrs: string[] = [];
  const methods: string[] = [];
  for (const member of node.members) {
    if (/\(.*\)/u.test(member)) {
      methods.push(member);
    } else {
      attrs.push(member);
    }
  }

  if (attrs.length > 0 && methods.length > 0) {
    return `${header}\n${attrs.join("\n")}\n\n${methods.join("\n")}`;
  }

  return `${header}\n${node.members.join("\n")}`;
}

function parseClassDiagram(
  lines: string[],
  source: string,
  title: string | undefined,
  layoutHints: { nodeSpacing?: number; rankSpacing?: number },
): DiagramAst {
  const nodes: ParsedNode[] = [];
  const edges: ParsedEdge[] = [];
  const subgraphs: ParsedSubgraph[] = [];
  const classDefs: Record<string, Record<string, string>> = {};
  const classAssignments: Record<string, string[]> = {};
  const styleOverrides: Record<string, Record<string, string>> = {};
  const classNodeMap = new Map<string, ClassNodeBuffer>();
  const classOrder: string[] = [];
  const namespaceStack: string[] = [];
  let direction: DiagramDirection = "TD";
  let currentClassBlockId: string | null = null;
  let noteCounter = 0;
  let namespaceCounter = 0;

  const registerClasses = (id: string, names?: string[]): void => {
    if (!names || names.length === 0) {
      return;
    }
    const current = classAssignments[id] ?? [];
    classAssignments[id] = [...new Set([...current, ...names])];
  };

  const ensureClassNode = (id: string, label?: string): ClassNodeBuffer => {
    const existing = classNodeMap.get(id);
    if (existing) {
      if (label && (existing.label === existing.id || label.length >= existing.label.length)) {
        existing.label = label;
      }
      if (!existing.subgraphId && namespaceStack.length > 0) {
        existing.subgraphId = namespaceStack[namespaceStack.length - 1];
      }
      return existing;
    }

    const created: ClassNodeBuffer = {
      id,
      label: label || id,
      members: [],
      subgraphId: namespaceStack[namespaceStack.length - 1],
    };
    classNodeMap.set(id, created);
    classOrder.push(id);
    return created;
  };

  for (let i = 0; i < lines.length; i += 1) {
    const lineNumber = i + 1;
    const rawLine = lines[i];

    const titleMatch = rawLine.match(/^\s*%%\s*title\s*:?\s*(.+?)\s*$/iu);
    if (titleMatch && !title) {
      title = cleanLabel(titleMatch[1]);
    }

    const initHint = parseInitDirective(rawLine);
    if (initHint?.nodeSpacing !== undefined) {
      layoutHints.nodeSpacing = initHint.nodeSpacing;
    }
    if (initHint?.rankSpacing !== undefined) {
      layoutHints.rankSpacing = initHint.rankSpacing;
    }

    const trimmed = stripComment(rawLine).trim();
    if (!trimmed) {
      continue;
    }

    if (CLASS_HEADER_RE.test(trimmed)) {
      continue;
    }

    if (currentClassBlockId) {
      if (trimmed === "}") {
        currentClassBlockId = null;
        continue;
      }

      const current = classNodeMap.get(currentClassBlockId);
      if (!current) {
        currentClassBlockId = null;
        continue;
      }

      const annotationMatch = trimmed.match(/^<<\s*([^>]+)\s*>>$/u);
      if (annotationMatch) {
        current.annotation = cleanLabel(annotationMatch[1]);
        continue;
      }

      current.members.push(cleanLabel(trimmed));
      continue;
    }

    const dirMatch = trimmed.match(/^direction\s+([A-Za-z]{2})$/iu);
    if (dirMatch) {
      direction = toDirection(dirMatch[1]);
      continue;
    }

    const namespaceStart = trimmed.match(/^namespace\s+(.+?)\s*\{$/iu);
    if (namespaceStart) {
      namespaceCounter += 1;
      const titleText = cleanLabel(stripQuotesOrBackticks(namespaceStart[1]));
      const idSlug = titleText
        .toLowerCase()
        .replace(/[^a-z0-9]+/gu, "_")
        .replace(/^_+|_+$/gu, "")
        .slice(0, 40) || `ns_${namespaceCounter}`;
      const namespaceId = `ns_${idSlug}_${namespaceCounter}`;
      subgraphs.push({
        id: namespaceId,
        title: titleText || namespaceId,
        line: lineNumber,
      });
      namespaceStack.push(namespaceId);
      continue;
    }

    if (trimmed === "}") {
      namespaceStack.pop();
      continue;
    }

    if (/^classDef\b/i.test(trimmed)) {
      const parsed = parseClassDefStatement(trimmed);
      if (parsed) {
        for (const name of parsed.names) {
          classDefs[name] = mergeStyleMap(classDefs[name] ?? {}, parsed.styles);
        }
      }
      continue;
    }

    if (/^class\b/i.test(trimmed)) {
      const decl = parseClassDeclarationStatement(trimmed);
      if (decl) {
        const classNode = ensureClassNode(decl.id, decl.label);
        if (decl.annotation) {
          classNode.annotation = decl.annotation;
        }
        registerClasses(decl.id, decl.classes);
        if (decl.openBlock) {
          currentClassBlockId = decl.id;
        }
        continue;
      }

      const parsed = parseClassStatement(trimmed);
      if (parsed) {
        for (const id of parsed.ids) {
          registerClasses(id, parsed.classes);
        }
      }
      continue;
    }

    if (/^cssClass\b/i.test(trimmed)) {
      const parsed = parseCssClassStatement(trimmed);
      if (parsed) {
        for (const id of parsed.ids) {
          registerClasses(id, parsed.classes);
        }
      }
      continue;
    }

    if (/^style\b/i.test(trimmed)) {
      const parsed = parseStyleStatement(trimmed);
      if (parsed) {
        styleOverrides[parsed.id] = mergeStyleMap(styleOverrides[parsed.id] ?? {}, parsed.styles);
      }
      continue;
    }

    if (/^(linkStyle|click|link|callback)\b/i.test(trimmed)) {
      continue;
    }

    const noteForMatch = trimmed.match(/^note\s+for\s+(.+?)\s+"([\s\S]+)"\s*$/iu);
    if (noteForMatch) {
      const targetId = normalizeClassId(noteForMatch[1]);
      if (!targetId) {
        continue;
      }
      const targetNode = ensureClassNode(targetId, classDisplayName(noteForMatch[1], targetId));
      noteCounter += 1;
      const noteId = `note_${noteCounter}`;
      nodes.push({
        id: noteId,
        label: cleanLabel(noteForMatch[2]),
        shape: "roundRect",
        subgraphId: targetNode.subgraphId,
        line: lineNumber,
        raw: trimmed,
      });
      registerClasses(noteId, ["__class_note"]);
      edges.push({
        from: noteId,
        to: targetId,
        style: "dotted",
        arrow: "none",
        startMarker: "none",
        endMarker: "none",
        line: lineNumber,
        raw: trimmed,
      });
      continue;
    }

    const noteMatch = trimmed.match(/^note\s+"([\s\S]+)"\s*$/iu);
    if (noteMatch) {
      noteCounter += 1;
      const noteId = `note_${noteCounter}`;
      nodes.push({
        id: noteId,
        label: cleanLabel(noteMatch[1]),
        shape: "roundRect",
        subgraphId: namespaceStack[namespaceStack.length - 1],
        line: lineNumber,
        raw: trimmed,
      });
      registerClasses(noteId, ["__class_note"]);
      continue;
    }

    const relation = parseClassRelationStatement(trimmed, lineNumber);
    if (relation) {
      const [fromNode, toNode] = relation.nodes;
      ensureClassNode(fromNode.id, fromNode.label);
      ensureClassNode(toNode.id, toNode.label);
      edges.push(relation.edge);
      continue;
    }

    const memberMatch = trimmed.match(/^([^\s:][^:]*?)\s*:\s*(.+)$/u);
    if (memberMatch) {
      const classId = normalizeClassId(memberMatch[1]);
      const member = cleanLabel(memberMatch[2]);
      if (classId) {
        const classNode = ensureClassNode(classId, classDisplayName(memberMatch[1], classId));
        const annotationMatch = member.match(/^<<\s*([^>]+)\s*>>$/u);
        if (annotationMatch) {
          classNode.annotation = cleanLabel(annotationMatch[1]);
        } else {
          classNode.members.push(member);
        }
        continue;
      }
    }

    const standaloneAnnotation = trimmed.match(/^(.+?)\s+<<\s*([^>]+)\s*>>$/u);
    if (standaloneAnnotation) {
      const classId = normalizeClassId(standaloneAnnotation[1]);
      if (classId) {
        const classNode = ensureClassNode(classId, classDisplayName(standaloneAnnotation[1], classId));
        classNode.annotation = cleanLabel(standaloneAnnotation[2]);
        continue;
      }
    }

    const prefixAnnotation = trimmed.match(/^<<\s*([^>]+)\s*>>\s+(.+)$/u);
    if (prefixAnnotation) {
      const classId = normalizeClassId(prefixAnnotation[2]);
      if (classId) {
        const classNode = ensureClassNode(classId, classDisplayName(prefixAnnotation[2], classId));
        classNode.annotation = cleanLabel(prefixAnnotation[1]);
        continue;
      }
    }

    const bareClass = parseNodeExpr(trimmed);
    if (bareClass) {
      const id = normalizeClassId(bareClass.id);
      if (id) {
        ensureClassNode(id, classDisplayName(bareClass.id, id));
      }
    }
  }

  classDefs.__class_note = mergeStyleMap(classDefs.__class_note ?? {}, {
    fill: "#FFFDE7",
    stroke: "#D4AF37",
    "stroke-width": "1px",
  });

  for (const id of classOrder) {
    const classNode = classNodeMap.get(id);
    if (!classNode) {
      continue;
    }
    nodes.push({
      id: classNode.id,
      label: classNodeLabel(classNode),
      shape: "rect",
      subgraphId: classNode.subgraphId,
      line: 1,
      raw: classNode.id,
    });
  }

  return {
    source,
    direction,
    title,
    nodes,
    edges,
    subgraphs,
    classDefs,
    classAssignments,
    styleOverrides,
    layoutHints,
  };
}

export function parseMermaid(source: string): DiagramAst {
  const normalizedSource = source.replace(/\r\n/g, "\n");
  const allLines = normalizedSource.split("\n");
  let lines = allLines;

  let firstNonEmpty = 0;
  while (firstNonEmpty < lines.length && lines[firstNonEmpty].trim() === "") {
    firstNonEmpty += 1;
  }
  if (firstNonEmpty < lines.length && lines[firstNonEmpty].trim() === "---") {
    let closing = -1;
    for (let i = firstNonEmpty + 1; i < lines.length; i += 1) {
      if (lines[i].trim() === "---") {
        closing = i;
        break;
      }
    }
    if (closing !== -1) {
      lines = [...lines.slice(0, firstNonEmpty), ...lines.slice(closing + 1)];
    }
  }

  const classHeaderLine = lines.find((line) => {
    const trimmed = stripComment(line).trim();
    return CLASS_HEADER_RE.test(trimmed);
  });
  if (classHeaderLine) {
    return parseClassDiagram(lines, source, undefined, {});
  }

  const logicalLines = buildLogicalLines(lines);
  let direction: DiagramDirection = "TD";
  let title: string | undefined;
  let hasHeader = false;

  const nodes: ParsedNode[] = [];
  const edges: ParsedEdge[] = [];
  const subgraphs: ParsedSubgraph[] = [];
  const classDefs: Record<string, Record<string, string>> = {};
  const classAssignments: Record<string, string[]> = {};
  const styleOverrides: Record<string, Record<string, string>> = {};
  const layoutHints: { nodeSpacing?: number; rankSpacing?: number } = {};

  const subgraphStack: string[] = [];
  const registerClasses = (id: string, names?: string[]): void => {
    if (!Array.isArray(names) || names.length === 0) {
      return;
    }

    const current = classAssignments[id] ?? [];
    const merged = [...new Set([...current, ...names])];
    classAssignments[id] = merged;
  };

  for (const logical of logicalLines) {
    const lineNumber = logical.lineNumber;
    const rawLine = logical.text;

    const titleMatch = rawLine.match(/^\s*%%\s*title\s*:?\s*(.+?)\s*$/iu);
    if (titleMatch) {
      title = cleanLabel(titleMatch[1]);
    }

    const initHint = parseInitDirective(rawLine);
    if (initHint?.nodeSpacing !== undefined) {
      layoutHints.nodeSpacing = initHint.nodeSpacing;
    }
    if (initHint?.rankSpacing !== undefined) {
      layoutHints.rankSpacing = initHint.rankSpacing;
    }

    const commentStripped = stripComment(rawLine);
    const statements = splitStatements(commentStripped);

    for (const statement of statements) {
      const trimmed = statement.trim();
      if (!trimmed) {
        continue;
      }

      if (!hasHeader) {
        const header = trimmed.match(HEADER_RE);
        if (header) {
          hasHeader = true;
          direction = toDirection(header[2]);
          continue;
        }
      }

      if (/^subgraph\b/i.test(trimmed)) {
        const parsed = parseSubgraphHeader(trimmed);
        subgraphs.push({
          id: parsed.id,
          title: parsed.title,
          line: lineNumber,
        });
        subgraphStack.push(parsed.id);
        continue;
      }

      if (/^end$/i.test(trimmed)) {
        subgraphStack.pop();
        continue;
      }

      if (/^classDef\b/i.test(trimmed)) {
        const parsed = parseClassDefStatement(trimmed);
        if (parsed) {
          for (const name of parsed.names) {
            classDefs[name] = mergeStyleMap(classDefs[name] ?? {}, parsed.styles);
          }
        }
        continue;
      }

      if (/^class\b/i.test(trimmed)) {
        const parsed = parseClassStatement(trimmed);
        if (parsed) {
          for (const id of parsed.ids) {
            registerClasses(id, parsed.classes);
          }
        }
        continue;
      }

      if (/^style\b/i.test(trimmed)) {
        const parsed = parseStyleStatement(trimmed);
        if (parsed) {
          styleOverrides[parsed.id] = mergeStyleMap(styleOverrides[parsed.id] ?? {}, parsed.styles);
        }
        continue;
      }

      if (/^(linkStyle|click)\b/i.test(trimmed)) {
        continue;
      }

      if (isEdgeMetadataDirective(trimmed)) {
        continue;
      }

      const edgeParse = parseEdgeStatement(trimmed, lineNumber);
      if (edgeParse) {
        for (const node of edgeParse.nodes) {
          node.subgraphId = subgraphStack[subgraphStack.length - 1];
          registerClasses(node.id, node.inlineClasses);
          nodes.push(node);
        }
        edges.push(...edgeParse.edges);
        continue;
      }

      const nodeExpr = parseNodeExpr(trimmed);
      if (nodeExpr) {
        registerClasses(nodeExpr.id, nodeExpr.classes);
        nodes.push({
          id: nodeExpr.id,
          label: nodeExpr.label,
          shape: nodeExpr.shape,
          inlineClasses: nodeExpr.classes,
          subgraphId: subgraphStack[subgraphStack.length - 1],
          line: lineNumber,
          raw: trimmed,
        });
      }
    }
  }

  return {
    source,
    direction,
    title,
    nodes,
    edges,
    subgraphs,
    classDefs,
    classAssignments,
    styleOverrides,
    layoutHints,
  };
}
