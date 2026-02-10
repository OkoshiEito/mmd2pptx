import type { DiagramIr, IrNode } from "../types.js";

interface MeasureConfig {
  minWidth: number;
  maxWidth: number;
  minHeight: number;
  maxHeight: number;
  paddingX: number;
  paddingY: number;
  fontSize: number;
  lineHeightRatio: number;
}

interface MeasureOptions {
  targetAspectRatio?: number;
}

const DEFAULT_MEASURE_CONFIG: MeasureConfig = {
  minWidth: 128,
  maxWidth: 1200,
  minHeight: 68,
  maxHeight: 620,
  paddingX: 28,
  paddingY: 20,
  fontSize: 14,
  lineHeightRatio: 1.30,
};

const MIN_LAYOUT_ASPECT = 4 / 3;
const MAX_LAYOUT_ASPECT = 16 / 9;
const SOFT_BREAK_CHARS = new Set(["/", "\\", "_", "-", "+", "{", "}", "(", ")", "[", "]", ",", ":", ";", "|"]);

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function effectiveTextLength(text: string): number {
  let total = 0;
  for (const ch of text) {
    if (/\s/u.test(ch)) {
      total += 0.35;
      continue;
    }

    if (/[\u3000-\u9FFF\uF900-\uFAFF]/u.test(ch)) {
      total += 1.8;
      continue;
    }

    if (/[A-Z]/u.test(ch)) {
      total += 1.05;
      continue;
    }

    total += 0.95;
  }

  return total;
}

function splitLines(text: string): string[] {
  const lines = text.replace(/\r\n/g, "\n").split("\n");
  return lines.length > 0 ? lines : [text];
}

function isBreakChar(ch: string): boolean {
  return /\s/u.test(ch) || SOFT_BREAK_CHARS.has(ch);
}

function recomputeBreakState(chars: string[]): { units: number; lastBreakAt: number } {
  let units = 0;
  let lastBreakAt = -1;

  for (let i = 0; i < chars.length; i += 1) {
    const ch = chars[i];
    units += effectiveTextLength(ch);
    if (isBreakChar(ch)) {
      lastBreakAt = i + 1;
    }
  }

  return { units, lastBreakAt };
}

function wrapLineByUnits(line: string, maxUnits: number): string[] {
  const input = line.trimEnd();
  if (input.length === 0) {
    return [""];
  }

  const out: string[] = [];
  let chars: string[] = [];
  let units = 0;
  let lastBreakAt = -1;

  for (const ch of input) {
    chars.push(ch);
    units += effectiveTextLength(ch);

    if (isBreakChar(ch)) {
      lastBreakAt = chars.length;
    }

    if (units <= maxUnits) {
      continue;
    }

    if (lastBreakAt > 0) {
      const head = chars.slice(0, lastBreakAt).join("").trimEnd();
      if (head.length > 0) {
        out.push(head);
      }
      chars = chars.slice(lastBreakAt);
      const state = recomputeBreakState(chars);
      units = state.units;
      lastBreakAt = state.lastBreakAt;
      continue;
    }

    if (chars.length > 1) {
      const head = chars.slice(0, -1).join("").trimEnd();
      if (head.length > 0) {
        out.push(head);
      }
      chars = chars.slice(-1);
      const state = recomputeBreakState(chars);
      units = state.units;
      lastBreakAt = state.lastBreakAt;
    }
  }

  const tail = chars.join("").trimEnd();
  if (tail.length > 0 || out.length === 0) {
    out.push(tail);
  }

  return out;
}

function wrapLines(lines: string[], maxUnits: number): string[] {
  const wrapped = lines.flatMap((line) => wrapLineByUnits(line, maxUnits));
  return wrapped.length > 0 ? wrapped : [""];
}

function clampAspect(value: number | undefined): number {
  if (!Number.isFinite(value)) {
    return MAX_LAYOUT_ASPECT;
  }

  return clamp(value as number, MIN_LAYOUT_ASPECT, MAX_LAYOUT_ASPECT);
}

function lerp(min: number, max: number, t: number): number {
  return min + (max - min) * t;
}

function aspectRatioToT(aspect: number): number {
  return clamp((aspect - MIN_LAYOUT_ASPECT) / (MAX_LAYOUT_ASPECT - MIN_LAYOUT_ASPECT), 0, 1);
}

function maxWidthForAspect(aspect: number): number {
  const t = aspectRatioToT(aspect);
  // 4:3 is less tolerant to ultra-wide nodes, so cap width lower.
  return Math.round(lerp(500, 760, t));
}

function heightScaleForAspect(aspect: number): number {
  const t = aspectRatioToT(aspect);
  // 4:3 needs more vertical room per node to keep text inside after overall scaling.
  return lerp(1.28, 1.05, t);
}

function applyShapeAdjustment(node: IrNode): void {
  if (node.shape === "circle") {
    const side = Math.max(node.width, node.height);
    node.width = side;
    node.height = side;
    return;
  }

  if (node.shape === "diamond") {
    node.height = Math.max(node.height, node.width * 0.55);
  }
}

export function measureDiagram(ir: DiagramIr, options: MeasureOptions = {}): void {
  const aspect = clampAspect(options.targetAspectRatio);
  const maxWidth = Math.min(DEFAULT_MEASURE_CONFIG.maxWidth, maxWidthForAspect(aspect));
  const heightScale = heightScaleForAspect(aspect);
  const textFitSafety = 1.12;

  for (const node of ir.nodes) {
    if (node.isJunction) {
      node.label = "";
      node.width = 2;
      node.height = 2;
      continue;
    }

    const fontSize = node.style.fontSize || DEFAULT_MEASURE_CONFIG.fontSize;
    const lineHeight = fontSize * DEFAULT_MEASURE_CONFIG.lineHeightRatio;
    const lines = splitLines(node.label || node.id);

    const lineLengths = lines.map((line) => effectiveTextLength(line));
    const longestLine = Math.max(...lineLengths, effectiveTextLength(node.id));
    const rawTextWidth = longestLine * fontSize * 0.68;

    const hasIcon = Boolean(node.icon && node.icon.trim().length > 0);
    let width =
      node.width > 0
        ? clamp(node.width, DEFAULT_MEASURE_CONFIG.minWidth, maxWidth)
        : clamp(rawTextWidth + DEFAULT_MEASURE_CONFIG.paddingX * 2, DEFAULT_MEASURE_CONFIG.minWidth, maxWidth);

    if (hasIcon) {
      width = Math.max(width, 148);
    }

    const usableWidth = Math.max(24, width - DEFAULT_MEASURE_CONFIG.paddingX * 2);
    const maxUnitsPerLine = Math.max(6, usableWidth / (fontSize * 0.68));
    const wrappedLines = wrapLines(lines, maxUnitsPerLine);

    if (node.width <= 0) {
      const wrappedLengths = wrappedLines.map((line) => effectiveTextLength(line));
      const wrappedLongest = Math.max(...wrappedLengths, effectiveTextLength(node.id));
      width = clamp(
        (wrappedLongest * fontSize * 0.70 + DEFAULT_MEASURE_CONFIG.paddingX * 2) * textFitSafety,
        DEFAULT_MEASURE_CONFIG.minWidth,
        maxWidth,
      );
    }

    const lineCount = Math.max(1, wrappedLines.length);
    const iconHeadroom = hasIcon ? Math.max(20, fontSize * 1.3) : 0;
    const desiredHeight =
      (lineCount * (lineHeight * 1.03) + DEFAULT_MEASURE_CONFIG.paddingY * 2 + iconHeadroom) * heightScale * textFitSafety;

    const height =
      node.height > 0
        ? clamp(node.height, DEFAULT_MEASURE_CONFIG.minHeight, DEFAULT_MEASURE_CONFIG.maxHeight)
        : clamp(desiredHeight, DEFAULT_MEASURE_CONFIG.minHeight, DEFAULT_MEASURE_CONFIG.maxHeight);

    node.label = wrappedLines.join("\n");
    node.width = width;
    node.height = height;

    applyShapeAdjustment(node);
  }
}
