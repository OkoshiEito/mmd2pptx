import { createRequire } from "node:module";
import type { DiagramIr, NodeShape } from "../types.js";

const PX_PER_INCH = 96;
const SLIDE_MARGIN_IN = 0.35;
const require = createRequire(import.meta.url);
const PptxGenJS = require("pptxgenjs");

function pxToIn(value: number): number {
  return value / PX_PER_INCH;
}

function parseSlideSize(slideSize?: string): { width: number; height: number; layout: string } {
  const text = String(slideSize ?? "16:9").trim().toLowerCase();
  if (text === "4:3" || text === "standard") {
    return { width: 10, height: 7.5, layout: "LAYOUT_STANDARD" };
  }
  const custom = text.match(/^([0-9]+(?:\.[0-9]+)?)x([0-9]+(?:\.[0-9]+)?)$/u);
  if (custom) {
    const width = Number(custom[1]);
    const height = Number(custom[2]);
    if (Number.isFinite(width) && Number.isFinite(height) && width > 0 && height > 0) {
      return { width, height, layout: "CUSTOM_SLIDE" };
    }
  }
  return { width: 13.333, height: 7.5, layout: "LAYOUT_WIDE" };
}

function resolveScale(
  ir: DiagramIr,
  slideSize?: string,
): { scale: number; offsetX: number; offsetY: number; slideWidth: number; slideHeight: number; layout: string } {
  const slide = parseSlideSize(slideSize);
  const widthIn = pxToIn(ir.bounds.width || 1);
  const heightIn = pxToIn(ir.bounds.height || 1);

  const availableW = slide.width - SLIDE_MARGIN_IN * 2;
  const availableH = slide.height - SLIDE_MARGIN_IN * 2;

  const scale = Math.min(availableW / widthIn, availableH / heightIn);

  const minXIn = pxToIn(ir.bounds.minX);
  const minYIn = pxToIn(ir.bounds.minY);

  const contentW = widthIn * scale;
  const contentH = heightIn * scale;
  const offsetX = (slide.width - contentW) / 2 - minXIn * scale;
  const offsetY = (slide.height - contentH) / 2 - minYIn * scale;

  return {
    scale,
    offsetX,
    offsetY,
    slideWidth: slide.width,
    slideHeight: slide.height,
    layout: slide.layout,
  };
}

function shapeType(pptx: any, shape: NodeShape): string {
  switch (shape) {
    case "roundRect":
      return pptx.ShapeType.roundRect;
    case "circle":
      return pptx.ShapeType.ellipse;
    case "diamond":
      return pptx.ShapeType.diamond;
    case "parallelogram":
      return pptx.ShapeType.parallelogram;
    case "parallelogramAlt":
      return pptx.ShapeType.parallelogram;
    case "trapezoid":
      return pptx.ShapeType.trapezoid;
    case "trapezoidAlt":
      return pptx.ShapeType.nonIsoscelesTrapezoid;
    case "cloud":
      return pptx.ShapeType.cloud;
    case "lightningBolt":
      return pptx.ShapeType.lightningBolt;
    case "explosion":
      return pptx.ShapeType.chartStar;
    case "notchedRightArrow":
      return pptx.ShapeType.notchedRightArrow;
    case "hourglass":
      return pptx.ShapeType.flowChartCollate;
    case "hCylinder":
      return pptx.ShapeType.flowChartOnlineStorage;
    case "curvedTrapezoid":
      return pptx.ShapeType.flowChartDisplay;
    case "forkBar":
      return pptx.ShapeType.rect;
    case "windowPane":
      return pptx.ShapeType.flowChartInternalStorage;
    case "filledCircle":
      return pptx.ShapeType.ellipse;
    case "smallCircle":
      return pptx.ShapeType.ellipse;
    case "framedCircle":
      return pptx.ShapeType.donut;
    case "linedDocument":
      return pptx.ShapeType.flowChartDocument;
    case "linedRect":
      return pptx.ShapeType.flowChartInternalStorage;
    case "wave":
      return pptx.ShapeType.doubleWave;
    case "stackedRect":
      return pptx.ShapeType.flowChartPredefinedProcess;
    case "framedRect":
      return pptx.ShapeType.flowChartPredefinedProcess;
    case "braceLeft":
      return pptx.ShapeType.leftBrace;
    case "braceRight":
      return pptx.ShapeType.rightBrace;
    case "bracePair":
      return pptx.ShapeType.bracePair;
    case "card":
      return pptx.ShapeType.flowChartCard;
    case "delay":
      return pptx.ShapeType.flowChartDelay;
    case "internalStorage":
      return pptx.ShapeType.flowChartInternalStorage;
    case "document":
      return pptx.ShapeType.flowChartDocument;
    case "multiDocument":
      return pptx.ShapeType.flowChartMultidocument;
    case "triangle":
      return pptx.ShapeType.triangle;
    case "rightTriangle":
      return pptx.ShapeType.rtTriangle;
    case "chevron":
      return pptx.ShapeType.chevron;
    case "plaqueTabs":
      return pptx.ShapeType.plaqueTabs;
    case "pentagon":
      return pptx.ShapeType.pentagon;
    case "decagon":
      return pptx.ShapeType.decagon;
    case "foldedCorner":
      return pptx.ShapeType.folderCorner;
    case "donut":
      return pptx.ShapeType.donut;
    case "summingJunction":
      return pptx.ShapeType.flowChartSummingJunction;
    case "rect":
    default:
      return pptx.ShapeType.rect;
  }
}

function transformPoint(x: number, y: number, scale: number, offsetX: number, offsetY: number): { x: number; y: number } {
  return {
    x: pxToIn(x) * scale + offsetX,
    y: pxToIn(y) * scale + offsetY,
  };
}

function normalizeLineShape(
  from: { x: number; y: number },
  to: { x: number; y: number },
): { x: number; y: number; w: number; h: number; flipH?: boolean; flipV?: boolean } {
  const dx = to.x - from.x;
  const dy = to.y - from.y;
  const x = dx >= 0 ? from.x : to.x;
  const y = dy >= 0 ? from.y : to.y;
  const w = Math.abs(dx);
  const h = Math.abs(dy);

  return {
    x,
    y,
    w,
    h,
    flipH: dx < 0 || undefined,
    flipV: dy < 0 || undefined,
  };
}

function makeEmbedPayload(source: string, patchText?: string): string {
  return [
    "__MMD2PPTX_EMBED_START__",
    JSON.stringify(
      {
        sourceMmd: source,
        patchYaml: patchText ?? null,
        version: 1,
      },
      null,
      2,
    ),
    "__MMD2PPTX_EMBED_END__",
  ].join("\n");
}

interface RenderOptions {
  outputPath: string;
  sourceMmd: string;
  patchText?: string;
  slideSize?: string;
  edgeRouting?: string;
}

export async function renderPptx(ir: DiagramIr, options: RenderOptions): Promise<void> {
  const pptx: any = new (PptxGenJS as any)();
  const { scale, offsetX, offsetY, slideWidth, slideHeight, layout } = resolveScale(ir, options.slideSize);
  const routingMode = String(options.edgeRouting ?? "straight").toLowerCase() === "elbow" ? "elbow" : "straight";
  if (layout === "CUSTOM_SLIDE") {
    pptx.defineLayout({ name: "CUSTOM_SLIDE", width: slideWidth, height: slideHeight });
    pptx.layout = "CUSTOM_SLIDE";
  } else {
    pptx.layout = layout;
  }
  pptx.author = "mmd2pptx";
  pptx.subject = "Mermaid diagram";
  pptx.title = "mmd2pptx output";

  const slide = pptx.addSlide();

  for (const subgraph of ir.subgraphs) {
    const tl = transformPoint(subgraph.x, subgraph.y, scale, offsetX, offsetY);
    const w = pxToIn(subgraph.width) * scale;
    const h = pxToIn(subgraph.height) * scale;

    slide.addShape(pptx.ShapeType.roundRect, {
      x: tl.x,
      y: tl.y,
      w,
      h,
      line: {
        color: subgraph.style.stroke,
        pt: subgraph.style.strokeWidth,
        dash: subgraph.style.dash === "dash" ? "dash" : "solid",
      },
      fill: {
        color: subgraph.style.fill,
        transparency: 90,
      },
      radius: 0.04,
    });

    slide.addText(subgraph.title, {
      x: tl.x + 0.06,
      y: tl.y + 0.02,
      w: Math.max(0.4, w - 0.12),
      h: 0.24,
      fontFace: ir.config.fontFamily,
      fontSize: 10,
      bold: true,
      color: subgraph.style.text,
      align: "left",
      valign: "mid",
      lang: ir.config.lang,
    });
  }

  for (const edge of ir.edges) {
    if (edge.style.lineStyle === "invisible") {
      continue;
    }
    if (edge.points.length < 2) {
      continue;
    }

    const start = edge.points[0];
    const end = edge.points[edge.points.length - 1];
    const from = transformPoint(start.x, start.y, scale, offsetX, offsetY);
    const to = transformPoint(end.x, end.y, scale, offsetX, offsetY);

    const lineOptions: Record<string, unknown> = {
      color: edge.style.color,
      pt: edge.style.width,
      dash: edge.style.lineStyle === "dotted" ? "dot" : "solid",
    };

    if (edge.style.arrow === "start" || edge.style.arrow === "both") {
      lineOptions.beginArrowType = "triangle";
    }

    if (edge.style.arrow === "end" || edge.style.arrow === "both") {
      lineOptions.endArrowType = "triangle";
    }

    const drawLineSegment = (
      p0: { x: number; y: number },
      p1: { x: number; y: number },
      segmentLineOptions: Record<string, unknown>,
    ): void => {
      if (Math.abs(p1.x - p0.x) < 1e-6 && Math.abs(p1.y - p0.y) < 1e-6) {
        return;
      }
      const lineShape = normalizeLineShape(p0, p1);
      slide.addShape(pptx.ShapeType.line, {
        x: lineShape.x,
        y: lineShape.y,
        w: lineShape.w,
        h: lineShape.h,
        flipH: lineShape.flipH,
        flipV: lineShape.flipV,
        line: segmentLineOptions as any,
      });
    };

    if (routingMode === "elbow") {
      const dx = to.x - from.x;
      const dy = to.y - from.y;
      let points: Array<{ x: number; y: number }>;
      if (Math.abs(dx) >= Math.abs(dy)) {
        const midX = from.x + dx / 2;
        points = [from, { x: midX, y: from.y }, { x: midX, y: to.y }, to];
      } else {
        const midY = from.y + dy / 2;
        points = [from, { x: from.x, y: midY }, { x: to.x, y: midY }, to];
      }

      for (let i = 0; i < points.length - 1; i += 1) {
        const segmentLineOptions: Record<string, unknown> = {
          color: edge.style.color,
          pt: edge.style.width,
          dash: edge.style.lineStyle === "dotted" ? "dot" : "solid",
        };
        if (i === 0 && (edge.style.arrow === "start" || edge.style.arrow === "both")) {
          segmentLineOptions.beginArrowType = "triangle";
        }
        if (i === points.length - 2 && (edge.style.arrow === "end" || edge.style.arrow === "both")) {
          segmentLineOptions.endArrowType = "triangle";
        }
        drawLineSegment(points[i], points[i + 1], segmentLineOptions);
      }
    } else {
      drawLineSegment(from, to, lineOptions);
    }

    if (edge.label && edge.labelPosition) {
      const label = transformPoint(edge.labelPosition.x, edge.labelPosition.y, scale, offsetX, offsetY);
      slide.addText(edge.label, {
        x: label.x - 0.4,
        y: label.y - 0.12,
        w: 0.8,
        h: 0.24,
        fontFace: ir.config.fontFamily,
        fontSize: edge.style.fontSize,
        color: edge.style.color,
        bold: false,
        align: "center",
        valign: "mid",
        margin: 0,
        lang: ir.config.lang,
      });
    }
  }

  for (const node of ir.nodes) {
    const tl = transformPoint(node.x, node.y, scale, offsetX, offsetY);
    const w = pxToIn(node.width) * scale;
    const h = pxToIn(node.height) * scale;

    slide.addShape(shapeType(pptx, node.shape), {
      x: tl.x,
      y: tl.y,
      w,
      h,
      line: {
        color: node.style.stroke,
        pt: node.style.strokeWidth,
      },
      fill: {
        color: node.style.fill,
      },
      radius: node.style.radius,
    });

    slide.addText(node.label, {
      x: tl.x,
      y: tl.y,
      w,
      h,
      fontFace: ir.config.fontFamily,
      fontSize: node.style.fontSize,
      color: node.style.text,
      bold: false,
      align: "center",
      valign: "mid",
      breakLine: false,
      fit: "shrink",
      margin: 0.05,
      lang: ir.config.lang,
    });
  }

  const titleText = ir.meta.title?.trim();
  if (titleText) {
    slide.addText(titleText, {
      x: 0.2,
      y: 0.08,
      w: Math.max(0.8, slideWidth - 0.4),
      h: 0.24,
      fontFace: ir.config.fontFamily,
      fontSize: 10,
      bold: true,
      color: "334155",
      align: "left",
      valign: "top",
      margin: 0,
      lang: ir.config.lang,
    });
  }

  const embedText = makeEmbedPayload(options.sourceMmd, options.patchText);
  const noteSlide = slide as unknown as { addNotes?: (notes: string[]) => void };
  if (typeof noteSlide.addNotes === "function") {
    noteSlide.addNotes([embedText]);
  } else {
    slide.addText(embedText, {
      x: 0.01,
      y: slideHeight - 0.04,
      w: 0.01,
      h: 0.01,
      fontFace: "Courier New",
      fontSize: 1,
      color: "FFFFFF",
      transparency: 100,
    });
  }

  await pptx.writeFile({ fileName: options.outputPath });
}
