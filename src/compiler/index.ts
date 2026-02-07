import { layoutDiagram } from "./layout.js";
import { measureDiagram } from "./measure.js";
import { normalizeDiagram } from "./normalize.js";
import { parseMermaid } from "./parse.js";
import { applyPatchPostLayout, applyPatchPreLayout } from "./patch.js";
import type { BuildOptions, DiagramIr } from "../types.js";

export function compileMmdToIr(source: string, options: BuildOptions = {}): DiagramIr {
  const ast = parseMermaid(source);
  const ir = normalizeDiagram(ast, {
    fontFamily: options.fontFamily,
    lang: options.lang,
  });

  measureDiagram(ir, {
    targetAspectRatio: options.targetAspectRatio,
  });
  applyPatchPreLayout(ir, options.patch);
  layoutDiagram(ir, {
    targetAspectRatio: options.targetAspectRatio,
  });
  applyPatchPostLayout(ir, options.patch);

  return ir;
}
