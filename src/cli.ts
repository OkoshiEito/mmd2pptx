#!/usr/bin/env node
import { spawn } from "node:child_process";
import { Command } from "commander";
import fs from "node:fs/promises";
import path from "node:path";
import yaml from "js-yaml";
import { compileMmdToIr } from "./compiler/index.js";
import { parsePatchYaml } from "./compiler/patch.js";
import { renderPptx } from "./render/pptx.js";
import { renderPptxPython, renderSequencePptxPython } from "./render/python.js";
import { evaluateReadability } from "./quality/readability.js";
import type { DiagramPatch, LayoutConfig } from "./types.js";

const program = new Command();

interface BuildCliOptions {
  output?: string;
  patch?: string;
  renderer?: string;
  slideSize?: string;
  edgeRouting?: string;
  irOut?: string;
  fontFamily?: string;
  lang?: string;
}

interface PdcaCliOptions {
  patch?: string;
  slideSize?: string;
  maxTrials?: string;
  reportOut?: string;
  bestPatchOut?: string;
  emitBestPptx?: boolean;
  renderer?: string;
  edgeRouting?: string;
  fontFamily?: string;
  lang?: string;
}

function defaultOutputPath(input: string): string {
  const parsed = path.parse(input);
  return path.join(parsed.dir, `${parsed.name}.pptx`);
}

function defaultMergedOutputPath(inputs: string[]): string {
  const first = inputs[0];
  const parsed = path.parse(first);
  return path.join(parsed.dir, `${parsed.name}.merged.pptx`);
}

function defaultPdcaReportPath(input: string): string {
  const parsed = path.parse(input);
  return path.join(parsed.dir, `${parsed.name}.pdca.report.json`);
}

function defaultPdcaPatchPath(input: string): string {
  const parsed = path.parse(input);
  return path.join(parsed.dir, `${parsed.name}.pdca.patch.yml`);
}

function defaultPdcaBestOutputPath(input: string): string {
  const parsed = path.parse(input);
  return path.join(parsed.dir, `${parsed.name}.pdca-best.pptx`);
}

function clamp(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function slideSizeToAspectRatio(slideSize: string | undefined): number {
  const text = String(slideSize ?? "16:9").trim().toLowerCase();
  if (text === "4:3" || text === "standard") {
    return 4 / 3;
  }
  if (text === "16:9" || text === "wide" || text === "widescreen") {
    return 16 / 9;
  }

  const matched = text.match(/^([0-9]+(?:\.[0-9]+)?)x([0-9]+(?:\.[0-9]+)?)$/u);
  if (!matched) {
    return 16 / 9;
  }

  const width = Number(matched[1]);
  const height = Number(matched[2]);
  if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0) {
    return 16 / 9;
  }
  return width / height;
}

function isSequenceDiagramSource(source: string): boolean {
  const lines = source.replace(/\r\n/g, "\n").split("\n");
  for (const raw of lines) {
    const line = raw.trim();
    if (!line) {
      continue;
    }
    if (line.startsWith("%%")) {
      continue;
    }
    return /^sequenceDiagram\b/i.test(line);
  }
  return false;
}

function normalizeRendererOption(input: string | undefined): string {
  return String(input ?? "auto").trim().toLowerCase();
}

function commandExists(command: string): Promise<boolean> {
  return new Promise((resolve) => {
    const check = process.platform === "win32" ? "where" : "which";
    const child = spawn(check, [command], { stdio: "ignore" });
    child.on("error", () => resolve(false));
    child.on("close", (code) => resolve(code === 0));
  });
}

async function resolveRenderer(rendererOption: string | undefined): Promise<"python" | "js"> {
  const renderer = normalizeRendererOption(rendererOption);
  if (renderer === "python" || renderer === "js") {
    return renderer;
  }
  if (renderer !== "auto") {
    throw new Error(`Unknown renderer: ${rendererOption}. Use 'auto', 'python', or 'js'.`);
  }

  const hasUv = await commandExists("uv");
  return hasUv ? "python" : "js";
}

async function runBuild(inputs: string[], opts: BuildCliOptions): Promise<void> {
  if (!Array.isArray(inputs) || inputs.length === 0) {
    throw new Error("At least one input .mmd file is required.");
  }

  const multiInput = inputs.length > 1;
  const outputPath = opts.output ?? (multiInput ? defaultMergedOutputPath(inputs) : defaultOutputPath(inputs[0]));

  let patchText: string | undefined;
  let patch;
  if (opts.patch) {
    if (multiInput) {
      throw new Error("--patch is not supported when multiple input files are provided.");
    }
    patchText = await fs.readFile(opts.patch, "utf8");
    patch = parsePatchYaml(patchText);
  }

  const rendererOption = normalizeRendererOption(opts.renderer);
  const renderer = await resolveRenderer(opts.renderer);
  if (multiInput && renderer !== "python") {
    throw new Error("Multiple input files require python renderer. Install uv or set --renderer python.");
  }
  if (multiInput && opts.irOut) {
    throw new Error("--ir-out is not supported when multiple input files are provided.");
  }
  if (rendererOption === "auto") {
    process.stdout.write(`Renderer(auto): ${renderer}\n`);
  }

  for (let index = 0; index < inputs.length; index += 1) {
    const input = inputs[index];
    const sourceMmd = await fs.readFile(input, "utf8");
    const sequenceMode = isSequenceDiagramSource(sourceMmd);
    const appendToPath = multiInput && index > 0 ? outputPath : undefined;

    if (sequenceMode) {
      if (renderer !== "python") {
        throw new Error("sequenceDiagram currently supports only python renderer (install uv).");
      }
      await renderSequencePptxPython(sourceMmd, {
        outputPath,
        patchText,
        slideSize: opts.slideSize,
        edgeRouting: opts.edgeRouting,
        appendToPath,
      });
      if (opts.irOut) {
        await fs.writeFile(
          opts.irOut,
          `${JSON.stringify({ diagramType: "sequence", note: "sequence mode has no flowchart IR output" }, null, 2)}\n`,
          "utf8",
        );
      }
      continue;
    }

    const ir = compileMmdToIr(sourceMmd, {
      patch,
      fontFamily: opts.fontFamily,
      lang: opts.lang,
      targetAspectRatio: slideSizeToAspectRatio(opts.slideSize),
    });
    if (opts.irOut) {
      await fs.writeFile(opts.irOut, `${JSON.stringify(ir, null, 2)}\n`, "utf8");
    }

    if (renderer === "js") {
      await renderPptx(ir, {
        outputPath,
        sourceMmd,
        patchText,
        slideSize: opts.slideSize,
        edgeRouting: opts.edgeRouting,
      });
      continue;
    }

    await renderPptxPython(ir, {
      outputPath,
      patchText,
      slideSize: opts.slideSize,
      edgeRouting: opts.edgeRouting,
      appendToPath,
    });
  }

  process.stdout.write(`Generated: ${outputPath}\n`);
  if (opts.irOut) {
    process.stdout.write(`IR JSON: ${opts.irOut}\n`);
  }
}

function mergePatchWithLayout(basePatch: DiagramPatch | undefined, layout: Partial<LayoutConfig>): DiagramPatch {
  return {
    ...(basePatch ?? {}),
    layout: {
      ...(basePatch?.layout ?? {}),
      ...layout,
    },
  };
}

function buildPdcaLayoutTrials(baseLayout: LayoutConfig, maxTrials: number): Array<{ id: string; layout: LayoutConfig }> {
  const nodeScales = [0.78, 0.88, 0.96, 1.0, 1.08, 1.18, 1.3];
  const rankScales = [0.78, 0.88, 0.96, 1.0, 1.08, 1.18, 1.3];
  const edgeScales = [0.86, 0.95, 1.0, 1.1, 1.2];
  const marginScales = [0.86, 1.0, 1.14];

  const candidates: Array<{ key: string; priority: number; layout: LayoutConfig }> = [];

  for (const ns of nodeScales) {
    for (const rs of rankScales) {
      for (const es of edgeScales) {
        for (const ms of marginScales) {
          const layout: LayoutConfig = {
            nodesep: Math.round(clamp(baseLayout.nodesep * ns, 22, 360)),
            ranksep: Math.round(clamp(baseLayout.ranksep * rs, 30, 460)),
            edgesep: Math.round(clamp(baseLayout.edgesep * es, 16, 220)),
            marginx: Math.round(clamp(baseLayout.marginx * ms, 8, 180)),
            marginy: Math.round(clamp(baseLayout.marginy * ms, 8, 180)),
          };

          const anisotropy = Math.abs(Math.log((layout.nodesep + 1) / (layout.ranksep + 1)));
          const distance =
            Math.abs(layout.nodesep - baseLayout.nodesep) / Math.max(1, baseLayout.nodesep) +
            Math.abs(layout.ranksep - baseLayout.ranksep) / Math.max(1, baseLayout.ranksep) +
            Math.abs(layout.edgesep - baseLayout.edgesep) / Math.max(1, baseLayout.edgesep) * 0.7 +
            Math.abs(layout.marginx - baseLayout.marginx) / Math.max(1, baseLayout.marginx) * 0.4;
          const priority = anisotropy * 0.55 + distance;

          const key = `${layout.nodesep}:${layout.ranksep}:${layout.edgesep}:${layout.marginx}:${layout.marginy}`;
          candidates.push({ key, priority, layout });
        }
      }
    }
  }

  const unique = new Map<string, { key: string; priority: number; layout: LayoutConfig }>();
  for (const candidate of candidates) {
    if (!unique.has(candidate.key) || candidate.priority < (unique.get(candidate.key)?.priority ?? Number.POSITIVE_INFINITY)) {
      unique.set(candidate.key, candidate);
    }
  }

  const sorted = [...unique.values()].sort((a, b) => a.priority - b.priority);
  const baselineKey = `${baseLayout.nodesep}:${baseLayout.ranksep}:${baseLayout.edgesep}:${baseLayout.marginx}:${baseLayout.marginy}`;
  const baseline = sorted.find((item) => item.key === baselineKey) ?? {
    key: baselineKey,
    priority: -1,
    layout: { ...baseLayout },
  };

  const count = Math.max(3, maxTrials);
  const nearCount = Math.max(1, Math.floor(count * 0.72));
  const farCount = Math.max(1, count - nearCount);
  const near = sorted.slice(0, nearCount);
  const far = sorted.slice(Math.max(0, sorted.length - farCount));

  const picked = new Map<string, { id: string; layout: LayoutConfig }>();
  picked.set(baseline.key, { id: "baseline", layout: baseline.layout });
  for (const item of [...near, ...far]) {
    if (!picked.has(item.key)) {
      picked.set(item.key, { id: `trial-${picked.size}`, layout: item.layout });
    }
    if (picked.size >= count) {
      break;
    }
  }

  return [...picked.values()];
}

async function runPdca(inputs: string[], opts: PdcaCliOptions): Promise<void> {
  if (!Array.isArray(inputs) || inputs.length === 0) {
    throw new Error("At least one input .mmd file is required.");
  }

  let basePatch: DiagramPatch | undefined;
  let patchText: string | undefined;
  if (opts.patch) {
    if (inputs.length > 1) {
      throw new Error("--patch is supported only for single-input PDCA.");
    }
    patchText = await fs.readFile(opts.patch, "utf8");
    basePatch = parsePatchYaml(patchText);
  }

  const sources: Array<{ input: string; source: string }> = [];
  for (const input of inputs) {
    const source = await fs.readFile(input, "utf8");
    if (isSequenceDiagramSource(source)) {
      throw new Error(`PDCA currently supports flow/class/architecture layouts only: ${input}`);
    }
    sources.push({ input, source });
  }

  const targetAspectRatio = slideSizeToAspectRatio(opts.slideSize);
  const baselineIr = compileMmdToIr(sources[0].source, {
    patch: basePatch,
    fontFamily: opts.fontFamily,
    lang: opts.lang,
    targetAspectRatio,
  });
  const trialCount = clamp(Number.parseInt(String(opts.maxTrials ?? "36"), 10) || 36, 3, 120);
  const trials = buildPdcaLayoutTrials(baselineIr.config.layout, trialCount);

  process.stdout.write(`PDCA: evaluating ${trials.length} layout candidates across ${sources.length} file(s)\n`);

  const trialResults: Array<{
    id: string;
    layout: LayoutConfig;
    penalty: number;
    score: number;
    perInput: Array<{ input: string; penalty: number; score: number }>;
  }> = [];

  for (const trial of trials) {
    const patch = mergePatchWithLayout(basePatch, trial.layout);
    const perInput: Array<{ input: string; penalty: number; score: number }> = [];
    let penaltySum = 0;
    let scoreSum = 0;

    for (const entry of sources) {
      const ir = compileMmdToIr(entry.source, {
        patch,
        fontFamily: opts.fontFamily,
        lang: opts.lang,
        targetAspectRatio,
      });
      const evalResult = evaluateReadability(ir);
      perInput.push({
        input: entry.input,
        penalty: evalResult.penalty,
        score: evalResult.score,
      });
      penaltySum += evalResult.penalty;
      scoreSum += evalResult.score;
    }

    trialResults.push({
      id: trial.id,
      layout: trial.layout,
      penalty: penaltySum / Math.max(1, sources.length),
      score: scoreSum / Math.max(1, sources.length),
      perInput,
    });
  }

  trialResults.sort((a, b) => a.penalty - b.penalty);
  const best = trialResults[0];
  const baseline = trialResults.find((trial) => trial.id === "baseline") ?? trialResults[0];
  const improvementPct = baseline.penalty > 0 ? ((baseline.penalty - best.penalty) / baseline.penalty) * 100 : 0;

  const bestPatch = mergePatchWithLayout(basePatch, best.layout);
  const bestPatchText = yaml.dump(bestPatch, {
    noRefs: true,
    lineWidth: 120,
  });

  const report = {
    generatedAt: new Date().toISOString(),
    inputs,
    trials: trialResults.length,
    baseline: {
      id: baseline.id,
      penalty: baseline.penalty,
      score: baseline.score,
      layout: baseline.layout,
    },
    best: {
      id: best.id,
      penalty: best.penalty,
      score: best.score,
      layout: best.layout,
    },
    improvementPct,
    topTrials: trialResults.slice(0, Math.min(10, trialResults.length)),
  };

  const reportOut = opts.reportOut ?? defaultPdcaReportPath(inputs[0]);
  const bestPatchOut = opts.bestPatchOut ?? defaultPdcaPatchPath(inputs[0]);
  await fs.writeFile(reportOut, `${JSON.stringify(report, null, 2)}\n`, "utf8");
  await fs.writeFile(bestPatchOut, bestPatchText, "utf8");

  process.stdout.write(`PDCA baseline penalty: ${baseline.penalty.toFixed(2)} (score ${baseline.score.toFixed(2)})\n`);
  process.stdout.write(`PDCA best penalty: ${best.penalty.toFixed(2)} (score ${best.score.toFixed(2)})\n`);
  process.stdout.write(`PDCA improvement: ${improvementPct.toFixed(2)}%\n`);
  process.stdout.write(`PDCA report: ${reportOut}\n`);
  process.stdout.write(`PDCA best patch: ${bestPatchOut}\n`);

  if (opts.emitBestPptx) {
    const renderer = await resolveRenderer(opts.renderer);
    if (normalizeRendererOption(opts.renderer) === "auto") {
      process.stdout.write(`Renderer(auto): ${renderer}\n`);
    }

    for (const entry of sources) {
      const outputPath = defaultPdcaBestOutputPath(entry.input);
      const ir = compileMmdToIr(entry.source, {
        patch: bestPatch,
        fontFamily: opts.fontFamily,
        lang: opts.lang,
        targetAspectRatio,
      });

      if (renderer === "js") {
        await renderPptx(ir, {
          outputPath,
          sourceMmd: entry.source,
          patchText: bestPatchText,
          slideSize: opts.slideSize,
          edgeRouting: opts.edgeRouting,
        });
      } else {
        await renderPptxPython(ir, {
          outputPath,
          patchText: bestPatchText,
          slideSize: opts.slideSize,
          edgeRouting: opts.edgeRouting,
        });
      }
      process.stdout.write(`PDCA best pptx: ${outputPath}\n`);
    }
  }
}

function normalizeArgvForDefaultBuild(argv: string[]): string[] {
  if (argv.length < 3) {
    return argv;
  }

  const first = argv[2];
  if (first.startsWith("-")) {
    return argv;
  }

  const knownCommands = new Set(["build", "pdca", "extract", "doctor", "help"]);
  if (knownCommands.has(first)) {
    return argv;
  }

  return [argv[0], argv[1], "build", ...argv.slice(2)];
}

program
  .name("mmd2pptx")
  .description("Compile Mermaid .mmd into editable .pptx")
  .version("0.1.0");

program
  .command("build")
  .argument("<inputs...>", "input .mmd file(s)")
  .option("-o, --output <path>", "output .pptx path")
  .option("-p, --patch <path>", "patch yaml path")
  .option("-r, --renderer <backend>", "render backend: auto|python|js", "auto")
  .option("-s, --slide-size <size>", "slide size: 16:9 | 4:3 | <width>x<height> (inches)", "16:9")
  .option("-e, --edge-routing <mode>", "edge routing: straight | elbow", "straight")
  .option("--ir-out <path>", "write normalized/layouted IR JSON")
  .option("--font-family <name>", "override rendering font family")
  .option("--lang <code>", "language tag for text rendering", "ja-JP")
  .action(async (inputs: string[], opts: BuildCliOptions) => runBuild(inputs, opts));

program
  .command("pdca")
  .description("Run readability PDCA loop (measure -> search -> report -> best patch)")
  .argument("<inputs...>", "input .mmd file(s)")
  .option("-p, --patch <path>", "base patch yaml path (single input only)")
  .option("-s, --slide-size <size>", "slide size: 16:9 | 4:3 | <width>x<height> (inches)", "16:9")
  .option("--max-trials <n>", "max layout candidates to evaluate", "36")
  .option("--report-out <path>", "output PDCA report json path")
  .option("--best-patch-out <path>", "output best patch yaml path")
  .option("--emit-best-pptx", "also generate best .pptx output(s)")
  .option("-r, --renderer <backend>", "render backend when --emit-best-pptx: auto|python|js", "auto")
  .option("-e, --edge-routing <mode>", "edge routing for --emit-best-pptx: straight | elbow", "straight")
  .option("--font-family <name>", "override rendering font family")
  .option("--lang <code>", "language tag for text rendering", "ja-JP")
  .action(async (inputs: string[], opts: PdcaCliOptions) => runPdca(inputs, opts));

program
  .command("doctor")
  .description("Check runtime dependencies for renderer selection")
  .action(async () => {
    const hasUv = await commandExists("uv");
    process.stdout.write(`Node: ${process.version}\n`);
    process.stdout.write(`uv: ${hasUv ? "found" : "not found"}\n`);
    process.stdout.write(`default renderer(auto): ${hasUv ? "python" : "js"}\n`);
    if (!hasUv) {
      process.stdout.write("tip: install uv to use python renderer (sequenceDiagram and best fidelity)\n");
    }
  });

program
  .command("extract")
  .argument("<input>", "input .pptx file")
  .description("Reserved for future: extract embedded mmd/patch")
  .action((input: string) => {
    process.stdout.write(`extract is not implemented yet: ${input}\n`);
    process.exitCode = 1;
  });

const argv = normalizeArgvForDefaultBuild([...process.argv]);
program.parseAsync(argv).catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack ?? error.message : String(error)}\n`);
  process.exit(1);
});
