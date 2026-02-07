#!/usr/bin/env node
import { Command } from "commander";
import fs from "node:fs/promises";
import path from "node:path";
import { compileMmdToIr } from "./compiler/index.js";
import { parsePatchYaml } from "./compiler/patch.js";
import { renderPptx } from "./render/pptx.js";
import { renderPptxPython, renderSequencePptxPython } from "./render/python.js";

const program = new Command();

function defaultOutputPath(input: string): string {
  const parsed = path.parse(input);
  return path.join(parsed.dir, `${parsed.name}.pptx`);
}

function defaultMergedOutputPath(inputs: string[]): string {
  const first = inputs[0];
  const parsed = path.parse(first);
  return path.join(parsed.dir, `${parsed.name}.merged.pptx`);
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

program
  .name("mmd2pptx")
  .description("Compile Mermaid .mmd into editable .pptx")
  .version("0.1.0");

program
  .command("build")
  .argument("<inputs...>", "input .mmd file(s)")
  .option("-o, --output <path>", "output .pptx path")
  .option("--patch <path>", "patch yaml path")
  .option("--renderer <backend>", "render backend: python|js", "python")
  .option("--slide-size <size>", "slide size: 16:9 | 4:3 | <width>x<height> (inches)", "16:9")
  .option("--edge-routing <mode>", "edge routing: straight | elbow", "straight")
  .option("--ir-out <path>", "write normalized/layouted IR JSON")
  .option("--font-family <name>", "override rendering font family")
  .option("--lang <code>", "language tag for text rendering", "ja-JP")
  .action(async (inputs: string[], opts) => {
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

    const renderer = String(opts.renderer ?? "python").toLowerCase();
    if (multiInput && renderer !== "python") {
      throw new Error("Multiple input files currently support only --renderer python.");
    }
    if (multiInput && opts.irOut) {
      throw new Error("--ir-out is not supported when multiple input files are provided.");
    }

    for (let index = 0; index < inputs.length; index += 1) {
      const input = inputs[index];
      const sourceMmd = await fs.readFile(input, "utf8");
      const sequenceMode = isSequenceDiagramSource(sourceMmd);
      const appendToPath = multiInput && index > 0 ? outputPath : undefined;

      if (sequenceMode) {
        if (renderer !== "python") {
          throw new Error("sequenceDiagram currently supports only --renderer python");
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

      if (renderer === "js") {
        const ir = compileMmdToIr(sourceMmd, {
          patch,
          fontFamily: opts.fontFamily,
          lang: opts.lang,
          targetAspectRatio: slideSizeToAspectRatio(opts.slideSize),
        });
        if (opts.irOut) {
          await fs.writeFile(opts.irOut, `${JSON.stringify(ir, null, 2)}\n`, "utf8");
        }
        await renderPptx(ir, {
          outputPath,
          sourceMmd,
          patchText,
          slideSize: opts.slideSize,
          edgeRouting: opts.edgeRouting,
        });
        continue;
      }

      if (renderer === "python") {
        const ir = compileMmdToIr(sourceMmd, {
          patch,
          fontFamily: opts.fontFamily,
          lang: opts.lang,
          targetAspectRatio: slideSizeToAspectRatio(opts.slideSize),
        });
        if (opts.irOut) {
          await fs.writeFile(opts.irOut, `${JSON.stringify(ir, null, 2)}\n`, "utf8");
        }
        await renderPptxPython(ir, {
          outputPath,
          patchText,
          slideSize: opts.slideSize,
          edgeRouting: opts.edgeRouting,
          appendToPath,
        });
        continue;
      }

      throw new Error(`Unknown renderer: ${opts.renderer}. Use 'python' or 'js'.`);
    }

    process.stdout.write(`Generated: ${outputPath}\n`);
    if (opts.irOut) {
      process.stdout.write(`IR JSON: ${opts.irOut}\n`);
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

program.parseAsync(process.argv).catch((error) => {
  process.stderr.write(`${error instanceof Error ? error.stack ?? error.message : String(error)}\n`);
  process.exit(1);
});
