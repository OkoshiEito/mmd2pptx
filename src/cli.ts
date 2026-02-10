#!/usr/bin/env node
import { spawn } from "node:child_process";
import { Command, Option } from "commander";
import fs from "node:fs/promises";
import path from "node:path";
import { compileMmdToIr } from "./compiler/index.js";
import { parsePatchYaml } from "./compiler/patch.js";
import { renderPptxPython, renderSequencePptxPython } from "./render/python.js";

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

interface ResolvedBuildInputs {
  files: string[];
  usedDirectoryInput: boolean;
  firstInputWasDirectory: boolean;
  firstDirectoryPath?: string;
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

function isMmdFilePath(filePath: string): boolean {
  return /\.mmd$/iu.test(filePath);
}

async function collectMmdFilesRecursively(directory: string): Promise<string[]> {
  const out: string[] = [];
  const entries = await fs.readdir(directory, { withFileTypes: true });
  entries.sort((a, b) => a.name.localeCompare(b.name, "en"));

  for (const entry of entries) {
    const fullPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      out.push(...(await collectMmdFilesRecursively(fullPath)));
      continue;
    }
    if (!entry.isFile()) {
      continue;
    }
    if (!isMmdFilePath(entry.name)) {
      continue;
    }
    out.push(fullPath);
  }

  return out;
}

async function resolveBuildInputs(inputs: string[]): Promise<ResolvedBuildInputs> {
  const files: string[] = [];
  let usedDirectoryInput = false;
  let firstInputWasDirectory = false;
  let firstDirectoryPath: string | undefined;

  for (let i = 0; i < inputs.length; i += 1) {
    const input = path.resolve(inputs[i]);
    let stat;
    try {
      stat = await fs.stat(input);
    } catch {
      throw new Error(`Input not found: ${inputs[i]}`);
    }

    if (stat.isDirectory()) {
      usedDirectoryInput = true;
      if (i === 0) {
        firstInputWasDirectory = true;
        firstDirectoryPath = input;
      }
      const discovered = await collectMmdFilesRecursively(input);
      if (discovered.length === 0) {
        throw new Error(`No .mmd files found under directory: ${inputs[i]}`);
      }
      files.push(...discovered);
      continue;
    }

    if (!stat.isFile()) {
      throw new Error(`Unsupported input type: ${inputs[i]}`);
    }
    if (!isMmdFilePath(input)) {
      throw new Error(`Input must be .mmd or a directory: ${inputs[i]}`);
    }
    files.push(input);
  }

  if (files.length === 0) {
    throw new Error("No .mmd input files were resolved.");
  }

  return {
    files,
    usedDirectoryInput,
    firstInputWasDirectory,
    firstDirectoryPath,
  };
}

function defaultMergedOutputPathForResolvedInputs(resolved: ResolvedBuildInputs): string {
  if (resolved.firstInputWasDirectory && resolved.firstDirectoryPath) {
    const dirName = path.basename(resolved.firstDirectoryPath);
    return path.join(resolved.firstDirectoryPath, `${dirName}.merged.pptx`);
  }
  return defaultMergedOutputPath(resolved.files);
}

function applySequenceFileTitle(sourceMmd: string, fileTitle: string): string {
  const lines = sourceMmd.replace(/\r\n/g, "\n").split("\n");
  const out: string[] = [];
  let headerSeen = false;
  let titleInjected = false;

  for (const raw of lines) {
    const trimmed = raw.trim();
    if (!headerSeen) {
      out.push(raw);
      if (/^sequenceDiagram\b/iu.test(trimmed)) {
        out.push(`title ${fileTitle}`);
        headerSeen = true;
        titleInjected = true;
      }
      continue;
    }

    if (/^title\s*:?.*$/iu.test(trimmed)) {
      continue;
    }

    out.push(raw);
  }

  if (!titleInjected) {
    return sourceMmd;
  }
  return out.join("\n");
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

function commandExists(command: string): Promise<boolean> {
  return new Promise((resolve) => {
    const check = process.platform === "win32" ? "where" : "which";
    const child = spawn(check, [command], { stdio: "ignore" });
    child.on("error", () => resolve(false));
    child.on("close", (code) => resolve(code === 0));
  });
}

async function runBuild(inputs: string[], opts: BuildCliOptions): Promise<void> {
  if (!Array.isArray(inputs) || inputs.length === 0) {
    throw new Error("At least one input .mmd file is required.");
  }

  const resolved = await resolveBuildInputs(inputs);
  const inputFiles = resolved.files;
  const multiInput = inputFiles.length > 1;
  const outputPath = opts.output ?? (multiInput ? defaultMergedOutputPathForResolvedInputs(resolved) : defaultOutputPath(inputFiles[0]));
  const shouldStampFilenameTitle = multiInput || resolved.usedDirectoryInput;

  if (resolved.usedDirectoryInput) {
    process.stdout.write(`Resolved ${inputFiles.length} .mmd files from directory input\n`);
  }

  let patchText: string | undefined;
  let patch;
  if (opts.patch) {
    if (multiInput) {
      throw new Error("--patch is not supported when multiple input files are provided.");
    }
    patchText = await fs.readFile(opts.patch, "utf8");
    patch = parsePatchYaml(patchText);
  }

  const rendererOption = String(opts.renderer ?? "").trim().toLowerCase();
  if (rendererOption && rendererOption !== "python" && rendererOption !== "auto") {
    process.stdout.write("Warning: --renderer is deprecated and ignored. Using python renderer.\n");
  }
  if (multiInput && opts.irOut) {
    throw new Error("--ir-out is not supported when multiple input files are provided.");
  }

  for (let index = 0; index < inputFiles.length; index += 1) {
    const input = inputFiles[index];
    const fileTitle = path.basename(input);
    const sourceMmd = await fs.readFile(input, "utf8");
    const sequenceMode = isSequenceDiagramSource(sourceMmd);
    const appendToPath = multiInput && index > 0 ? outputPath : undefined;

    if (sequenceMode) {
      const sequenceSource = shouldStampFilenameTitle ? applySequenceFileTitle(sourceMmd, fileTitle) : sourceMmd;
      await renderSequencePptxPython(sequenceSource, {
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
    if (shouldStampFilenameTitle) {
      ir.meta.title = fileTitle;
    }
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
  }

  process.stdout.write(`Generated: ${outputPath}\n`);
  if (opts.irOut) {
    process.stdout.write(`IR JSON: ${opts.irOut}\n`);
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

  const knownCommands = new Set(["build", "extract", "doctor", "help"]);
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
  .argument("<inputs...>", "input .mmd file(s) or directories")
  .option("-o, --output <path>", "output .pptx path")
  .option("-p, --patch <path>", "patch yaml path")
  .addOption(new Option("-r, --renderer <backend>").hideHelp())
  .option("-s, --slide-size <size>", "slide size: 16:9 | 4:3 | <width>x<height> (inches)", "16:9")
  .option("-e, --edge-routing <mode>", "edge routing: straight | elbow", "straight")
  .option("--ir-out <path>", "write normalized/layouted IR JSON")
  .option("--font-family <name>", "override rendering font family")
  .option("--lang <code>", "language tag for text rendering", "ja-JP")
  .action(async (inputs: string[], opts: BuildCliOptions) => runBuild(inputs, opts));

program
  .command("doctor")
  .description("Check runtime dependencies")
  .action(async () => {
    const hasPython3 = await commandExists("python3");
    const hasPython = hasPython3 ? true : await commandExists("python");
    process.stdout.write(`Node: ${process.version}\n`);
    process.stdout.write(`renderer: python (fixed)\n`);
    process.stdout.write(`python3: ${hasPython3 ? "found" : "not found"}\n`);
    if (!hasPython3) {
      process.stdout.write(`python: ${hasPython ? "found" : "not found"}\n`);
    }
    if (!hasPython) {
      process.stdout.write("tip: install Python 3 to use mmd2pptx\n");
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
