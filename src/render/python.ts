import { spawn } from "node:child_process";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import type { DiagramIr } from "../types.js";

interface RenderOptions {
  outputPath: string;
  patchText?: string;
  slideSize?: string;
  edgeRouting?: string;
  appendToPath?: string;
}

function runCommand(command: string, args: string[]): Promise<void> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { stdio: ["ignore", "pipe", "pipe"] });
    let stdout = "";
    let stderr = "";

    child.stdout.on("data", (chunk) => {
      stdout += String(chunk);
    });

    child.stderr.on("data", (chunk) => {
      stderr += String(chunk);
    });

    child.on("error", (error) => {
      reject(error);
    });

    child.on("close", (code) => {
      if (code === 0) {
        resolve();
        return;
      }

      const output = [stdout.trim(), stderr.trim()].filter(Boolean).join("\n");
      reject(new Error(output || `Command failed: ${command} ${args.join(" ")}`));
    });
  });
}

export async function renderPptxPython(ir: DiagramIr, options: RenderOptions): Promise<void> {
  const tempRoot = await mkdtemp(path.join(tmpdir(), "mmd2pptx-"));
  const irPath = path.join(tempRoot, "diagram.ir.json");
  const patchPath = path.join(tempRoot, "diagram.patch.yml");

  try {
    await writeFile(irPath, `${JSON.stringify(ir, null, 2)}\n`, "utf8");

    const scriptPath = fileURLToPath(new URL("../../scripts/render_pptx.py", import.meta.url));
    const args = [
      "run",
      "--with",
      "python-pptx",
      "python",
      scriptPath,
      "--ir",
      irPath,
      "--output",
      path.resolve(options.outputPath),
    ];

    if (options.slideSize) {
      args.push("--slide-size", options.slideSize);
    }
    if (options.edgeRouting) {
      args.push("--edge-routing", options.edgeRouting);
    }
    if (options.appendToPath) {
      args.push("--append-to", path.resolve(options.appendToPath));
    }

    if (options.patchText) {
      await writeFile(patchPath, options.patchText, "utf8");
      args.push("--patch", patchPath);
    }

    await runCommand("uv", args);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`python renderer failed: ${message}`);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
}

export async function renderSequencePptxPython(sourceMmd: string, options: RenderOptions): Promise<void> {
  const tempRoot = await mkdtemp(path.join(tmpdir(), "mmd2pptx-seq-"));
  const sourcePath = path.join(tempRoot, "diagram.sequence.mmd");
  const patchPath = path.join(tempRoot, "diagram.patch.yml");

  try {
    await writeFile(sourcePath, sourceMmd, "utf8");

    const scriptPath = fileURLToPath(new URL("../../scripts/render_sequence_pptx.py", import.meta.url));
    const args = [
      "run",
      "--with",
      "python-pptx",
      "python",
      scriptPath,
      "--source",
      sourcePath,
      "--output",
      path.resolve(options.outputPath),
    ];

    if (options.slideSize) {
      args.push("--slide-size", options.slideSize);
    }
    if (options.edgeRouting) {
      args.push("--edge-routing", options.edgeRouting);
    }
    if (options.appendToPath) {
      args.push("--append-to", path.resolve(options.appendToPath));
    }

    if (options.patchText) {
      await writeFile(patchPath, options.patchText, "utf8");
      args.push("--patch", patchPath);
    }

    await runCommand("uv", args);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`python sequence renderer failed: ${message}`);
  } finally {
    await rm(tempRoot, { recursive: true, force: true });
  }
}
