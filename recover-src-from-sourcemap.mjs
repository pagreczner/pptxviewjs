#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";
import process from "node:process";

const repoRoot = process.cwd();
const mapPath = path.resolve(repoRoot, "dist/PptxViewJS.js.map");

if (!fs.existsSync(mapPath)) {
  console.error(`Source map not found: ${mapPath}`);
  process.exit(1);
}

const map = JSON.parse(fs.readFileSync(mapPath, "utf8"));

if (!Array.isArray(map.sources) || !Array.isArray(map.sourcesContent)) {
  console.error("Invalid source map: missing sources or sourcesContent arrays.");
  process.exit(1);
}

if (map.sources.length !== map.sourcesContent.length) {
  console.error(
    `Invalid source map: sources (${map.sources.length}) and sourcesContent (${map.sourcesContent.length}) length mismatch.`,
  );
  process.exit(1);
}

const distDir = path.resolve(repoRoot, "dist");
let writtenCount = 0;

for (let i = 0; i < map.sources.length; i += 1) {
  const sourceRef = map.sources[i];
  const sourceContent = map.sourcesContent[i];

  if (typeof sourceRef !== "string" || !sourceRef.trim()) {
    console.error(`Invalid source entry at index ${i}.`);
    process.exit(1);
  }

  if (typeof sourceContent !== "string") {
    console.error(`Missing sourcesContent for ${sourceRef} at index ${i}.`);
    process.exit(1);
  }

  const absoluteSourcePath = path.resolve(distDir, sourceRef);
  const relativeSourcePath = path.relative(repoRoot, absoluteSourcePath);

  if (relativeSourcePath.startsWith("..")) {
    console.error(`Resolved source path escapes repository root: ${sourceRef}`);
    process.exit(1);
  }

  fs.mkdirSync(path.dirname(absoluteSourcePath), { recursive: true });
  fs.writeFileSync(absoluteSourcePath, sourceContent, "utf8");
  writtenCount += 1;
}

console.log(`Recovered ${writtenCount} source files from ${path.relative(repoRoot, mapPath)}.`);
