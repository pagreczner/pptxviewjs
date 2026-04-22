#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";

const repoRoot = process.cwd();
const sourcePath = path.resolve(repoRoot, "types/index.d.ts");
const destinationPath = path.resolve(repoRoot, "dist/PptxViewJS.d.ts");

if (!fs.existsSync(sourcePath)) {
  console.error(`Type definitions not found: ${sourcePath}`);
  process.exit(1);
}

fs.mkdirSync(path.dirname(destinationPath), { recursive: true });
fs.copyFileSync(sourcePath, destinationPath);

console.log(
  `Copied type definitions: ${path.relative(repoRoot, sourcePath)} -> ${path.relative(repoRoot, destinationPath)}`,
);
