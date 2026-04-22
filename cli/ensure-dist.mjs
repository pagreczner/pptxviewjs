#!/usr/bin/env node

import fs from "node:fs";
import path from "node:path";

const distDir = path.resolve(process.cwd(), "dist");
fs.mkdirSync(distDir, { recursive: true });

console.log(`Ensured directory exists: ${distDir}`);
