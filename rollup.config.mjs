import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

import commonjs from "@rollup/plugin-commonjs";
import json from "@rollup/plugin-json";
import nodeResolve from "@rollup/plugin-node-resolve";
import terser from "@rollup/plugin-terser";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const pkg = JSON.parse(
  fs.readFileSync(path.join(__dirname, "package.json"), "utf8"),
);

const banner = `/**
 * ${pkg.name} v${pkg.version}
 * JavaScript library for viewing PowerPoint presentations in web browsers
 * 
 * Copyright ${new Date().getFullYear()} gptsci.com
 * Licensed under the MIT License
 */`;

const external = ["chart.js/auto", "jszip"];
const globals = {
  "chart.js/auto": "Chart",
  jszip: "JSZip",
};

const basePlugins = [
  nodeResolve({
    browser: true,
    preferBuiltins: false,
  }),
  commonjs(),
  json(),
];

const sharedConfig = {
  input: "src/index.js",
  external,
  treeshake: false,
};

const baseBuild = {
  ...sharedConfig,
  plugins: basePlugins,
  output: [
    {
      file: "dist/PptxViewJS.es.js",
      format: "es",
      banner,
      sourcemap: true,
      globals,
    },
    {
      file: "dist/PptxViewJS.cjs.js",
      format: "cjs",
      banner,
      sourcemap: true,
      exports: "named",
      globals,
    },
    {
      file: "dist/PptxViewJS.js",
      format: "umd",
      name: "PptxViewJS",
      banner,
      sourcemap: true,
      globals,
    },
  ],
};

const minifiedBuild = {
  ...sharedConfig,
  plugins: [
    ...basePlugins,
    terser({
      format: {
        comments: false,
      },
    }),
  ],
  output: {
    file: "dist/PptxViewJS.min.js",
    format: "umd",
    name: "PptxViewJS",
    banner,
    sourcemap: false,
    globals,
  },
};

const minifiedOnly = (process.env.BUILD || "").toLowerCase() === "minified";

export default minifiedOnly ? [minifiedBuild] : [baseBuild, minifiedBuild];
