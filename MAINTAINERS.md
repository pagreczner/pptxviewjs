# Maintainer Notes

This fork was reconstructed into a buildable source repository from the published source maps.

## Source Of Truth

- Edit files in `src/`.
- Treat `dist/` as generated output.
- Do not hand-edit bundled files in `dist/`.

## Rebuild Commands

- `npm run build` generates ESM, CJS, UMD, minified UMD, source maps, and `dist/PptxViewJS.d.ts`.
- `npm run build:min` generates only the minified UMD build plus `dist/PptxViewJS.d.ts`.
- `npm run test:smoke` runs a basic API-level smoke test against `dist/PptxViewJS.cjs.js`.
- `npm pack --dry-run` verifies the publishable package layout.

## Recovery Utility

- `npm run recover:src` rewrites `src/` from `dist/PptxViewJS.js.map` (`sourcesContent`).
- Use recovery only when needed (for example, if source files are accidentally removed).
- After recovery, run `npm run build` and re-run tests.
