# Maintainer Notes

This fork was reconstructed into a buildable source repository from the published source maps.

## Source Of Truth

- Edit files in `src/`.
- Treat `dist/` as generated output.
- Do not hand-edit bundled files in `dist/`.

## Layout

- `src/` - library source, ES modules.
- `src/index.js` - library entrypoint.
- `src/utils/font-loader.js` - Carlito font loader (see "Font substitution" below).
- `fonts/` - bundled Carlito woff2 files + SIL OFL license, served as-is to consumers via jsDelivr CDN.
- `types/index.d.ts` - source of truth for TypeScript definitions; copied to `dist/PptxViewJS.d.ts` by the build.
- `dist/` - generated; do not hand-edit.
- `harness/` - local manual test harness (Vite). Not published.
- `test/` - Jest tests.

## Rebuild Commands

- `npm run build` generates ESM, CJS, UMD, minified UMD, source maps, and `dist/PptxViewJS.d.ts`.
- `npm run build:min` generates only the minified UMD build plus `dist/PptxViewJS.d.ts`.
- `npm run test:smoke` runs a basic API-level smoke test against `dist/PptxViewJS.cjs.js`.
- `npm pack --dry-run` verifies the publishable package layout. The tarball must include `fonts/Carlito-*.woff2` and `fonts/LICENSE.txt`.
- `npm run harness:dev` starts the local viewer harness at `http://localhost:3000/harness/` for manual testing.

## Font Substitution

We ship Carlito (SIL OFL 1.1) as a metric-compatible Calibri substitute.

- Fonts live in `fonts/` and are bundled in the npm tarball (see `files` in `package.json`).
- `src/utils/font-loader.js` registers Carlito as `PptxViewJS-Calibri` via the `FontFace` API and defaults to a pinned jsDelivr CDN URL.
- `src/graphics/graphics-adapter.js` `setupStandardFont` prepends `PptxViewJS-Calibri` to the canvas font stack for Calibri runs only.
- Scoped family name prevents collision with host-app `font-family: Calibri` usages.

### Font release checklist

When bumping `package.json` version:

1. Update `PPTXVIEWJS_PKG_VERSION` in `src/utils/font-loader.js` to match.
2. Update `LIB_VERSION` in `src/index.js` to match.
3. Run `npm run build` and verify dist bundle contains the new version.
4. Run `npm pack --dry-run` and confirm `fonts/` is included in the tarball.
5. After publishing, confirm jsDelivr serves the new version at:
   `https://cdn.jsdelivr.net/npm/@petepetepete/pptxviewjs@<version>/fonts/Carlito-Regular-latin.woff2`
6. Update `CHANGELOG.md` with the release entry before publishing.

## Recovery Utility

- `npm run recover:src` rewrites `src/` from `dist/PptxViewJS.js.map` (`sourcesContent`).
- Use recovery only when needed (for example, if source files are accidentally removed).
- After recovery, run `npm run build` and re-run tests.
