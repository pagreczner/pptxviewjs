# Changelog

All notable changes to `@petepetepete/pptxviewjs` are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.3.0] - 2026-04-24

### Added

- **Calibri metric substitution.** Bundles Carlito (SIL OFL 1.1), a metric-compatible replacement for Calibri. Registered under the library-scoped family `PptxViewJS-Calibri` so canvas measurement and rendering use Calibri-equivalent widths without affecting host-app UI that references `font-family: Calibri`. Loaded by default from a pinned jsDelivr CDN URL so it works out of the box in React, Next.js, Vite, Webpack, and UMD-CDN setups.
- New `PPTXViewerOptions.substituteCalibri` (default `true`) to opt out of Carlito substitution.
- New `PPTXViewerOptions.fontBaseUrl` to override the font base URL (site-relative paths or absolute URLs). Useful for offline / CSP-restricted / airgapped deployments.
- **`normAutofit` support.** Parses `<a:normAutofit fontScale="..." lnSpcReduction="..."/>` on text body properties and applies `fontScale` to run font sizes and `(1 - lnSpcReduction)` to paragraph line spacing. Matches PowerPoint's "Shrink text on overflow" behavior.
- Parses `<a:noAutofit/>` and `<a:spAutoFit/>` body properties (stored for future use; `spAutoFit` not yet applied).
- **`wrap="none"` honored in regular shapes.** Text bodies with `<a:bodyPr wrap="none">` no longer auto-wrap on canvas, matching PowerPoint.
- TypeScript definitions for new options (`substituteCalibri`, `fontBaseUrl`, `autoRenderFirstSlide`, `autoExposeGlobals`, `autoChartRerenderDelayMs`).

### Fixed

- **Bullet / indent overflow.** Wrap width for bulleted and indented paragraphs now accounts for `marL` (paragraph left margin). Previously text could render past the right edge of its text body by up to the indent amount; notably visible in narrow callout boxes.
- Line height calculation correctness (carry-over from 1.2.0): per-wrapped-line height driven by the tallest run on that line, using a natural 1.2 single-line spacing factor that matches PowerPoint and Google Slides.

### Changed

- Library now registers Carlito via `FontFace` at viewer construction time when `substituteCalibri` is enabled. First render awaits font readiness so initial paint uses Calibri-compatible metrics.

### Notes

- For airgapped or strict-CSP deployments, either copy `node_modules/@petepetepete/pptxviewjs/fonts/` into your public assets and pass `fontBaseUrl: "/fonts/"`, or set `substituteCalibri: false`.

## [1.2.0] - 2026-04-22

Internal recovery and tooling release. Reconstructed source tree from published source maps, added Rollup build pipeline and local manual harness. No API changes vs the upstream `pptxviewjs`.
