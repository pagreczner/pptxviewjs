/**
 * PptxViewJS font loader.
 *
 * Registers Carlito (an SIL-OFL metric-compatible replacement for Calibri)
 * under the library-scoped font-family name "PptxViewJS-Calibri". Rendering
 * code prepends this scoped name ahead of "Calibri" in the canvas font stack,
 * so measurements and glyph widths align with PowerPoint/Google Slides while
 * the host application's other `font-family: Calibri` usages are unaffected.
 *
 * The loader is idempotent: only the first call performs network work; all
 * later calls return the same cached Promise.
 *
 * Default source: the woff2 files bundled alongside this package, served from
 * jsDelivr pinned to the installed package version. This gives us a reliable
 * zero-config default across bundlers (Vite, Webpack, Next.js, etc.) where
 * runtime `import.meta.url`-based resolution cannot reliably reach the
 * package's `fonts/` directory. Consumers can override with the `fontBaseUrl`
 * option to point at their own hosted copy (for offline / CSP-restricted /
 * airgapped deployments) or disable substitution entirely.
 */

// Package version used to pin the jsDelivr CDN URL. Kept in sync with
// package.json at release time. See MAINTAINERS.md for the release checklist.
const PPTXVIEWJS_PKG_VERSION = '1.3.0';

// Default base URL served from jsDelivr, pinned to the package version so a
// later release cannot change font assets under an older installed client.
// Consumers that cannot reach jsDelivr (strict CSP, offline, airgapped) should
// pass a local `fontBaseUrl` or set `substituteCalibri: false`.
const DEFAULT_JSDELIVR_BASE_URL =
  `https://cdn.jsdelivr.net/npm/@petepetepete/pptxviewjs@${PPTXVIEWJS_PKG_VERSION}/fonts/`;

// Google Fonts subset unicode-ranges for Carlito. These match the fragments
// published as separate woff2 files by Google and allow the browser to pick
// the correct subset based on the characters actually rendered.
const CARLITO_LATIN_RANGE =
  'U+0000-00FF, U+0131, U+0152-0153, U+02BB-02BC, U+02C6, U+02DA, U+02DC, ' +
  'U+0304, U+0308, U+0329, U+2000-206F, U+20AC, U+2122, U+2191, U+2193, ' +
  'U+2212, U+2215, U+FEFF, U+FFFD';

const CARLITO_LATIN_EXT_RANGE =
  'U+0100-02BA, U+02BD-02C5, U+02C7-02CC, U+02CE-02D7, U+02DD-02FF, ' +
  'U+0304, U+0308, U+0329, U+1D00-1DBF, U+1E00-1E9F, U+1EF2-1EFF, U+2020, ' +
  'U+20A0-20AB, U+20AD-20C0, U+2113, U+2C60-2C7F, U+A720-A7FF';

const FONT_DEFS = [
  { weight: '400', style: 'normal', file: 'Carlito-Regular-latin.woff2', unicodeRange: CARLITO_LATIN_RANGE },
  { weight: '400', style: 'normal', file: 'Carlito-Regular-latin-ext.woff2', unicodeRange: CARLITO_LATIN_EXT_RANGE },
  { weight: '700', style: 'normal', file: 'Carlito-Bold-latin.woff2', unicodeRange: CARLITO_LATIN_RANGE },
  { weight: '700', style: 'normal', file: 'Carlito-Bold-latin-ext.woff2', unicodeRange: CARLITO_LATIN_EXT_RANGE }
];

export const PPTX_CALIBRI_FAMILY = 'PptxViewJS-Calibri';

let loadPromise = null;

/**
 * Register Carlito with the document under the scoped family name so canvas
 * measurement and rendering both use Calibri-compatible metrics.
 *
 * Safe to call multiple times; subsequent calls resolve immediately.
 *
 * @param {{ fontBaseUrl?: string }} [options] Optional base URL pointing at a
 *   `fonts/` directory containing the Carlito woff2 files. Accepts absolute
 *   URLs (`https://my.cdn/fonts/`) or site-relative paths (`/fonts/`). When
 *   omitted, defaults to the jsDelivr CDN URL pinned to this package version.
 * @returns {Promise<void>} Resolves after font registration attempts complete.
 *   Never rejects: on failure the viewer falls back to the native font stack.
 */
export function ensurePptxViewFonts(options = {}) {
  if (loadPromise) {
    return loadPromise;
  }

  if (
    typeof document === 'undefined' ||
    typeof FontFace === 'undefined' ||
    !document.fonts ||
    typeof document.fonts.add !== 'function'
  ) {
    loadPromise = Promise.resolve();
    return loadPromise;
  }

  const providedBase = typeof options.fontBaseUrl === 'string' && options.fontBaseUrl.length > 0
    ? options.fontBaseUrl.replace(/\/$/, '') + '/'
    : null;
  const rawBaseUrl = providedBase || DEFAULT_JSDELIVR_BASE_URL;

  // `new URL(relative, base)` requires `base` to be an absolute URL. For
  // site-relative path overrides (e.g. "/fonts/"), resolve against
  // `document.baseURI` first. The jsDelivr default is already absolute.
  let baseUrl;
  try {
    baseUrl = new URL(rawBaseUrl, document.baseURI).href;
  } catch (_err) {
    loadPromise = Promise.resolve();
    return loadPromise;
  }

  loadPromise = Promise.all(
    FONT_DEFS.map(async def => {
      try {
        const url = new URL(def.file, baseUrl).href;
        const face = new FontFace(
          PPTX_CALIBRI_FAMILY,
          `url(${url}) format('woff2')`,
          {
            weight: def.weight,
            style: def.style,
            display: 'block',
            unicodeRange: def.unicodeRange
          }
        );
        const loaded = await face.load();
        document.fonts.add(loaded);
      } catch (_err) {
        // Silently ignore; viewer falls back to the native font stack.
      }
    })
  ).then(() => undefined);

  return loadPromise;
}

/**
 * For tests: reset the memoized load promise so the next call re-runs.
 * Not part of the public API.
 *
 * @private
 */
export function _resetPptxViewFontsForTests() {
  loadPromise = null;
}
