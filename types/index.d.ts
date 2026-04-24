/**
 * PptxViewJS TypeScript Definitions
 * Matches the runtime exports from src/index.js
 */

export interface PPTXViewerOptions {
  canvas?: HTMLCanvasElement | null;
  debug?: boolean;
  enableThumbnails?: boolean;
  slideSizeMode?: 'fit' | 'actual' | 'custom';
  backgroundColor?: string;
  logger?: Console;
  /**
   * Auto-render slide 0 after `loadFile()` completes. Defaults to `true`.
   */
  autoRenderFirstSlide?: boolean;
  /**
   * Auto-expose PPTX processor globals to simplify chart integration. Defaults to `true`.
   */
  autoExposeGlobals?: boolean;
  /**
   * Schedule one delayed re-render after load to catch late-parsing charts
   * (milliseconds). Set to `0` to disable. Defaults to `200`.
   */
  autoChartRerenderDelayMs?: number;
  /**
   * Bundle and apply Carlito (metric-compatible Calibri substitute) so Calibri
   * runs measure and render with PowerPoint-equivalent widths. Defaults to
   * `true`. Set to `false` to disable (useful for strict CSP / offline deploys
   * or when you manage Calibri yourself).
   */
  substituteCalibri?: boolean;
  /**
   * Override the base URL used to fetch the bundled Carlito woff2 files.
   * Accepts absolute URLs (`"https://my.cdn/fonts/"`) or site-relative paths
   * (`"/fonts/"`). When omitted, defaults to a pinned jsDelivr CDN URL that
   * tracks the installed package version. Useful for offline / airgapped /
   * CSP-restricted deployments.
   */
  fontBaseUrl?: string | null;
  [key: string]: unknown;
}

export interface RenderOptions {
  slideIndex?: number;
  scale?: number;
  quality?: 'low' | 'medium' | 'high';
  [key: string]: unknown;
}

export type PPTXViewerEventCallback = (...args: unknown[]) => void;

export class PPTXViewer {
  constructor(options?: PPTXViewerOptions);

  /** Load PPTX content supplied as File/ArrayBuffer/Uint8Array */
  loadFile(input: File | ArrayBuffer | Uint8Array, options?: Record<string, unknown>): Promise<PPTXViewer>;

  /** Fetch a PPTX file from the provided URL and load it */
  loadFromUrl(url: string): Promise<PPTXViewer>;

  /** Render the current or specified slide to a canvas element */
  render(canvas?: HTMLCanvasElement | null, options?: RenderOptions): Promise<PPTXViewer>;

  /** Convenience alias for rendering a specific slide */
  renderSlide(slideIndex: number, canvas?: HTMLCanvasElement | null, options?: RenderOptions): Promise<PPTXViewer>;

  /** Advance to the next slide if possible */
  nextSlide(canvas?: HTMLCanvasElement | null): Promise<PPTXViewer>;

  /** Move to the previous slide if possible */
  previousSlide(canvas?: HTMLCanvasElement | null): Promise<PPTXViewer>;

  /** Jump to a specific slide index */
  goToSlide(slideIndex: number, canvas?: HTMLCanvasElement | null): Promise<PPTXViewer>;

  /** Get the total number of slides in the loaded presentation */
  getSlideCount(): number;

  /** Get the current slide index */
  getCurrentSlideIndex(): number;

  /** Replace the canvas element used for rendering */
  setCanvas(canvas: HTMLCanvasElement | null): PPTXViewer;

  /** Register an event listener */
  on(event: string, callback: PPTXViewerEventCallback): void;

  /** Remove an event listener */
  off(event: string, callback: PPTXViewerEventCallback): void;

  /** Release references and reset internal state */
  destroy(): void;
}

export declare const version: string;

export interface PptxViewJSNamespace {
  PPTXViewer: typeof PPTXViewer;
  version: typeof version;
  [key: string]: unknown;
}

declare const namespace: PptxViewJSNamespace;

export default namespace;

declare global {
  interface Window {
    PptxViewJS: PptxViewJSNamespace;
  }
}

export as namespace PptxViewJS;
