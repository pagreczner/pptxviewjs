/**
 * @jest-environment jsdom
 */

const createMock2DContext = () => ({
  clearRect: () => {},
  drawImage: () => {},
  fillRect: () => {},
  fillText: () => {},
  getImageData: () => ({ data: new Uint8ClampedArray(4) }),
  measureText: () => ({ width: 0 }),
  putImageData: () => {},
  restore: () => {},
  save: () => {},
  setTransform: () => {},
  strokeRect: () => {},
});

describe("PptxViewJS distribution smoke test", () => {
  let moduleExports;

  beforeAll(() => {
    if (globalThis.HTMLCanvasElement) {
      Object.defineProperty(globalThis.HTMLCanvasElement.prototype, "getContext", {
        configurable: true,
        writable: true,
        value: () => createMock2DContext(),
      });
    }

    moduleExports = require("../dist/PptxViewJS.cjs.js");
  });

  it("exposes expected top-level exports", () => {
    expect(typeof moduleExports.PPTXViewer).toBe("function");
    expect(moduleExports.default).toBeDefined();
    expect(moduleExports.default.PPTXViewer).toBe(moduleExports.PPTXViewer);
    expect(typeof moduleExports.default.version).toBe("string");
    expect(typeof moduleExports.default.mountSimpleViewer).toBe("function");
  });

  it("constructs a viewer with core methods", () => {
    const viewer = new moduleExports.PPTXViewer();

    expect(viewer).toBeInstanceOf(moduleExports.PPTXViewer);
    expect(typeof viewer.loadFile).toBe("function");
    expect(typeof viewer.loadFromUrl).toBe("function");
    expect(typeof viewer.render).toBe("function");
    expect(typeof viewer.nextSlide).toBe("function");
    expect(typeof viewer.previousSlide).toBe("function");
    expect(typeof viewer.goToSlide).toBe("function");
    expect(typeof viewer.on).toBe("function");
    expect(typeof viewer.off).toBe("function");
    expect(typeof viewer.destroy).toBe("function");
  });
});
