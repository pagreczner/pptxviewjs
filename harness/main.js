const RUNTIMES = {
  source: {
    label: "Source (`src/index.js`)",
    importPath: "/src/index.js",
  },
  dist: {
    label: "Dist (`dist/PptxViewJS.es.js`)",
    importPath: "/dist/PptxViewJS.es.js",
  },
};

const state = {
  mode: "source",
  viewer: null,
  loadedFile: null,
  busy: false,
};

const refs = {
  modeSelect: document.getElementById("modeSelect"),
  fileInput: document.getElementById("fileInput"),
  prevBtn: document.getElementById("prevBtn"),
  nextBtn: document.getElementById("nextBtn"),
  fullscreenBtn: document.getElementById("fullscreenBtn"),
  slideStatus: document.getElementById("slideStatus"),
  runtimeStatus: document.getElementById("runtimeStatus"),
  viewerShell: document.getElementById("viewerShell"),
  canvas: document.getElementById("viewerCanvas"),
  message: document.getElementById("message"),
  overlayPrevBtn: document.getElementById("overlayPrevBtn"),
  overlayNextBtn: document.getElementById("overlayNextBtn"),
  overlayExitBtn: document.getElementById("overlayExitBtn"),
  overlaySlideStatus: document.getElementById("overlaySlideStatus"),
};

function runtimeLabel(mode = state.mode) {
  return RUNTIMES[mode]?.label ?? mode;
}

function setMessage(text, tone = "") {
  refs.message.textContent = text;
  refs.message.classList.remove("error", "success");
  if (tone) {
    refs.message.classList.add(tone);
  }
}

function updateFullscreenButton() {
  refs.fullscreenBtn.textContent = document.fullscreenElement
    ? "Exit Fullscreen"
    : "Fullscreen";
}

function clearCanvas() {
  const context = refs.canvas.getContext("2d");
  if (context) {
    context.clearRect(0, 0, refs.canvas.width, refs.canvas.height);
  }
}

function updateStatus() {
  refs.runtimeStatus.textContent = runtimeLabel();

  if (!state.viewer) {
    refs.slideStatus.textContent = "No presentation loaded";
    refs.overlaySlideStatus.textContent = "Slide 0 / 0";
    refs.prevBtn.disabled = true;
    refs.nextBtn.disabled = true;
    refs.fullscreenBtn.disabled = true;
    refs.overlayPrevBtn.disabled = true;
    refs.overlayNextBtn.disabled = true;
    return;
  }

  const total = state.viewer.getSlideCount();
  const index = state.viewer.getCurrentSlideIndex();
  const hasDeck = total > 0;

  const statusText = hasDeck
    ? `Slide ${index + 1} / ${total}`
    : "No presentation loaded";
  refs.slideStatus.textContent = statusText;
  refs.overlaySlideStatus.textContent = hasDeck
    ? `Slide ${index + 1} / ${total}`
    : "Slide 0 / 0";

  const prevDisabled = state.busy || !hasDeck || index <= 0;
  const nextDisabled = state.busy || !hasDeck || index >= total - 1;
  refs.prevBtn.disabled = prevDisabled;
  refs.nextBtn.disabled = nextDisabled;
  refs.overlayPrevBtn.disabled = prevDisabled;
  refs.overlayNextBtn.disabled = nextDisabled;
  refs.fullscreenBtn.disabled = state.busy || !hasDeck;
}

function setBusy(value) {
  state.busy = value;
  refs.modeSelect.disabled = value;
  refs.fileInput.disabled = value;
  updateStatus();
}

function destroyViewer() {
  if (state.viewer && typeof state.viewer.destroy === "function") {
    try {
      state.viewer.destroy();
    } catch (_error) {
      // Ignore teardown errors in manual harness mode.
    }
  }
  state.viewer = null;
  clearCanvas();
  updateStatus();
}

async function loadViewerClass(mode) {
  const runtime = RUNTIMES[mode];
  if (!runtime) {
    throw new Error(`Unknown runtime mode: ${mode}`);
  }

  const imported = await import(/* @vite-ignore */ runtime.importPath);
  const ViewerClass = imported.PPTXViewer ?? imported.default?.PPTXViewer;

  if (typeof ViewerClass !== "function") {
    throw new Error(`PPTXViewer export missing from ${runtime.importPath}`);
  }

  return ViewerClass;
}

async function initializeViewer({ reloadCurrentFile = false } = {}) {
  setBusy(true);
  try {
    destroyViewer();
    const ViewerClass = await loadViewerClass(state.mode);
    // In dev the source module lives at /src/utils/font-loader.js so the
    // default `../fonts/` resolution lands in /src/fonts/. Override the base
    // URL to the project-root /fonts/ directory so both Source and Dist
    // runtimes find the bundled Carlito woff2 files during manual testing.
    state.viewer = new ViewerClass({
      canvas: refs.canvas,
      fontBaseUrl: "/fonts/",
    });
    updateStatus();

    // Deterministic first paint: wait for any registered fonts to load before
    // asking the viewer to render. The viewer also awaits this internally, but
    // awaiting here keeps test timing predictable.
    if (typeof document !== "undefined" && document.fonts && document.fonts.ready) {
      try {
        await document.fonts.ready;
      } catch (_err) {
        // No-op; viewer will still fall back to the native stack if needed.
      }
    }

    if (reloadCurrentFile && state.loadedFile) {
      await state.viewer.loadFile(state.loadedFile);
      await state.viewer.render(refs.canvas, { slideIndex: 0 });
      setMessage(
        `Switched to ${runtimeLabel()} and reloaded "${state.loadedFile.name}".`,
        "success",
      );
    } else {
      setMessage(
        `Ready in ${runtimeLabel()} mode. Select a local .pptx file to render.`,
      );
    }
  } catch (error) {
    destroyViewer();
    setMessage(
      `Failed to initialize ${runtimeLabel()}: ${error?.message ?? String(error)}`,
      "error",
    );
  } finally {
    setBusy(false);
  }
}

async function loadSelectedFile(file) {
  if (!file) {
    return;
  }

  if (!state.viewer) {
    await initializeViewer();
    if (!state.viewer) {
      return;
    }
  }

  setBusy(true);
  try {
    state.loadedFile = file;
    await state.viewer.loadFile(file);
    await state.viewer.render(refs.canvas, { slideIndex: 0 });
    const total = state.viewer.getSlideCount();
    setMessage(`Loaded "${file.name}" (${total} slides).`, "success");
  } catch (error) {
    setMessage(`Load/render failed: ${error?.message ?? String(error)}`, "error");
  } finally {
    setBusy(false);
  }
}

async function navigate(direction) {
  if (!state.viewer) {
    return;
  }

  setBusy(true);
  try {
    if (direction === "next") {
      await state.viewer.nextSlide(refs.canvas);
    } else {
      await state.viewer.previousSlide(refs.canvas);
    }
  } catch (error) {
    setMessage(`Navigation failed: ${error?.message ?? String(error)}`, "error");
  } finally {
    setBusy(false);
  }
}

async function toggleFullscreen() {
  try {
    if (document.fullscreenElement) {
      await document.exitFullscreen();
    } else {
      await refs.viewerShell.requestFullscreen();
    }
  } catch (error) {
    setMessage(`Fullscreen failed: ${error?.message ?? String(error)}`, "error");
  } finally {
    updateFullscreenButton();
  }
}

function bindEvents() {
  refs.modeSelect.addEventListener("change", async (event) => {
    const nextMode = event.target.value;
    if (nextMode === state.mode) {
      return;
    }
    state.mode = nextMode;
    await initializeViewer({ reloadCurrentFile: true });
  });

  refs.fileInput.addEventListener("change", async (event) => {
    const [file] = event.target.files || [];
    await loadSelectedFile(file);
  });

  refs.prevBtn.addEventListener("click", async () => navigate("previous"));
  refs.nextBtn.addEventListener("click", async () => navigate("next"));
  refs.overlayPrevBtn.addEventListener("click", async () => navigate("previous"));
  refs.overlayNextBtn.addEventListener("click", async () => navigate("next"));
  refs.fullscreenBtn.addEventListener("click", toggleFullscreen);
  refs.overlayExitBtn.addEventListener("click", async () => {
    if (document.fullscreenElement) {
      try {
        await document.exitFullscreen();
      } catch (_error) {
        // Ignore; browser will surface the error in console if needed.
      }
    }
  });

  refs.canvas.addEventListener("click", async () => navigate("next"));

  document.addEventListener("fullscreenchange", updateFullscreenButton);

  document.addEventListener("keydown", (event) => {
    const target = event.target;
    const tag = target && target.tagName ? target.tagName.toLowerCase() : "";
    const isFormControl =
      tag === "input" ||
      tag === "select" ||
      tag === "textarea" ||
      (target && target.isContentEditable);

    if (isFormControl) {
      return;
    }

    if (event.key === "ArrowRight") {
      event.preventDefault();
      navigate("next");
    } else if (event.key === "ArrowLeft") {
      event.preventDefault();
      navigate("previous");
    }
  });
}

async function startHarness() {
  refs.modeSelect.value = state.mode;
  refs.runtimeStatus.textContent = runtimeLabel();
  bindEvents();
  updateFullscreenButton();
  updateStatus();
  await initializeViewer();
}

startHarness();
