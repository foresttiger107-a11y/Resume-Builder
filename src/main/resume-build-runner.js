/** @type {Map<string, { resolve: (value: unknown) => void, reject: (reason?: unknown) => void, timeout: NodeJS.Timeout }>} */
const pendingBuilds = new Map();

/** @type {Set<number>} */
const readyRendererIds = new Set();

const DEFAULT_BUILD_TIMEOUT_MS = 6 * 60 * 1000;

/** @param {number} webContentsId */
function markRendererReady(webContentsId) {
  if (Number.isFinite(webContentsId)) {
    readyRendererIds.add(webContentsId);
  }
}

/** @param {import("electron").BrowserWindow} win @param {number} [timeoutMs] */
async function waitForRendererReady(win, timeoutMs = 120000) {
  if (!win || win.isDestroyed()) {
    throw new Error("Profile window is not available.");
  }
  const wcId = win.webContents.id;
  if (readyRendererIds.has(wcId)) {
    return;
  }
  if (win.webContents.isLoading()) {
    await new Promise((resolve) => {
      win.webContents.once("did-finish-load", resolve);
    });
  }
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (win.isDestroyed()) {
      throw new Error("Profile window closed before build started.");
    }
    if (readyRendererIds.has(wcId)) {
      return;
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  throw new Error("Profile window did not become ready in time.");
}

/**
 * @param {string} requestId
 * @param {number} [timeoutMs]
 */
function registerResumeBuild(requestId, timeoutMs = DEFAULT_BUILD_TIMEOUT_MS) {
  if (pendingBuilds.has(requestId)) {
    throw new Error("Duplicate resume build request.");
  }
  return new Promise((resolve, reject) => {
    const timeout = setTimeout(() => {
      pendingBuilds.delete(requestId);
      reject(new Error("Resume build timed out."));
    }, timeoutMs);
    pendingBuilds.set(requestId, { resolve, reject, timeout });
  });
}

/** @param {string} requestId @param {Record<string, unknown>} payload */
function completeResumeBuild(requestId, payload) {
  const pending = pendingBuilds.get(requestId);
  if (!pending) {
    return false;
  }
  clearTimeout(pending.timeout);
  pendingBuilds.delete(requestId);
  if (payload?.ok) {
    pending.resolve(payload);
  } else {
    pending.reject(
      new Error(String(payload?.error || "Resume build failed.")),
    );
  }
  return true;
}

function makeResumeBuildRequestId() {
  return `build-${Date.now()}-${process.pid}-${Math.random().toString(36).slice(2, 8)}`;
}

module.exports = {
  DEFAULT_BUILD_TIMEOUT_MS,
  markRendererReady,
  waitForRendererReady,
  registerResumeBuild,
  completeResumeBuild,
  makeResumeBuildRequestId,
};
