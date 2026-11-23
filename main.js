const {
  app,
  BrowserWindow,
  ipcMain,
  dialog,
  Tray,
  Menu,
  nativeImage,
  session,
  shell,
  globalShortcut,
  clipboard,
} = require("electron");
const path = require("path");
const { version: APP_VERSION } = require("./package.json");
const { spawn } = require("child_process");
const fs = require("fs/promises");
const fsSync = require("fs");
const { fileURLToPath, pathToFileURL } = require("node:url");
const { buildJustinTemplateResumeDocxBuffer } = require("./resume-justin-docx.js");
const {
  buildResumeFromJustinStructureDocx,
  buildResumeFromResumeStructureJson,
  formatDocxtemplaterError,
} = require("./resume-justin-template-docx.js");
const {
  extractResumeJsonObjectFromGptText,
  isCompleteResumeStructureJsonForExport,
  buildWorkExperienceWordPreviewRows,
  parseResumeStructureJson,
  resumeJsonObjectToJustinFields,
  suggestJustinMcgheeDocxFileStem,
  suggestResumeDocxFileStemForProfile,
} = require("./resume-json-to-justin.js");

const appRoot = __dirname;

/** Electron's normal user agent (a spoofed Chrome UA mismatches the real Chromium). */
const CHROME_USER_AGENT = app.userAgentFallback;

const CHATGPT_PARTITIONS = [
  "persist:chatgpt",
  "persist:chatgpt-oluwaseun",
  "persist:chatgpt-justin",
];

/** Permissions that trigger Windows passkey / security-key UI in webviews. */
const CHATGPT_DENIED_PERMISSIONS = new Set([
  "hid",
  "serial",
  "usb",
  "bluetooth",
]);

/** @param {string} url */
const ALLOW_FEDERATED_SIGNIN = true;

function isBlockedFederatedAuthUrl(url) {
  // Google / Microsoft / Apple sign-in is allowed in the app (no blocking).
  if (ALLOW_FEDERATED_SIGNIN) {
    return false;
  }
  if (!url || url.startsWith("about:")) {
    return false;
  }
  try {
    const u = new URL(url);
    const host = u.hostname.toLowerCase();
    const pathQuery = `${u.pathname}${u.search}`.toLowerCase();

    if (host === "accounts.google.com") {
      return true;
    }
    if (host === "google.com" || host.endsWith(".google.com")) {
      if (/signin|oauth|servicelogin|openid|gsi|federated/i.test(pathQuery)) {
        return true;
      }
    }
    if (host === "appleid.apple.com") {
      return true;
    }
    if (host.endsWith(".apple.com") && /auth|signin|oauth/i.test(pathQuery)) {
      return true;
    }
    if (
      host === "login.microsoftonline.com" ||
      host === "login.live.com" ||
      host.endsWith(".microsoftonline.com")
    ) {
      return true;
    }
  } catch (_e) {
    return /accounts\.google\.com|login\.microsoftonline\.com|appleid\.apple\.com/i.test(
      url,
    );
  }
  return false;
}

const configuredChatGptPartitions = new Set();

/**
 * @param {string} partition
 */
function configureChatGptSession(partition) {
  if (configuredChatGptPartitions.has(partition)) {
    return;
  }
  configuredChatGptPartitions.add(partition);

  const ses = session.fromPartition(partition, { cache: true });
  ses.setUserAgent(CHROME_USER_AGENT);

  ses.setPermissionCheckHandler((_webContents, permission) => {
    if (CHATGPT_DENIED_PERMISSIONS.has(permission)) {
      return false;
    }
    return true;
  });
  ses.setPermissionRequestHandler((_webContents, permission, callback) => {
    callback(!CHATGPT_DENIED_PERMISSIONS.has(permission));
  });
}

const CHATGPT_EMAIL_LOGIN_URL = "https://auth.openai.com/log-in";
/** @type {Map<string, import("electron").BrowserWindow>} */
const chatGptLoginWindows = new Map();

/**
 * @param {import("electron").WebContents} contents
 */
function attachChatGptFederatedAuthBlock(contents) {
  if (contents._pbChatgptAuthBlock) {
    return;
  }
  contents._pbChatgptAuthBlock = true;
  const redirectToEmailLogin = (event, url) => {
    if (!isBlockedFederatedAuthUrl(url)) {
      return;
    }
    event.preventDefault();
    contents.loadURL(CHATGPT_EMAIL_LOGIN_URL);
  };
  contents.on("will-navigate", redirectToEmailLogin);
  contents.on("will-redirect", redirectToEmailLogin);
  // Safety net for navigations the events above miss (JS location changes, etc.).
  contents.on("did-start-navigation", (details) => {
    if (details.isMainFrame && isBlockedFederatedAuthUrl(details.url)) {
      contents.loadURL(CHATGPT_EMAIL_LOGIN_URL);
    }
  });
}

/**
 * @param {string} partition
 * @returns {import("electron").BrowserWindow}
 */
function openChatGptLoginWindow(partition) {
  const part = partition || "persist:chatgpt";
  const existing = chatGptLoginWindows.get(part);
  if (existing && !existing.isDestroyed()) {
    existing.focus();
    return existing;
  }

  const win = new BrowserWindow({
    width: 480,
    height: 760,
    show: true,
    autoHideMenuBar: true,
    title: "Sign in to ChatGPT",
    parent: mainWindow && !mainWindow.isDestroyed() ? mainWindow : undefined,
    modal: !!(mainWindow && !mainWindow.isDestroyed()),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
      partition: part,
    },
  });

  chatGptLoginWindows.set(part, win);
  configureChatGptSession(part);
  win.webContents.setUserAgent(CHROME_USER_AGENT);
  attachChatGptFederatedAuthBlock(win.webContents);
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (isBlockedFederatedAuthUrl(url)) {
      return { action: "deny" };
    }
    return {
      action: "allow",
      overrideBrowserWindowOptions: chatGptPopupWindowOptions(part),
    };
  });

  win.webContents.on("did-navigate", (_event, url) => {
    if (/^https:\/\/(chatgpt\.com|chat\.openai\.com)(\/|$)/i.test(url)) {
      if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send("chatgpt-login-complete", { partition: part });
      }
    }
  });

  win.on("closed", () => {
    chatGptLoginWindows.delete(part);
    if (mainWindow && !mainWindow.isDestroyed()) {
      mainWindow.webContents.send("chatgpt-login-complete", { partition: part });
    }
  });

  void win.loadURL(CHATGPT_EMAIL_LOGIN_URL);
  return win;
}

function chatGptPopupWindowOptions(partition) {
  const part = partition || "persist:chatgpt";
  return {
    width: 520,
    height: 780,
    show: true,
    autoHideMenuBar: true,
    parent: mainWindow && !mainWindow.isDestroyed() ? mainWindow : undefined,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      sandbox: false,
      partition: part,
    },
  };
}

/** @param {import("electron").WebContents} contents */
function enableChatGptPopups(contents) {
  const part = contents.session?.partition || "persist:chatgpt";
  contents.setUserAgent(CHROME_USER_AGENT);
  contents.setWindowOpenHandler(({ url }) => {
    if (isBlockedFederatedAuthUrl(url)) {
      return { action: "deny" };
    }
    return {
      action: "allow",
      overrideBrowserWindowOptions: chatGptPopupWindowOptions(part),
    };
  });
  attachChatGptFederatedAuthBlock(contents);
}

function setupChatGptWebviewSecurity() {
  app.on("web-contents-created", (_event, contents) => {
    const type = contents.getType();
    const part = contents.session?.partition || "";
    if (!part.includes("chatgpt")) {
      return;
    }
    if (type === "webview") {
      configureChatGptSession(part);
    }
    if (type === "webview" || type === "window") {
      enableChatGptPopups(contents);
    }
  });
}

/** @param {"olu" | "justin"} profile */
function chatGptPartitionForProfile(profile) {
  if (profile === "justin") {
    return "persist:chatgpt-justin";
  }
  return "persist:chatgpt";
}

app.commandLine.appendSwitch("disable-blink-features", "AutomationControlled");
// ChatGPT over QUIC/HTTP2 often fails in embedded <webview> on Windows.
app.commandLine.appendSwitch("disable-quic");
app.commandLine.appendSwitch("disable-http2");
setupChatGptWebviewSecurity();
/** Bundled defaults (inside app.asar when packaged — read-only). */
const bundledPromptDir = path.join(appRoot, "prompt");

/** Ordered dirs to read factory prompts from (asar, then asar.unpacked if used). */
function getBundledPromptDirs() {
  const dirs = [bundledPromptDir];
  if (app.isPackaged) {
    dirs.push(path.join(process.resourcesPath, "app.asar.unpacked", "prompt"));
  }
  return dirs;
}

/**
 * Writable prompt directory. Packaged apps cannot write inside app.asar;
 * use userData so Save works in installed builds.
 * @returns {string}
 */
function getWritablePromptDir() {
  if (app.isPackaged) {
    return path.join(app.getPath("userData"), "prompt");
  }
  return bundledPromptDir;
}

/** @returns {string} */
function appSettingsPath() {
  return path.join(app.getPath("userData"), "settings.json");
}

/** @param {unknown} profile */
function normalizeResumeProfile(profile) {
  return profile === "justin" ? "justin" : "olu";
}

/** @param {unknown} arg IPC payload or profile string */
function ipcResumeProfile(arg) {
  if (arg && typeof arg === "object" && !Array.isArray(arg)) {
    const rec = /** @type {Record<string, unknown>} */ (arg);
    return normalizeResumeProfile(
      rec.profile ?? rec.account ?? rec.who ?? rec.name,
    );
  }
  return normalizeResumeProfile(arg);
}

/** @param {unknown} arg IPC payload, prompt id, or legacy profile */
function ipcDownloadKey(arg) {
  if (arg && typeof arg === "object" && !Array.isArray(arg)) {
    const rec = /** @type {Record<string, unknown>} */ (arg);
    if (typeof rec.downloadKey === "string" && rec.downloadKey.trim()) {
      return normalizeDownloadKey(rec.downloadKey);
    }
    if (typeof rec.promptId === "string" && rec.promptId.trim()) {
      return normalizeDownloadKey(rec.promptId);
    }
    if (typeof rec.promptKind === "string" && rec.promptKind.trim()) {
      return normalizeDownloadKey(rec.promptKind);
    }
    return normalizeDownloadKey(rec.profile);
  }
  return normalizeDownloadKey(arg);
}

/** @param {unknown} key */
function normalizeDownloadKey(key) {
  const k = String(key ?? "").trim();
  if (k === "justin") {
    return "justin";
  }
  if (k === "olu" || k === "other") {
    return "other";
  }
  if (k.startsWith("custom:")) {
    return k;
  }
  return k || "other";
}

/**
 * @param {unknown} raw
 * @returns {Record<string, string>}
 */
function readDownloadDirectoriesMap(raw) {
  /** @type {Record<string, string>} */
  const out = {};
  if (raw && typeof raw === "object" && !Array.isArray(raw)) {
    const rec = /** @type {Record<string, unknown>} */ (raw);
    for (const [key, v] of Object.entries(rec)) {
      if (typeof v !== "string" || !v.trim()) {
        continue;
      }
      out[normalizeDownloadKey(key)] = path.normalize(v.trim());
    }
  }
  if (out.olu && !out.other) {
    out.other = out.olu;
  }
  return out;
}

/**
 * @param {Record<string, unknown>} settings
 * @returns {Record<string, string>}
 */
function getDownloadDirectoriesFromSettings(settings) {
  const map = readDownloadDirectoriesMap(settings.downloadDirectories);
  const legacy = String(settings.downloadDirectory || "").trim();
  if (legacy && !map.other && !map.olu) {
    map.other = path.normalize(legacy);
  }
  return map;
}

/**
 * @param {{ olu?: string, justin?: string }} partial
 */
/**
 * @typedef {{
 *   id: string,
 *   label: string,
 *   promptPath: string,
 *   structurePath: string,
 *   chatgptProfile: "olu" | "justin",
 * }} CustomPromptEntry
 */

/** @param {unknown} raw */
function normalizeCustomPromptEntry(raw) {
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    return null;
  }
  const rec = /** @type {Record<string, unknown>} */ (raw);
  const id = String(rec.id || "").trim();
  const label = String(rec.label || "").trim();
  const promptPath = path.normalize(String(rec.promptPath || "").trim());
  const structurePath = path.normalize(String(rec.structurePath || "").trim());
  const chatgptProfile = rec.chatgptProfile === "justin" ? "justin" : "olu";
  if (!id.startsWith("custom:") || !label || !promptPath || !structurePath) {
    return null;
  }
  return /** @type {CustomPromptEntry} */ ({
    id,
    label,
    promptPath,
    structurePath,
    chatgptProfile,
  });
}

/** @param {CustomPromptEntry[]} entries */
function dedupeCustomPromptEntries(entries) {
  /** @type {Map<string, CustomPromptEntry>} */
  const byId = new Map();
  for (const e of entries) {
    if (!e?.id || byId.has(e.id)) {
      continue;
    }
    byId.set(e.id, e);
  }
  return [...byId.values()];
}

/** @returns {Promise<CustomPromptEntry[]>} */
async function readCustomPrompts() {
  const settings = await readAppSettings();
  const raw = settings.customPrompts;
  if (!Array.isArray(raw)) {
    return [];
  }
  return dedupeCustomPromptEntries(
    raw.map(normalizeCustomPromptEntry).filter(Boolean),
  );
}

/** @param {CustomPromptEntry[]} entries */
async function writeCustomPrompts(entries) {
  const cur = await readAppSettings();
  const next = {
    ...cur,
    customPrompts: entries,
  };
  await fs.writeFile(appSettingsPath(), JSON.stringify(next, null, 2), "utf8");
}

/** @param {string} label @param {Set<string>} usedIds */
function makeCustomPromptId(label, usedIds) {
  const base =
    "custom:" +
    String(label || "prompt")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, "-")
      .replace(/^-+|-+$/g, "")
      .slice(0, 48) || "prompt";
  let id = base;
  let n = 2;
  while (usedIds.has(id)) {
    id = `${base}-${n}`;
    n += 1;
  }
  usedIds.add(id);
  return id;
}

/** @param {Record<string, string>} partial */
async function writeDownloadDirectoriesPartial(partial) {
  const cur = await readAppSettings();
  /** @type {Record<string, string>} */
  const map = { ...getDownloadDirectoriesFromSettings(cur) };
  for (const [key, val] of Object.entries(partial)) {
    const nk = normalizeDownloadKey(key);
    const trimmed = String(val || "").trim();
    if (trimmed) {
      map[nk] = path.normalize(trimmed);
    } else {
      delete map[nk];
    }
  }
  const next = { ...cur, downloadDirectories: { ...map } };
  delete next.downloadDirectory;
  await fs.writeFile(appSettingsPath(), JSON.stringify(next, null, 2), "utf8");
}

/** @returns {Promise<Record<string, unknown>>} */
async function readAppSettings() {
  try {
    const raw = await fs.readFile(appSettingsPath(), "utf8");
    const data = JSON.parse(raw);
    return data && typeof data === "object" ? data : {};
  } catch {
    return {};
  }
}

/**
 * @param {Record<string, unknown>} settings
 * @returns {Promise<string>}
 */
async function getGlobalResumeSaveRoot() {
  const settings = await readAppSettings();
  const direct = String(settings.resumeSaveRoot || "").trim();
  if (direct) {
    return path.normalize(direct);
  }
  const map = getDownloadDirectoriesFromSettings(settings);
  for (const candidate of [
    map.other,
    map.olu,
    map.justin,
    ...Object.values(map),
  ]) {
    if (candidate) {
      return candidate;
    }
  }
  return "";
}

/**
 * @param {string} directory
 * @returns {Promise<{ ok: true } | { ok: false, error: string }>}
 */
async function setGlobalResumeSaveRoot(directory) {
  const normalized = path.normalize(String(directory || "").trim());
  if (!normalized) {
    return { ok: false, error: "Folder path is empty." };
  }
  await writeAppSettings({ resumeSaveRoot: normalized });
  return { ok: true };
}

/**
 * @param {unknown} downloadKey
 * @returns {Promise<string>}
 */
async function getDownloadDirectoryForKey(_downloadKey) {
  return getGlobalResumeSaveRoot();
}

/**
 * @param {unknown} downloadKey
 * @param {string} directory
 * @returns {Promise<{ ok: true } | { ok: false, error: string }>}
 */
async function setDownloadDirectoryForKey(_downloadKey, directory) {
  return setGlobalResumeSaveRoot(directory);
}

/** Repair legacy / corrupted settings (e.g. downloadDirectories stored as a string). */
async function migrateLegacyDownloadSettings() {
  const settings = await readAppSettings();
  const raw = settings.downloadDirectories;
  const legacy = String(settings.downloadDirectory || "").trim();
  const needsRepair =
    typeof raw === "string" ||
    (legacy && (!raw || typeof raw !== "object"));
  if (!needsRepair && !legacy) {
    return;
  }
  const map = getDownloadDirectoriesFromSettings(settings);
  const next = {
    ...settings,
    downloadDirectories: map,
  };
  delete next.downloadDirectory;
  await fs.writeFile(appSettingsPath(), JSON.stringify(next, null, 2), "utf8");
}

/** @param {Record<string, unknown>} patch */
async function writeAppSettings(patch) {
  const cur = await readAppSettings();
  await fs.writeFile(
    appSettingsPath(),
    JSON.stringify({ ...cur, ...patch }, null, 2),
    "utf8",
  );
}

/**
 * @param {string} dir
 * @param {string} fileName
 * @returns {Promise<string>}
 */
async function uniqueDocxPathInDir(dir, fileName) {
  const safeName = fileName.toLowerCase().endsWith(".docx")
    ? fileName
    : `${fileName}.docx`;
  let candidate = path.join(dir, safeName);
  if (!fsSync.existsSync(candidate)) {
    return candidate;
  }
  const ext = path.extname(safeName);
  const stem = path.basename(safeName, ext);
  for (let i = 2; i < 200; i++) {
    candidate = path.join(dir, `${stem} (${i})${ext}`);
    if (!fsSync.existsSync(candidate)) {
      return candidate;
    }
  }
  return path.join(dir, `${stem} (${Date.now()})${ext}`);
}

/**
 * Keep existing JD.txt; additional saves use JD (2).txt, JD (3).txt, …
 * @param {string} dir
 * @returns {Promise<string>}
 */
async function uniqueJdTxtPathInDir(dir) {
  let candidate = path.join(dir, "JD.txt");
  if (!fsSync.existsSync(candidate)) {
    return candidate;
  }
  for (let i = 2; i < 200; i++) {
    candidate = path.join(dir, `JD (${i}).txt`);
    if (!fsSync.existsSync(candidate)) {
      return candidate;
    }
  }
  return path.join(dir, `JD (${Date.now()}).txt`);
}

/** Absolute path to branding asset `assets/icon.png` (taskbar, window, tray, builds). */
const iconPath = path.resolve(appRoot, "assets", "icon.png");

/** @returns {import("electron").NativeImage | string} */
function windowIconFromDisk() {
  try {
    const buf = fsSync.readFileSync(iconPath);
    const img = nativeImage.createFromBuffer(buf);
    if (!img.isEmpty()) return img;
  } catch {
    /* fall through */
  }
  const img = nativeImage.createFromPath(iconPath);
  return img.isEmpty() ? iconPath : img;
}

/** @type {BrowserWindow | null} */
let mainWindow = null;

/** @type {Set<import("electron").BrowserWindow>} */
const childWindows = new Set();

/** @type {Map<string, import("electron").BrowserWindow>} */
const promptWindowsByKind = new Map();

let lastPromptPopupHotkeyOpenAt = 0;
/** @type {string} */
let lastPromptPopupHotkeyOpenKind = "";

/** @type {Tray | null} */
let tray = null;

let isQuitting = false;

/** @param {import("electron").IpcMainInvokeEvent | { sender?: import("electron").WebContents }} event */
function windowFromEvent(event) {
  if (!event?.sender) {
    return null;
  }
  const win = BrowserWindow.fromWebContents(event.sender);
  if (win && !win.isDestroyed()) {
    return win;
  }
  return null;
}

/** @param {import("electron").IpcMainInvokeEvent} [event] */
function dialogParentWindow(event) {
  return (
    windowFromEvent(event) ||
    BrowserWindow.getFocusedWindow() ||
    (mainWindow && !mainWindow.isDestroyed() ? mainWindow : null)
  );
}

function trayIconImage() {
  const icon = windowIconFromDisk();
  if (typeof icon === "string") return icon;
  return icon.resize({ width: 16, height: 16 });
}

function createTray() {
  if (tray) return;
  tray = new Tray(trayIconImage());
  tray.setToolTip("Resume Builder");
  const menu = Menu.buildFromTemplate([
    {
      label: "Show Resume Builder",
      click: () => {
        if (!mainWindow || mainWindow.isDestroyed()) {
          createWindow();
          return;
        }
        if (mainWindow.isMinimized()) {
          mainWindow.restore();
        }
        mainWindow.show();
        mainWindow.focus();
      },
    },
    { type: "separator" },
    {
      label: "Quit",
      click: () => {
        isQuitting = true;
        app.quit();
      },
    },
  ]);
  tray.setContextMenu(menu);
  tray.on("click", () => {
    if (!mainWindow || mainWindow.isDestroyed()) {
      createWindow();
      return;
    }
    if (mainWindow.isVisible()) {
      mainWindow.focus();
    } else {
      mainWindow.show();
      mainWindow.focus();
    }
  });
}

const ALLOWED_PROMPTS = new Set([
  "prompt/skill extractor.txt",
  "prompt/exp updator.txt",
  "prompt/summary generator.txt",
  "prompt/olu prompt.txt",
  "prompt/justin prompt.txt",
]);

/** If the canonical file is missing/empty in the bundle, try these names (older installs). */
const PROMPT_READ_ALIASES = {
  "olu prompt.txt": ["other prompt.txt"],
  "justin prompt.txt": ["olu prompt.txt", "other prompt.txt"],
};

/**
 * @param {string} filePath
 * @returns {Promise<string | null>}
 */
async function readUtf8IfNonempty(filePath) {
  try {
    const text = await fs.readFile(filePath, "utf8");
    if (String(text).trim().length > 0) {
      return text;
    }
  } catch (e) {
    if (!e || typeof e !== "object" || e.code !== "ENOENT") {
      console.warn("Resume Builder: read failed:", filePath, e);
    }
  }
  return null;
}

/**
 * @param {string} name basename under prompt/
 * @returns {Promise<string | null>}
 */
async function findPromptSourceText(name) {
  const candidates = [name, ...(PROMPT_READ_ALIASES[name] || [])];
  const dirs = [];
  const writable = getWritablePromptDir();
  dirs.push(writable);
  for (const dir of getBundledPromptDirs()) {
    if (dir !== writable) {
      dirs.push(dir);
    }
  }
  for (const tryName of candidates) {
    for (const dir of dirs) {
      const text = await readUtf8IfNonempty(path.join(dir, tryName));
      if (text) {
        return text;
      }
    }
  }
  return null;
}

/**
 * @param {string} relPath e.g. prompt/olu prompt.txt
 * @returns {Promise<string>}
 */
async function readAllowedPromptContent(relPath) {
  let name;
  try {
    name = promptFileNameFromRel(relPath);
  } catch {
    return "";
  }
  const text = await findPromptSourceText(name);
  if (!text) {
    console.error("Resume Builder: could not read prompt anywhere:", name);
    return "";
  }
  return text;
}

/** Copy bundled defaults into the writable prompt dir when missing or empty. */
async function seedWritablePromptsIfNeeded() {
  if (!app.isPackaged) {
    return;
  }
  const destDir = getWritablePromptDir();
  if (destDir === bundledPromptDir) {
    return;
  }
  await fs.mkdir(destDir, { recursive: true });
  for (const relPath of ALLOWED_PROMPTS) {
    let name;
    try {
      name = promptFileNameFromRel(relPath);
    } catch {
      continue;
    }
    const dest = path.join(destDir, name);
    const existing = await readUtf8IfNonempty(dest);
    if (existing) {
      continue;
    }
    const text = await findPromptSourceText(name);
    if (!text) {
      console.error("Resume Builder: could not seed prompt file (no bundle copy):", name);
      continue;
    }
    try {
      await fs.writeFile(dest, text, "utf8");
    } catch (e) {
      console.error("Resume Builder: could not seed prompt file", name, e);
    }
  }
}

/** @param {import("electron").BrowserWindow} win */
function attachAppWindowWebviewHandlers(win) {
  win.webContents.on(
    "will-attach-webview",
    (_event, webPreferences, params) => {
      const part =
        typeof params.partition === "string" && params.partition
          ? params.partition
          : "persist:chatgpt";
      configureChatGptSession(part);
      webPreferences.contextIsolation = true;
      webPreferences.javascript = true;
      webPreferences.nativeWindowOpen = true;
      webPreferences.sandbox = false;
    },
  );

  win.webContents.on("before-input-event", (event, input) => {
    if (input.type === "keyDown" && input.key === "F12") {
      win.webContents.toggleDevTools();
      event.preventDefault();
    }
  });
}

/** @param {import("electron").BrowserWindow} win */
function attachAppWindowTitlebarHooks(win) {
  if (process.platform !== "win32") {
    return;
  }
  const WM_NCLBUTTONDBLCLK = 0x00a3;
  const HTCAPTION = 2;
  /** @param {Buffer | number} wParam */
  const readNcHitTest = (wParam) => {
    if (typeof wParam === "number") {
      return wParam >>> 0;
    }
    if (Buffer.isBuffer(wParam)) {
      if (wParam.length >= 8) {
        return Number(wParam.readBigUInt64LE(0));
      }
      if (wParam.length >= 4) {
        return wParam.readUInt32LE(0);
      }
    }
    return -1;
  };
  win.hookWindowMessage(WM_NCLBUTTONDBLCLK, (wParam) => {
    if (win.isDestroyed()) {
      return;
    }
    if (readNcHitTest(wParam) !== HTCAPTION) {
      return;
    }
    if (win.isMaximized()) {
      win.unmaximize();
    } else {
      win.maximize();
    }
  });
}

/**
 * @param {{ isMain?: boolean, promptKind?: string }} [options]
 * @returns {import("electron").BrowserWindow}
 */
function createAppWindow(options = {}) {
  const isMain = options.isMain === true;
  const promptKind =
    typeof options.promptKind === "string" && options.promptKind.trim()
      ? options.promptKind.trim()
      : null;

  /** @type {import("electron").BrowserWindowConstructorOptions} */
  const opts = {
    width: isMain ? 1440 : 1280,
    height: isMain ? 900 : 820,
    minWidth: 1200,
    minHeight: 680,
    resizable: true,
    maximizable: true,
    show: false,
    frame: false,
    icon: windowIconFromDisk(),
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
      webviewTag: true,
    },
  };

  const win = new BrowserWindow(opts);
  win.setOpacity(1);
  win.setMenuBarVisibility(false);
  win.removeMenu();
  win.setIcon(windowIconFromDisk());
  attachAppWindowWebviewHandlers(win);
  attachAppWindowTitlebarHooks(win);

  if (isMain) {
    mainWindow = win;
    win.on("close", (e) => {
      if (!isQuitting) {
        e.preventDefault();
        win.hide();
      }
    });
  } else {
    childWindows.add(win);
    if (promptKind) {
      promptWindowsByKind.set(promptKind, win);
    }
    win.on("closed", () => {
      childWindows.delete(win);
      if (promptKind && promptWindowsByKind.get(promptKind) === win) {
        promptWindowsByKind.delete(promptKind);
      }
    });
    if (mainWindow && !mainWindow.isDestroyed()) {
      const [mx, my] = mainWindow.getPosition();
      const offset = childWindows.size * 24;
      win.setPosition(mx + 32 + offset, my + 32 + offset);
    }
  }

  const sendMaxState = () => {
    if (win.isDestroyed()) {
      return;
    }
    win.webContents.send("window-max-state", win.isMaximized());
  };
  win.on("maximize", sendMaxState);
  win.on("unmaximize", sendMaxState);

  win.once("ready-to-show", () => {
    win.show();
  });

  if (promptKind) {
    const onPromptPageReady = () => {
      if (!win.isDestroyed()) {
        win.webContents.send("window-prompt-kind", promptKind);
      }
    };
    win.webContents.once("did-finish-load", onPromptPageReady);
    win.loadFile(path.join(__dirname, "index.html"), {
      query: { prompt: promptKind, dedicated: "1" },
    });
  } else {
    win.loadFile(path.join(__dirname, "index.html"));
  }

  return win;
}

function createWindow() {
  createAppWindow({ isMain: true });
}

/** @param {unknown} promptKind */
function focusOrOpenPromptWindow(promptKind) {
  const kind =
    typeof promptKind === "string" && promptKind.trim()
      ? promptKind.trim()
      : "other";
  const now = Date.now();
  if (kind === lastPromptPopupHotkeyOpenKind && now - lastPromptPopupHotkeyOpenAt < 450) {
    const existing = promptWindowsByKind.get(kind);
    if (existing && !existing.isDestroyed()) {
      return existing;
    }
  }
  lastPromptPopupHotkeyOpenKind = kind;
  lastPromptPopupHotkeyOpenAt = now;

  const existing = promptWindowsByKind.get(kind);
  if (existing && !existing.isDestroyed()) {
    if (existing.isMinimized()) {
      existing.restore();
    }
    existing.show();
    existing.focus();
    return { win: existing, alreadyOpen: true };
  }
  const win = createAppWindow({ promptKind: kind, isMain: false });
  promptWindowsByKind.set(kind, win);
  return { win, alreadyOpen: false };
}

/** @param {unknown} promptKind */
function createPromptWindow(promptKind) {
  const r = focusOrOpenPromptWindow(promptKind);
  return r.win;
}

/** @type {Map<string, string>} accelerator -> prompt kind */
const registeredPromptPopupHotkeys = new Map();

/** @type {import("node-global-key-listener").GlobalKeyboardListener | null} */
let promptPopupGlobalKeyListener = null;

/** @type {{ combo: string, kind: string }[]} */
let promptPopupGlobalKeyListenerBindings = [];

function unregisterAllPromptPopupHotkeys() {
  for (const accelerator of registeredPromptPopupHotkeys.keys()) {
    try {
      globalShortcut.unregister(accelerator);
    } catch (_e) {
      /* ignore */
    }
  }
  registeredPromptPopupHotkeys.clear();
  stopPromptPopupGlobalKeyListener();
}

function stopPromptPopupGlobalKeyListener() {
  if (promptPopupGlobalKeyListener) {
    try {
      promptPopupGlobalKeyListener.kill();
    } catch (_e) {
      /* ignore */
    }
    promptPopupGlobalKeyListener = null;
  }
  promptPopupGlobalKeyListenerBindings = [];
}

/** @param {string} mod */
function isGlobalKeyModifierDown(/** @type {import("node-global-key-listener").IGlobalKeyDownMap} */ down, mod) {
  switch (mod) {
    case "Ctrl":
      return !!(down["LEFT CTRL"] || down["RIGHT CTRL"]);
    case "Alt":
      return !!(down["LEFT ALT"] || down["RIGHT ALT"]);
    case "Shift":
      return !!(down["LEFT SHIFT"] || down["RIGHT SHIFT"]);
    case "Meta":
      return !!(
        down["LEFT META"] ||
        down["RIGHT META"] ||
        down["LEFT SUPER"] ||
        down["RIGHT SUPER"]
      );
    default:
      return false;
  }
}

/** @param {string} keyPart */
function comboKeyPartToGlobalKeyName(keyPart) {
  const k = String(keyPart || "").trim();
  if (!k) {
    return "";
  }
  if (k === "Space") {
    return "SPACE";
  }
  if (k.length === 1 && /[0-9]/.test(k)) {
    return k;
  }
  if (k.length === 1) {
    return k.toUpperCase();
  }
  return k.toUpperCase();
}

/**
 * @param {string} combo
 * @param {import("node-global-key-listener").IGlobalKeyEvent} e
 * @param {import("node-global-key-listener").IGlobalKeyDownMap} down
 */
function comboMatchesGlobalKeyEvent(combo, e, down) {
  if (e.state !== "DOWN") {
    return false;
  }
  const parts = String(combo || "")
    .split("+")
    .map((p) => p.trim())
    .filter(Boolean);
  if (parts.length < 2) {
    return false;
  }
  const keyPart = parts[parts.length - 1];
  const modifiers = parts.slice(0, -1);
  const expectedKey = comboKeyPartToGlobalKeyName(keyPart);
  if (!expectedKey || e.name !== expectedKey) {
    return false;
  }
  for (const mod of ["Ctrl", "Alt", "Shift", "Meta"]) {
    const needed = modifiers.includes(mod);
    if (isGlobalKeyModifierDown(down, mod) !== needed) {
      return false;
    }
  }
  return true;
}

/** @param {{ combo: string, kind: string }[]} bindings */
function startPromptPopupGlobalKeyListener(bindings) {
  stopPromptPopupGlobalKeyListener();
  if (!bindings.length) {
    return;
  }
  let GlobalKeyboardListener;
  try {
    ({ GlobalKeyboardListener } = require("node-global-key-listener"));
  } catch (_e) {
    return;
  }
  promptPopupGlobalKeyListenerBindings = bindings.slice();
  promptPopupGlobalKeyListener = new GlobalKeyboardListener();
  promptPopupGlobalKeyListener.addListener((e, down) => {
    for (const entry of promptPopupGlobalKeyListenerBindings) {
      if (comboMatchesGlobalKeyEvent(entry.combo, e, down)) {
        focusOrOpenPromptWindow(entry.kind);
        return true;
      }
    }
    return false;
  });
}

/** @param {string} combo e.g. Ctrl+Shift+1 */
function comboToElectronAccelerator(combo) {
  return String(combo || "")
    .split("+")
    .map((part) => part.trim())
    .filter(Boolean)
    .map((part) => {
      if (part === "Ctrl") {
        return "CommandOrControl";
      }
      if (part === "Alt") {
        return "Alt";
      }
      if (part === "Shift") {
        return "Shift";
      }
      if (part === "Meta") {
        return "Super";
      }
      if (part.length === 1) {
        return part.toUpperCase();
      }
      return part;
    })
    .join("+");
}

/** @param {unknown} arg */
function registerPromptPopupHotkeysFromPayload(arg) {
  unregisterAllPromptPopupHotkeys();
  const bindings =
    arg && typeof arg === "object" && Array.isArray(arg.bindings) ? arg.bindings : [];
  /** @type {string[]} */
  const failed = [];
  /** @type {string[]} */
  const shortcutCombos = [];
  /** @type {string[]} */
  const listenerCombos = [];
  /** @type {Set<string>} */
  const usedAccelerators = new Set();
  /** @type {{ combo: string, kind: string }[]} */
  const listenerBindings = [];
  for (const entry of bindings) {
    const kind =
      entry && typeof entry === "object" && typeof entry.kind === "string"
        ? entry.kind.trim()
        : "";
    const combo =
      entry && typeof entry === "object" && typeof entry.combo === "string"
        ? entry.combo.trim()
        : "";
    if (!kind || !combo) {
      continue;
    }
    const accelerator = comboToElectronAccelerator(combo);
    let registered = false;
    if (accelerator && !usedAccelerators.has(accelerator)) {
      try {
        const ok = globalShortcut.register(accelerator, () => {
          focusOrOpenPromptWindow(kind);
        });
        if (ok) {
          registeredPromptPopupHotkeys.set(accelerator, kind);
          usedAccelerators.add(accelerator);
          shortcutCombos.push(combo);
          registered = true;
        }
      } catch (_e) {
        /* try listener fallback */
      }
    }
    if (!registered) {
      listenerBindings.push({ combo, kind });
    }
  }
  if (listenerBindings.length) {
    startPromptPopupGlobalKeyListener(listenerBindings);
    for (const b of listenerBindings) {
      listenerCombos.push(b.combo);
    }
  }
  for (const entry of bindings) {
    const combo =
      entry && typeof entry === "object" && typeof entry.combo === "string"
        ? entry.combo.trim()
        : "";
    if (
      combo &&
      !shortcutCombos.includes(combo) &&
      !listenerCombos.includes(combo)
    ) {
      failed.push(combo);
    }
  }
  return { ok: true, failed, shortcutCombos, listenerCombos };
}

function promptFileNameFromRel(relPath) {
  const normalized = String(relPath || "").replace(/\\/g, "/");
  if (!ALLOWED_PROMPTS.has(normalized)) {
    throw new Error("Invalid prompt path");
  }
  return normalized.slice("prompt/".length);
}

/** Path used for read/write of allowed prompts (writable in packaged builds). */
function resolveAllowedPrompt(relPath) {
  const name = promptFileNameFromRel(relPath);
  return path.join(getWritablePromptDir(), name);
}

/** Safe Windows/macOS filename stem (no path separators or reserved chars). */
function sanitizeDocxDefaultStem(name) {
  const s = String(name || "")
    .replace(/[\\/:*?"<>|]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\.docx$/i, "");
  return s.slice(0, 180) || "ChatGPT Reply";
}

/** Safe folder name segment for download path folders. */
function sanitizeCompanyFolderSegment(name) {
  const s = String(name || "")
    .replace(/[\\/:*?"<>|]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^\.+/, "")
    .slice(0, 80);
  return s;
}

/** @returns {string} YYYY-MM-DD for dated download folders */
function todayDatePathSegment() {
  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, "0");
  const d = String(now.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

/**
 * `{downloadDir}/{profileName}/{date}/{company}`
 * @param {string} downloadDir
 * @param {string} profileName
 * @param {string} companyName
 * @returns {string | null}
 */
function buildProfileDatedCompanySaveDir(downloadDir, profileName, companyName) {
  const root = String(downloadDir || "").trim();
  const profileSeg = sanitizeCompanyFolderSegment(profileName);
  const companySeg = sanitizeCompanyFolderSegment(companyName);
  if (!root || !profileSeg || !companySeg) {
    return null;
  }
  return path.join(root, profileSeg, todayDatePathSegment(), companySeg);
}

const GPT_DOC_SECTION_HEADERS =
  /^(professional summary|summary|profile|objective|work experience|experience|employment|education|skills|technical skills|core competencies|projects|certifications|certificates|awards|publications|languages|contact|references|key achievements)$/i;

/** Section headings whose body must not appear in exported .docx */
function isExcludedCvSectionHeader(line) {
  const t = String(line || "").trim();
  if (!t) {
    return false;
  }
  if (/^#{1,6}\s*(cover\s*letter|recruiter\s*messages?)\s*:?\s*$/i.test(t)) {
    return true;
  }
  if (/^[*_\s`]*cover\s*letter\s*:?\s*[*_`\s]*$/i.test(t) && t.length < 48) {
    return true;
  }
  if (/^[*_\s`]*recruiter\s*messages?\s*:?\s*[*_`\s]*$/i.test(t) && t.length < 56) {
    return true;
  }
  if (/^(cover\s*letter|recruiter\s*messages?)\s*:?\s*$/i.test(t)) {
    return true;
  }
  const u = t.toUpperCase().replace(/\s+/g, " ").trim();
  if (u === "COVER LETTER" || u === "COVER LETTER:") {
    return true;
  }
  if (u === "RECRUITER MESSAGE" || u === "RECRUITER MESSAGES" || u === "RECRUITER MESSAGES:") {
    return true;
  }
  if (/^COVER\s+LETTER\b/i.test(t) && t.length < 52) {
    return true;
  }
  if (/^RECRUITER\s+MESSAGES?\b/i.test(t) && t.length < 64) {
    return true;
  }
  return false;
}

/** Start of a new major section (ends an excluded block). */
function looksLikeMajorSectionHeaderForStrip(line) {
  const L = String(line || "").trim();
  if (!L) {
    return false;
  }
  if (/^#{1,6}\s+\S/.test(L)) {
    return true;
  }
  if (GPT_DOC_SECTION_HEADERS.test(L)) {
    return true;
  }
  if (
    L.length >= 4 &&
    L.length <= 58 &&
    L === L.toUpperCase() &&
    /[A-Z]/.test(L) &&
    !L.includes("@")
  ) {
    return true;
  }
  return false;
}

/** True when this line begins a new resume block after an excluded section. */
function looksLikeResumeBoundaryAfterExcluded(line) {
  if (isExcludedCvSectionHeader(line)) {
    return false;
  }
  const L = String(line || "").trim();
  if (!L) {
    return false;
  }
  if (looksLikeMajorSectionHeaderForStrip(line)) {
    return true;
  }
  if (/\s-\s/.test(L) && L.length < 160 && L.length > 4) {
    return true;
  }
  return false;
}

/**
 * Remove excluded sections and everything until the next major header.
 * @param {string[]} lines
 * @returns {string[]}
 */
function stripExcludedCvSections(lines) {
  const out = [];
  let skip = false;
  for (const line of lines) {
    const t = line.trim();
    if (!skip) {
      if (isExcludedCvSectionHeader(line)) {
        skip = true;
        continue;
      }
      out.push(line);
      continue;
    }
    if (t && looksLikeResumeBoundaryAfterExcluded(line)) {
      skip = false;
      out.push(line);
    }
  }
  return out;
}

/**
 * Strip assistant preamble/postamble and prefer fenced resume blocks so the .docx
 * contains only CV-style content.
 * @param {string} raw
 * @returns {string}
 */
function extractResumePlainTextForDocx(raw) {
  const s0 = String(raw || "").replace(/\r\n/g, "\n").trim();
  if (!s0) {
    return "";
  }

  let bestFence = "";
  const fenceRe = /```(?:[a-z0-9_-]+)?\s*\n?([\s\S]*?)```/gi;
  let fm;
  while ((fm = fenceRe.exec(s0)) !== null) {
    const inner = fm[1].trim();
    if (inner.length > bestFence.length) {
      bestFence = inner;
    }
  }
  if (bestFence.length >= 100) {
    const fl = bestFence.replace(/\r\n/g, "\n").split("\n");
    return stripExcludedCvSections(fl).join("\n").trim();
  }

  const lines = s0.split("\n");
  const introLine = /^(certainly|sure[,!]|hi[,!]?|hello[,!]?|okay[,!]?|ok[,!]?|great[,!]?|here'?s (your|the)|here is|below (is|you)|i'?ve (created|prepared|drafted|put together|included)|as (you )?requested|thank you|thanks for)/i;

  function looksLikeResumeLine(L) {
    const t = L.trim();
    if (!t) {
      return false;
    }
    if (/\s-\s/.test(t) && t.length < 160) {
      return true;
    }
    if (/^#{1,3}\s+\S/.test(t)) {
      return true;
    }
    if (GPT_DOC_SECTION_HEADERS.test(t)) {
      return true;
    }
    if (/^[-•*·]\s/.test(t) || /^\d+\.\s/.test(t)) {
      return true;
    }
    if (t.length >= 4 && t.length <= 58 && t === t.toUpperCase() && /[A-Z]/.test(t) && !t.includes("@")) {
      return true;
    }
    return false;
  }

  let start = lines.findIndex((ln) => looksLikeResumeLine(ln));
  if (start === -1) {
    let i = 0;
    for (; i < Math.min(lines.length, 22); i++) {
      const L = lines[i].trim();
      if (!L) {
        continue;
      }
      if (introLine.test(L)) {
        continue;
      }
      break;
    }
    start = i;
  }

  let end = lines.length;
  const outro = /^(let me know|hope (this|that)|feel free|if you|would you like|best regards|regards|need any|don'?t hesitate|happy to|source:|sources:|references:|\[\d+\])/i;
  while (end > start + 4) {
    const L = lines[end - 1].trim();
    if (!L) {
      end--;
      continue;
    }
    if (L.length < 160 && outro.test(L)) {
      end--;
      continue;
    }
    if (L.length < 90 && /\?\s*$/.test(L)) {
      end--;
      continue;
    }
    break;
  }

  const bodyLines = stripExcludedCvSections(lines.slice(start, end));
  return bodyLines.join("\n").trim();
}

function resolveUserFilePath(raw) {
  const s = String(raw || "").trim();
  if (!s) {
    throw new Error("Empty path");
  }
  if (/^file:/i.test(s)) {
    try {
      return path.normalize(fileURLToPath(s));
    } catch {
      throw new Error("Invalid file URL");
    }
  }
  if (path.isAbsolute(s)) {
    return path.normalize(s);
  }
  const fromCwd = path.normalize(path.resolve(s));
  if (fsSync.existsSync(fromCwd)) {
    return fromCwd;
  }
  const fromAppRoot = path.normalize(path.join(appRoot, s.replace(/\//g, path.sep)));
  if (fsSync.existsSync(fromAppRoot)) {
    return fromAppRoot;
  }
  return fromCwd;
}

function applyLoginStartup() {
  if (!app.isPackaged) return;
  try {
    app.setLoginItemSettings({
      openAtLogin: true,
      enabled: true,
      name: "Resume Builder",
      path: process.execPath,
    });
  } catch {
    /* ignore if OS denies or API unavailable */
  }
}

/**
 * @param {{ filePath?: string, folderPath?: string }} payload
 * @returns {Promise<{ ok: true } | { ok: false, error: string }>}
 */
async function openGptDownloadFolderImpl(payload) {
  const filePath = path.normalize(String(payload?.filePath ?? "").trim());
  let folderPath = path.normalize(String(payload?.folderPath ?? "").trim());
  if (!folderPath && filePath) {
    folderPath = path.dirname(filePath);
  }
  if (filePath && fsSync.existsSync(filePath)) {
    shell.showItemInFolder(filePath);
    return { ok: true };
  }
  if (folderPath && fsSync.existsSync(folderPath)) {
    const errMsg = await shell.openPath(folderPath);
    if (errMsg) {
      return { ok: false, error: errMsg };
    }
    return { ok: true };
  }
  return { ok: false, error: "Download folder not found." };
}

/**
 * @param {{ filePath?: string }} payload
 * @returns {Promise<{ ok: true } | { ok: false, error: string }>}
 */
async function openGptDownloadFileImpl(payload) {
  const filePath = path.normalize(String(payload?.filePath ?? "").trim());
  if (!filePath) {
    return { ok: false, error: "No file path." };
  }
  if (!fsSync.existsSync(filePath)) {
    return { ok: false, error: "File not found." };
  }
  const errMsg = await shell.openPath(filePath);
  if (errMsg) {
    return { ok: false, error: errMsg };
  }
  return { ok: true };
}

/** @param {string[]} filePaths */
function buildMacFilenamesClipboardBuffer(filePaths) {
  const resolved = filePaths.map((p) => path.resolve(p));
  const plist = `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><array>
${resolved
  .map(
    (p) =>
      `  <string>${p
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")}</string>`,
  )
  .join("\n")}
</array></plist>`;
  return Buffer.from(plist, "utf8");
}

/** @param {string[]} filePaths */
function buildWindowsCfHdropBuffer(filePaths) {
  const paths = filePaths.map((p) => path.win32.normalize(path.resolve(p)));
  const list = `${paths.join("\0")}\0\0`;
  const listBuf = Buffer.from(list, "utf16le");
  const header = Buffer.alloc(20);
  header.writeUInt32LE(20, 0);
  header.writeInt32LE(0, 4);
  header.writeInt32LE(0, 8);
  header.writeUInt32LE(0, 12);
  header.writeUInt32LE(1, 16);
  return Buffer.concat([header, listBuf]);
}

/** @param {string} value */
function psSingleQuoted(value) {
  return `'${String(value).replace(/'/g, "''")}'`;
}

/** @param {string} scriptPath */
function runPowerShellStaFile(scriptPath) {
  return new Promise((resolve, reject) => {
    const proc = spawn(
      "powershell.exe",
      [
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-STA",
        "-File",
        scriptPath,
      ],
      { windowsHide: true },
    );
    let stderr = "";
    proc.stderr.on("data", (chunk) => {
      stderr += String(chunk);
    });
    proc.on("error", reject);
    proc.on("close", (code) => {
      if (code === 0) {
        resolve();
        return;
      }
      reject(new Error(stderr.trim() || `PowerShell exited with code ${code}`));
    });
  });
}

/** @param {string[]} filePaths */
async function copyFilesToClipboardWindows(filePaths) {
  const paths = filePaths.map((p) => path.resolve(p));
  const lines = [
    "Add-Type -AssemblyName System.Windows.Forms",
    "$col = New-Object System.Collections.Specialized.StringCollection",
    ...paths.map((p) => `$col.Add(${psSingleQuoted(p)})`),
    "[System.Windows.Forms.Clipboard]::SetFileDropList($col)",
    "exit 0",
  ];
  const tmpPath = path.join(
    app.getPath("temp"),
    `resume-builder-clipboard-${Date.now()}-${process.pid}.ps1`,
  );
  await fs.writeFile(tmpPath, `${lines.join("\r\n")}\r\n`, "utf8");
  try {
    await runPowerShellStaFile(tmpPath);
  } finally {
    await fs.unlink(tmpPath).catch(() => {});
  }
}

/** @param {string[]} filePaths */
async function copyFilesToClipboardWin32(filePaths) {
  const paths = filePaths.map((p) => path.resolve(p));
  const errors = [];
  try {
    clipboard.writeBuffer("CF_HDROP", buildWindowsCfHdropBuffer(paths));
    return;
  } catch (e) {
    errors.push(String(e && e.message ? e.message : e));
  }
  try {
    await copyFilesToClipboardWindows(paths);
    return;
  } catch (e) {
    errors.push(String(e && e.message ? e.message : e));
  }
  throw new Error(errors.join(" | ") || "Clipboard copy failed.");
}

/** @param {string} rawPath */
async function copyResumeFileToClipboardImpl(rawPath) {
  const filePath = path.normalize(
    String(
      rawPath && typeof rawPath === "object"
        ? rawPath.filePath ?? rawPath.path ?? ""
        : rawPath ?? "",
    ).trim(),
  );
  if (!filePath) {
    return { ok: false, error: "No file path." };
  }
  if (!fsSync.existsSync(filePath)) {
    return { ok: false, error: "File not found." };
  }

  const paths = [filePath];
  try {
    if (process.platform === "win32") {
      await copyFilesToClipboardWin32(paths);
      return { ok: true };
    }
    if (process.platform === "darwin") {
      clipboard.writeBuffer(
        "NSFilenamesPboardType",
        buildMacFilenamesClipboardBuffer(paths),
      );
      return { ok: true };
    }
    clipboard.writeText(filePath);
    return { ok: true, pathOnly: true };
  } catch (e) {
    return {
      ok: false,
      error: String(e && e.message ? e.message : e),
    };
  }
}

function registerGptDownloadFolderIpc() {
  ipcMain.removeHandler("open-gpt-download-folder");
  ipcMain.removeHandler("show-path-in-folder");
  ipcMain.removeHandler("open-gpt-download-file");
  ipcMain.removeHandler("read-gpt-docx-file");
  ipcMain.removeHandler("copy-resume-file-to-clipboard");
  ipcMain.handle("open-gpt-download-folder", async (_event, payload) =>
    openGptDownloadFolderImpl(
      payload && typeof payload === "object" ? payload : {},
    ),
  );
  ipcMain.handle("show-path-in-folder", async (_event, rawPath) => {
    const filePath =
      rawPath && typeof rawPath === "object"
        ? String(rawPath.filePath ?? rawPath.path ?? "")
        : String(rawPath ?? "");
    return openGptDownloadFolderImpl({ filePath });
  });
  ipcMain.handle("open-gpt-download-file", async (_event, payload) =>
    openGptDownloadFileImpl(
      payload && typeof payload === "object" ? payload : {},
    ),
  );
  ipcMain.handle("read-gpt-docx-file", async (_event, rawPath) => {
    const filePath = path.normalize(
      String(
        rawPath && typeof rawPath === "object"
          ? rawPath.filePath ?? rawPath.path ?? ""
          : rawPath ?? "",
      ).trim(),
    );
    if (!filePath) {
      return { ok: false, error: "No file path." };
    }
    if (!fsSync.existsSync(filePath)) {
      return { ok: false, error: "File not found." };
    }
    try {
      const buf = await fs.readFile(filePath);
      return { ok: true, base64: buf.toString("base64") };
    } catch (e) {
      return {
        ok: false,
        error: String(e && e.message ? e.message : e),
      };
    }
  });
  ipcMain.handle("copy-resume-file-to-clipboard", async (_event, rawPath) =>
    copyResumeFileToClipboardImpl(rawPath),
  );
}

const gotLock = app.requestSingleInstanceLock();
if (!gotLock) {
  app.quit();
} else {
  app.on("second-instance", () => {
    if (!mainWindow || mainWindow.isDestroyed()) {
      createWindow();
      return;
    }
    if (mainWindow.isMinimized()) {
      mainWindow.restore();
    }
    mainWindow.show();
    mainWindow.focus();
  });

  registerGptDownloadFolderIpc();

  ipcMain.handle("read-prompt", async (_event, relPath) => {
    return readAllowedPromptContent(relPath);
  });

  ipcMain.handle("save-prompt", async (_event, relPath, content) => {
    const full = resolveAllowedPrompt(relPath);
    const dir = path.dirname(full);
    await fs.mkdir(dir, { recursive: true });
    await fs.writeFile(full, content, "utf8");
    return { ok: true };
  });

  ipcMain.handle("pick-prompt-file", async (event) => {
    const parent = dialogParentWindow(event);
    if (!parent) {
      return { canceled: true };
    }
    const { canceled, filePaths } = await dialog.showOpenDialog(parent, {
      title: "Open prompt file",
      properties: ["openFile"],
      filters: [
        { name: "Text", extensions: ["txt", "md"] },
        { name: "All files", extensions: ["*"] },
      ],
    });
    if (canceled || !filePaths?.[0]) {
      return { canceled: true };
    }
    const filePath = path.normalize(filePaths[0]);
    const text = await fs.readFile(filePath, "utf8");
    const fileUrl = pathToFileURL(filePath).href;
    return { canceled: false, path: filePath, fileUrl, content: text };
  });

  ipcMain.handle("pick-structure-docx", async (event) => {
    const parent = dialogParentWindow(event);
    if (!parent) {
      return { canceled: true };
    }
    const { canceled, filePaths } = await dialog.showOpenDialog(parent, {
      title: "Choose Word structure template",
      properties: ["openFile"],
      filters: [{ name: "Word document", extensions: ["docx"] }],
    });
    if (canceled || !filePaths?.[0]) {
      return { canceled: true };
    }
    const filePath = path.normalize(filePaths[0]);
    const fileUrl = pathToFileURL(filePath).href;
    return { canceled: false, path: filePath, fileUrl };
  });

  ipcMain.handle("get-custom-prompts", async () => {
    const prompts = await readCustomPrompts();
    return { ok: true, prompts };
  });

  ipcMain.handle("save-custom-prompt", async (_event, raw) => {
    const rec = raw && typeof raw === "object" ? raw : {};
    const label = String(rec.label || "").trim();
    const promptPath = path.normalize(String(rec.promptPath || "").trim());
    const structurePath = path.normalize(String(rec.structurePath || "").trim());
    const chatgptProfile = rec.chatgptProfile === "justin" ? "justin" : "olu";
    const existingId = String(rec.id || "").trim();

    if (!label) {
      return { ok: false, error: "Enter a name for this prompt." };
    }
    if (!promptPath || !fsSync.existsSync(promptPath)) {
      return { ok: false, error: "Choose a valid prompt .txt file." };
    }
    if (!structurePath || !fsSync.existsSync(structurePath)) {
      return { ok: false, error: "Choose a valid structure .docx file." };
    }
    if (!structurePath.toLowerCase().endsWith(".docx")) {
      return { ok: false, error: "Structure file must be a .docx template." };
    }

    const list = await readCustomPrompts();
    const usedIds = new Set(list.map((p) => p.id));
    let id = existingId.startsWith("custom:") ? existingId : "";
    if (id && !list.some((p) => p.id === id)) {
      usedIds.delete(id);
    }
    if (!id) {
      id = makeCustomPromptId(label, usedIds);
    } else {
      usedIds.add(id);
    }

    const entry = normalizeCustomPromptEntry({
      id,
      label,
      promptPath,
      structurePath,
      chatgptProfile,
    });
    if (!entry) {
      return { ok: false, error: "Could not save prompt." };
    }

    const next = list.filter((p) => p.id !== id);
    next.push(entry);
    await writeCustomPrompts(next);
    return { ok: true, prompt: entry };
  });

  ipcMain.handle("update-custom-prompt-path", async (_event, raw) => {
    const rec = raw && typeof raw === "object" ? raw : {};
    const id = String(rec.id || "").trim();
    const promptPath = path.normalize(String(rec.promptPath || "").trim());
    if (!id.startsWith("custom:")) {
      return { ok: false, error: "Invalid prompt id." };
    }
    if (!promptPath || !fsSync.existsSync(promptPath)) {
      return { ok: false, error: "Choose a valid prompt .txt file." };
    }
    const list = await readCustomPrompts();
    const idx = list.findIndex((p) => p.id === id);
    if (idx < 0) {
      return { ok: false, error: "Prompt not found." };
    }
    const entry = normalizeCustomPromptEntry({
      ...list[idx],
      promptPath,
    });
    if (!entry) {
      return { ok: false, error: "Could not save prompt path." };
    }
    const next = [...list];
    next[idx] = entry;
    await writeCustomPrompts(next);
    return { ok: true, prompt: entry };
  });

  ipcMain.handle("close-prompt-window", async (_event, raw) => {
    const id = String(
      raw && typeof raw === "object" && raw.id != null ? raw.id : raw || "",
    ).trim();
    if (!id) {
      return { ok: false, error: "Invalid prompt id." };
    }
    const popupWin = promptWindowsByKind.get(id);
    if (popupWin && !popupWin.isDestroyed()) {
      popupWin.close();
    }
    return { ok: true };
  });

  ipcMain.handle("delete-custom-prompt", async (_event, raw) => {
    const id = String(
      raw && typeof raw === "object" && raw.id != null ? raw.id : raw || "",
    ).trim();
    if (!id.startsWith("custom:")) {
      return { ok: false, error: "Invalid prompt id." };
    }
    const list = await readCustomPrompts();
    const next = list.filter((p) => p.id !== id);
    if (next.length === list.length) {
      return { ok: false, error: "Prompt not found." };
    }
    const popupWin = promptWindowsByKind.get(id);
    if (popupWin && !popupWin.isDestroyed()) {
      popupWin.close();
    }
    await writeCustomPrompts(next);
    return { ok: true };
  });

  ipcMain.handle("read-prompt-file-path", async (_event, raw) => {
    try {
      const full = resolveUserFilePath(raw);
      if (!fsSync.existsSync(full)) {
        return { ok: false, code: "ENOENT", error: "File not found", path: full };
      }
      const content = await fs.readFile(full, "utf8");
      const fileUrl = pathToFileURL(full).href;
      return { ok: true, content, path: full, fileUrl };
    } catch (e) {
      const err = e && typeof e === "object" ? e : {};
      return {
        ok: false,
        code: err.code || "READ_FAILED",
        error: err.message || "Could not read file",
        path: String(raw || "").trim(),
      };
    }
  });

  ipcMain.handle("save-prompt-external", async (_event, raw, content) => {
    const full = resolveUserFilePath(raw);
    await fs.writeFile(full, String(content ?? ""), "utf8");
    return { ok: true };
  });

  ipcMain.handle("get-download-directory", async (_event, arg) => {
    const downloadKey = ipcDownloadKey(arg);
    const dir = await getDownloadDirectoryForKey(downloadKey);
    return {
      downloadKey,
      directory: dir,
      exists: !!(dir && fsSync.existsSync(dir)),
    };
  });

  ipcMain.handle("pick-download-directory", async (event, arg) => {
    const win = dialogParentWindow(event);
    if (!win || win.isDestroyed()) {
      return { ok: false, error: "No window" };
    }
    const downloadKey = ipcDownloadKey(arg);
    const current = await getDownloadDirectoryForKey(downloadKey);
    const { canceled, filePaths } = await dialog.showOpenDialog(win, {
      title: "Choose global resume save folder",
      properties: ["openDirectory", "createDirectory"],
      defaultPath: current || app.getPath("documents"),
    });
    if (canceled || !filePaths?.[0]) {
      return { ok: false, canceled: true };
    }
    const directory = path.normalize(filePaths[0]);
    const saved = await setDownloadDirectoryForKey(downloadKey, directory);
    if (!saved.ok) {
      return { ok: false, error: saved.error };
    }
    return { ok: true, downloadKey, directory };
  });

  /**
   * @param {string} raw
   * @returns {{ ok: true, format: "json" } | { ok: false, error: string }}
   */
  function describeGptDocxExportable(raw) {
    const text = String(raw ?? "");
    const jsonObj = extractResumeJsonObjectFromGptText(text);
    if (jsonObj && isCompleteResumeStructureJsonForExport(jsonObj)) {
      return { ok: true, format: "json" };
    }
    if (jsonObj) {
      return {
        ok: false,
        error:
          "Resume JSON is incomplete — include work_experience with at least one job (resume_structure format).",
      };
    }
    return {
      ok: false,
      error:
        "Download requires complete resume_structure JSON (usually in a ```json``` block). Plain-text or markdown CV replies cannot be downloaded.",
    };
  }

  ipcMain.handle("check-gpt-docx-export", (_event, raw) => {
    return describeGptDocxExportable(raw);
  });

  /* Debug trace for the send → reply → Word build flow. */
  ipcMain.handle("gpt-debug-log", (_event, msg) => {
    try {
      fsSync.appendFileSync(
        path.join(app.getPath("userData"), "gpt-debug.log"),
        `${new Date().toISOString()} ${String(msg)}\n`,
      );
    } catch (_e) {
      /* ignore */
    }
  });

  ipcMain.handle("save-gpt-docx", async (event, payload) => {
    const win = dialogParentWindow(event);
    if (!win || win.isDestroyed()) {
      return { ok: false, error: "No window" };
    }
    const raw = String(payload?.text ?? "");
    const profile = payload?.profile === "olu" ? "olu" : "justin";
    const structureRaw = String(payload?.structureTemplatePath || "").trim();
    const structureTemplatePath =
      structureRaw && fsSync.existsSync(structureRaw)
        ? path.normalize(structureRaw)
        : null;
    const exportCheck = describeGptDocxExportable(raw);
    if (!exportCheck.ok) {
      return { ok: false, error: exportCheck.error };
    }
    const jsonObj = extractResumeJsonObjectFromGptText(raw);
    if (!jsonObj || !isCompleteResumeStructureJsonForExport(jsonObj)) {
      return { ok: false, error: exportCheck.error };
    }

    const downloadKey = ipcDownloadKey({
      downloadKey: payload?.downloadKey,
      profile: payload?.profile,
    });
    const downloadDir = await getDownloadDirectoryForKey(downloadKey);
    const companyRaw = String(payload?.companyFolder ?? "").trim();
    const companySeg = sanitizeCompanyFolderSegment(companyRaw);
    if (!companyRaw || !companySeg) {
      return {
        ok: false,
        error: "Enter a company name before downloading (saved under that folder).",
      };
    }
    if (!downloadDir) {
      return {
        ok: false,
        needsDirectory: true,
        downloadKey,
        error: "Set a global resume download folder first (Settings → Browse).",
      };
    }
    const profileName =
      sanitizeCompanyFolderSegment(String(payload?.profileName ?? "").trim()) ||
      sanitizeCompanyFolderSegment(downloadKey) ||
      "Profile";
    const saveDir = buildProfileDatedCompanySaveDir(
      downloadDir,
      profileName,
      companyRaw,
    );
    if (!saveDir) {
      return {
        ok: false,
        error: "Set a profile name and company before downloading.",
      };
    }
    const defaultFileName = `${profileName}.docx`;
    await fs.mkdir(saveDir, { recursive: true });

    const profileLabel = "Resume";
    const wantsAutoSave =
      payload?.useDefaultDirectory === true ||
      (payload?.useDefaultDirectory !== false &&
        payload?.showSaveDialog !== true);
    const preferDirectory =
      wantsAutoSave && !!downloadDir && fsSync.existsSync(downloadDir);

    if (wantsAutoSave && payload?.useDefaultDirectory === true && !preferDirectory) {
      return {
        ok: false,
        needsDirectory: true,
        downloadKey,
        error: "Set a global resume download folder first (Settings → Browse).",
      };
    }

    let outPath = "";
    if (preferDirectory) {
      outPath = await uniqueDocxPathInDir(saveDir, defaultFileName);
    } else {
      const defaultPath = saveDir
        ? path.join(saveDir, defaultFileName)
        : downloadDir
          ? path.join(downloadDir, defaultFileName)
          : defaultFileName;
      const { canceled, filePath } = await dialog.showSaveDialog(win, {
        title: `Save as Word — ${profileLabel}`,
        defaultPath,
        filters: [{ name: "Word document", extensions: ["docx"] }],
      });
      if (canceled || !filePath) {
        return { ok: false, canceled: true };
      }
      outPath = filePath;
      if (!outPath.toLowerCase().endsWith(".docx")) {
        outPath += ".docx";
      }
    }

    try {
      let buf;
      try {
        buf = await buildResumeFromResumeStructureJson(
          jsonObj,
          appRoot,
          profile,
          structureTemplatePath,
        );
      } catch (e) {
        console.error(
          "[save-gpt-docx] JSON template failed:",
          formatDocxtemplaterError(e),
        );
        const parsed = resumeJsonObjectToJustinFields(jsonObj);
        const {
          buildJustinTemplateResumeDocxBufferParsed,
        } = require("./resume-justin-docx.js");
        buf = await buildJustinTemplateResumeDocxBufferParsed(parsed);
      }
      if (!buf) {
        return {
          ok: false,
          error:
            "Could not build Word file. Check your Word structure template in Manage users (matching {tags}), or paste resume JSON.",
        };
      }
      await fs.writeFile(outPath, buf);

      let jdSaved = false;
      let jdPath = "";
      const jdRaw = String(payload?.jdText ?? "").trim();
      if (jdRaw) {
        const jdDir = path.dirname(outPath);
        jdPath = await uniqueJdTxtPathInDir(jdDir);
        await fs.writeFile(jdPath, jdRaw, "utf8");
        jdSaved = true;
      }

      const folderPath = path.dirname(outPath);
      return {
        ok: true,
        filePath: outPath,
        folderPath,
        profile,
        directory: preferDirectory ? downloadDir : "",
        savedToDirectory: preferDirectory,
        jdSaved,
        jdPath,
      };
    } catch (e) {
      return {
        ok: false,
        error: formatDocxtemplaterError(e),
      };
    }
  });

  /**
   * For the ChatGPT reply modal: extract resume JSON (if any) and return work_experience
   * for display. Renderer cannot require resume-json-to-justin.js (no nodeIntegration).
   */
  ipcMain.handle("resolve-work-experience-preview", (_event, raw) => {
    const obj = extractResumeJsonObjectFromGptText(String(raw ?? ""));
    if (!obj || typeof obj !== "object" || Array.isArray(obj)) {
      return { ok: false };
    }
    const rec = /** @type {Record<string, unknown>} */ (obj);
    if (!Object.prototype.hasOwnProperty.call(rec, "work_experience") &&
        !Object.prototype.hasOwnProperty.call(rec, "work_xperience")) {
      return { ok: false };
    }
    const wx = Array.isArray(rec.work_experience)
      ? rec.work_experience
      : Array.isArray(rec.work_xperience)
        ? rec.work_xperience
        : [];
    return {
      ok: true,
      work_experience: wx,
      word_preview_rows: buildWorkExperienceWordPreviewRows(wx),
    };
  });

  ipcMain.handle("export-resume-json-to-docx", async (event) => {
    const win = dialogParentWindow(event);
    if (!win || win.isDestroyed()) {
      return { ok: false, error: "No window" };
    }
    const { canceled, filePaths } = await dialog.showOpenDialog(win, {
      title: "Open resume_structure JSON (GPT export)",
      filters: [{ name: "JSON", extensions: ["json"] }],
      properties: ["openFile"],
    });
    if (canceled || !filePaths?.[0]) {
      return { ok: false, canceled: true };
    }
    const jsonPath = path.normalize(filePaths[0]);
    let raw;
    try {
      raw = await fs.readFile(jsonPath, "utf8");
    } catch {
      return { ok: false, error: "Could not read the JSON file." };
    }
    let buf;
    try {
      buf = await buildResumeFromResumeStructureJson(raw, appRoot);
    } catch (e) {
      return {
        ok: false,
        error: e && typeof e === "object" && "message" in e ? e.message : String(e),
      };
    }
    const jsonObj = parseResumeStructureJson(raw);
    const stem = sanitizeDocxDefaultStem(
      jsonObj
        ? suggestJustinMcgheeDocxFileStem(jsonObj)
        : path.basename(jsonPath, path.extname(jsonPath)) || "Resume",
    );
    const { canceled: saveCanceled, filePath: savePath } = await dialog.showSaveDialog(win, {
      title: "Save resume as Word",
      defaultPath: `${stem}.docx`,
      filters: [{ name: "Word document", extensions: ["docx"] }],
    });
    if (saveCanceled || !savePath) {
      return { ok: false, canceled: true };
    }
    let outPath = savePath;
    if (!outPath.toLowerCase().endsWith(".docx")) {
      outPath += ".docx";
    }
    try {
      await fs.writeFile(outPath, buf);
      return { ok: true, filePath: outPath };
    } catch (e) {
      return {
        ok: false,
        error: e && typeof e === "object" && "message" in e ? e.message : String(e),
      };
    }
  });

  ipcMain.handle("get-app-version", () => APP_VERSION);

  ipcMain.handle("window-set-title", (event, title) => {
    const win = windowFromEvent(event);
    if (win) {
      const t = String(title ?? "").trim();
      if (t) {
        win.setTitle(t);
      }
    }
    return { ok: true };
  });

  ipcMain.handle("window-set-always-on-top", (event, flag) => {
    const win = windowFromEvent(event);
    if (win) {
      win.setAlwaysOnTop(!!flag);
    }
  });

  ipcMain.handle("window-minimize", (event) => {
    const win = windowFromEvent(event);
    if (win) {
      win.minimize();
    }
  });

  ipcMain.handle("window-maximize-toggle", (event) => {
    const win = windowFromEvent(event);
    if (!win) {
      return false;
    }
    if (win.isMaximized()) {
      win.unmaximize();
    } else {
      win.maximize();
    }
    return win.isMaximized();
  });

  ipcMain.handle("window-is-maximized", (event) => {
    const win = windowFromEvent(event);
    return !!(win && win.isMaximized());
  });

  ipcMain.handle("window-close", (event) => {
    const win = windowFromEvent(event);
    if (win) {
      win.close();
    }
  });

  ipcMain.handle("open-chatgpt-external", async (_event, profile) => {
    const part = chatGptPartitionForProfile(profile === "justin" ? "justin" : "olu");
    await shell.openExternal(CHATGPT_EMAIL_LOGIN_URL);
    return { ok: true, partition: part };
  });

  ipcMain.handle("open-chatgpt-login-window", async (_event, payload) => {
    const part =
      payload && typeof payload.partition === "string" && payload.partition
        ? payload.partition
        : chatGptPartitionForProfile(
            payload && payload.profile === "justin" ? "justin" : "olu",
          );
    openChatGptLoginWindow(part);
    return { ok: true, partition: part };
  });

  ipcMain.handle("clear-chatgpt-session", async (_event, payload) => {
    const part =
      payload && typeof payload.partition === "string" && payload.partition
        ? payload.partition
        : chatGptPartitionForProfile(
            payload && payload.profile === "justin" ? "justin" : "olu",
          );
    const ses = session.fromPartition(part, { cache: true });
    await ses.clearStorageData();
    return { ok: true, partition: part };
  });

  ipcMain.handle("open-prompt-window", (_event, arg) => {
    const kind =
      arg && typeof arg === "object" && typeof arg.prompt === "string"
        ? arg.prompt
        : "other";
    const r = focusOrOpenPromptWindow(kind);
    return { ok: true, prompt: kind, alreadyOpen: !!r.alreadyOpen };
  });

  ipcMain.handle("sync-prompt-popup-hotkeys", (_event, arg) => {
    return registerPromptPopupHotkeysFromPayload(arg);
  });

  app.whenReady().then(async () => {
    if (process.platform === "win32") {
      app.setAppUserModelId("com.promptbuilder.app");
    }
    try {
      await migrateLegacyDownloadSettings();
    } catch (e) {
      console.error("Resume Builder: migrateLegacyDownloadSettings", e);
    }
    for (const part of CHATGPT_PARTITIONS) {
      configureChatGptSession(part);
    }
    applyLoginStartup();
    try {
      await seedWritablePromptsIfNeeded();
    } catch (e) {
      console.error("Resume Builder: seedWritablePromptsIfNeeded", e);
    }
    if (app.dock) {
      const dockImg = nativeImage.createFromPath(iconPath);
      if (!dockImg.isEmpty()) {
        app.dock.setIcon(dockImg);
      }
    }
    createWindow();
    createTray();

    app.on("activate", () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
      } else if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.show();
      }
    });
  });

  app.on("before-quit", () => {
    isQuitting = true;
    unregisterAllPromptPopupHotkeys();
    stopPromptPopupGlobalKeyListener();
    if (tray) {
      tray.destroy();
      tray = null;
    }
  });
}
