/* global windowControls — exposed from preload in Electron */

const jdInput = document.getElementById("jdInput");
const mainCompanyInput = document.getElementById("mainCompanyInput");
const editPromptSettings = document.getElementById("editPromptSettings");
const editPromptUserNameInput = document.getElementById("editPromptUserNameInput");
const editPromptPathDisplay = document.getElementById("editPromptPathDisplay");
const editPromptPathBrowseBtn = document.getElementById("editPromptPathBrowseBtn");
const editPromptStructureDisplay = document.getElementById("editPromptStructureDisplay");
const editPromptStructureBrowseBtn = document.getElementById("editPromptStructureBrowseBtn");
const editPromptPopupHotkeyDisplay = document.getElementById("editPromptPopupHotkeyDisplay");
const editPromptPopupHotkeySetBtn = document.getElementById("editPromptPopupHotkeySetBtn");
const editPromptPopupHotkeyClearBtn = document.getElementById("editPromptPopupHotkeyClearBtn");
const editPromptDeleteField = document.getElementById("editPromptDeleteField");
const editPromptDeleteBtn = document.getElementById("editPromptDeleteBtn");
const mainPromptStructureDisplay = document.getElementById("mainPromptStructureDisplay");
const mainPromptStructureBrowseBtn = document.getElementById("mainPromptStructureBrowseBtn");
const dedicatedPromptDirsPanel = document.getElementById("dedicatedPromptDirsPanel");
const dedicatedPromptStructureDisplay = document.getElementById(
  "dedicatedPromptStructureDisplay",
);
const dedicatedPromptStructureBrowseBtn = document.getElementById(
  "dedicatedPromptStructureBrowseBtn",
);
const workspaceBrowserSplit = document.getElementById("workspaceBrowserSplit");
const workspaceLeft = document.getElementById("workspaceLeft");
const workspaceSplitResizer = document.getElementById("workspaceSplitResizer");
const dedicatedBrowserLayout = document.querySelector(".dedicated-browser-layout");
const dedicatedBrowserSide = document.getElementById("dedicatedBrowserSide");
const dedicatedBrowserSplitResizer = document.getElementById(
  "dedicatedBrowserSplitResizer",
);
const previewOutput = document.getElementById("previewOutput");
const previewMeta = document.getElementById("previewMeta");
const reloadAllBtn = document.getElementById("reloadAllBtn");
const exportResumeJsonBtn = document.getElementById("exportResumeJsonBtn");
const reloadBuildBtn = document.getElementById("reloadBuildBtn");
const inputGptBtn = document.getElementById("inputGptBtn");
const gptResultReadyBar = document.getElementById("gptResultReadyBar");
const gptResultReadyLabel = document.getElementById("gptResultReadyLabel");
const viewGptResultBtn = document.getElementById("viewGptResultBtn");
const downloadGptResultBtn = document.getElementById("downloadGptResultBtn");
const fetchGptResultAgainBtn = document.getElementById("fetchGptResultAgainBtn");
const togglePromptPanelBtn = document.getElementById("togglePromptPanelBtn");
const hideBrowserPanelBtn = document.getElementById("hideBrowserPanelBtn");
const gptResultModal = document.getElementById("gptResultModal");
const gptResultModalStatus = document.getElementById("gptResultModalStatus");
const gptResultModalBody = document.getElementById("gptResultModalBody");
const gptResultModalClose = document.getElementById("gptResultModalClose");
const gptResultModalFetchBtn = document.getElementById("gptResultModalFetchBtn");
const gptResultModalDownloadBtn = document.getElementById(
  "gptResultModalDownloadBtn",
);
const gptResultModalCopy = document.getElementById("gptResultModalCopy");
const gptResultModalDownloadWord = document.getElementById(
  "gptResultModalDownloadWord",
);
const gptResultModalWxWrap = document.getElementById("gptResultModalWxWrap");
const gptResultModalWxWordPreview = document.getElementById(
  "gptResultModalWxWordPreview",
);
const gptResultModalWorkExperience = document.getElementById(
  "gptResultModalWorkExperience",
);

const pageBuild = document.getElementById("pageBuild");
const pageBrowser = document.getElementById("pageBrowser");
const tabBuild = document.getElementById("tabBuild");
const tabDedicatedPrompt = document.getElementById("tabDedicatedPrompt");
const tabsListMain = document.getElementById("tabsListMain");
const tabsListDedicated = document.getElementById("tabsListDedicated");
const dedicatedBrowserTabStrip = document.getElementById("dedicatedBrowserTabStrip");
const dedicatedBrowserPages = document.getElementById("dedicatedBrowserPages");
const dedicatedBrowserSideInner = document.getElementById("dedicatedBrowserSideInner");
const dedicatedAddBrowserTabBtn = document.getElementById("dedicatedAddBrowserTabBtn");
const dedicatedCompanySection = document.getElementById("dedicatedCompanySection");
const dedicatedCompanyInput = document.getElementById("dedicatedCompanyInput");

const DEDICATED_COMPANY_STORAGE_PREFIX = "prompt-builder.dedicated-company.";
/** One company name for all prompts (not per user). */
const SHARED_COMPANY_STORAGE_KEY = "resume-builder.company";
const MAIN_COMPANY_STORAGE_PREFIX = "resume-builder.company.";
const PROMPT_USER_NAME_STORAGE_PREFIX = "resume-builder.prompt-user-name.";
const PROMPT_STRUCTURE_STORAGE_PREFIX = "resume-builder.prompt-structure.";
const PROMPT_PATH_STORAGE_PREFIX = "resume-builder.prompt-path.";
const PROMPT_POPUP_HOTKEY_STORAGE_PREFIX = "resume-builder.prompt-popup-hotkey.";
const HIDDEN_BUILTIN_PROMPTS_STORAGE_KEY = "pb_hidden_builtin_prompts";
const REMOVE_BUILTIN_SAMPLES_KEY = "pb_remove_builtin_samples_v2";
const DEFAULT_HIDDEN_BUILTIN_PROMPTS = /** @type {const} */ (["other", "justin"]);

/** @type {boolean} */
let capturingPromptPopupHotkey = false;

/** @type {string | null} Prompt id for Edit-tab user name / structure (other, justin, custom:…). */
let editPromptSettingsKind = null;
const DEDICATED_BROWSER_TAB_STORAGE_PREFIX = "prompt-builder.dedicated-browser-tab.";

/**
 * @typedef {object} DedicatedBrowserTabEntry
 * @property {string} id
 * @property {"olu" | "justin"} profile
 * @property {HTMLElement} pageEl
 * @property {HTMLElement} tabEl
 * @property {HTMLElement | null} paneEl
 * @property {string} jd
 * @property {string} company
 */

/** @type {DedicatedBrowserTabEntry[]} */
let dedicatedBrowserTabs = [];
let dedicatedBrowserTabSeq = 0;
/** @type {"prompt" | string} */
let dedicatedActiveWorkspacePage = "prompt";

const DEDICATED_PAGE_TRANSITION_MS = 220;

/** @type {WeakMap<HTMLElement, number>} */
const dedicatedPageHideGeneration = new WeakMap();

/** @param {HTMLElement | null} el @param {boolean} visible */
function setDedicatedAnimatedPageVisible(el, visible) {
  if (!el) {
    return;
  }
  const nextGen = (dedicatedPageHideGeneration.get(el) || 0) + 1;
  dedicatedPageHideGeneration.set(el, nextGen);
  if (visible) {
    el.hidden = false;
    el.classList.remove("page-hidden");
    el.setAttribute("aria-hidden", "false");
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        if (dedicatedPageHideGeneration.get(el) === nextGen) {
          el.classList.add("dedicated-tab-page-active");
        }
      });
    });
    return;
  }
  el.classList.remove("dedicated-tab-page-active");
  window.setTimeout(() => {
    if (dedicatedPageHideGeneration.get(el) !== nextGen) {
      return;
    }
    if (!el.classList.contains("dedicated-tab-page-active")) {
      el.classList.add("page-hidden");
      el.hidden = true;
      el.setAttribute("aria-hidden", "true");
    }
  }, DEDICATED_PAGE_TRANSITION_MS);
}

const editTemplate = document.getElementById("editTemplate");
const editPanelTitle = document.getElementById("editPanelTitle");
const editSaveBtn = document.getElementById("editSaveBtn");
const editReloadBtn = document.getElementById("editReloadBtn");
/** @type {HTMLInputElement | null} */
const editFileUrlInput = document.querySelector("#editFileUrlInput");
const editFileBrowseBtn = document.getElementById("editFileBrowseBtn");
const editFileLoadBtn = document.getElementById("editFileLoadBtn");
const editFileBtns = document.querySelectorAll(".edit-file-btn");

const OTHER_ID = "prompt/olu prompt.txt";
const JUSTIN_ID = "prompt/justin prompt.txt";

const EDIT_LABELS = {
  other: "olu prompt.txt",
  justin: "justin prompt.txt",
};

let otherTemplateStr = "";
let justinTemplateStr = "";

/**
 * @typedef {{
 *   id: string,
 *   label: string,
 *   promptPath: string,
 *   structurePath: string,
 *   chatgptProfile: "olu" | "justin",
 * }} CustomPromptEntry
 */

/** @type {CustomPromptEntry[]} */
let customPrompts = [];

/** @type {Record<string, string>} */
const customTemplateStrById = {};

const editCustomPromptNav = document.getElementById("editCustomPromptNav");
const buildPromptCardsEl = document.getElementById("buildPromptCards");
const addPromptNavBtn = document.getElementById("addPromptNavBtn");
const addPromptModal = document.getElementById("addPromptModal");
const addPromptForm = document.getElementById("addPromptForm");
const addPromptLabelInput = document.getElementById("addPromptLabel");
const addPromptPathDisplay = document.getElementById("addPromptPathDisplay");
const addPromptStructureDisplay = document.getElementById("addPromptStructureDisplay");
const addPromptPathHidden = document.getElementById("addPromptPathHidden");
const addPromptStructureHidden = document.getElementById("addPromptStructureHidden");
const addPromptEditId = document.getElementById("addPromptEditId");
const addPromptModalTitle = document.getElementById("addPromptModalTitle");
const addPromptModalClose = document.getElementById("addPromptModalClose");
const addPromptCancelBtn = document.getElementById("addPromptCancelBtn");
const addPromptDeleteBtn = document.getElementById("addPromptDeleteBtn");
const addPromptBrowseTxtBtn = document.getElementById("addPromptBrowseTxtBtn");
const addPromptBrowseDocxBtn = document.getElementById("addPromptBrowseDocxBtn");

/** @param {string} s */
function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/** @param {string | null | undefined} k */
function isBuiltinPromptKind(k) {
  return k === "other" || k === "justin";
}

/** @returns {Set<"other" | "justin">} */
function readHiddenBuiltinPromptSet() {
  /** @type {Set<"other" | "justin">} */
  const hidden = new Set();
  try {
    const raw = localStorage.getItem(HIDDEN_BUILTIN_PROMPTS_STORAGE_KEY);
    if (!raw) {
      return new Set(DEFAULT_HIDDEN_BUILTIN_PROMPTS);
    }
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed) || parsed.length === 0) {
      return new Set(DEFAULT_HIDDEN_BUILTIN_PROMPTS);
    }
    for (const id of parsed) {
      if (id === "other" || id === "justin") {
        hidden.add(id);
      }
    }
  } catch (_e) {
    return new Set(DEFAULT_HIDDEN_BUILTIN_PROMPTS);
  }
  return hidden;
}

/** Remove Oluwaseun / Justin sample users from the app (files on disk stay). */
async function removeDefaultSampleUsers() {
  try {
    if (localStorage.getItem(REMOVE_BUILTIN_SAMPLES_KEY) === "1") {
      return;
    }
    writeHiddenBuiltinPromptSet(new Set(DEFAULT_HIDDEN_BUILTIN_PROMPTS));
    for (const kind of DEFAULT_HIDDEN_BUILTIN_PROMPTS) {
      clearPromptLocalSettingsForKind(kind);
      await closePromptWindowForKind(kind);
    }
    localStorage.setItem(REMOVE_BUILTIN_SAMPLES_KEY, "1");
  } catch (_e) {
    /* ignore */
  }
}

/** @param {Set<"other" | "justin">} hidden */
function writeHiddenBuiltinPromptSet(hidden) {
  try {
    localStorage.setItem(
      HIDDEN_BUILTIN_PROMPTS_STORAGE_KEY,
      JSON.stringify([...hidden]),
    );
  } catch (_e) {
    /* ignore */
  }
}

/** @param {"other" | "justin"} kind */
function hideBuiltinPromptKind(kind) {
  const hidden = readHiddenBuiltinPromptSet();
  hidden.add(kind);
  writeHiddenBuiltinPromptSet(hidden);
}

/** @param {string | null | undefined} kind */
function isPromptKindAvailable(kind) {
  const k = String(kind || "").trim();
  if (isCustomPromptKind(k)) {
    return !!getCustomPromptByKind(k);
  }
  if (isBuiltinPromptKind(k)) {
    return !readHiddenBuiltinPromptSet().has(k);
  }
  return false;
}

/** @param {string | null | undefined} k */
function isCustomPromptKind(k) {
  return typeof k === "string" && k.startsWith("custom:");
}

/** @param {string | null | undefined} kind */
function getCustomPromptByKind(kind) {
  if (!isCustomPromptKind(kind)) {
    return null;
  }
  return customPrompts.find((p) => p.id === kind) || null;
}

/** @param {string} promptId */
function promptStructureStorageKey(promptId) {
  return `${PROMPT_STRUCTURE_STORAGE_PREFIX}${promptId}`;
}

/** @param {string} promptId */
function promptPathStorageKey(promptId) {
  return `${PROMPT_PATH_STORAGE_PREFIX}${promptId}`;
}

/** @param {string} promptId */
function promptPopupHotkeyStorageKey(promptId) {
  return `${PROMPT_POPUP_HOTKEY_STORAGE_PREFIX}${promptId}`;
}

/** @param {string} promptId */
function getPromptPopupHotkey(promptId) {
  const id = String(promptId || "").trim();
  if (!id) {
    return "";
  }
  try {
    return localStorage.getItem(promptPopupHotkeyStorageKey(id))?.trim() || "";
  } catch (_e) {
    return "";
  }
}

/** @param {string} promptId @param {string} combo */
function setPromptPopupHotkey(promptId, combo) {
  const id = String(promptId || "").trim();
  if (!id) {
    return;
  }
  const val = String(combo || "").trim();
  try {
    if (val) {
      localStorage.setItem(promptPopupHotkeyStorageKey(id), val);
    } else {
      localStorage.removeItem(promptPopupHotkeyStorageKey(id));
    }
  } catch (_e) {
    /* ignore */
  }
  refreshPromptPopupHotkeyTooltips();
  void syncPromptPopupHotkeysToMain();
}

/** @param {KeyboardEvent} e */
function formatHotkeyFromKeyboardEvent(e) {
  const key = e.key;
  if (!key || key === "Escape") {
    return "";
  }
  const modifierOnly = new Set(["Control", "Alt", "Shift", "Meta"]);
  if (modifierOnly.has(key)) {
    return "";
  }
  const parts = [];
  if (e.ctrlKey) {
    parts.push("Ctrl");
  }
  if (e.altKey) {
    parts.push("Alt");
  }
  if (e.shiftKey) {
    parts.push("Shift");
  }
  if (e.metaKey) {
    parts.push("Meta");
  }
  if (!parts.length) {
    return "";
  }
  let keyLabel = key;
  if (keyLabel === " ") {
    keyLabel = "Space";
  } else if (keyLabel.length === 1) {
    keyLabel = keyLabel.toUpperCase();
  }
  parts.push(keyLabel);
  return parts.join("+");
}

/** @param {string} key */
function normalizeHotkeyKeyLabel(key) {
  const k = String(key || "");
  if (k === " ") {
    return "Space";
  }
  if (k.length === 1) {
    return k.toUpperCase();
  }
  return k;
}

/** @param {KeyboardEvent} e @param {string} combo */
function keyboardEventMatchesHotkeyCombo(e, combo) {
  const target = String(combo || "").trim();
  if (!target) {
    return false;
  }
  const parts = target.split("+").map((p) => p.trim());
  const keyPart = parts[parts.length - 1];
  const needCtrl = parts.includes("Ctrl");
  const needAlt = parts.includes("Alt");
  const needShift = parts.includes("Shift");
  const needMeta = parts.includes("Meta");
  if (
    e.ctrlKey !== needCtrl ||
    e.altKey !== needAlt ||
    e.shiftKey !== needShift ||
    e.metaKey !== needMeta
  ) {
    return false;
  }
  return normalizeHotkeyKeyLabel(e.key) === keyPart;
}

/** @param {KeyboardEvent} e */
function shouldIgnorePromptPopupHotkeyTrigger(e) {
  if (capturingPromptPopupHotkey) {
    return true;
  }
  const t = e.target;
  if (!(t instanceof HTMLElement)) {
    return false;
  }
  if (t.isContentEditable) {
    return true;
  }
  const tag = t.tagName;
  if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") {
    if (t.id === "editPromptPopupHotkeyDisplay") {
      return false;
    }
    return true;
  }
  return false;
}

let lastPromptPopupHotkeyOpenAt = 0;
/** @type {string} */
let lastPromptPopupHotkeyOpenKind = "";

/** @param {string} promptId */
function triggerPromptPopupHotkeyForKind(promptId) {
  const kind = String(promptId || "").trim();
  if (!kind) {
    return;
  }
  const now = Date.now();
  if (kind === lastPromptPopupHotkeyOpenKind && now - lastPromptPopupHotkeyOpenAt < 450) {
    return;
  }
  lastPromptPopupHotkeyOpenKind = kind;
  lastPromptPopupHotkeyOpenAt = now;
  void openPromptWindowForKind(kind);
}

/** @param {string} combo @param {string} exceptKind */
function findPromptKindWithPopupHotkey(combo, exceptKind = "") {
  const target = String(combo || "").trim();
  if (!target) {
    return null;
  }
  for (const p of listAllMainPrompts()) {
    if (p.id === exceptKind) {
      continue;
    }
    if (getPromptPopupHotkey(p.id) === target) {
      return p.id;
    }
  }
  return null;
}

/** @param {string} promptId */
async function openPromptWindowForKind(promptId) {
  if (!isPromptKindAvailable(promptId)) {
    toast("That user was removed from Resume Builder.", "error", 4500);
    return;
  }
  const api = window.promptApi;
  if (!api || typeof api.openPromptWindow !== "function") {
    toast("New window is not available.", "error");
    return;
  }
  try {
    const r = await api.openPromptWindow(promptId);
    if (r?.ok && !r.alreadyOpen) {
      toast(`Opened ${promptLabelForKind(promptId)} in a new window.`, "success", 2500);
    }
  } catch (_e) {
    toast("Could not open a new window.", "error");
  }
}

/** @returns {Promise<{ failed: string[] } | null>} */
async function syncPromptPopupHotkeysToMain() {
  const api = window.promptApi;
  if (!api?.syncPromptPopupHotkeys) {
    return null;
  }
  const bindings = listAllMainPrompts()
    .map((p) => ({ kind: p.id, combo: getPromptPopupHotkey(p.id) }))
    .filter((b) => b.combo);
  try {
    const r = await api.syncPromptPopupHotkeys({ bindings });
    const failed = Array.isArray(r?.failed) ? r.failed : [];
    const shortcutCombos = Array.isArray(r?.shortcutCombos) ? r.shortcutCombos : [];
    const listenerCombos = Array.isArray(r?.listenerCombos) ? r.listenerCombos : [];
    return { failed, shortcutCombos, listenerCombos };
  } catch (_e) {
    return { failed: bindings.map((b) => b.combo), shortcutCombos: [], listenerCombos: [] };
  }
}

function refreshPromptPopupHotkeyTooltips() {
  document.querySelectorAll(".prompt-new-window-btn[data-prompt-new-window]").forEach((btn) => {
    const id = btn.getAttribute("data-prompt-new-window");
    if (!id) {
      return;
    }
    const label = getPromptUserName(id) || promptLabelForKind(id);
    const hotkey = getPromptPopupHotkey(id);
    const suffix = hotkey ? ` — ${hotkey}` : "";
    const title = `Open ${label} in a new window${suffix}`;
    btn.title = title;
    btn.setAttribute("aria-label", title);
  });
}

function refreshEditPromptPopupHotkeyDisplay() {
  if (!editPromptPopupHotkeyDisplay) {
    return;
  }
  const kind = getEditPromptSettingsKind();
  if (capturingPromptPopupHotkey) {
    editPromptPopupHotkeyDisplay.value = "";
    editPromptPopupHotkeyDisplay.placeholder = "Press keys…";
    editPromptPopupHotkeyDisplay.classList.add("is-capturing");
    return;
  }
  editPromptPopupHotkeyDisplay.classList.remove("is-capturing");
  const hotkey = kind ? getPromptPopupHotkey(kind) : "";
  editPromptPopupHotkeyDisplay.value = hotkey;
  editPromptPopupHotkeyDisplay.placeholder = hotkey ? "" : "Not set";
}

/** @param {string} kind */
function defaultPromptPathForKind(kind) {
  return kind === "justin" ? JUSTIN_ID : OTHER_ID;
}

/** @param {string | null | undefined} kind */
function getPromptPathForKind(kind) {
  const id = String(kind || "").trim();
  if (!id) {
    return "";
  }
  if (isCustomPromptKind(id)) {
    return getCustomPromptByKind(id)?.promptPath || "";
  }
  if (isBuiltinPromptKind(id)) {
    try {
      const stored = localStorage.getItem(promptPathStorageKey(id))?.trim();
      if (stored) {
        return stored;
      }
    } catch (_e) {
      /* ignore */
    }
    return defaultPromptPathForKind(id);
  }
  return "";
}

/** @param {string | null | undefined} kind */
function getPromptStructurePathForKind(kind) {
  const id = String(kind || "").trim();
  if (!id) {
    return "";
  }
  if (isCustomPromptKind(id)) {
    return getCustomPromptByKind(id)?.structurePath || "";
  }
  try {
    return localStorage.getItem(promptStructureStorageKey(id))?.trim() || "";
  } catch (_e) {
    return "";
  }
}

/** @param {string | null | undefined} kind */
function getStructureTemplatePathForKind(kind) {
  const path = getPromptStructurePathForKind(kind);
  return path || null;
}

/** @returns {string | null} */
function getEditPromptSettingsKind() {
  return editPromptSettingsKind;
}

/** @type {"olu" | "justin"} */
let activeGptResumeProfile = "olu";

/** @type {Record<"olu" | "justin", boolean>} */
const gptProfileBusy = { olu: false, justin: false };

/** Per dedicated browser tab — allows parallel GPT in multiple browser tabs. */
/** @type {Record<string, boolean>} */
const gptBrowserTabBusy = {};

/** Browser tab showed a successful Word download for the current GPT run. */
/** @type {Record<string, boolean>} */
const gptBrowserTabDownloadDone = {};

/**
 * @typedef {{ text: string, isError: boolean }} GptStoredResult
 */

/** @type {Record<"olu" | "justin", GptStoredResult | null>} */
const lastGptResults = { olu: null, justin: null };

/** @type {Record<string, GptStoredResult | null>} */
const lastGptResultsByBrowserTab = {};

/** Context for the open GPT result modal (pane that opened it). */
let gptResultModalContext = {
  profile: /** @type {"olu" | "justin"} */ ("olu"),
  browserTabId: /** @type {string | null} */ (null),
  downloadKey: /** @type {string | null} */ (null),
};

/** @param {"olu" | "justin"} profile */
function gptProfileLabel(profile) {
  return profile === "justin" ? "Justin" : "Oluwaseun";
}

/** @param {"olu" | "justin"} profile */
function isGptProfileBusy(profile) {
  return !!gptProfileBusy[profile];
}

/** @param {string} browserTabId */
function isGptBrowserTabBusy(browserTabId) {
  return !!gptBrowserTabBusy[browserTabId];
}

/**
 * @param {"olu" | "justin"} profile
 * @param {{ browserTabId?: string | null }} [opts]
 */
function isGptBusyForSendTarget(profile, opts = {}) {
  const tabId = opts.browserTabId ?? null;
  if (tabId) {
    return isGptBrowserTabBusy(tabId);
  }
  return isGptProfileBusy(profile);
}

/**
 * @param {"olu" | "justin"} profile
 * @param {boolean} busy
 * @param {{ browserTabId?: string | null, paneEl?: HTMLElement | null }} [opts]
 */
function setGptProfileBusy(profile, busy, opts = {}) {
  const tabId = opts.browserTabId ?? null;
  if (tabId) {
    gptBrowserTabBusy[tabId] = busy;
    if (!busy) {
      delete gptBrowserTabBusy[tabId];
    } else {
      delete gptBrowserTabDownloadDone[tabId];
    }
    const pane =
      opts.paneEl ??
      dedicatedBrowserTabs.find((t) => t.id === tabId)?.paneEl ??
      null;
    pane?.classList.toggle("is-generating", busy);
    updateDedicatedBrowserTabStatus(tabId);
  } else {
    gptProfileBusy[profile] = busy;
    if (dedicatedPromptKind) {
      document
        .querySelectorAll(`.chatgpt-pane[data-chatgpt-profile="${profile}"]`)
        .forEach((pane) => {
          pane.classList.toggle("is-generating", busy);
        });
    } else {
      const kind = getSelectedPromptKind();
      if (kind) {
        document
          .querySelector(
            `#chatgptPanesHost .chatgpt-pane[data-prompt="${CSS.escape(kind)}"]`,
          )
          ?.classList.toggle("is-generating", busy);
      }
    }
  }
  updatePreviewCopyUi();
  updateGptResultReadyUi();
  if (tabId) {
    const pane =
      opts.paneEl ??
      dedicatedBrowserTabs.find((t) => t.id === tabId)?.paneEl ??
      null;
    if (pane) {
      updateChatgptPaneDownloadStatusUi(/** @type {HTMLElement} */ (pane));
    }
  } else if (dedicatedPromptKind) {
    document.querySelectorAll(".chatgpt-pane").forEach((pane) => {
      updateChatgptPaneDownloadStatusUi(/** @type {HTMLElement} */ (pane));
    });
  } else {
    const kind = getSelectedPromptKind();
    if (kind) {
      const pane = document.querySelector(
        `#chatgptPanesHost .chatgpt-pane[data-prompt="${CSS.escape(kind)}"]`,
      );
      if (pane) {
        updateChatgptPaneDownloadStatusUi(/** @type {HTMLElement} */ (pane));
      }
    }
  }
}

const MAIN_CHATGPT_HIDDEN_KEY = "resume-builder.main-chatgpt-hidden";

const PROMPT_NEW_WINDOW_ICON_SVG = `<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M15 3h6v6"/><path d="M10 14 21 3"/><path d="M18 13v6a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h6"/></svg>`;

const PROMPT_GPT_PANE_EYE_ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/></svg>';

const PROMPT_GPT_PANE_EYE_OFF_ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M9.88 9.88a3 3 0 1 0 4.24 4.24"/><path d="M10.73 5.08A10.43 10.43 0 0 1 12 5c7 0 10 7 10 7a13.16 13.16 0 0 1-1.67 2.68"/><path d="M6.61 6.61A13.526 13.526 0 0 0 2 12s3 7 10 7a9.74 9.74 0 0 0 5.39-1.61"/><line x1="2" y1="2" x2="22" y2="22"/></svg>';

/** @param {boolean} shown */
function gptPaneToggleIconHtml(shown) {
  return shown ? PROMPT_GPT_PANE_EYE_OFF_ICON_SVG : PROMPT_GPT_PANE_EYE_ICON_SVG;
}

const CHATGPT_PANE_DOWNLOAD_ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>';

const CHATGPT_PANE_INPUT_ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m22 2-7 20-4-9-9-4Z"/><path d="M22 2 11 13"/></svg>';

const CHATGPT_PANE_FETCH_ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 12a9 9 0 0 0-9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/><path d="M21 12a9 9 0 0 1-9 9 9.75 9.75 0 0 1-6.74-2.74L21 16"/><path d="M16 16h5v5"/></svg>';

const CHATGPT_PANE_VIEW_ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M2 12s3-7 10-7 10 7 10 7-3 7-10 7-10-7-10-7Z"/><circle cx="12" cy="12" r="3"/></svg>';

const CHATGPT_PANE_FOLDER_OPEN_ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M20 20a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.9a2 2 0 0 1-1.69-.9L9.6 3.9A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13a2 2 0 0 0 2 2Z"/><path d="M2 10h20"/></svg>';

const CHATGPT_PANE_RESUME_FILE_ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/></svg>';

const CHATGPT_PANE_JD_FILE_ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="15" height="15" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="16" y1="13" x2="8" y2="13"/><line x1="16" y1="17" x2="8" y2="17"/><line x1="10" y1="9" x2="8" y2="9"/></svg>';

const CHATGPT_PANE_COPY_ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>';

const CHROME_TAB_BROWSER_ICON_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.75" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="10"/><path d="M2 12h20"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/></svg>';

/** @param {HTMLElement} pane */
function downloadKeyForPane(pane) {
  const promptId = pane.getAttribute("data-prompt");
  if (promptId) {
    return promptId;
  }
  if (pane.hasAttribute("data-browser-tab") && dedicatedPromptKind) {
    return dedicatedPromptKind;
  }
  return null;
}

/** @param {HTMLElement} pane */
function ensureChatgptPaneHeaderMeta(pane) {
  pane.querySelector(".chatgpt-pane-header .chatgpt-pane-header-meta")?.remove();
  pane.querySelector(".chatgpt-pane-word-dir-bar")?.remove();
}

/** @param {HTMLInputElement | null} input @param {string} path */
function syncPreviewDirInputDisplay(input, path) {
  if (!input) {
    return;
  }
  const text = String(path || "").trim();
  input.value = text;
  input.title = text || input.placeholder || "";
}

/** @param {string} kind */
async function loadMainPromptStructureDirectory(kind) {
  const path = getPromptStructurePathForKind(kind);
  syncPreviewDirInputDisplay(mainPromptStructureDisplay, path);
  syncPreviewDirInputDisplay(dedicatedPromptStructureDisplay, path);
  if (editPromptStructureDisplay && getEditPromptSettingsKind() === kind) {
    editPromptStructureDisplay.value = path;
    editPromptStructureDisplay.title = path;
  }
}

/** @param {string} kind */
async function loadMainPreviewDirectories(kind) {
  await loadMainPromptStructureDirectory(kind);
  if (dedicatedPromptKind === kind) {
    await loadDedicatedPromptStructureDirectory(kind);
  }
  await syncGlobalResumeSaveRootDisplays();
}

/** @param {string} _kind */
async function refreshChatgptPaneWordDirectoriesForKind(_kind) {
  const panes = document.querySelectorAll(
    "#chatgptPanesHost .chatgpt-pane, .dedicated-browser-host .chatgpt-pane",
  );
  for (const node of panes) {
    ensureChatgptPaneHeaderMeta(/** @type {HTMLElement} */ (node));
  }
}

/** @returns {Set<string>} */
function readMainChatgptHiddenSet() {
  try {
    const raw = localStorage.getItem(MAIN_CHATGPT_HIDDEN_KEY);
    const arr = raw ? JSON.parse(raw) : [];
    return new Set(Array.isArray(arr) ? arr.filter((x) => typeof x === "string") : []);
  } catch {
    return new Set();
  }
}

/** @param {Set<string>} set */
function writeMainChatgptHiddenSet(set) {
  try {
    localStorage.setItem(MAIN_CHATGPT_HIDDEN_KEY, JSON.stringify([...set]));
  } catch {
    /* ignore */
  }
}

/** @param {string} promptId */
function chatgptPartitionForPromptId(promptId) {
  if (promptId === "other") {
    return "persist:chatgpt";
  }
  if (promptId === "justin") {
    return "persist:chatgpt-justin";
  }
  if (isCustomPromptKind(promptId)) {
    const slug = promptId
      .slice(7)
      .replace(/[^a-zA-Z0-9_-]/g, "-")
      .slice(0, 48) || "custom";
    return `persist:chatgpt-${slug}`;
  }
  return "persist:chatgpt";
}

/** @param {string} promptId */
function webviewDomIdForPromptId(promptId) {
  if (promptId === "other") {
    return "chatgptWebview-olu";
  }
  if (promptId === "justin") {
    return "chatgptWebview-justin";
  }
  return `chatgptWebview-${promptId.replace(/:/g, "-")}`;
}

/** @param {string} promptId */
function promptLabelForKind(promptId) {
  if (promptId === "justin") {
    return "Justin";
  }
  if (promptId === "other") {
    return "Oluwaseun";
  }
  return getCustomPromptByKind(promptId)?.label || "Prompt";
}

/** @param {string} promptId */
function promptUserNameStorageKey(promptId) {
  return `${PROMPT_USER_NAME_STORAGE_PREFIX}${promptId}`;
}

/** @param {string} promptId */
function defaultPromptUserName(promptId) {
  return promptLabelForKind(promptId);
}

/** @param {string} promptId */
function getPromptUserName(promptId) {
  const id = String(promptId || "").trim();
  if (!id) {
    return "";
  }
  if (isCustomPromptKind(id)) {
    return getCustomPromptByKind(id)?.label || defaultPromptUserName(id);
  }
  try {
    const stored = localStorage.getItem(promptUserNameStorageKey(id))?.trim();
    if (stored) {
      return stored;
    }
  } catch (_e) {
    /* ignore */
  }
  return defaultPromptUserName(id);
}

/** @param {string} name */
function getUserInitials(name) {
  const parts = String(name || "")
    .trim()
    .split(/\s+/)
    .filter(Boolean);
  if (parts.length === 0) {
    return "?";
  }
  if (parts.length === 1) {
    return parts[0].slice(0, 2).toUpperCase();
  }
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

/** @param {string} name @param {boolean} [small] */
function userAvatarHtml(name, small) {
  const cls = small ? "user-avatar user-avatar--sm" : "user-avatar";
  return `<span class="${cls}" aria-hidden="true">${escapeHtml(getUserInitials(name))}</span>`;
}

/** @param {string} filePath */
function normalizePathForCompare(filePath) {
  return String(filePath || "")
    .trim()
    .replace(/\\/g, "/")
    .toLowerCase();
}

/** @param {string} filePath */
function resolveEditPromptKindForPath(filePath) {
  const norm = normalizePathForCompare(filePath);
  if (!norm) {
    return null;
  }
  const otherTail = OTHER_ID.replace(/\\/g, "/").toLowerCase();
  const justinTail = JUSTIN_ID.replace(/\\/g, "/").toLowerCase();
  if (norm.endsWith(otherTail) || norm.endsWith("/olu prompt.txt")) {
    return "other";
  }
  if (norm.endsWith(justinTail) || norm.endsWith("/justin prompt.txt")) {
    return "justin";
  }
  for (const id of ["other", "justin"]) {
    if (normalizePathForCompare(getPromptPathForKind(id)) === norm) {
      return id;
    }
  }
  for (const entry of customPrompts) {
    if (normalizePathForCompare(entry.promptPath) === norm) {
      return entry.id;
    }
  }
  return null;
}

/** @param {string | null | undefined} kind */
function isKnownEditPromptKind(kind) {
  return isBuiltinPromptKind(kind) || isCustomPromptKind(kind);
}

/** @returns {string} */
function promptUserNameForContext() {
  const kind = dedicatedPromptKind || getSelectedPromptKind();
  return kind ? getPromptUserName(kind) : "";
}

/** Display name for toasts and labels (Edit-tab user name, else prompt card name). */
function userDisplayNameForPromptKind(kind) {
  return getPromptUserName(kind);
}

/** @param {string} kind */
function dedicatedWindowTitleForKind(kind) {
  return (
    String(userDisplayNameForPromptKind(kind) || promptLabelForKind(kind) || "")
      .trim() || "Resume Builder"
  );
}

function syncAppVersionLabel() {
  const el = document.getElementById("appVersionLabel");
  const api = window.promptApi;
  if (!el || !api || typeof api.getAppVersion !== "function") {
    return;
  }
  void api.getAppVersion().then((v) => {
    const ver = String(v ?? "").trim();
    el.textContent = ver ? `v${ver}` : "";
    el.setAttribute("aria-label", ver ? `Version ${ver}` : "");
  });
}

/** @param {string} [kind] */
function syncDedicatedWindowTitle(kind) {
  const id = kind || dedicatedPromptKind;
  if (!id || !dedicatedPromptKind) {
    return;
  }
  const name = dedicatedWindowTitleForKind(id);
  document.title = name;
  const titleEl = document.querySelector(".titlebar-brand .titlebar-title");
  if (titleEl) {
    titleEl.textContent = name;
  }
  const api = typeof windowControls !== "undefined" ? windowControls : null;
  if (api && typeof api.setTitle === "function") {
    void api.setTitle(name);
  }
}

/**
 * @param {{ promptKind?: string | null, browserTabId?: string | null, profile?: "olu" | "justin" }} [opts]
 * @returns {string}
 */
function contextLabelForGptUi(opts = {}) {
  const browserTabId = opts.browserTabId ?? null;
  if (browserTabId) {
    return dedicatedBrowserTabLabel(browserTabId);
  }
  const kind =
    opts.promptKind ||
    (dedicatedPromptKind ? dedicatedPromptKind : null) ||
    getSelectedPromptKind();
  if (kind) {
    return userDisplayNameForPromptKind(kind);
  }
  if (opts.profile) {
    return gptProfileLabel(opts.profile);
  }
  return "";
}

/** @param {string} kind */
function syncPromptUserNameShellForInput(/** @type {HTMLInputElement | null} */ input) {
  if (!input) {
    return;
  }
  const shell = input.closest(".build-input-shell");
  if (!shell) {
    return;
  }
  const filled = !!input.value.trim();
  shell.classList.toggle("is-empty", !filled);
  shell.classList.toggle("is-filled", filled);
}

function loadEditPromptSettings() {
  const kind = getEditPromptSettingsKind();
  const fileUrlRow = document.querySelector(".edit-file-url-row");
  if (!editPromptSettings || dedicatedPromptKind) {
    return;
  }
  if (!isKnownEditPromptKind(kind)) {
    editPromptSettings.hidden = true;
    if (fileUrlRow) {
      fileUrlRow.hidden = listAllMainPrompts().length === 0;
    }
    syncEditLayoutEmptyState();
    return;
  }
  editPromptSettings.hidden = false;
  if (fileUrlRow) {
    fileUrlRow.hidden = true;
  }
  if (editPromptUserNameInput) {
    editPromptUserNameInput.value = getPromptUserName(kind);
    if (kind === "other") {
      editPromptUserNameInput.placeholder = "e.g. Oluwaseun Adeniji";
    } else if (kind === "justin") {
      editPromptUserNameInput.placeholder = "e.g. Justin Mcghee";
    } else {
      editPromptUserNameInput.placeholder = "Name on downloaded .docx file";
    }
  }
  if (editPromptPathDisplay) {
    const promptPath = getPromptPathForKind(kind);
    editPromptPathDisplay.value = promptPath;
    editPromptPathDisplay.title = promptPath;
  }
  if (editPromptStructureDisplay) {
    editPromptStructureDisplay.value = getPromptStructurePathForKind(kind);
    editPromptStructureDisplay.title = editPromptStructureDisplay.value;
  }
  if (kind) {
    void refreshChatgptPaneWordDirectoriesForKind(kind);
  }
  refreshEditPromptPopupHotkeyDisplay();
  if (editPromptDeleteField) {
    editPromptDeleteField.hidden = !kind;
  }
  syncEditUserHeader(kind);
  syncEditSavePathToday(kind);
  syncEditPanelHeading(kind);
  syncMainResumeSaveRootDisplay();
  syncEditLayoutEmptyState();
}

function syncEditPanelHeading(kind) {
  if (!editPanelTitle) {
    return;
  }
  editPanelTitle.textContent = "Prompt content";
  if (!kind) {
    editPanelTitle.title = "";
    return;
  }
  const path = getPromptPathForKind(kind);
  editPanelTitle.title = path ? fileNameFromPath(path) : "";
}

function syncEditUserHeader(kind) {
  const header = document.getElementById("editUserHeader");
  const avatar = document.getElementById("editUserHeaderAvatar");
  const nameEl = document.getElementById("editUserHeaderName");
  if (!(header instanceof HTMLElement)) {
    return;
  }
  if (!kind) {
    header.hidden = true;
    return;
  }
  const name = getPromptUserName(kind);
  header.hidden = false;
  if (avatar instanceof HTMLElement) {
    avatar.textContent = getUserInitials(name);
  }
  if (nameEl instanceof HTMLElement) {
    nameEl.textContent = name;
    nameEl.title = name;
  }
}

/** @param {string} [kind] */
function buildTodaySavePathPreview(kind) {
  const id = kind || getEditPromptSettingsKind();
  if (!id) {
    return "";
  }
  const mainEl = document.getElementById("mainResumeSaveRootDisplay");
  const editEl = document.getElementById("editResumeSaveRootDisplay");
  const root = String(
    (mainEl instanceof HTMLInputElement ? mainEl.value : "") ||
      (editEl instanceof HTMLInputElement ? editEl.value : ""),
  ).trim();
  if (!root) {
    return "";
  }
  const now = new Date();
  const date = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-${String(now.getDate()).padStart(2, "0")}`;
  const company = "{company}";
  const sep = root.includes("\\") ? "\\" : "/";
  return `${root}${sep}${date}${sep}${company}${sep}{profile name}.docx`;
}

/** @param {string} [kind] */
function syncEditSavePathToday(kind) {
  const el = document.getElementById("editPromptSavePathToday");
  if (!(el instanceof HTMLInputElement)) {
    return;
  }
  const preview = buildTodaySavePathPreview(kind);
  el.value = preview;
  el.title = preview;
}

async function getGlobalResumeSaveRootPath() {
  const api = window.promptApi;
  if (!api || typeof api.getDownloadDirectory !== "function") {
    return "";
  }
  try {
    const r = await api.getDownloadDirectory("global");
    return String(r?.directory || "").trim();
  } catch (_e) {
    return "";
  }
}

async function syncGlobalResumeSaveRootDisplays() {
  const root = await getGlobalResumeSaveRootPath();
  const el = document.getElementById("mainResumeSaveRootDisplay");
  const editEl = document.getElementById("editResumeSaveRootDisplay");
  if (el instanceof HTMLInputElement) {
    el.value = root;
    el.title = root;
    el.placeholder = root ? "" : "Choose global resume save folder…";
  }
  if (editEl instanceof HTMLInputElement) {
    editEl.value = root;
    editEl.title = root;
    editEl.placeholder = root ? "" : "Choose global resume save folder…";
  }
  syncEditSavePathToday(getEditPromptSettingsKind() || undefined);
}

function syncMainResumeSaveRootDisplay() {
  void syncGlobalResumeSaveRootDisplays();
}

/** @param {string} kind */
function clearPromptLocalSettingsForKind(kind) {
  const id = String(kind || "").trim();
  if (!id) {
    return;
  }
  try {
    localStorage.removeItem(promptUserNameStorageKey(id));
    localStorage.removeItem(promptStructureStorageKey(id));
    localStorage.removeItem(promptPathStorageKey(id));
    localStorage.removeItem(promptPopupHotkeyStorageKey(id));
    localStorage.removeItem(`${DEDICATED_COMPANY_STORAGE_PREFIX}${id}`);
    localStorage.removeItem(`${MAIN_COMPANY_STORAGE_PREFIX}${id}`);
  } catch (_e) {
    /* ignore */
  }
  delete customTemplateStrById[id];
}

/** @param {string} kind */
async function closePromptWindowForKind(kind) {
  const id = String(kind || "").trim();
  if (!id) {
    return;
  }
  try {
    if (typeof window.promptApi?.closePromptWindow === "function") {
      await window.promptApi.closePromptWindow(id);
    }
  } catch (_e) {
    /* ignore */
  }
}

async function openFirstAvailablePromptInEditor() {
  const next = listAllMainPrompts()[0];
  if (!next) {
    editPromptSettingsKind = null;
    clearEditFileNavActive();
    if (editTemplate) {
      editTemplate.value = "";
    }
    if (editPanelTitle) {
      editPanelTitle.textContent = "No users";
      editPanelTitle.title = "";
    }
    loadEditPromptSettings();
    return;
  }
  if (isBuiltinPromptKind(next.id)) {
    await openEditKey(/** @type {"other" | "justin"} */ (next.id));
  } else if (isCustomPromptKind(next.id)) {
    await openEditCustomPrompt(next.id);
  }
}

function syncEditLayoutEmptyState() {
  const layout = getEditLayoutElement();
  if (!(layout instanceof HTMLElement)) {
    return;
  }
  const empty = listAllMainPrompts().length === 0;
  layout.classList.toggle("is-empty", empty);
  const panel = layout.querySelector(".edit-editor-panel");
  if (panel instanceof HTMLElement) {
    panel.hidden = empty;
    if (empty) {
      panel.setAttribute("aria-hidden", "true");
    } else {
      panel.removeAttribute("aria-hidden");
    }
  }
  const emptyHint = document.getElementById("editNavEmptyHint");
  if (emptyHint instanceof HTMLElement) {
    emptyHint.hidden = !empty;
  }
  const manageDialog = document.querySelector(
    "#usersManageModal .users-manage-dialog",
  );
  if (manageDialog instanceof HTMLElement) {
    manageDialog.classList.toggle("users-manage-dialog--empty", empty);
  }
}

function refreshEditUsersNav() {
  applyBuiltinPromptVisibilityToUi();
  renderEditPromptNav();
  refreshPromptDisplayNamesAcrossUi();
  syncEditLayoutEmptyState();
}

/** @param {"other" | "justin"} kind */
async function deleteBuiltinPromptForKind(kind) {
  if (!isBuiltinPromptKind(kind)) {
    return false;
  }
  const name = getPromptUserName(kind) || (kind === "justin" ? "Justin" : "Oluwaseun");
  const ok = window.confirm(
    `Delete "${name}" from Resume Builder?\n\nThis removes the user from the app. Your .txt and .docx files on disk are not deleted.`,
  );
  if (!ok) {
    return false;
  }
  hideBuiltinPromptKind(kind);
  clearPromptLocalSettingsForKind(kind);
  const hiddenMain = readMainChatgptHiddenSet();
  if (hiddenMain.has(kind)) {
    hiddenMain.delete(kind);
    writeMainChatgptHiddenSet(hiddenMain);
  }
  await closePromptWindowForKind(kind);
  refreshEditUsersNav();
  syncMainWindowChatgptPanes();
  renderChatgptShowPanesBar();

  const highlighted = document.querySelector(
    `.prompt-card.is-highlighted[data-prompt="${CSS.escape(kind)}"]`,
  );
  if (highlighted) {
    highlighted.classList.remove("is-highlighted");
    selectDefaultPromptCard();
  }

  await openFirstAvailablePromptInEditor();
  void syncPromptPopupHotkeysToMain();
  refreshBuildPreviewIfCardSelected();
  updateGptResultReadyUi();
  toast(`Deleted ${name}.`, "success", 3500);
  return true;
}

/** @param {string} kind */
async function deletePromptForKind(kind) {
  if (isCustomPromptKind(kind)) {
    return deleteCustomPromptForKind(kind);
  }
  if (isBuiltinPromptKind(kind)) {
    return deleteBuiltinPromptForKind(kind);
  }
  return false;
}

/** @param {string} kind */
async function deleteCustomPromptForKind(kind) {
  if (!isCustomPromptKind(kind)) {
    return false;
  }
  const entry = getCustomPromptByKind(kind);
  const name = entry?.label || getPromptUserName(kind) || "this user";
  const ok = window.confirm(
    `Delete "${name}" from Resume Builder?\n\nThis removes the user from the app. Your .txt and .docx files on disk are not deleted.`,
  );
  if (!ok) {
    return false;
  }
  const api = window.promptApi;
  if (!api?.deleteCustomPrompt) {
    toast("Delete is not available.", "error");
    return false;
  }
  try {
    const r = await api.deleteCustomPrompt(kind);
    if (!r?.ok) {
      toast(String(r?.error || "Could not delete user."), "error");
      return false;
    }
    clearPromptLocalSettingsForKind(kind);
    if (editPromptSettingsKind === kind) {
      editPromptSettingsKind = null;
    }
    closeAddPromptModal();
    await refreshCustomPrompts();
    renderBuildPromptCards();
    void syncPromptPopupHotkeysToMain();

    const highlighted = document.querySelector(
      `.prompt-card.is-highlighted[data-prompt="${CSS.escape(kind)}"]`,
    );
    if (highlighted) {
      highlighted.classList.remove("is-highlighted");
      selectDefaultPromptCard();
    }

    await openFirstAvailablePromptInEditor();
    refreshBuildPreviewIfCardSelected();
    updateGptResultReadyUi();
    toast(`Deleted ${name}.`, "success", 3500);
    return true;
  } catch (_e) {
    toast("Could not delete user.", "error");
    return false;
  }
}

/** @param {unknown} r */
function isPromptReadSuccess(r) {
  return !!(r && typeof r === "object" && typeof r.content === "string");
}

/** @param {string | null | undefined} promptPath */
async function tryReadPromptFileAtPath(promptPath) {
  const p = String(promptPath || "").trim();
  if (!p || !window.promptApi?.readPromptFilePath) {
    return null;
  }
  try {
    const r = await window.promptApi.readPromptFilePath(p);
    if (!r || typeof r !== "object") {
      return null;
    }
    if (r.ok === false) {
      return r;
    }
    if (typeof r.content === "string") {
      return { content: r.content, path: r.path, fileUrl: r.fileUrl };
    }
    return null;
  } catch (_e) {
    return null;
  }
}

/** @returns {HTMLTextAreaElement | null} */
function getEditTemplateEl() {
  return document.getElementById("editTemplate");
}

/** @param {string} text @param {{ loading?: boolean }} [opts] */
function setPromptEditorContent(text, opts = {}) {
  const el = getEditTemplateEl();
  if (!el) {
    return;
  }
  el.value = text;
  el.placeholder = opts.loading
    ? "Loading prompt…"
    : "Choose a prompt file or click Reload after fixing the path.";
}

/** @param {string} content */
function applyPromptEditorContent(content) {
  setPromptEditorContent(content);
}

/** @param {string} kind @param {string} path */
function promptLoadFailureMessage(kind, path) {
  const displayPath = String(path || getPromptPathForKind(kind) || "").trim();
  const lines = [
    "Could not load prompt file.",
    displayPath ? `\nPath:\n${displayPath}` : "",
    "\nUse the folder button next to Prompt file to choose a .txt file, then click Reload.",
  ];
  return lines.join("");
}

/**
 * Load prompt text into the editor for a user.
 * @param {string} kind
 * @returns {Promise<boolean>}
 */
async function loadPromptIntoEditor(kind) {
  if (!isBuiltinPromptKind(kind) && !isCustomPromptKind(kind)) {
    return false;
  }
  setPromptEditorContent("", { loading: true });
  try {
    editPromptSettingsKind = kind;
    if (isBuiltinPromptKind(kind)) {
      editSelection = /** @type {"other" | "justin"} */ (kind);
    }
    const promptPath = getPromptPathForKind(kind);
    const r = await tryReadPromptFileAtPath(promptPath);
    if (isPromptReadSuccess(r)) {
      if (isBuiltinPromptKind(kind)) {
        setStrForEdit(kind, r.content);
      } else if (isCustomPromptKind(kind)) {
        customTemplateStrById[kind] = r.content;
      }
      applyExternalEditor(r.path, r.fileUrl || "", r.content, { preserveSettingsKind: true });
      syncEditPanelHeading(kind);
      setEditFileNavActive(kind);
      loadEditPromptSettings();
      refreshPromptDisplayNamesAcrossUi();
      return true;
    }
    if (isCustomPromptKind(kind)) {
      delete customTemplateStrById[kind];
    }
    const failPath = (r && typeof r === "object" && r.path) || promptPath;
    editMode = "external";
    externalAbsPath = failPath || "";
    if (editFileUrlInput) {
      editFileUrlInput.value = "";
    }
    applyPromptEditorContent(promptLoadFailureMessage(kind, failPath));
    syncEditPanelHeading(kind);
    if (editPanelTitle && failPath) {
      editPanelTitle.title = fileNameFromPath(failPath);
    }
    setEditFileNavActive(kind);
    loadEditPromptSettings();
    refreshPromptDisplayNamesAcrossUi();
    toast(
      failPath
        ? "Prompt file not found. Use the folder button next to Prompt file to choose a .txt file."
        : "No prompt file set. Use the folder button next to Prompt file to choose a .txt file.",
      "error",
      6500,
    );
    return false;
  } catch (_e) {
    applyPromptEditorContent(promptLoadFailureMessage(kind, getPromptPathForKind(kind)));
    return false;
  }
}

/** @param {string} kind @param {{ silent?: boolean }} [opts] */
async function reloadPromptTemplateForKind(kind, opts = {}) {
  const path = getPromptPathForKind(kind);
  if (!path) {
    if (!opts.silent) {
      toast("No prompt file set.", "error");
    }
    return;
  }
  try {
    const r = await tryReadPromptFileAtPath(path);
    if (!isPromptReadSuccess(r)) {
      throw new Error("missing");
    }
    if (isBuiltinPromptKind(kind)) {
      setStrForEdit(kind, r.content);
    } else if (isCustomPromptKind(kind)) {
      customTemplateStrById[kind] = r.content;
    }
    if (getEditPromptSettingsKind() === kind && editTemplate) {
      editTemplate.value = r.content;
      editMode = "external";
      externalAbsPath = r.path;
      if (editFileUrlInput) {
        editFileUrlInput.value = r.fileUrl;
      }
    }
    refreshBuildPreviewIfCardSelected();
  } catch (_e) {
    if (!opts.silent) {
      toast("Could not load prompt file.", "error");
    }
  }
}

/** @param {string} kind @param {string} promptPath */
function setPromptPathForKindInMemory(kind, promptPath) {
  const normalized = String(promptPath || "").trim();
  if (!normalized) {
    return;
  }
  if (isCustomPromptKind(kind)) {
    const entry = getCustomPromptByKind(kind);
    if (entry) {
      entry.promptPath = normalized;
    }
    return;
  }
  if (isBuiltinPromptKind(kind)) {
    try {
      localStorage.setItem(promptPathStorageKey(kind), normalized);
    } catch (_e) {
      /* ignore */
    }
  }
}

/** @param {string} kind @param {string} promptPath */
function syncPromptPathDisplay(kind, promptPath) {
  const normalized = String(promptPath || getPromptPathForKind(kind) || "").trim();
  if (editPromptPathDisplay) {
    editPromptPathDisplay.value = normalized;
    editPromptPathDisplay.title = normalized;
  }
  if (kind) {
    syncEditPanelHeading(kind);
  }
}

/** @param {string} kind @param {{ path: string, fileUrl?: string, content: string }} pick */
function applyPromptFilePickForKind(kind, pick) {
  const path = String(pick.path || "").trim();
  const content = String(pick.content ?? "");
  if (!path) {
    return;
  }
  setPromptPathForKindInMemory(kind, path);
  syncPromptPathDisplay(kind, path);
  if (isCustomPromptKind(kind)) {
    customTemplateStrById[kind] = content;
  } else if (isBuiltinPromptKind(kind)) {
    setStrForEdit(kind, content);
  }
  applyExternalEditor(path, pick.fileUrl || "", content, { preserveSettingsKind: true });
  syncEditPanelHeading(kind);
  setEditFileNavActive(kind);
  loadEditPromptSettings();
  refreshBuildPreviewIfCardSelected();
  refreshPromptDisplayNamesAcrossUi();
}

/** @param {string} kind @param {string} promptPath @returns {Promise<{ ok: boolean, error?: string }>} */
async function persistPromptPathForKind(kind, promptPath) {
  const id = String(kind || "").trim();
  const normalized = String(promptPath || "").trim();
  if (!id || !normalized) {
    return { ok: false, error: "No prompt path provided." };
  }
  if (isBuiltinPromptKind(id)) {
    try {
      localStorage.setItem(promptPathStorageKey(id), normalized);
    } catch (_e) {
      return { ok: false, error: "Could not save prompt path." };
    }
    await reloadPromptTemplateForKind(id);
    return { ok: true };
  }
  const entry = getCustomPromptByKind(id);
  const api = window.promptApi;
  if (!entry || !api?.updateCustomPromptPath) {
    return { ok: false, error: "Could not save prompt path." };
  }
  try {
    const r = await api.updateCustomPromptPath({
      id: entry.id,
      promptPath: normalized,
    });
    if (!r?.ok) {
      return { ok: false, error: String(r?.error || "Could not save prompt path.") };
    }
    if (r.prompt?.promptPath) {
      setPromptPathForKindInMemory(id, r.prompt.promptPath);
    }
    await refreshCustomPrompts();
    await reloadPromptTemplateForKind(id, { silent: true });
    return { ok: true };
  } catch (_e) {
    return { ok: false, error: "Could not save prompt path." };
  }
}

/** @param {string} kind */
async function openPromptEditorForKind(kind) {
  if (isBuiltinPromptKind(kind)) {
    await openEditKey(/** @type {"other" | "justin"} */ (kind));
  } else if (isCustomPromptKind(kind)) {
    await openEditCustomPrompt(kind);
  }
}

/** @param {string} kind @param {string} structurePath */
async function persistPromptStructureForKind(kind, structurePath) {
  const id = String(kind || "").trim();
  const normalized = String(structurePath || "").trim();
  if (!id) {
    return;
  }
  if (isBuiltinPromptKind(id)) {
    try {
      if (normalized) {
        localStorage.setItem(promptStructureStorageKey(id), normalized);
      } else {
        localStorage.removeItem(promptStructureStorageKey(id));
      }
    } catch (_e) {
      /* ignore */
    }
    return;
  }
  const entry = getCustomPromptByKind(id);
  const api = window.promptApi;
  if (!entry || !api?.saveCustomPrompt) {
    return;
  }
  try {
    const r = await api.saveCustomPrompt({
      id: entry.id,
      label: entry.label,
      promptPath: entry.promptPath,
      structurePath: normalized,
      chatgptProfile: entry.chatgptProfile,
    });
    if (r?.ok) {
      await refreshCustomPrompts();
    }
  } catch (_e) {
    /* ignore */
  }
}

/** @type {ReturnType<typeof setTimeout> | null} */
let customPromptLabelPersistTimer = null;

function persistEditPromptSettingsFromInputs() {
  const kind = getEditPromptSettingsKind();
  if (!kind) {
    return;
  }
  if (editPromptUserNameInput) {
    void persistPromptDisplayNameForKind(kind, editPromptUserNameInput.value);
  }
}

/** @param {string} kind @param {string} raw */
async function persistPromptDisplayNameForKind(kind, raw) {
  if (!kind) {
    return;
  }
  const val = String(raw ?? "").trim();
  if (isBuiltinPromptKind(kind)) {
    persistPromptUserNameForKind(kind, val);
    return;
  }
  const entry = getCustomPromptByKind(kind);
  const api = window.promptApi;
  if (!entry || !api?.saveCustomPrompt) {
    return;
  }
  if (!val || val === entry.label) {
    return;
  }
  if (customPromptLabelPersistTimer) {
    clearTimeout(customPromptLabelPersistTimer);
  }
  customPromptLabelPersistTimer = setTimeout(async () => {
    customPromptLabelPersistTimer = null;
    try {
      const r = await api.saveCustomPrompt({
        id: entry.id,
        label: val,
        promptPath: entry.promptPath,
        structurePath: entry.structurePath,
        chatgptProfile: entry.chatgptProfile,
      });
      if (!r?.ok) {
        return;
      }
      await refreshCustomPrompts();
      refreshBuildPreviewIfCardSelected();
      if (getEditPromptSettingsKind() === kind) {
        if (editPanelTitle) {
          editPanelTitle.innerHTML = `<code>${escapeHtml(val)}</code>`;
        }
        if (editPromptUserNameInput) {
          editPromptUserNameInput.value = val;
        }
      }
    } catch (_e) {
      /* ignore */
    }
  }, 400);
}

/** @param {string} kind @param {string} raw */
function persistPromptUserNameForKind(kind, raw) {
  if (!kind) {
    return;
  }
  const val = String(raw ?? "").trim();
  try {
    if (val && val !== defaultPromptUserName(kind)) {
      localStorage.setItem(promptUserNameStorageKey(kind), val);
    } else {
      localStorage.removeItem(promptUserNameStorageKey(kind));
    }
  } catch (_e) {
    /* ignore */
  }
  refreshPromptDisplayNamesAcrossUi();
}

/** Sync user-facing names after Edit-tab name changes. */
function refreshPromptDisplayNamesAcrossUi() {
  document.querySelectorAll(".edit-file-btn[data-edit]").forEach((btn) => {
    const kind = btn.getAttribute("data-edit");
    if (!kind) {
      return;
    }
    const labelEl = btn.querySelector(".edit-file-label");
    const hotkeyEl = btn.querySelector(".edit-file-hotkey");
    const avatarEl = btn.querySelector(".user-avatar");
    if (isBuiltinPromptKind(kind) || isCustomPromptKind(kind)) {
      const name = getPromptUserName(kind);
      if (labelEl) {
        labelEl.textContent = name;
      }
      if (avatarEl) {
        avatarEl.textContent = getUserInitials(name);
      }
      const hotkey = getPromptPopupHotkey(kind);
      if (hotkeyEl) {
        hotkeyEl.textContent = hotkey;
        hotkeyEl.hidden = !hotkey;
      } else if (hotkey) {
        const textWrap = btn.querySelector(".edit-file-btn-text");
        if (textWrap) {
          const span = document.createElement("span");
          span.className = "edit-file-hotkey";
          span.textContent = hotkey;
          textWrap.appendChild(span);
        }
      }
      btn.title = "Edit prompt and settings";
    }
  });

  if (!dedicatedPromptKind && buildPromptCardsEl) {
    renderBuildPromptCards();
  }

  for (const kind of ["other", "justin"]) {
    const name = getPromptUserName(kind);
    const card = document.querySelector(
      `.prompt-card[data-prompt="${CSS.escape(kind)}"] .prompt-card-topic`,
    );
    if (card) {
      card.textContent = name;
    }
    const heading = document.getElementById(
      kind === "other" ? "other-heading" : "justin-heading",
    );
    if (heading) {
      heading.textContent = name;
    }
    const pane = document.querySelector(
      `#chatgptPanesHost .chatgpt-pane[data-prompt="${CSS.escape(kind)}"]`,
    );
    if (pane) {
      const titleEl = pane.querySelector(".chatgpt-pane-title");
      if (titleEl) {
        titleEl.textContent = name;
      }
      pane.setAttribute("aria-label", `ChatGPT — ${name}`);
    }
  }

  renderChatgptShowPanesBar();
  refreshPromptPopupHotkeyTooltips();
  if (dedicatedPromptKind) {
    syncDedicatedWindowTitle(dedicatedPromptKind);
  }
}

function persistPromptUserNameFromInputs() {
  persistEditPromptSettingsFromInputs();
}

/** @returns {{ id: string, label: string }[]} */
function listAllMainPrompts() {
  const hiddenBuiltin = readHiddenBuiltinPromptSet();
  /** @type {{ id: string, label: string }[]} */
  const prompts = [];
  if (!hiddenBuiltin.has("other")) {
    prompts.push({ id: "other", label: getPromptUserName("other") });
  }
  if (!hiddenBuiltin.has("justin")) {
    prompts.push({ id: "justin", label: getPromptUserName("justin") });
  }
  const seenIds = new Set(prompts.map((p) => p.id));
  for (const p of customPrompts) {
    if (!p?.id || seenIds.has(p.id)) {
      continue;
    }
    seenIds.add(p.id);
    prompts.push({ id: p.id, label: getPromptUserName(p.id) });
  }
  return prompts;
}

function applyBuiltinPromptVisibilityToUi() {
  const hidden = readHiddenBuiltinPromptSet();
  for (const kind of /** @type {const} */ (["other", "justin"])) {
    const isHidden = hidden.has(kind);
    document
      .querySelectorAll(`.prompt-card[data-prompt="${CSS.escape(kind)}"]`)
      .forEach((node) => {
        /** @type {HTMLElement} */ (node).hidden = isHidden;
      });
    document
      .querySelectorAll(`.edit-nav .edit-file-btn[data-edit="${CSS.escape(kind)}"]`)
      .forEach((node) => {
        /** @type {HTMLElement} */ (node).hidden = isHidden;
      });
  }
}

/** @param {string} label */
function promptNewWindowButtonHtml(promptId, label) {
  const safeLabel = escapeHtml(label);
  const safeId = escapeHtml(promptId);
  const hotkey = getPromptPopupHotkey(promptId);
  const suffix = hotkey ? ` — ${escapeHtml(hotkey)}` : "";
  const title = `Open ${safeLabel} in a new window${suffix}`;
  return `<button type="button" class="prompt-new-window-btn" data-prompt-new-window="${safeId}" title="${title}" aria-label="${title}">${PROMPT_NEW_WINDOW_ICON_SVG}</button>`;
}

/** @param {string} promptId @param {string} label */
function promptGptPaneToggleButtonHtml(promptId, label) {
  const safeLabel = escapeHtml(label);
  const safeId = escapeHtml(promptId);
  return `<button type="button" class="prompt-gpt-pane-toggle btn-icon" data-prompt-gpt-toggle="${safeId}" aria-pressed="true" title="Hide ${safeLabel} ChatGPT panel" aria-label="Hide ${safeLabel} ChatGPT panel">${PROMPT_GPT_PANE_EYE_OFF_ICON_SVG}</button>`;
}

/** @param {string} promptId @param {string} label */
function promptCardActionsHtml(promptId, label) {
  return `${promptGptPaneToggleButtonHtml(promptId, label)}${promptNewWindowButtonHtml(promptId, label)}`;
}

/** @param {{ id: string, label: string }} prompt */
function buildPromptCardElement(prompt) {
  const section = document.createElement("section");
  section.className = "panel prompt-card";
  section.tabIndex = 0;
  section.dataset.prompt = prompt.id;
  const topicId = isBuiltinPromptKind(prompt.id)
    ? prompt.id === "justin"
      ? "justin-heading"
      : "other-heading"
    : "";
  const headingAttr = topicId ? ` id="${topicId}"` : "";
  section.innerHTML = `<div class="prompt-card-body"><div class="prompt-card-text">${userAvatarHtml(prompt.label)}<div class="prompt-card-name-wrap"><h3 class="prompt-card-topic"${headingAttr}>${escapeHtml(prompt.label)}</h3><p class="prompt-card-hotkey">${escapeHtml(getPromptPopupHotkey(prompt.id) || "No hotkey")}</p></div></div><div class="prompt-card-actions">${promptCardActionsHtml(prompt.id, prompt.label)}</div></div>`;
  return section;
}

function renderBuildPromptCards() {
  if (!buildPromptCardsEl || dedicatedPromptKind) {
    return;
  }
  const prompts = listAllMainPrompts();
  const highlighted = getSelectedPromptKind();
  buildPromptCardsEl.textContent = "";
  for (const p of prompts) {
    const card = buildPromptCardElement(p);
    if (highlighted === p.id || (!highlighted && prompts[0]?.id === p.id)) {
      card.classList.add("is-highlighted");
    }
    buildPromptCardsEl.appendChild(card);
  }
  if (!buildPromptCardsEl.querySelector(".prompt-card.is-highlighted")) {
    buildPromptCardsEl.querySelector(".prompt-card")?.classList.add("is-highlighted");
  }
  syncPromptCardGptPaneToggles();
  reorderPromptCardsByGptVisibility();
}

function gptPaneVisibilitySortKey(promptId, hidden, prompts) {
  const idx = prompts.findIndex((p) => p.id === promptId);
  const order = idx >= 0 ? idx : 999;
  return (hidden.has(promptId) ? 1000 : 0) + order;
}

/** Reorder prompt cards: shown ChatGPT panels first (top), hidden below (all users together). */
function reorderPromptCardsByGptVisibility() {
  if (dedicatedPromptKind) {
    return;
  }
  const hidden = readMainChatgptHiddenSet();
  const prompts = listAllMainPrompts();
  const host = buildPromptCardsEl;
  if (!host) {
    return;
  }

  const allCards = [...host.querySelectorAll(".prompt-card[data-prompt]")];
  allCards.sort(
    (a, b) =>
      gptPaneVisibilitySortKey(
        a.getAttribute("data-prompt") || "",
        hidden,
        prompts,
      ) -
      gptPaneVisibilitySortKey(
        b.getAttribute("data-prompt") || "",
        hidden,
        prompts,
      ),
  );

  for (const card of allCards) {
    host.appendChild(card);
  }
}

/** Reorder ChatGPT panes: shown panels first (left), hidden to the right. */
function reorderChatgptPanesInHost() {
  const host = document.getElementById("chatgptPanesHost");
  if (!host || dedicatedPromptKind) {
    return;
  }
  const hidden = readMainChatgptHiddenSet();
  const prompts = listAllMainPrompts();
  const panes = [...host.querySelectorAll(".chatgpt-pane[data-prompt]")];
  panes.sort(
    (a, b) =>
      gptPaneVisibilitySortKey(
        a.getAttribute("data-prompt") || "",
        hidden,
        prompts,
      ) -
      gptPaneVisibilitySortKey(
        b.getAttribute("data-prompt") || "",
        hidden,
        prompts,
      ),
  );
  for (const pane of panes) {
    host.appendChild(pane);
  }
}

function syncPromptCardGptPaneToggles() {
  const hidden = readMainChatgptHiddenSet();
  document
    .querySelectorAll(".prompt-gpt-pane-toggle[data-prompt-gpt-toggle]")
    .forEach((node) => {
      const btn = /** @type {HTMLButtonElement} */ (node);
      const id = btn.getAttribute("data-prompt-gpt-toggle") || "";
      const shown = !hidden.has(id);
      btn.setAttribute("aria-pressed", shown ? "true" : "false");
      const name = userDisplayNameForPromptKind(id) || id;
      const title = shown
        ? `Hide ${name} ChatGPT panel`
        : `Show ${name} ChatGPT panel`;
      btn.title = title;
      btn.setAttribute("aria-label", title);
      btn.innerHTML = gptPaneToggleIconHtml(shown);
    });
  document.querySelectorAll(".prompt-card[data-prompt]").forEach((node) => {
    const card = /** @type {HTMLElement} */ (node);
    const id = card.getAttribute("data-prompt") || "";
    const shown = !hidden.has(id);
    card.classList.toggle("prompt-card--gpt-shown", shown);
    card.classList.toggle("prompt-card--gpt-hidden", !shown);
  });
  reorderPromptCardsByGptVisibility();
  reorderChatgptPanesInHost();
}

const CHATGPT_HOME_URL = "https://chatgpt.com/";
/** OpenAI login page — pick “Continue with email”; avoid Google (passkeys fail in webview). */
const CHATGPT_EMAIL_LOGIN_URL = "https://auth.openai.com/log-in";

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
    if (
      host === "google.com" ||
      host.endsWith(".google.com") ||
      host.endsWith(".googleusercontent.com")
    ) {
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

/** @param {HTMLElement} wv @param {string} url */
function navigateChatgptWebviewTo(wv, url) {
  if (!wv || !url) {
    return;
  }
  const current = (wv.getAttribute("src") || "").trim();
  if (current === url) {
    return;
  }
  wv.dataset.pbNavTarget = url;
  wv.dataset.pbIntentionalNav = "1";
  try {
    if (typeof wv.loadURL === "function") {
      wv.loadURL(url);
    } else {
      wv.setAttribute("src", url);
    }
  } catch (_e) {
    wv.setAttribute("src", url);
  }
  setTimeout(() => {
    delete wv.dataset.pbIntentionalNav;
  }, 800);
}

/** @param {HTMLElement} wv */
function setupChatgptAuthNavigationGuard(wv) {
  if (wv.dataset.pbAuthGuard === "1") {
    return;
  }
  wv.dataset.pbAuthGuard = "1";

  let blockToastTimer = null;

  const redirectFromBlockedAuth = () => {
    if (wv.dataset.pbAuthRedirecting === "1") {
      return;
    }
    wv.dataset.pbAuthRedirecting = "1";
    navigateChatgptWebviewToEmailLogin(wv);
    if (!blockToastTimer) {
      blockToastTimer = setTimeout(() => {
        blockToastTimer = null;
        toast(
          "Google sign-in is blocked in the app. Use Email login, then email and password.",
          "error",
          9000,
        );
      }, 300);
    }
    setTimeout(() => {
      delete wv.dataset.pbAuthRedirecting;
    }, 1200);
  };

  wv.addEventListener("will-navigate", (e) => {
    if (isBlockedFederatedAuthUrl(e.url)) {
      e.preventDefault();
      redirectFromBlockedAuth();
    }
  });

  wv.addEventListener("did-navigate", (e) => {
    if (!isBlockedFederatedAuthUrl(e.url)) {
      return;
    }
    // Stay in this pane: no separate login window.
    redirectFromBlockedAuth();
  });

  wv.addEventListener("new-window", (e) => {
    if (isBlockedFederatedAuthUrl(e.url)) {
      e.preventDefault();
      toast(
        "Opened Google sign-in was blocked. Use Email login in this pane instead.",
        "error",
        7000,
      );
    }
  });
}

/** @param {HTMLElement} wv */
function bindChatgptWebviewChrome(wv) {
  if (wv.dataset.pbChatgptBound === "1") {
    return;
  }
  wv.dataset.pbChatgptBound = "1";
  const pre = getChatgptWebviewPreloadPath();
  if (pre && !wv.getAttribute("preload")) {
    wv.setAttribute("preload", pre);
  }
  setupChatgptWebviewLoading(wv);
  setupChatgptZoomForWebview(wv);
  setupChatgptAuthNavigationGuard(wv);
}

/**
 * @param {HTMLElement} wv
 * @param {{ force?: boolean }} [opts]
 */
function loadChatgptWebviewIfNeeded(wv, opts = {}) {
  if (!wv) {
    return;
  }
  bindChatgptWebviewChrome(wv);
  const src = (wv.getAttribute("src") || "").trim();
  if (!opts.force && src && src !== "about:blank") {
    if (wv.dataset.pbNeedsReloadOnShow === "1") {
      delete wv.dataset.pbNeedsReloadOnShow;
      try {
        if (typeof wv.reload === "function") {
          wv.reload();
        }
      } catch (_e) {
        /* reload below */
      }
    }
    return;
  }
  navigateChatgptWebviewTo(wv, CHATGPT_HOME_URL);
}

/**
 * @param {HTMLElement} wv
 */
function initChatgptWebviewElement(wv) {
  loadChatgptWebviewIfNeeded(wv);
}

/**
 * @param {string} promptId
 * @param {string} label
 * @returns {HTMLElement}
 */
function createMainChatgptPaneElement(promptId, label) {
  const profile = gptResumeProfileForPromptKind(promptId);
  const pane = document.createElement("section");
  pane.className = "chatgpt-pane";
  pane.dataset.prompt = promptId;
  pane.dataset.chatgptProfile = profile;
  pane.setAttribute("aria-label", `ChatGPT — ${label}`);

  const header = document.createElement("div");
  header.className = "chatgpt-pane-header";
  pane.appendChild(header);

  const title = document.createElement("span");
  title.className = "chatgpt-pane-title";
  title.textContent = label;
  header.appendChild(title);

  const actions = document.createElement("div");
  actions.className = "chatgpt-pane-actions";
  header.appendChild(actions);

  const refreshBtn = document.createElement("button");
  refreshBtn.type = "button";
  refreshBtn.className = "btn-icon-reload chatgpt-pane-refresh";
  refreshBtn.dataset.chatgptProfile = profile;
  refreshBtn.dataset.prompt = promptId;
  refreshBtn.title = `Refresh ${label} ChatGPT`;
  refreshBtn.setAttribute("aria-label", `Refresh ${label} ChatGPT`);
  refreshBtn.innerHTML =
    '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M21 12a9 9 0 1 1-6.219-8.56"/><polyline points="21 3 21 9 15 9"/></svg>';
  actions.appendChild(refreshBtn);

  appendChatgptPaneToolbar(pane, {
    promptId,
    profile,
  });

  ensureChatgptPaneHeaderMeta(pane);
  ensureChatgptPaneDownloadStatus(pane);
  updateChatgptPaneDownloadStatusUi(pane);

  const wv = document.createElement("webview");
  wv.id = webviewDomIdForPromptId(promptId);
  wv.className = "chatgpt-webview";
  wv.setAttribute("partition", chatgptPartitionForPromptId(promptId));
  wv.setAttribute("allowpopups", "");
  wv.setAttribute(
    "webpreferences",
    "contextIsolation=yes, nativeWindowOpen=yes, javascript=yes",
  );
  pane.appendChild(wv);

  return pane;
}

/**
 * @param {{ promptId?: string | null, browserTabId?: string | null, profile: "olu" | "justin" }} spec
 * @returns {HTMLElement}
 */
function buildChatgptPaneToolbar(spec) {
  const toolbar = document.createElement("div");
  toolbar.className = "chatgpt-pane-toolbar";

  const inputBtn = document.createElement("button");
  inputBtn.type = "button";
  inputBtn.className = "btn-icon chatgpt-pane-input-gpt";
  inputBtn.innerHTML = CHATGPT_PANE_INPUT_ICON_SVG;
  inputBtn.title = "Insert preview into ChatGPT and send";
  inputBtn.setAttribute("aria-label", "Input to GPT");
  if (spec.promptId) {
    inputBtn.dataset.prompt = spec.promptId;
  }
  if (spec.browserTabId) {
    inputBtn.dataset.dedicatedBrowserTab = spec.browserTabId;
  }
  toolbar.appendChild(inputBtn);

  const viewBtn = document.createElement("button");
  viewBtn.type = "button";
  viewBtn.className = "btn-icon chatgpt-pane-view";
  viewBtn.innerHTML = CHATGPT_PANE_VIEW_ICON_SVG;
  viewBtn.title = "View retrieved result";
  viewBtn.setAttribute("aria-label", "View retrieved result");
  viewBtn.disabled = true;
  if (spec.promptId) {
    viewBtn.dataset.prompt = spec.promptId;
  }
  if (spec.browserTabId) {
    viewBtn.dataset.dedicatedBrowserTab = spec.browserTabId;
  }
  toolbar.appendChild(viewBtn);

  const downloadBtn = document.createElement("button");
  downloadBtn.type = "button";
  downloadBtn.className = "btn-icon chatgpt-pane-download";
  downloadBtn.innerHTML = CHATGPT_PANE_DOWNLOAD_ICON_SVG;
  downloadBtn.title = "Download Word";
  downloadBtn.setAttribute("aria-label", "Download Word");
  downloadBtn.disabled = true;
  if (spec.promptId) {
    downloadBtn.dataset.prompt = spec.promptId;
  }
  if (spec.browserTabId) {
    downloadBtn.dataset.dedicatedBrowserTab = spec.browserTabId;
  }
  toolbar.appendChild(downloadBtn);

  const fetchBtn = document.createElement("button");
  fetchBtn.type = "button";
  fetchBtn.className = "btn-icon chatgpt-pane-fetch";
  fetchBtn.innerHTML = CHATGPT_PANE_FETCH_ICON_SVG;
  fetchBtn.title = "Get result again from ChatGPT";
  fetchBtn.setAttribute("aria-label", "Get result again from ChatGPT");
  if (spec.promptId) {
    fetchBtn.dataset.prompt = spec.promptId;
  }
  if (spec.browserTabId) {
    fetchBtn.dataset.dedicatedBrowserTab = spec.browserTabId;
  }
  toolbar.appendChild(fetchBtn);

  return toolbar;
}

/** @returns {HTMLElement | null} */
function getDedicatedToolbarTargetPane() {
  if (!dedicatedPromptKind) {
    return null;
  }
  if (dedicatedActiveWorkspacePage !== "prompt") {
    const entry = dedicatedBrowserTabs.find((t) => t.id === dedicatedActiveWorkspacePage);
    return entry?.paneEl ?? dedicatedBrowserTabs[0]?.paneEl ?? null;
  }
  return dedicatedBrowserTabs[0]?.paneEl ?? null;
}

/** @param {Element | null | undefined} el @returns {HTMLElement | null} */
function resolveChatgptPaneForToolbarClick(el) {
  if (!(el instanceof Element)) {
    return null;
  }
  const fromPane = el.closest(".chatgpt-pane");
  if (fromPane instanceof HTMLElement) {
    return fromPane;
  }
  if (!el.closest("#dedicatedJdToolbar")) {
    return null;
  }
  return getDedicatedToolbarTargetPane();
}

function updateDedicatedJdToolbarContext() {
  const toolbar = document.getElementById("dedicatedJdToolbar");
  if (!toolbar || !dedicatedPromptKind) {
    return;
  }
  const tabId =
    dedicatedActiveWorkspacePage !== "prompt"
      ? dedicatedActiveWorkspacePage
      : dedicatedBrowserTabs[0]?.id ?? null;
  toolbar.querySelectorAll("button").forEach((node) => {
    const btn = /** @type {HTMLButtonElement} */ (node);
    btn.dataset.prompt = dedicatedPromptKind;
    if (tabId) {
      btn.dataset.dedicatedBrowserTab = tabId;
    } else {
      delete btn.dataset.dedicatedBrowserTab;
    }
  });
}

/** @param {HTMLElement} pane */
function mirrorDedicatedJdToolbarFromPane(pane) {
  const dedicated = document.getElementById("dedicatedJdToolbar");
  if (!dedicated || !dedicatedPromptKind) {
    return;
  }
  for (const cls of [
    "chatgpt-pane-input-gpt",
    "chatgpt-pane-view",
    "chatgpt-pane-download",
    "chatgpt-pane-fetch",
  ]) {
    const src = pane.querySelector(`.${cls}`);
    const dst = dedicated.querySelector(`.${cls}`);
    if (src instanceof HTMLButtonElement && dst instanceof HTMLButtonElement) {
      dst.disabled = src.disabled;
      dst.title = src.title;
      const aria = src.getAttribute("aria-label");
      if (aria) {
        dst.setAttribute("aria-label", aria);
      }
    }
  }
}

/** @returns {HTMLElement | null} */
function getDedicatedJdToolbarAnchor() {
  return (
    jdInput?.closest(".build-input-group") ??
    getDedicatedBuildTopRow()
  );
}

function ensureDedicatedJdToolbar() {
  const anchor = getDedicatedJdToolbarAnchor();
  let toolbar = document.getElementById("dedicatedJdToolbar");
  if (!dedicatedPromptKind) {
    if (toolbar) {
      toolbar.hidden = true;
    }
    return;
  }
  if (!anchor) {
    return;
  }
  if (!toolbar) {
    toolbar = document.createElement("div");
    toolbar.id = "dedicatedJdToolbar";
    toolbar.className = "dedicated-jd-toolbar chatgpt-pane-toolbar";
    toolbar.setAttribute("aria-label", "ChatGPT actions");
    anchor.appendChild(toolbar);
  } else if (toolbar.parentElement !== anchor) {
    anchor.appendChild(toolbar);
  }
  if (!toolbar.querySelector(".chatgpt-pane-input-gpt")) {
    const built = buildChatgptPaneToolbar({
      promptId: dedicatedPromptKind,
      profile: gptResumeProfileForPromptKind(dedicatedPromptKind),
    });
    toolbar.replaceChildren(...built.childNodes);
  }
  updateDedicatedJdToolbarContext();
  syncDedicatedJdToolbarUi();
}

function syncDedicatedJdToolbarUi() {
  const toolbar = document.getElementById("dedicatedJdToolbar");
  if (!toolbar || !dedicatedPromptKind) {
    return;
  }
  const inBrowserSidebar = !!toolbar.closest(".dedicated-browser-side-inner");
  toolbar.hidden = !inBrowserSidebar;
  if (!inBrowserSidebar) {
    return;
  }
  const pane = getDedicatedToolbarTargetPane();
  if (pane) {
    updateChatgptPaneToolbarUi(pane);
  }
}

/**
 * @param {HTMLElement} pane
 * @param {{ promptId?: string | null, browserTabId?: string | null, profile: "olu" | "justin" }} spec
 */
function appendChatgptPaneToolbar(pane, spec) {
  const header = pane.querySelector(".chatgpt-pane-header");
  if (!header) {
    return;
  }
  ensureChatgptPaneHeaderMeta(pane);
  let existingToolbar = pane.querySelector(".chatgpt-pane-toolbar");
  if (existingToolbar && !header.contains(existingToolbar)) {
    existingToolbar.remove();
    existingToolbar = null;
  }
  if (existingToolbar) {
    const hasInputIcon = !!existingToolbar.querySelector(".chatgpt-pane-input-gpt");
    const hasViewIcon = !!existingToolbar.querySelector(".chatgpt-pane-view");
    const hasFetchIcon = !!existingToolbar.querySelector(".chatgpt-pane-fetch");
    if (hasInputIcon && hasViewIcon && hasFetchIcon) {
      updateChatgptPaneToolbarUi(pane);
      return;
    }
    existingToolbar.remove();
  }
  const toolbar = buildChatgptPaneToolbar(spec);
  const actions = header.querySelector(".chatgpt-pane-actions");
  if (actions) {
    header.insertBefore(toolbar, actions);
  } else {
    header.appendChild(toolbar);
  }
  ensureChatgptPaneDownloadStatus(pane);
  updateChatgptPaneToolbarUi(pane);
}

/**
 * @param {HTMLElement} pane
 * @returns {{
 *   pane: HTMLElement;
 *   promptId: string | null;
 *   browserTabId: string | null;
 *   profile: "olu" | "justin";
 *   kind: string | null;
 *   wv: HTMLElement | null;
 * }}
 */
function getChatgptPaneContextFromPane(pane) {
  const promptId = pane.getAttribute("data-prompt");
  const browserTabId = pane.getAttribute("data-browser-tab");
  const profile =
    pane.getAttribute("data-chatgpt-profile") === "justin" ? "justin" : "olu";
  const kind = browserTabId ? dedicatedPromptKind : promptId;
  const wv = pane.querySelector(".chatgpt-webview");
  return {
    pane,
    promptId,
    browserTabId,
    profile,
    kind,
    wv:
      wv && typeof wv.executeJavaScript === "function"
        ? /** @type {HTMLElement} */ (wv)
        : null,
  };
}

/**
 * @param {ReturnType<typeof getChatgptPaneContextFromPane>} ctx
 * @returns {string | null} Disabled reason, or null if enabled.
 */
function inputGptDisabledReasonForPaneContext(ctx) {
  const kind = ctx.browserTabId ? dedicatedPromptKind : ctx.promptId;
  if (!kind) {
    return "Select a prompt";
  }
  if (
    isGptBusyForSendTarget(ctx.profile, { browserTabId: ctx.browserTabId })
  ) {
    return ctx.browserTabId
      ? "This browser tab is still generating."
      : "ChatGPT is still generating.";
  }
  if (!ctx.browserTabId && !dedicatedPromptKind && !isPasteCompleteForPromptKind(kind)) {
    return "Paste a job description (JD) first.";
  }
  if (!getFullPromptTextForKind(kind).length) {
    return "Nothing in preview yet";
  }
  return null;
}

/** @param {ReturnType<typeof getChatgptPaneContextFromPane>} ctx */
function getStoredGptResultForPaneContext(ctx) {
  if (ctx.browserTabId) {
    return lastGptResultsByBrowserTab[ctx.browserTabId] ?? null;
  }
  return getLastGptResult(ctx.profile);
}

/** @param {string} text */
async function isGptResultExportable(text) {
  const api = window.promptApi;
  if (!api || typeof api.checkGptDocxExport !== "function") {
    return false;
  }
  try {
    const r = await api.checkGptDocxExport(String(text ?? ""));
    return !!r?.ok;
  } catch (_e) {
    return false;
  }
}

/** @returns {HTMLElement | null} */
function getChatgptPaneForModalContext() {
  const { browserTabId, downloadKey } = gptResultModalContext;
  if (browserTabId) {
    return (
      dedicatedBrowserTabs.find((t) => t.id === browserTabId)?.paneEl ?? null
    );
  }
  if (downloadKey) {
    return document.querySelector(
      `#chatgptPanesHost .chatgpt-pane[data-prompt="${CSS.escape(downloadKey)}"]`,
    );
  }
  return null;
}

/** @param {HTMLElement} pane */
function updateChatgptPaneToolbarUi(pane) {
  const ctx = getChatgptPaneContextFromPane(pane);
  const stored = getStoredGptResultForPaneContext(ctx);
  const inputBtn = pane.querySelector(".chatgpt-pane-input-gpt");
  const viewBtn = pane.querySelector(".chatgpt-pane-view");
  const downloadBtn = pane.querySelector(".chatgpt-pane-download");
  const fetchBtn = pane.querySelector(".chatgpt-pane-fetch");
  const busy = isGptBusyForSendTarget(ctx.profile, { browserTabId: ctx.browserTabId });
  const inputBlocked = inputGptDisabledReasonForPaneContext(ctx);
  const mirrorDedicatedToolbarIfActive = () => {
    if (dedicatedPromptKind && pane === getDedicatedToolbarTargetPane()) {
      mirrorDedicatedJdToolbarFromPane(pane);
    }
  };
  if (inputBtn instanceof HTMLButtonElement) {
    inputBtn.disabled = !!inputBlocked;
    inputBtn.title =
      inputBlocked ||
      "Insert preview into ChatGPT, wait if uploading, then Enter to send";
    inputBtn.setAttribute("aria-label", inputBlocked || "Input to GPT");
  }
  if (viewBtn instanceof HTMLButtonElement) {
    viewBtn.disabled = !stored || stored.isError || !stored.text.trim();
    viewBtn.title = stored?.text?.trim()
      ? "View retrieved result"
      : "No result yet";
  }
  if (fetchBtn instanceof HTMLButtonElement) {
    fetchBtn.disabled = busy || !ctx.wv;
    fetchBtn.title = busy
      ? "Wait until ChatGPT finishes generating"
      : "Get result again from ChatGPT";
  }
  if (downloadBtn instanceof HTMLButtonElement) {
    if (busy) {
      downloadBtn.disabled = true;
      downloadBtn.title = "Wait until ChatGPT finishes generating";
    } else if (!stored || stored.isError || !stored.text.trim()) {
      downloadBtn.disabled = true;
      downloadBtn.title = "No JSON resume yet — use View and refresh";
    } else {
      const snapshot = stored.text;
      downloadBtn.disabled = true;
      downloadBtn.title = "Checking for resume_structure JSON…";
      void isGptResultExportable(snapshot).then((exportable) => {
        if (!pane.isConnected) {
          return;
        }
        const btn = pane.querySelector(".chatgpt-pane-download");
        if (!(btn instanceof HTMLButtonElement)) {
          return;
        }
        const latestCtx = getChatgptPaneContextFromPane(pane);
        const latest = getStoredGptResultForPaneContext(latestCtx);
        if (!latest || latest.text !== snapshot) {
          return;
        }
        if (
          isGptBusyForSendTarget(latestCtx.profile, {
            browserTabId: latestCtx.browserTabId,
          })
        ) {
          return;
        }
        btn.disabled = !exportable;
        btn.title = exportable
          ? "Download Word"
          : "Reply must be complete resume_structure JSON — open View, refresh, then try again";
        mirrorDedicatedToolbarIfActive();
      });
    }
  }
  updateChatgptPaneDownloadStatusUi(pane);
  mirrorDedicatedToolbarIfActive();
}

function updateAllChatgptPaneToolbarsUi() {
  document.querySelectorAll(".chatgpt-pane").forEach((pane) => {
    updateChatgptPaneToolbarUi(/** @type {HTMLElement} */ (pane));
  });
}

function updateChatgptPanesHostLayoutClasses() {
  const host = document.getElementById("chatgptPanesHost");
  if (!host) {
    return;
  }
  const visibleCount = listAllMainPrompts().filter(
    (p) => !readMainChatgptHiddenSet().has(p.id),
  ).length;
  host.classList.toggle("chatgpt-single", visibleCount === 1);
  host.classList.toggle("chatgpt-multi", visibleCount > 1);
}

/** @param {string} promptId */
function toggleMainChatgptPaneVisibility(promptId) {
  const hidden = readMainChatgptHiddenSet();
  const isHidden = hidden.has(promptId);
  if (!isHidden) {
    const visibleCount = listAllMainPrompts().filter(
      (p) => !hidden.has(p.id),
    ).length;
    if (visibleCount <= 1) {
      toast("Keep at least one ChatGPT panel visible.", "error", 3500);
      return;
    }
  }
  setMainChatgptPaneHidden(promptId, !isHidden);
  reorderPromptCardsByGptVisibility();
  const name = userDisplayNameForPromptKind(promptId);
  toast(
    isHidden ? `${name}: ChatGPT panel shown` : `${name}: ChatGPT panel hidden`,
    "success",
    3500,
  );
}

/** @param {string} promptId */
function setMainChatgptPaneHidden(promptId, hidden) {
  const set = readMainChatgptHiddenSet();
  if (hidden) {
    set.add(promptId);
  } else {
    set.delete(promptId);
  }
  writeMainChatgptHiddenSet(set);
  const pane = document.querySelector(
    `#chatgptPanesHost .chatgpt-pane[data-prompt="${CSS.escape(promptId)}"]`,
  );
  if (pane) {
    updateMainChatgptPaneHideUi(/** @type {HTMLElement} */ (pane));
  }
  removeMainChatgptPaneChrome();
  updateChatgptPanesHostLayoutClasses();
  syncPromptCardGptPaneToggles();
}

/** @param {HTMLElement} pane */
function updateMainChatgptPaneHideUi(pane) {
  const promptId = pane.getAttribute("data-prompt") || "";
  if (!promptId || dedicatedPromptKind) {
    pane.querySelector(".chatgpt-pane-hide")?.remove();
    return;
  }
  const hidden = readMainChatgptHiddenSet().has(promptId);
  pane.classList.toggle("chatgpt-pane--user-hidden", hidden);

  const actions = pane.querySelector(".chatgpt-pane-actions");
  if (!actions) {
    return;
  }
  let hideBtn = pane.querySelector(".chatgpt-pane-hide");
  if (hidden) {
    hideBtn?.remove();
    return;
  }
  const name = userDisplayNameForPromptKind(promptId) || promptId;
  const title = `Hide ${name} ChatGPT panel`;
  if (!hideBtn) {
    hideBtn = document.createElement("button");
    hideBtn.type = "button";
    hideBtn.className = "btn-icon chatgpt-pane-hide";
    hideBtn.dataset.promptGptHide = promptId;
    const refreshBtn = actions.querySelector(".chatgpt-pane-refresh");
    if (refreshBtn) {
      refreshBtn.insertAdjacentElement("afterend", hideBtn);
    } else {
      actions.appendChild(hideBtn);
    }
  }
  hideBtn.title = title;
  hideBtn.setAttribute("aria-label", title);
  hideBtn.innerHTML = PROMPT_GPT_PANE_EYE_OFF_ICON_SVG;
}

function removeMainChatgptPaneChrome() {
  document
    .getElementById("chatgptPanesHost")
    ?.querySelector(".chatgpt-panes-empty")
    ?.remove();
}

function renderChatgptShowPanesBar() {
  syncPromptCardGptPaneToggles();
}

function syncMainChatgptPaneHighlights() {
  const kind = getSelectedPromptKind();
  document.querySelectorAll("#chatgptPanesHost .chatgpt-pane").forEach((pane) => {
    const id = pane.getAttribute("data-prompt");
    pane.classList.toggle("chatgpt-pane--active", !!kind && id === kind);
  });
}

function syncMainWindowChatgptPanes() {
  const host = document.getElementById("chatgptPanesHost");
  if (!host || dedicatedPromptKind) {
    return;
  }

  const prompts = listAllMainPrompts();
  const wanted = new Set(prompts.map((p) => p.id));

  for (const pane of [...host.querySelectorAll(".chatgpt-pane")]) {
    const id = pane.getAttribute("data-prompt");
    if (id && !wanted.has(id)) {
      pane.remove();
    }
  }

  for (const { id, label } of prompts) {
    let pane = host.querySelector(
      `.chatgpt-pane[data-prompt="${CSS.escape(id)}"]`,
    );
    if (!pane) {
      pane = createMainChatgptPaneElement(id, label);
      host.appendChild(pane);
      initChatgptWebviewElement(pane.querySelector(".chatgpt-webview"));
    } else {
      const titleEl = pane.querySelector(".chatgpt-pane-title");
      if (titleEl) {
        titleEl.textContent = label;
      }
      pane.setAttribute("aria-label", `ChatGPT — ${label}`);
    }
    updateMainChatgptPaneHideUi(/** @type {HTMLElement} */ (pane));
    ensureChatgptPaneHeaderMeta(/** @type {HTMLElement} */ (pane));
    appendChatgptPaneToolbar(/** @type {HTMLElement} */ (pane), {
      promptId: id,
      profile: gptResumeProfileForPromptKind(id),
    });
  }

  removeMainChatgptPaneChrome();
  updateChatgptPanesHostLayoutClasses();
  syncMainChatgptPaneHighlights();
  renderChatgptShowPanesBar();
}

function setupChatgptShowPanesHandlers() {
  if (document.body.dataset.chatgptShowPanesBound) {
    return;
  }
  document.body.dataset.chatgptShowPanesBound = "1";

  document.addEventListener("click", (e) => {
    const hidePaneBtn = e.target.closest(".chatgpt-pane-hide");
    if (hidePaneBtn) {
      e.preventDefault();
      e.stopPropagation();
      const promptId =
        hidePaneBtn.getAttribute("data-prompt-gpt-hide") ||
        hidePaneBtn.closest(".chatgpt-pane")?.getAttribute("data-prompt");
      if (promptId) {
        toggleMainChatgptPaneVisibility(promptId);
      }
      return;
    }
    const toggleBtn = e.target.closest(".prompt-gpt-pane-toggle");
    if (toggleBtn) {
      e.preventDefault();
      e.stopPropagation();
      const promptId = toggleBtn.getAttribute("data-prompt-gpt-toggle");
      if (promptId) {
        toggleMainChatgptPaneVisibility(promptId);
      }
      return;
    }
  });
}

function setupMainChatgptPanesHost() {
  const panel = document.getElementById("workspaceBrowserPanel");
  if (!panel || panel.dataset.chatgptBrowserBound) {
    return;
  }
  panel.dataset.chatgptBrowserBound = "1";
}

/**
 * @param {string | null | undefined} kind
 * @returns {"olu" | "justin"}
 */
function gptResumeProfileForPromptKind(kind) {
  if (kind === "justin") {
    return "justin";
  }
  const custom = getCustomPromptByKind(kind);
  if (custom) {
    return custom.chatgptProfile;
  }
  if (kind === "other") {
    return "olu";
  }
  return "olu";
}

/**
 * @param {"olu" | "justin"} profile
 * @returns {HTMLElement | null}
 */
function getChatgptWebviewForProfile(profile) {
  const tabId = getActiveDedicatedBrowserTabId();
  if (tabId) {
    const dedicated = document.getElementById(`chatgptWebviewDedicated-${tabId}`);
    if (dedicated) {
      return dedicated;
    }
  }
  if (!dedicatedPromptKind) {
    const pane = document.querySelector(
      `#chatgptPanesHost .chatgpt-pane[data-chatgpt-profile="${profile === "justin" ? "justin" : "olu"}"]`,
    );
    const wv = pane?.querySelector(".chatgpt-webview");
    if (wv && typeof wv.executeJavaScript === "function") {
      return /** @type {HTMLElement} */ (wv);
    }
  }
  return document.getElementById(webviewDomIdForPromptId("other"));
}

/**
 * @param {string | null | undefined} kind
 * @returns {HTMLElement | null}
 */
function getChatgptWebviewForKind(kind) {
  if (!kind) {
    return null;
  }
  const tabId = getActiveDedicatedBrowserTabId();
  if (tabId) {
    const dedicated = document.getElementById(`chatgptWebviewDedicated-${tabId}`);
    if (dedicated) {
      return dedicated;
    }
  }
  if (dedicatedPromptKind) {
    return getChatgptWebviewForProfile(gptResumeProfileForPromptKind(kind));
  }
  const pane = document.querySelector(
    `#chatgptPanesHost .chatgpt-pane[data-prompt="${CSS.escape(kind)}"]`,
  );
  const wv = pane?.querySelector(".chatgpt-webview");
  return wv && typeof wv.executeJavaScript === "function"
    ? /** @type {HTMLElement} */ (wv)
    : null;
}

/** @returns {HTMLElement[]} */
function getAllChatgptWebviews() {
  const seen = new Set();
  /** @type {HTMLElement[]} */
  const list = [];
  const add = (/** @type {HTMLElement | null} */ el) => {
    if (!el || seen.has(el) || typeof el.executeJavaScript !== "function") {
      return;
    }
    seen.add(el);
    list.push(el);
  };
  document
    .querySelectorAll("#chatgptPanesHost .chatgpt-webview")
    .forEach((el) => {
      add(/** @type {HTMLElement} */ (el));
    });
  document.querySelectorAll(".dedicated-browser-host .chatgpt-webview").forEach((el) => {
    add(/** @type {HTMLElement} */ (el));
  });
  return list;
}

/** @type {"other" | "justin"} */
let editSelection = "other";

/** @type {"builtin" | "external"} */
let editMode = "builtin";

/** Absolute path of file opened in external mode (for save / reload). */
let externalAbsPath = "";

const JD_PLACEHOLDERS = [
  "{PASTE JD HERE}",
  "[PASTE JD HERE]",
  "[PASTE JOB DESCRIPTION HERE]",
];

const WORK_EXP_PLACEHOLDERS = [
  "{PASTE EXISTING RESUME OR LEAVE EMPTY}",
  "[PASTE RESUME HERE]",
  "[PASTE ORIGIN WORK EXPERIENCE HERE]",
];

function clearBuildFieldHighlights() {
  jdInput?.closest(".build-input-group")?.classList.remove("build-input-group--highlight");
}

function syncBuildInputShellState(el) {
  const shell = el.closest(".build-input-shell");
  if (!shell) {
    return;
  }
  const filled = el.value.trim().length > 0;
  shell.classList.toggle("is-empty", !filled);
  shell.classList.toggle("is-filled", filled);
}

function syncAllBuildInputShellStates() {
  if (jdInput) {
    syncBuildInputShellState(jdInput);
  }
  syncMainCompanyShell();
}

function readSharedCompanyFromStorage() {
  // Main-window company should not persist between app sessions.
  return "";
}

function writeSharedCompanyToStorage(value) {
  // Keep in-memory only; do not persist session input history.
  void value;
}

function purgePersistedInputHistory() {
  try {
    const keysToRemove = [SHARED_COMPANY_STORAGE_KEY];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (!key) {
        continue;
      }
      if (
        key.startsWith(MAIN_COMPANY_STORAGE_PREFIX) ||
        key.startsWith(DEDICATED_BROWSER_TAB_STORAGE_PREFIX)
      ) {
        keysToRemove.push(key);
      }
    }
    for (const key of keysToRemove) {
      localStorage.removeItem(key);
    }
  } catch (_e) {
    /* ignore */
  }
}

function syncMainCompanyShell() {
  if (!mainCompanyInput) {
    return;
  }
  const shell = mainCompanyInput.closest(".build-input-shell");
  if (!shell) {
    return;
  }
  const filled = !!mainCompanyInput.value.trim();
  shell.classList.toggle("is-empty", !filled);
  shell.classList.toggle("is-filled", filled);
}

function loadMainCompanyName() {
  if (!mainCompanyInput || dedicatedPromptKind) {
    return;
  }
  mainCompanyInput.value = readSharedCompanyFromStorage();
  syncMainCompanyShell();
}

function persistMainCompanyName() {
  if (!mainCompanyInput || dedicatedPromptKind) {
    return;
  }
  writeSharedCompanyToStorage(mainCompanyInput.value);
}

function getMainCompanyName() {
  return mainCompanyInput?.value?.trim() || "";
}

/**
 * Browser tab entry with its form state up to date. The shared JD / company inputs only
 * reflect the tab currently shown, so flush them into that tab's entry before reading.
 * @param {string | null | undefined} browserTabId
 */
function getBrowserTabEntryForDownload(browserTabId) {
  if (!browserTabId) {
    return null;
  }
  if (browserTabId === dedicatedActiveWorkspacePage) {
    saveDedicatedBrowserTabFormState(browserTabId);
  }
  return dedicatedBrowserTabs.find((t) => t.id === browserTabId) ?? null;
}

/** @param {string | null | undefined} [browserTabId] owning tab of the result */
function getCompanyFolderForDownload(browserTabId) {
  const entry = getBrowserTabEntryForDownload(browserTabId);
  if (entry) {
    return String(entry.company ?? "").trim();
  }
  if (dedicatedPromptKind) {
    return getDedicatedCompanyName();
  }
  return getMainCompanyName();
}

/** @param {string | null | undefined} browserTabId */
function getJdTextForDownload(browserTabId) {
  const entry = getBrowserTabEntryForDownload(browserTabId);
  if (entry) {
    return String(entry.jd ?? "").trim();
  }
  return String(jdInput?.value ?? "").trim();
}

/**
 * @param {string | null} kind data-prompt value
 */
function syncBuildFieldHighlightsForKind(kind) {
  if (!jdInput) {
    return;
  }
  if (!isBuiltinPromptKind(kind) && !isCustomPromptKind(kind)) {
    clearBuildFieldHighlights();
    return;
  }
  clearBuildFieldHighlights();
  const jd = jdInput.value.trim();
  const jdGroup = jdInput.closest(".build-input-group");
  if (!jd) {
    jdGroup?.classList.add("build-input-group--highlight");
  }
}

function onBuildInputsChanged() {
  syncAllBuildInputShellStates();
  if (dedicatedPromptKind && dedicatedActiveWorkspacePage !== "prompt") {
    saveDedicatedBrowserTabFormState(dedicatedActiveWorkspacePage);
  }
  if (!(dedicatedPromptKind && dedicatedActiveWorkspacePage === "prompt")) {
    refreshBuildPreviewIfCardSelected();
  }
}

function afterPasteApply(el, onNonEmpty) {
  setTimeout(() => {
    if (el.value.trim()) {
      onNonEmpty();
    }
  }, 0);
}

function setupBuildPasteToasts() {
  if (!jdInput) {
    return;
  }
  jdInput.addEventListener("paste", () => {
    afterPasteApply(jdInput, () => toast("JD pasted."));
  });
}

/**
 * @param {string} message
 * @param {"success" | "error"} [kind]
 * @param {number} [durationMs]
 * @param {{ label: string, onClick: () => void }[]} [actions]
 */
function toast(message, kind = "success", durationMs, actions) {
  const text = message == null ? "" : String(message).trim();
  if (!text) {
    return;
  }
  const host = document.getElementById("toastHost");
  if (!host) {
    return;
  }
  const isError = kind === "error";
  const duration =
    typeof durationMs === "number" && durationMs > 0
      ? durationMs
      : isError
        ? 5200
        : 3600;

  const el = document.createElement("div");
  el.className = `toast toast--${isError ? "error" : "success"}`;
  el.setAttribute("role", "status");

  const icon = document.createElement("span");
  icon.className = "toast-icon";
  icon.setAttribute("aria-hidden", "true");
  icon.textContent = isError ? "!" : "✓";

  const msg = document.createElement("p");
  msg.className = "toast-message";
  if (text.includes("\n")) {
    msg.classList.add("toast-message--multiline");
  }
  msg.textContent = text;

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "toast-close";
  closeBtn.setAttribute("aria-label", "Dismiss notification");
  closeBtn.textContent = "×";

  const inner = document.createElement("div");
  inner.className = "toast-inner";
  inner.append(icon, msg, closeBtn);

  const actionList = Array.isArray(actions) ? actions.filter((a) => a?.label) : [];
  if (actionList.length > 0) {
    el.classList.add("toast--with-actions");
    if (text.includes("\n")) {
      el.classList.add("toast--wide");
    }
    const actionsEl = document.createElement("div");
    actionsEl.className = "toast-actions";
    for (const a of actionList) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "toast-action-btn";
      btn.textContent = a.label;
      btn.addEventListener("click", (ev) => {
        ev.stopPropagation();
        a.onClick();
      });
      actionsEl.append(btn);
    }
    inner.append(actionsEl);
  }

  el.append(inner);
  host.appendChild(el);

  let hideTimer = 0;
  let gone = false;
  const dismiss = () => {
    if (gone) {
      return;
    }
    gone = true;
    clearTimeout(hideTimer);
    el.classList.remove("toast--visible");
    el.classList.add("toast--leaving");
    setTimeout(() => el.remove(), 220);
  };

  closeBtn.addEventListener("click", dismiss);
  requestAnimationFrame(() => {
    el.classList.add("toast--visible");
  });
  hideTimer = setTimeout(dismiss, duration);
}

function applyPlaceholders(template, jdText, workExpText) {
  let out = template;
  for (const ph of JD_PLACEHOLDERS) {
    out = out.split(ph).join(jdText);
  }
  for (const ph of WORK_EXP_PLACEHOLDERS) {
    out = out.split(ph).join(workExpText);
  }
  return out;
}

function getStrForEdit(key) {
  if (key === "justin") return justinTemplateStr;
  return otherTemplateStr;
}

function setStrForEdit(key, value) {
  if (key === "justin") justinTemplateStr = value;
  else otherTemplateStr = value;
}

function getIdForEdit(key) {
  if (key === "justin") return JUSTIN_ID;
  return OTHER_ID;
}

function fileNameFromPath(p) {
  const s = String(p || "").replace(/\\/g, "/");
  const i = s.lastIndexOf("/");
  return i >= 0 ? s.slice(i + 1) : s || "file";
}

/** @type {{ folderPath: string, filePath: string, paneEl: HTMLElement | null } | null} */
let lastGptDownloadOpenTarget = null;

/**
 * @param {HTMLElement | null} host
 * @param {unknown} rows
 */
function renderWordPreviewRows(host, rows) {
  if (!host) {
    return;
  }
  host.textContent = "";
  if (!Array.isArray(rows) || rows.length === 0) {
    return;
  }
  for (const row of rows) {
    const rec = /** @type {Record<string, unknown>} */ (row || {});
    const c = String(rec.company ?? "").trim();
    const loc = String(rec.location ?? "").trim();
    const lt = String(rec.location_type ?? "").trim();
    const rt = String(rec.role_title ?? "").trim();
    const dur = String(rec.duration ?? "").trim();
    const locationLine = String(rec.location_line ?? "").trim();
    const roleCompany = [rt, c].filter(Boolean).join(" | ");
    const algoLine = locationLine
      ? `${roleCompany}${locationLine}${dur ? `\t${dur}` : ""}`
      : [c, loc, lt, rt, dur].filter(Boolean).join(" | ");
    const legacy = String(rec.experience_line_left ?? "").trim();
    const text = algoLine || legacy;
    if (!text) {
      continue;
    }
    const div = document.createElement("div");
    div.className = "gpt-result-wx-word-row";
    const line = document.createElement("span");
    line.className = "gpt-result-wx-word-algo-line";
    line.textContent = text;
    div.append(line);
    host.appendChild(div);
  }
}

/**
 * @param {HTMLElement} textEl
 * @param {string} savedName
 * @param {string} folderPath
 * @param {string} jdFileName
 * @param {string} who
 */
function fillGptDownloadStatusEl(textEl, savedName, folderPath, jdFileName, who) {
  textEl.replaceChildren();
  const line1 = document.createElement("span");
  line1.className = "chatgpt-pane-download-status-line";
  const strong = document.createElement("strong");
  const jdNote = jdFileName ? ` · ${jdFileName}` : "";
  strong.textContent = `${who}: saved ${savedName}${jdNote}`;
  line1.appendChild(strong);
  textEl.appendChild(line1);
  if (folderPath) {
    const pathRow = document.createElement("div");
    pathRow.className = "chatgpt-pane-download-status-path-row";

    const line2 = document.createElement("code");
    line2.className = "chatgpt-pane-download-status-path";
    line2.textContent = folderPath;

    const copyBtn = document.createElement("button");
    copyBtn.type = "button";
    copyBtn.className = "btn-icon chatgpt-pane-download-copy-path";
    copyBtn.innerHTML = CHATGPT_PANE_COPY_ICON_SVG;
    copyBtn.title = "Copy folder path";
    copyBtn.setAttribute("aria-label", "Copy folder path");

    pathRow.append(line2, copyBtn);
    textEl.appendChild(pathRow);
  }
}

/** @param {HTMLElement} bar */
function resolvePaneForDownloadStatusBar(bar) {
  const tabId = bar.getAttribute("data-dedicated-browser-tab");
  if (tabId) {
    return dedicatedBrowserTabs.find((t) => t.id === tabId)?.paneEl ?? null;
  }
  const pane = bar.closest(".chatgpt-pane");
  return pane instanceof HTMLElement ? pane : null;
}

/** @param {HTMLElement} bar @param {"folder" | "resume" | "jd"} kind */
async function openGptDownloadFromBar(bar, kind) {
  const folderPath = String(bar.dataset.statusFolderPath || "").trim();
  const filePath = String(bar.dataset.statusFilePath || "").trim();
  const jdPath = String(bar.dataset.statusJdPath || "").trim();
  const pane = resolvePaneForDownloadStatusBar(bar);
  const api = window.promptApi;

  if (kind === "folder") {
    if (!folderPath && !filePath) {
      showGptDownloadOpenFolderError("No download folder to open.");
      return;
    }
    lastGptDownloadOpenTarget = { folderPath, filePath, paneEl: pane };
    await openLastGptDownloadFolder();
    return;
  }

  const targetPath = kind === "jd" ? jdPath : filePath;
  if (!targetPath) {
    showGptDownloadOpenFolderError(
      kind === "jd" ? "No JD file to open." : "No resume file to open.",
    );
    return;
  }
  if (!api?.openGptDownloadFile) {
    showGptDownloadOpenFolderError(
      "Restart the app to enable Open file (reload main process).",
    );
    return;
  }
  try {
    const openR = await api.openGptDownloadFile({ filePath: targetPath });
    if (!openR?.ok) {
      showGptDownloadOpenFolderError(openR?.error || "Could not open file.");
    }
  } catch (e) {
    showGptDownloadOpenFolderError(String(e && e.message ? e.message : e));
  }
}

/** @param {HTMLElement} previewWrap */
function ensureDownloadResumePreviewHead(previewWrap) {
  if (!(previewWrap instanceof HTMLElement)) {
    return;
  }
  if (previewWrap.querySelector(".chatgpt-pane-download-resume-preview-head")) {
    previewWrap.querySelector(".chatgpt-pane-download-resume-preview-actions")?.remove();
    return;
  }
  const label = previewWrap.querySelector(".chatgpt-pane-download-resume-preview-label");
  const host = previewWrap.querySelector(".chatgpt-pane-download-resume-preview-body");
  if (!(label instanceof HTMLElement) || !(host instanceof HTMLElement)) {
    return;
  }
  const head = document.createElement("div");
  head.className = "chatgpt-pane-download-resume-preview-head";
  label.remove();
  head.append(label);
  previewWrap.insertBefore(head, host);
}

/** @param {HTMLElement} bar */
function getDocxPreviewHostFromBar(bar) {
  return bar.querySelector(".chatgpt-pane-download-resume-preview-body");
}

/** @param {HTMLElement} bar */
function updateDocxPreviewInteractivity(bar) {
  const wrap = bar.querySelector(".chatgpt-pane-download-resume-preview");
  if (!(wrap instanceof HTMLElement)) {
    return;
  }
  ensureDownloadResumePreviewHead(wrap);
  const host = getDocxPreviewHostFromBar(bar);
  const hasDoc = !!host?.querySelector("section.docx-resume");
  const filePath = String(bar.dataset.statusFilePath || "").trim();
  const openable = hasDoc && !!filePath;
  if (host instanceof HTMLElement) {
    host.classList.toggle("docx-preview-openable", openable);
    if (openable) {
      host.setAttribute("role", "button");
      host.setAttribute("tabindex", "0");
      host.setAttribute("title", "Open document preview");
      host.setAttribute("aria-label", "Open document preview");
    } else {
      host.removeAttribute("role");
      host.removeAttribute("tabindex");
      host.removeAttribute("title");
      host.setAttribute("aria-label", "Resume preview");
    }
  }
}

/** @param {HTMLElement} bar */
async function openDocxPreviewModalFromBar(bar) {
  const filePath = String(bar.dataset.statusFilePath || "").trim();
  if (!filePath) {
    toast("No document to show.", "error");
    return;
  }
  const modal = document.getElementById("docxPreviewModal");
  const body = document.getElementById("docxPreviewModalBody");
  const title = document.getElementById("docxPreviewModalTitle");
  if (!(modal instanceof HTMLElement) || !(body instanceof HTMLElement)) {
    return;
  }
  if (title instanceof HTMLElement) {
    title.textContent = fileNameFromPath(filePath) || "Document preview";
  }
  modal.hidden = false;
  const dialog = modal.querySelector(".docx-preview-dialog");
  if (dialog instanceof HTMLElement) {
    dialog.focus();
  }
  ensureDocxPreviewAutoFit(body);
  await renderDocxFilePreview(body, filePath);
}

function closeDocxPreviewModal() {
  const modal = document.getElementById("docxPreviewModal");
  const body = document.getElementById("docxPreviewModalBody");
  if (modal instanceof HTMLElement) {
    modal.hidden = true;
  }
  if (body instanceof HTMLElement) {
    body.replaceChildren();
  }
}

function setupDocxPreviewModal() {
  const modal = document.getElementById("docxPreviewModal");
  const closeBtn = document.getElementById("docxPreviewModalClose");
  if (!(modal instanceof HTMLElement)) {
    return;
  }
  closeBtn?.addEventListener("click", closeDocxPreviewModal);
  modal.addEventListener("click", (e) => {
    if (e.target === modal) {
      closeDocxPreviewModal();
    }
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && !modal.hidden) {
      closeDocxPreviewModal();
    }
  });
}

/** @param {HTMLElement} bar */
function bindDocxPreviewOpenHandler(bar) {
  const host = getDocxPreviewHostFromBar(bar);
  if (!(host instanceof HTMLElement) || host.dataset.openHandlerBound === "1") {
    return;
  }
  host.dataset.openHandlerBound = "1";
  host.addEventListener("click", () => {
    if (!host.classList.contains("docx-preview-openable")) {
      return;
    }
    void openDocxPreviewModalFromBar(bar);
  });
  host.addEventListener("keydown", (e) => {
    if (!host.classList.contains("docx-preview-openable")) {
      return;
    }
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      void openDocxPreviewModalFromBar(bar);
    }
  });
}

/** @param {HTMLElement} bar */
function bindDownloadStatusBarActions(bar) {
  const previewWrap = bar.querySelector(".chatgpt-pane-download-resume-preview");
  if (previewWrap instanceof HTMLElement) {
    ensureDownloadResumePreviewHead(previewWrap);
    bindDocxPreviewOpenHandler(bar);
  }

  if (bar.dataset.actionsBound === "1") {
    updateDocxPreviewInteractivity(bar);
    return;
  }
  bar.dataset.actionsBound = "1";

  const folderBtn = bar.querySelector(".chatgpt-pane-download-open-folder");
  const resumeBtn = bar.querySelector(".chatgpt-pane-download-open-resume");
  const jdBtn = bar.querySelector(".chatgpt-pane-download-open-jd");

  folderBtn?.addEventListener("click", () => {
    void openGptDownloadFromBar(bar, "folder");
  });
  resumeBtn?.addEventListener("click", () => {
    void openGptDownloadFromBar(bar, "resume");
  });
  jdBtn?.addEventListener("click", () => {
    void openGptDownloadFromBar(bar, "jd");
  });

  if (bar.dataset.pathCopyBound !== "1") {
    bar.dataset.pathCopyBound = "1";
    bar.addEventListener("click", (e) => {
      const btn =
        e.target instanceof HTMLElement
          ? e.target.closest(".chatgpt-pane-download-copy-path")
          : null;
      if (!(btn instanceof HTMLElement) || !bar.contains(btn)) {
        return;
      }
      const folderPath = String(bar.dataset.statusFolderPath || "").trim();
      void copyStringToClipboard(folderPath, {
        emptyMsg: "No folder path to copy.",
        successMsg: "Copied folder path.",
      });
    });
  }
}

/** @param {HTMLElement} host */
function applyDocxPreviewNaturalLayout(host) {
  const mount = host.querySelector(".docx-resume-preview-mount");
  const firstPage = host.querySelector("section.docx-resume:not([hidden])");
  const hostWidth = Math.max(0, host.clientWidth);
  const pageWidth =
    firstPage instanceof HTMLElement ? firstPage.offsetWidth : 0;
  const availableWidth = Math.max(0, hostWidth - 32);
  const scale =
    pageWidth > 0 && availableWidth > 0
      ? Math.min(1, availableWidth / pageWidth)
      : 1;
  host.style.setProperty("--docx-fit-scale", String(scale));

  if (mount instanceof HTMLElement) {
    mount.style.zoom = scale < 0.999 ? String(scale) : "";
    mount.style.width = "100%";
    mount.style.marginBottom = "0";
    mount.style.transform = "";
  }

  host.querySelectorAll("section.docx-resume:not([hidden])").forEach((section) => {
    if (!(section instanceof HTMLElement)) {
      return;
    }
    if (section.style.width === "100%") {
      section.style.removeProperty("width");
    }
    section.style.maxWidth = "none";
    section.style.minHeight = "auto";
    section.style.height = "auto";
    section.style.flexShrink = "0";
  });
}

/** @param {HTMLElement} section */
function isDocxPreviewPageBlank(section) {
  if (section.querySelector("img[src]")) {
    return false;
  }
  const article = section.querySelector("article");
  const contentRoot = article instanceof HTMLElement ? article : section;
  const text = (contentRoot.textContent || "").replace(/\s+/g, " ").trim();
  if (text.replace(/[^a-zA-Z0-9]/g, "").length >= 3) {
    return false;
  }
  for (const cell of contentRoot.querySelectorAll("td, th, li")) {
    if ((cell.textContent || "").replace(/[^a-zA-Z0-9]/g, "").length >= 2) {
      return false;
    }
  }
  return true;
}

/** @param {HTMLElement} section */
function trimDocxPreviewTrailingEmptyNodes(section) {
  const article = section.querySelector("article");
  if (!(article instanceof HTMLElement)) {
    return;
  }
  while (article.lastElementChild) {
    const node = article.lastElementChild;
    const text = (node.textContent || "").replace(/\s+/g, "");
    if (text.length > 0 || node.querySelector("img, table")) {
      break;
    }
    node.remove();
  }
}

/** Hide blank Word pages; label visible pages for page-by-page preview. */
function normalizeDocxPreviewPages(host) {
  const sections = [...host.querySelectorAll("section.docx-resume")].filter(
    (el) => el instanceof HTMLElement,
  );
  let pageNum = 0;
  for (const section of sections) {
    section.style.minHeight = "auto";
    section.style.height = "auto";
    if (isDocxPreviewPageBlank(section)) {
      section.remove();
      continue;
    }
    trimDocxPreviewTrailingEmptyNodes(section);
    pageNum += 1;
    section.dataset.pageNumber = String(pageNum);
    section.hidden = false;
  }
  requestAnimationFrame(() => applyDocxPreviewNaturalLayout(host));
}

/** @type {WeakMap<HTMLElement, ResizeObserver>} */
const docxPreviewResizeObservers = new WeakMap();

/** @param {HTMLElement} host */
function ensureDocxPreviewAutoFit(host) {
  if (docxPreviewResizeObservers.has(host) || typeof ResizeObserver !== "function") {
    return;
  }
  let rafId = 0;
  const observer = new ResizeObserver(() => {
    if (rafId) {
      cancelAnimationFrame(rafId);
    }
    rafId = requestAnimationFrame(() => {
      rafId = 0;
      applyDocxPreviewNaturalLayout(host);
    });
  });
  observer.observe(host);
  docxPreviewResizeObservers.set(host, observer);
}

/** @param {Record<string, unknown>} payload */
function ipcPayloadToUint8Array(payload) {
  const base64 = String(payload?.base64 ?? "").trim();
  if (base64) {
    const binary = atob(base64);
    const out = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) {
      out[i] = binary.charCodeAt(i);
    }
    return out;
  }
  const raw = payload?.bytes;
  if (!raw) {
    return null;
  }
  if (raw instanceof Uint8Array) {
    return raw;
  }
  if (raw instanceof ArrayBuffer) {
    return new Uint8Array(raw);
  }
  const bufLike = /** @type {{ type?: string, data?: number[] }} */ (raw);
  if (bufLike.type === "Buffer" && Array.isArray(bufLike.data)) {
    return Uint8Array.from(bufLike.data);
  }
  try {
    return new Uint8Array(/** @type {ArrayLike<number>} */ (raw));
  } catch (_e) {
    return null;
  }
}

/** @param {HTMLElement} host @param {string} message */
function renderDocxPreviewPlaceholder(host, message) {
  host.replaceChildren();
  host.className = "chatgpt-pane-download-resume-preview-body docx-resume-preview-body";
  const note = document.createElement("p");
  note.className = "docx-resume-preview-placeholder";
  note.textContent = message;
  host.appendChild(note);
}

/**
 * @param {HTMLElement} host
 * @param {string} filePath
 */
async function renderDocxFilePreview(host, filePath) {
  host.replaceChildren();
  host.className = "chatgpt-pane-download-resume-preview-body docx-resume-preview-body";
  ensureDocxPreviewAutoFit(host);

  const docxApi = /** @type {{ renderAsync?: Function } | undefined} */ (
    globalThis.docx
  );
  if (!filePath) {
    return false;
  }
  if (typeof docxApi?.renderAsync !== "function") {
    renderDocxPreviewPlaceholder(
      host,
      "Document viewer is not loaded. Restart the app and try again.",
    );
    return true;
  }

  const api = window.promptApi;
  if (!api?.readGptDocxFile) {
    renderDocxPreviewPlaceholder(
      host,
      "Cannot read the saved file from this window.",
    );
    return true;
  }

  const loading = document.createElement("p");
  loading.className = "docx-resume-preview-placeholder docx-resume-preview-loading";
  loading.textContent = "Loading document…";
  host.appendChild(loading);

  let payload;
  try {
    payload = await api.readGptDocxFile(filePath);
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    if (/no handler registered/i.test(msg)) {
      renderDocxPreviewPlaceholder(
        host,
        "Restart the app (stop npm start, then start again) to enable document preview.",
      );
    } else {
      renderDocxPreviewPlaceholder(host, `Could not read file: ${msg}`);
    }
    return true;
  }
  if (!payload?.ok) {
    renderDocxPreviewPlaceholder(
      host,
      String(payload?.error || "Could not read the saved resume file."),
    );
    return true;
  }

  const bytes = ipcPayloadToUint8Array(
    /** @type {Record<string, unknown>} */ (payload),
  );
  if (!bytes?.length) {
    renderDocxPreviewPlaceholder(host, "The saved resume file is empty.");
    return true;
  }

  host.replaceChildren();

  const styleHost = document.createElement("div");
  styleHost.className = "docx-resume-preview-styles";
  host.appendChild(styleHost);

  const bodyHost = document.createElement("div");
  bodyHost.className = "docx-resume-preview-mount";
  host.appendChild(bodyHost);

  try {
    await docxApi.renderAsync(bytes, bodyHost, styleHost, {
      className: "docx-resume",
      inWrapper: true,
      ignoreFonts: false,
      breakPages: true,
      renderHeaders: true,
      renderFooters: true,
      ignoreWidth: false,
      ignoreHeight: true,
    });
    if (bodyHost.childElementCount === 0) {
      renderDocxPreviewPlaceholder(host, "No content in the saved document.");
      return true;
    }
    normalizeDocxPreviewPages(host);
    return true;
  } catch (e) {
    console.warn("DOCX preview failed:", e);
    renderDocxPreviewPlaceholder(
      host,
      "Could not render the Word document. Use Open resume to view it.",
    );
    return true;
  }
}

/** @param {HTMLElement} bar @param {HTMLElement | null} _pane */
async function refreshDownloadStatusResumePreview(bar, _pane) {
  const previewWrap = bar.querySelector(".chatgpt-pane-download-resume-preview");
  const previewHost = bar.querySelector(".chatgpt-pane-download-resume-preview-body");
  if (!(previewWrap instanceof HTMLElement) || !(previewHost instanceof HTMLElement)) {
    return;
  }

  const filePath = String(bar.dataset.statusFilePath || "").trim();
  if (!filePath) {
    previewWrap.hidden = true;
    return;
  }

  previewWrap.hidden = false;
  ensureDocxPreviewAutoFit(previewHost);
  await renderDocxFilePreview(previewHost, filePath);
  updateDocxPreviewInteractivity(bar);
}

/** @returns {HTMLElement} */
function createChatgptPaneDownloadStatusBarElement() {
  const bar = document.createElement("div");
  bar.className = "chatgpt-pane-download-status chatgpt-pane-download-status--idle";
  bar.dataset.statusState = "idle";

  const loading = document.createElement("span");
  loading.className = "chatgpt-pane-download-status-loading";
  loading.hidden = true;
  loading.setAttribute("aria-hidden", "true");
  const spinner = document.createElement("span");
  spinner.className = "chatgpt-pane-download-status-spinner";
  loading.appendChild(spinner);

  const content = document.createElement("div");
  content.className = "chatgpt-pane-download-status-content";
  content.hidden = true;

  const text = document.createElement("div");
  text.className = "chatgpt-pane-download-status-text";

  const actions = document.createElement("div");
  actions.className = "chatgpt-pane-download-status-actions";

  const folderBtn = document.createElement("button");
  folderBtn.type = "button";
  folderBtn.className = "btn-icon chatgpt-pane-download-open-folder";
  folderBtn.innerHTML = CHATGPT_PANE_FOLDER_OPEN_ICON_SVG;
  folderBtn.title = "Open folder";
  folderBtn.setAttribute("aria-label", "Open folder");

  const resumeBtn = document.createElement("button");
  resumeBtn.type = "button";
  resumeBtn.className = "btn-icon chatgpt-pane-download-open-resume";
  resumeBtn.innerHTML = CHATGPT_PANE_RESUME_FILE_ICON_SVG;
  resumeBtn.title = "Open resume (.docx)";
  resumeBtn.setAttribute("aria-label", "Open resume");

  const jdBtn = document.createElement("button");
  jdBtn.type = "button";
  jdBtn.className = "btn-icon chatgpt-pane-download-open-jd";
  jdBtn.innerHTML = CHATGPT_PANE_JD_FILE_ICON_SVG;
  jdBtn.title = "Open JD (.txt)";
  jdBtn.setAttribute("aria-label", "Open JD");
  jdBtn.hidden = true;

  actions.append(folderBtn, resumeBtn, jdBtn);

  const previewWrap = document.createElement("div");
  previewWrap.className = "chatgpt-pane-download-resume-preview";
  previewWrap.hidden = true;

  const previewHead = document.createElement("div");
  previewHead.className = "chatgpt-pane-download-resume-preview-head";

  const previewLabel = document.createElement("div");
  previewLabel.className = "chatgpt-pane-download-resume-preview-label";
  previewLabel.textContent = "Document preview";

  const previewHost = document.createElement("div");
  previewHost.className = "chatgpt-pane-download-resume-preview-body docx-resume-preview-body";
  previewHost.setAttribute("aria-label", "Resume preview");

  previewHead.append(previewLabel);
  previewWrap.append(previewHead, previewHost);
  content.append(text, actions, previewWrap);
  bar.append(loading, content);
  bindDownloadStatusBarActions(bar);
  return bar;
}

/** @param {HTMLElement} pane */
function isDedicatedBrowserChatgptPane(pane) {
  return !!(dedicatedPromptKind && pane.getAttribute("data-browser-tab"));
}

/** @returns {HTMLElement | null} */
function getDedicatedJdDownloadStatusHost() {
  if (!dedicatedPromptKind) {
    return null;
  }
  let host = document.getElementById("dedicatedJdDownloadStatusHost");
  if (!host) {
    host = document.createElement("div");
    host.id = "dedicatedJdDownloadStatusHost";
    host.className = "dedicated-jd-download-status-host";
    host.setAttribute("aria-live", "polite");
    host.hidden = true;
  }
  repositionDedicatedJdDownloadStatusHost();
  return host;
}

/** Keep download status near build preview settings. */
function repositionDedicatedJdDownloadStatusHost() {
  const host = document.getElementById("dedicatedJdDownloadStatusHost");
  if (!host) {
    return;
  }
  if (dedicatedPromptDirsPanel?.parentElement) {
    if (host.previousElementSibling !== dedicatedPromptDirsPanel) {
      dedicatedPromptDirsPanel.insertAdjacentElement("afterend", host);
    }
    return;
  }
  const block = getDedicatedBuildPreviewBlock();
  if (block && host.parentElement !== block) {
    block.appendChild(host);
  }
}

function parkDedicatedJdDownloadStatusHost() {
  const host = document.getElementById("dedicatedJdDownloadStatusHost");
  if (!host) {
    return;
  }
  repositionDedicatedJdDownloadStatusHost();
  host.hidden = true;
  host.querySelectorAll(".chatgpt-pane-download-status").forEach((node) => {
    const bar = /** @type {HTMLElement} */ (node);
    bar.hidden = true;
  });
}

function mountDedicatedJdDownloadStatusHostIn(parent) {
  if (!parent) {
    return;
  }
  const host = getDedicatedJdDownloadStatusHost();
  if (!host) {
    return;
  }
  if (dedicatedPromptDirsPanel && dedicatedPromptDirsPanel.parentElement === parent) {
    dedicatedPromptDirsPanel.insertAdjacentElement("afterend", host);
  } else if (host.parentElement !== parent) {
    parent.appendChild(host);
  }
}

function syncDedicatedJdDownloadStatusVisibility() {
  const host = document.getElementById("dedicatedJdDownloadStatusHost");
  if (!host) {
    return;
  }
  const onBrowser =
    !!dedicatedPromptKind && dedicatedActiveWorkspacePage !== "prompt";
  const activeId = onBrowser ? dedicatedActiveWorkspacePage : null;
  let anyVisible = false;
  host.querySelectorAll(".chatgpt-pane-download-status").forEach((node) => {
    const bar = /** @type {HTMLElement} */ (node);
    const tabId = bar.getAttribute("data-dedicated-browser-tab");
    const state = bar.dataset.statusState || "idle";
    const show = !!activeId && tabId === activeId && state !== "idle";
    bar.hidden = !show;
    if (show) {
      anyVisible = true;
    }
  });
  host.hidden = !onBrowser || !anyVisible;
}

/** @param {HTMLElement} pane */
function ensureChatgptPaneDownloadStatus(pane) {
  const browserTabId = pane.getAttribute("data-browser-tab");
  if (isDedicatedBrowserChatgptPane(pane) && browserTabId) {
    pane.querySelector(".chatgpt-pane-download-status")?.remove();
    const host = getDedicatedJdDownloadStatusHost();
    if (!host) {
      return /** @type {HTMLElement} */ (pane);
    }
    let bar = host.querySelector(
      `.chatgpt-pane-download-status[data-dedicated-browser-tab="${CSS.escape(browserTabId)}"]`,
    );
    if (bar && !bar.querySelector(".chatgpt-pane-download-status-content")) {
      bar.remove();
      bar = null;
    }
    if (bar && !bar.querySelector(".chatgpt-pane-download-status-actions")) {
      bar.remove();
      bar = null;
    }
    if (!bar) {
      bar = createChatgptPaneDownloadStatusBarElement();
      bar.setAttribute("data-dedicated-browser-tab", browserTabId);
      host.appendChild(bar);
    } else {
      bindDownloadStatusBarActions(/** @type {HTMLElement} */ (bar));
    }
    syncDedicatedJdDownloadStatusVisibility();
    return /** @type {HTMLElement} */ (bar);
  }

  let bar = pane.querySelector(".chatgpt-pane-download-status");
  if (bar && !bar.querySelector(".chatgpt-pane-download-status-content")) {
    bar.remove();
    bar = null;
  }
  if (bar && !bar.querySelector(".chatgpt-pane-download-status-actions")) {
    bar.remove();
    bar = null;
  }
  if (!bar) {
    bar = createChatgptPaneDownloadStatusBarElement();
    const header = pane.querySelector(".chatgpt-pane-header");
    const wordDirBar = pane.querySelector(".chatgpt-pane-word-dir-bar");
    if (wordDirBar) {
      wordDirBar.insertAdjacentElement("afterend", bar);
    } else if (header) {
      header.insertAdjacentElement("afterend", bar);
    } else {
      pane.insertBefore(bar, pane.firstChild);
    }
  } else {
    bindDownloadStatusBarActions(/** @type {HTMLElement} */ (bar));
  }
  return /** @type {HTMLElement} */ (bar);
}

/**
 * @param {HTMLElement} bar
 * @param {"idle" | "loading" | "success" | "error"} state
 * @param {{ who?: string, savedName?: string, folderPath?: string, jdFileName?: string, error?: string }} [data]
 */
function renderChatgptPaneDownloadStatusBar(bar, state, data = {}) {
  bar.dataset.statusState = state;
  bar.classList.remove(
    "chatgpt-pane-download-status--idle",
    "chatgpt-pane-download-status--loading",
    "chatgpt-pane-download-status--success",
    "chatgpt-pane-download-status--error",
  );
  bar.classList.add(`chatgpt-pane-download-status--${state}`);

  const loading = bar.querySelector(".chatgpt-pane-download-status-loading");
  const content = bar.querySelector(".chatgpt-pane-download-status-content");
  const textEl = bar.querySelector(".chatgpt-pane-download-status-text");

  if (state === "loading") {
    if (loading instanceof HTMLElement) {
      loading.hidden = false;
    }
    if (content instanceof HTMLElement) {
      content.hidden = true;
    }
    if (textEl instanceof HTMLElement) {
      textEl.replaceChildren();
    }
    return;
  }

  if (loading instanceof HTMLElement) {
    loading.hidden = true;
  }

  if (state === "success") {
    if (content instanceof HTMLElement) {
      content.hidden = false;
    }
    if (textEl instanceof HTMLElement) {
      fillGptDownloadStatusEl(
        textEl,
        String(data.savedName || ""),
        String(data.folderPath || ""),
        String(data.jdFileName || ""),
        String(data.who || ""),
      );
    }
    const jdBtn = bar.querySelector(".chatgpt-pane-download-open-jd");
    if (jdBtn instanceof HTMLElement) {
      jdBtn.hidden = !String(bar.dataset.statusJdPath || "").trim();
    }
    const pane = resolvePaneForDownloadStatusBar(bar);
    void refreshDownloadStatusResumePreview(bar, pane);
    return;
  }

  const previewWrap = bar.querySelector(".chatgpt-pane-download-resume-preview");
  if (previewWrap instanceof HTMLElement) {
    previewWrap.hidden = true;
  }

  if (content instanceof HTMLElement) {
    content.hidden = true;
  }
  if (textEl instanceof HTMLElement) {
    textEl.replaceChildren();
  }
}

/** @param {HTMLElement} pane */
function updateChatgptPaneDownloadStatusUi(pane) {
  const bar = ensureChatgptPaneDownloadStatus(pane);
  const ctx = getChatgptPaneContextFromPane(pane);
  const busy = isGptBusyForSendTarget(ctx.profile, { browserTabId: ctx.browserTabId });
  if (busy) {
    renderChatgptPaneDownloadStatusBar(bar, "idle");
    syncDedicatedJdDownloadStatusVisibility();
    return;
  }
  const state = bar.dataset.statusState || "idle";
  if (state === "success") {
    renderChatgptPaneDownloadStatusBar(bar, "success", {
      who: bar.dataset.statusWho || "",
      savedName: bar.dataset.statusSavedName || "",
      folderPath: bar.dataset.statusFolderPath || "",
      jdFileName: bar.dataset.statusJdFileName || "",
    });
    syncDedicatedJdDownloadStatusVisibility();
    return;
  }
  if (state === "error") {
    renderChatgptPaneDownloadStatusBar(bar, "idle");
    syncDedicatedJdDownloadStatusVisibility();
    return;
  }
  renderChatgptPaneDownloadStatusBar(bar, "idle");
  syncDedicatedJdDownloadStatusVisibility();
}

/**
 * @param {{ paneEl?: HTMLElement | null, browserTabId?: string | null, downloadKey?: string }} opts
 * @returns {HTMLElement | null}
 */
function resolvePaneForDownloadStatus(opts = {}) {
  if (opts.paneEl) {
    return opts.paneEl;
  }
  const browserTabId = opts.browserTabId ?? null;
  if (browserTabId) {
    return dedicatedBrowserTabs.find((t) => t.id === browserTabId)?.paneEl ?? null;
  }
  const downloadKey = opts.downloadKey ?? null;
  if (downloadKey) {
    return document.querySelector(
      `#chatgptPanesHost .chatgpt-pane[data-prompt="${CSS.escape(String(downloadKey))}"]`,
    );
  }
  const kind = getSelectedPromptKind();
  if (kind) {
    return document.querySelector(
      `#chatgptPanesHost .chatgpt-pane[data-prompt="${CSS.escape(kind)}"]`,
    );
  }
  return null;
}

/**
 * @param {Record<string, unknown>} r
 * @param {string} who
 * @param {{ paneEl?: HTMLElement | null, browserTabId?: string | null, downloadKey?: string }} [opts]
 */
function showGptDownloadStatus(r, who, opts = {}) {
  const folderPath = String(r.folderPath || "").trim();
  const filePath = String(r.filePath || "").trim();
  const savedName = fileNameFromPath(filePath);
  if (!folderPath && !filePath) {
    return;
  }
  const pane = resolvePaneForDownloadStatus(opts);
  if (!pane) {
    return;
  }
  lastGptDownloadOpenTarget = { folderPath, filePath, paneEl: pane };

  const bar = ensureChatgptPaneDownloadStatus(pane);
  bar.dataset.statusWho = who;
  bar.dataset.statusSavedName = savedName;
  bar.dataset.statusFolderPath = folderPath;
  bar.dataset.statusFilePath = filePath;
  bar.dataset.statusJdPath = r.jdSaved ? String(r.jdPath || "").trim() : "";
  bar.dataset.statusJdFileName = r.jdSaved
    ? fileNameFromPath(String(r.jdPath || "JD.txt"))
    : "";
  delete bar.dataset.statusError;
  bar.dataset.statusState = "success";
  updateChatgptPaneDownloadStatusUi(pane);
}

/** @param {string} message */
function showGptDownloadOpenFolderError(message) {
  const msg = String(message || "Could not open folder.").trim();
  toast(msg, "error", 5500);
}

async function openLastGptDownloadFolder() {
  const api = window.promptApi;
  if (!lastGptDownloadOpenTarget) {
    showGptDownloadOpenFolderError("No download folder to open.");
    return;
  }
  const target = lastGptDownloadOpenTarget;
  try {
    let openR = null;
    if (typeof api?.openGptDownloadFolder === "function") {
      openR = await api.openGptDownloadFolder(target);
    } else if (typeof api?.showPathInFolder === "function" && target.filePath) {
      openR = await api.showPathInFolder(target.filePath);
    } else {
      showGptDownloadOpenFolderError(
        "Restart the app to enable Open folder (reload main process).",
      );
      return;
    }
    if (!openR?.ok) {
      showGptDownloadOpenFolderError(openR?.error || "Could not open folder.");
    }
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    if (/no handler registered/i.test(msg)) {
      showGptDownloadOpenFolderError(
        "Restart the app fully (quit and run again) to enable Open folder.",
      );
      return;
    }
    showGptDownloadOpenFolderError(msg);
  }
}

function syncEditorToMemory() {
  if (editMode !== "builtin") return;
  setStrForEdit(editSelection, editTemplate.value);
}

function clearEditFileNavActive() {
  document.querySelectorAll(".edit-nav .edit-file-btn[data-edit]").forEach((btn) => {
    btn.classList.remove("edit-file-active");
    btn.setAttribute("aria-selected", "false");
  });
}

/** @param {string} key */
function setEditFileNavActive(key) {
  clearEditFileNavActive();
  const btn = document.querySelector(
    `.edit-nav .edit-file-btn[data-edit="${CSS.escape(key)}"]`,
  );
  if (btn) {
    btn.classList.add("edit-file-active");
    btn.setAttribute("aria-selected", "true");
  }
  syncPromptCardHighlightFromKind(key);
  syncMainChatgptPaneHighlights();
}

/**
 * @param {string} path absolute path
 * @param {string} fileUrl file:// URL for the field
 * @param {string} content file text
 * @param {{ preserveSettingsKind?: boolean }} [opts]
 */
function applyExternalEditor(path, fileUrl, content, opts = {}) {
  const el = getEditTemplateEl();
  if (editMode === "builtin" && el) {
    setStrForEdit(editSelection, el.value);
  }
  editMode = "external";
  externalAbsPath = path;
  if (!opts.preserveSettingsKind) {
    editPromptSettingsKind = resolveEditPromptKindForPath(path);
  }
  if (editFileUrlInput) {
    editFileUrlInput.value = fileUrl;
  }
  applyPromptEditorContent(content);
  const kind = getEditPromptSettingsKind();
  if (kind) {
    syncEditPanelHeading(kind);
    if (editPanelTitle) {
      editPanelTitle.title = fileNameFromPath(path);
    }
    setEditFileNavActive(kind);
  } else {
    syncEditPanelHeading(null);
    if (editPanelTitle) {
      editPanelTitle.title = "";
    }
    clearEditFileNavActive();
  }
  loadEditPromptSettings();
}

/** @param {"other" | "justin"} key */
async function openEditKey(key) {
  if (!isBuiltinPromptKind(key) || !isPromptKindAvailable(key)) {
    return;
  }
  if (editMode === "builtin" && isBuiltinPromptKind(editSelection) && getEditTemplateEl()) {
    setStrForEdit(editSelection, getEditTemplateEl().value);
  } else if (
    editMode === "external" &&
    isBuiltinPromptKind(editPromptSettingsKind) &&
    getEditTemplateEl()
  ) {
    setStrForEdit(editPromptSettingsKind, getEditTemplateEl().value);
  }
  await loadPromptIntoEditor(key);
}

function ensureMainWindowTabLayout() {
  if (dedicatedPromptKind) {
    return;
  }
  if (tabsListDedicated) {
    tabsListDedicated.hidden = true;
  }
  if (dedicatedAddBrowserTabBtn) {
    dedicatedAddBrowserTabBtn.hidden = true;
  }
  if (pageBrowser) {
    pageBrowser.hidden = true;
    pageBrowser.classList.add("page-hidden");
    pageBrowser.setAttribute("aria-hidden", "true");
  }
}

/** @returns {"olu" | "justin"} */
function getDedicatedGptProfile() {
  if (isCustomPromptKind(dedicatedPromptKind)) {
    return getCustomPromptByKind(dedicatedPromptKind)?.chatgptProfile ?? "olu";
  }
  return dedicatedPromptKind === "justin" ? "justin" : "olu";
}

/** @param {"olu" | "justin"} profile */
function getChatgptPartitionForProfile(profile) {
  return profile === "justin" ? "persist:chatgpt-justin" : "persist:chatgpt";
}

function getChatgptWebviewPreloadPath() {
  return typeof window.pbPaths !== "undefined" &&
    window.pbPaths &&
    window.pbPaths.chatgptWebviewZoomPreload
    ? window.pbPaths.chatgptWebviewZoomPreload
    : "";
}

/**
 * @param {"olu" | "justin"} profile
 * @param {string} tabId
 */
function getDedicatedChatgptPartition() {
  if (dedicatedPromptKind) {
    return chatgptPartitionForPromptId(dedicatedPromptKind);
  }
  return getChatgptPartitionForProfile(getDedicatedGptProfile());
}

function createDedicatedBrowserWebview(profile, tabId) {
  const wv = document.createElement("webview");
  wv.className = "chatgpt-webview";
  wv.setAttribute("partition", getDedicatedChatgptPartition());
  wv.setAttribute("allowpopups", "");
  wv.setAttribute(
    "webpreferences",
    "contextIsolation=yes, nativeWindowOpen=yes, javascript=yes",
  );
  wv.id = `chatgptWebviewDedicated-${tabId}`;
  bindChatgptWebviewChrome(wv);
  return wv;
}

/**
 * @param {"olu" | "justin"} profile
 * @param {string} tabId
 */
function createDedicatedBrowserPane(profile, tabId) {
  const label = dedicatedPromptKind
    ? promptLabelForKind(dedicatedPromptKind)
    : profile === "justin"
      ? "Justin"
      : "Oluwaseun";
  const pane = document.createElement("section");
  pane.className = "chatgpt-pane chatgpt-pane--embedded";
  pane.setAttribute("data-chatgpt-profile", profile);
  pane.setAttribute("data-browser-tab", tabId);
  pane.setAttribute("aria-label", `ChatGPT — ${label}`);
  if (dedicatedPromptKind) {
    pane.dataset.prompt = dedicatedPromptKind;
  }

  const header = document.createElement("div");
  header.className = "chatgpt-pane-header";

  const actions = document.createElement("div");
  actions.className = "chatgpt-pane-actions";

  const refreshBtn = document.createElement("button");
  refreshBtn.type = "button";
  refreshBtn.className = "btn-icon-reload chatgpt-pane-refresh";
  refreshBtn.setAttribute("data-chatgpt-profile", profile);
  refreshBtn.setAttribute("data-dedicated-browser-tab", tabId);
  refreshBtn.title = "Refresh ChatGPT";
  refreshBtn.setAttribute("aria-label", "Refresh ChatGPT");
  refreshBtn.innerHTML =
    '<svg xmlns="http://www.w3.org/2000/svg" width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true"><path d="M21 12a9 9 0 1 1-6.219-8.56"/><polyline points="21 3 21 9 15 9"/></svg>';

  actions.appendChild(refreshBtn);
  header.appendChild(actions);
  pane.appendChild(header);
  appendChatgptPaneToolbar(pane, {
    browserTabId: tabId,
    profile,
  });
  ensureChatgptPaneHeaderMeta(pane);
  ensureChatgptPaneDownloadStatus(pane);
  updateChatgptPaneDownloadStatusUi(pane);
  pane.appendChild(createDedicatedBrowserWebview(profile, tabId));
  return pane;
}

/** @param {HTMLElement} btn */
function getChatgptPartitionForPaneButton(btn) {
  const tabId = btn.getAttribute("data-dedicated-browser-tab");
  if (tabId) {
    return getDedicatedChatgptPartition();
  }
  const promptId = btn.getAttribute("data-prompt");
  if (promptId) {
    return chatgptPartitionForPromptId(promptId);
  }
  const profile = btn.getAttribute("data-chatgpt-profile");
  return profile === "justin" ? "persist:chatgpt-justin" : "persist:chatgpt";
}

/** @param {string} partition */
function reloadChatgptWebviewsForPartition(partition) {
  document.querySelectorAll(".chatgpt-webview").forEach((node) => {
    const wv = /** @type {HTMLElement} */ (node);
    if (wv.getAttribute("partition") === partition) {
      delete wv.dataset.pbLoadRetry;
      loadChatgptWebviewIfNeeded(wv, { force: true });
    }
  });
}

/** @param {HTMLElement} btn */
async function openChatgptSignInForPaneButton(btn) {
  // Sign in inside the pane itself (same session/partition) — no popup window.
  const tabId = btn.getAttribute("data-dedicated-browser-tab");
  const promptId = btn.getAttribute("data-prompt");
  const profile = btn.getAttribute("data-chatgpt-profile");
  /** @type {HTMLElement | null} */
  let wv = null;
  if (tabId) {
    const entry = dedicatedBrowserTabs.find((t) => t.id === tabId);
    wv = entry?.paneEl?.querySelector(".chatgpt-webview") ?? null;
  } else if (promptId) {
    wv = getChatgptWebviewForKind(promptId);
  } else {
    wv = getChatgptWebviewForProfile(profile === "justin" ? "justin" : "olu");
  }
  navigateChatgptWebviewToEmailLogin(wv);
  toast(
    "Sign in here in the pane: choose email (not Google), then email → password.",
    "success",
    11000,
  );
}

/** @param {HTMLElement | null} wv */
function navigateChatgptWebviewToEmailLogin(wv) {
  if (!wv) {
    return;
  }
  bindChatgptWebviewChrome(wv);
  delete wv.dataset.pbLoadRetry;
  navigateChatgptWebviewTo(wv, CHATGPT_EMAIL_LOGIN_URL);
}

/** @param {string} tabId */
function dedicatedBrowserTabLabel(tabId) {
  const entry = dedicatedBrowserTabs.find((t) => t.id === tabId);
  const company = entry?.company?.trim();
  if (company) {
    return company;
  }
  const idx = dedicatedBrowserTabs.findIndex((t) => t.id === tabId);
  if (idx <= 0) {
    return "Browser";
  }
  return `Browser ${idx + 1}`;
}

/** @param {string} tabId */
function refreshDedicatedBrowserTabLabels() {
  dedicatedBrowserTabs.forEach((entry) => {
    entry.tabEl.setAttribute("aria-label", dedicatedBrowserTabLabel(entry.id));
    refreshDedicatedBrowserTabResultBadge(entry.id);
  });
}

/**
 * @param {{ reuseExistingPane?: boolean, activate?: boolean }} [opts]
 * @returns {string | null}
 */
function addDedicatedBrowserTab(opts = {}) {
  if (!dedicatedBrowserPages || !dedicatedBrowserTabStrip) {
    return null;
  }
  const profile = getDedicatedGptProfile();
  dedicatedBrowserTabSeq += 1;
  const tabId = `browser-${dedicatedBrowserTabSeq}`;

  const pageEl = document.createElement("div");
  pageEl.className = "dedicated-browser-tab-page page-hidden";
  pageEl.dataset.browserTabId = tabId;
  pageEl.hidden = true;
  pageEl.setAttribute("role", "tabpanel");

  const host = document.createElement("div");
  host.className = "dedicated-browser-host";
  pageEl.appendChild(host);
  dedicatedBrowserPages.appendChild(pageEl);

  let pane = null;
    if (opts.reuseExistingPane) {
    pane = document.querySelector(`.chatgpt-pane[data-chatgpt-profile="${profile}"]`);
    if (pane) {
      host.appendChild(pane);
      pane.hidden = false;
      pane.classList.add("chatgpt-pane--embedded");
      pane.setAttribute("data-browser-tab", tabId);
      appendChatgptPaneToolbar(/** @type {HTMLElement} */ (pane), {
        browserTabId: tabId,
        profile,
      });
    }
  }
  if (!pane) {
    pane = createDedicatedBrowserPane(profile, tabId);
    host.appendChild(pane);
  }

  const tabEl = document.createElement("span");
  tabEl.className = "tab tab-chrome tab-dedicated-browser";
  tabEl.setAttribute("role", "tab");
  tabEl.dataset.browserTabId = tabId;
  tabEl.tabIndex = -1;
  tabEl.setAttribute("aria-controls", "pageBrowser");

  const iconSpan = document.createElement("span");
  iconSpan.className = "tab-chrome-icon tab-dedicated-browser-icon";
  iconSpan.setAttribute("aria-hidden", "true");
  iconSpan.innerHTML = CHROME_TAB_BROWSER_ICON_SVG;
  tabEl.appendChild(iconSpan);

  const labelSpan = document.createElement("span");
  labelSpan.className = "tab-chrome-label tab-dedicated-browser-label";
  labelSpan.textContent = "Browser";
  tabEl.appendChild(labelSpan);

  const statusSpan = document.createElement("span");
  statusSpan.className = "tab-dedicated-browser-status";
  statusSpan.setAttribute("aria-hidden", "true");
  tabEl.appendChild(statusSpan);

  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "tab-dedicated-browser-close";
  closeBtn.setAttribute("aria-label", "Close browser tab");
  closeBtn.title = "Close";
  closeBtn.textContent = "×";
  tabEl.appendChild(closeBtn);

  dedicatedBrowserTabStrip.appendChild(tabEl);

  const entry = {
    id: tabId,
    profile,
    pageEl,
    tabEl,
    paneEl: pane,
    jd: "",
    company: "",
  };
  loadDedicatedBrowserTabFormStateFromStorage(entry);
  dedicatedBrowserTabs.push(entry);
  refreshDedicatedBrowserTabLabels();
  refreshDedicatedBrowserTabCloseButtons();

  tabEl.addEventListener("click", (e) => {
    if (e.target instanceof HTMLElement && e.target.closest(".tab-dedicated-browser-close")) {
      return;
    }
    activateDedicatedBrowserTabById(tabId);
  });
  closeBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    closeDedicatedBrowserTab(tabId);
  });

  if (opts.activate !== false) {
    activateDedicatedBrowserTabById(tabId);
  }
  return tabId;
}

function refreshDedicatedBrowserTabCloseButtons() {
  /* Close buttons use Chrome-style hover visibility in CSS. */
}

/** @param {string} tabId */
function closeDedicatedBrowserTab(tabId) {
  const idx = dedicatedBrowserTabs.findIndex((t) => t.id === tabId);
  if (idx < 0) {
    return;
  }
  const wasActiveTab = dedicatedActiveWorkspacePage === tabId;
  const [removed] = dedicatedBrowserTabs.splice(idx, 1);
  removed.jd = "";
  removed.company = "";
  clearDedicatedBrowserTabStorage(removed.id);
  if (gptResultModalContext.browserTabId === removed.id) {
    gptResultModalContext.browserTabId = null;
    hideGptResultModal();
  }
  if (
    lastGptDownloadOpenTarget &&
    (lastGptDownloadOpenTarget.paneEl === removed.paneEl ||
      lastGptDownloadOpenTarget.paneEl?.getAttribute?.("data-browser-tab") === removed.id)
  ) {
    lastGptDownloadOpenTarget = null;
  }
  delete gptBrowserTabBusy[removed.id];
  delete gptBrowserTabDownloadDone[removed.id];
  delete lastGptResultsByBrowserTab[removed.id];
  removed.paneEl?.classList.remove("is-generating");
  removed.tabEl.remove();
  removed.pageEl.remove();
  document
    .getElementById("dedicatedJdDownloadStatusHost")
    ?.querySelector(
      `.chatgpt-pane-download-status[data-dedicated-browser-tab="${CSS.escape(removed.id)}"]`,
    )
    ?.remove();
  syncDedicatedJdDownloadStatusVisibility();
  refreshDedicatedBrowserTabLabels();
  refreshDedicatedBrowserTabCloseButtons();
  if (wasActiveTab) {
    const next = dedicatedBrowserTabs[Math.min(idx, dedicatedBrowserTabs.length - 1)];
    if (next) {
      activateDedicatedBrowserTabById(next.id);
    } else {
      if (jdInput) {
        jdInput.value = "";
        syncBuildInputShellState(jdInput);
      }
      if (dedicatedCompanyInput) {
        dedicatedCompanyInput.value = "";
        syncDedicatedCompanyShell();
      }
      activateDedicatedPromptTab();
    }
  }
}

function initDedicatedBrowserTabs() {
  if (dedicatedBrowserTabs.length > 0) {
    return;
  }
  addDedicatedBrowserTab({ reuseExistingPane: true, activate: false });
  if (dedicatedAddBrowserTabBtn) {
    dedicatedAddBrowserTabBtn.hidden = false;
  }
  refreshDedicatedBrowserTabCloseButtons();
}

/** @param {"prompt" | string} which — "prompt" or browser tab id */
function showDedicatedPage(which) {
  if (dedicatedPromptKind) {
    const prev = dedicatedActiveWorkspacePage;
    if (
      prev !== "prompt" &&
      prev !== which &&
      dedicatedBrowserTabs.some((t) => t.id === prev)
    ) {
      saveDedicatedBrowserTabFormState(prev);
    }
  }
  dedicatedActiveWorkspacePage = which;
  const isPrompt = which === "prompt";

  if (dedicatedPromptKind) {
    setDedicatedAnimatedPageVisible(pageBuild, isPrompt);
    setDedicatedAnimatedPageVisible(pageBrowser, !isPrompt);
  } else {
    if (pageBuild) {
      pageBuild.classList.toggle("page-hidden", !isPrompt);
      pageBuild.hidden = !isPrompt;
      pageBuild.setAttribute("aria-hidden", String(!isPrompt));
    }
    if (pageBrowser) {
      pageBrowser.classList.toggle("page-hidden", isPrompt);
      pageBrowser.hidden = isPrompt;
      pageBrowser.setAttribute("aria-hidden", String(isPrompt));
    }
  }

  dedicatedBrowserTabs.forEach((entry) => {
    const active = !isPrompt && entry.id === which;
    if (dedicatedPromptKind) {
      setDedicatedAnimatedPageVisible(entry.pageEl, active);
    } else {
      entry.pageEl.classList.toggle("page-hidden", !active);
      entry.pageEl.hidden = !active;
      entry.pageEl.setAttribute("aria-hidden", String(!active));
    }
    entry.tabEl.setAttribute("aria-selected", String(active));
    entry.tabEl.tabIndex = active ? 0 : -1;
    entry.tabEl.classList.toggle("tab-active", active);
  });

  if (tabDedicatedPrompt) {
    tabDedicatedPrompt.setAttribute("aria-selected", String(isPrompt));
    tabDedicatedPrompt.tabIndex = isPrompt ? 0 : -1;
    tabDedicatedPrompt.classList.toggle("tab-active", isPrompt);
  }

  if (dedicatedPromptKind) {
    document.body.classList.toggle("dedicated-workspace-page-prompt", isPrompt);
    document.body.classList.toggle("dedicated-workspace-page-browser", !isPrompt);
    if (isPrompt) {
      mountDedicatedPromptTabPanels();
    } else {
      mountDedicatedBrowserTabPanels();
      loadDedicatedBrowserTabFormState(which);
      const entry = dedicatedBrowserTabs.find((t) => t.id === which);
      requestAnimationFrame(() => {
        const wv = entry?.paneEl?.querySelector(".chatgpt-webview");
        loadChatgptWebviewIfNeeded(/** @type {HTMLElement} */ (wv));
      });
    }
    applyDedicatedDownloadDirVisibility();
    ensureDedicatedJdToolbar();
    updateDedicatedJdToolbarContext();
    syncDedicatedJdToolbarUi();
    syncDedicatedJdDownloadStatusVisibility();
    updatePreviewCopyUi();
    updateGptResultReadyUi();
  }
}

/** @param {string} tabId */
function activateDedicatedBrowserTabById(tabId) {
  if (!dedicatedBrowserTabs.some((t) => t.id === tabId)) {
    return;
  }
  showDedicatedPage(tabId);
  const entry = dedicatedBrowserTabs.find((t) => t.id === tabId);
  entry?.tabEl.focus();
}

function cycleDedicatedWorkspaceTabs() {
  const order = ["prompt", ...dedicatedBrowserTabs.map((t) => t.id)];
  if (!order.length) {
    return;
  }
  const current = dedicatedActiveWorkspacePage;
  const idx = order.indexOf(current);
  const next = order[(idx + 1) % order.length];
  if (next === "prompt") {
    activateDedicatedPromptTab();
  } else {
    activateDedicatedBrowserTabById(next);
  }
}

function showPage(which) {
  if (dedicatedPromptKind) {
    if (which === "browser") {
      const active =
        dedicatedActiveWorkspacePage !== "prompt"
          ? dedicatedActiveWorkspacePage
          : dedicatedBrowserTabs[0]?.id;
      if (active) {
        showDedicatedPage(active);
      } else {
        showDedicatedPage("prompt");
      }
    } else {
      showDedicatedPage("prompt");
    }
    return;
  }
  if (!pageBuild || !tabBuild) {
    return;
  }
  pageBuild.classList.remove("page-hidden");
  pageBuild.hidden = false;
  pageBuild.setAttribute("aria-hidden", "false");
  tabBuild.setAttribute("aria-selected", "true");
  tabBuild.tabIndex = 0;
  tabBuild.classList.add("tab-active");
  document.body.classList.add("workspace-tab-build");
}

function activateBuildTab() {
  if (isUsersManageModalOpen()) {
    closeUsersManageModal();
  } else if (editMode === "builtin") {
    setStrForEdit(editSelection, editTemplate.value);
  }
  showPage("build");
  (dedicatedPromptKind ? tabDedicatedPrompt : tabBuild)?.focus();
}

function activateDedicatedPromptTab() {
  showPage("prompt");
  tabDedicatedPrompt?.focus();
}

function activateDedicatedBrowserTab() {
  const tabId =
    dedicatedActiveWorkspacePage !== "prompt"
      ? dedicatedActiveWorkspacePage
      : dedicatedBrowserTabs[0]?.id;
  if (tabId) {
    activateDedicatedBrowserTabById(tabId);
  } else {
    addDedicatedBrowserTab({ activate: true });
  }
}

tabBuild?.addEventListener("click", activateBuildTab);

tabBuild?.addEventListener("keydown", (e) => {
  if (e.key === "Enter" || e.key === " ") {
    e.preventDefault();
    activateBuildTab();
  }
});

function setupMainCompanyInput() {
  mainCompanyInput?.addEventListener("input", () => {
    syncMainCompanyShell();
    persistMainCompanyName();
  });
}

function setupMainPromptWordDirectory() {
  mainPromptStructureBrowseBtn?.addEventListener("click", async () => {
    const kind = getSelectedPromptKind();
    if (!kind) {
      toast("Select a prompt first.", "error");
      return;
    }
    const api = window.promptApi;
    if (!api?.pickStructureDocx) {
      toast("Structure picker is not available.", "error");
      return;
    }
    try {
      const r = await api.pickStructureDocx();
      if (r?.canceled || !r?.path) {
        return;
      }
      await persistPromptStructureForKind(kind, r.path);
      await loadMainPromptStructureDirectory(kind);
      toast("Word skeleton directory updated.", "success", 4000);
    } catch (_e) {
      toast("Could not choose structure file.", "error");
    }
  });
  dedicatedPromptStructureBrowseBtn?.addEventListener("click", async () => {
    if (!dedicatedPromptKind) {
      return;
    }
    const api = window.promptApi;
    if (!api?.pickStructureDocx) {
      return;
    }
    try {
      const r = await api.pickStructureDocx();
      if (r?.canceled || !r?.path) {
        return;
      }
      await persistPromptStructureForKind(dedicatedPromptKind, r.path);
      await loadMainPromptStructureDirectory(dedicatedPromptKind);
      toast("Word skeleton directory updated.", "success", 4000);
    } catch (_e) {
      toast("Could not choose structure file.", "error");
    }
  });
}

function setupEditPromptSettings() {
  editPromptUserNameInput?.addEventListener("input", () => {
    syncEditUserHeader(getEditPromptSettingsKind());
    persistEditPromptSettingsFromInputs();
  });
  editPromptUserNameInput?.addEventListener("change", () => {
    persistEditPromptSettingsFromInputs();
  });
  editPromptPathDisplay?.addEventListener("click", () => {
    editPromptPathBrowseBtn?.click();
  });
  editPromptPathBrowseBtn?.addEventListener("click", async () => {
    const kind = getEditPromptSettingsKind();
    if (!kind) {
      toast("Select a user first.", "error");
      return;
    }
    const api = window.promptApi;
    if (!api?.pickPromptFile) {
      toast("Prompt picker is not available.", "error");
      return;
    }
    try {
      const r = await api.pickPromptFile();
      if (r?.canceled || !r?.path) {
        return;
      }
      const saved = await persistPromptPathForKind(kind, r.path);
      if (!saved.ok) {
        toast(saved.error || "Could not save prompt path.", "error", 6500);
        return;
      }
      applyPromptFilePickForKind(kind, r);
      toast("Prompt file updated.", "success", 4000);
    } catch (_e) {
      toast("Could not choose prompt file.", "error");
    }
  });

  editPromptStructureBrowseBtn?.addEventListener("click", async () => {
    const kind = getEditPromptSettingsKind();
    if (!kind) {
      return;
    }
    const api = window.promptApi;
    if (!api?.pickStructureDocx) {
      toast("Structure picker is not available.", "error");
      return;
    }
    try {
      const r = await api.pickStructureDocx();
      if (r?.canceled || !r?.path) {
        return;
      }
      await persistPromptStructureForKind(kind, r.path);
      if (editPromptStructureDisplay) {
        editPromptStructureDisplay.value = r.path;
        editPromptStructureDisplay.title = r.path;
      }
      renderEditPromptNav();
      await loadMainPromptStructureDirectory(kind);
      toast("Word skeleton directory updated.", "success", 4000);
    } catch (_e) {
      toast("Could not choose structure file.", "error");
    }
  });
}

function setupPromptPopupHotkeys() {
  editPromptPopupHotkeySetBtn?.addEventListener("click", () => {
    const kind = getEditPromptSettingsKind();
    if (!kind) {
      toast("Select a user first.", "error");
      return;
    }
    capturingPromptPopupHotkey = true;
    if (editPromptPopupHotkeySetBtn) {
      editPromptPopupHotkeySetBtn.textContent = "Listening…";
    }
    refreshEditPromptPopupHotkeyDisplay();
    editPromptPopupHotkeyDisplay?.focus();
  });

  editPromptPopupHotkeyClearBtn?.addEventListener("click", () => {
    const kind = getEditPromptSettingsKind();
    if (!kind) {
      return;
    }
    capturingPromptPopupHotkey = false;
    if (editPromptPopupHotkeySetBtn) {
      editPromptPopupHotkeySetBtn.textContent = "Set";
    }
    setPromptPopupHotkey(kind, "");
    refreshEditPromptPopupHotkeyDisplay();
    toast("Popup hotkey cleared.", "success", 2500);
  });

  document.addEventListener(
    "keydown",
    (e) => {
      if (!capturingPromptPopupHotkey) {
        return;
      }
      const kind = getEditPromptSettingsKind();
      if (!kind) {
        capturingPromptPopupHotkey = false;
        if (editPromptPopupHotkeySetBtn) {
          editPromptPopupHotkeySetBtn.textContent = "Set";
        }
        refreshEditPromptPopupHotkeyDisplay();
        return;
      }
      if (e.key === "Escape") {
        e.preventDefault();
        capturingPromptPopupHotkey = false;
        if (editPromptPopupHotkeySetBtn) {
          editPromptPopupHotkeySetBtn.textContent = "Set";
        }
        refreshEditPromptPopupHotkeyDisplay();
        return;
      }
      const combo = formatHotkeyFromKeyboardEvent(e);
      if (!combo) {
        return;
      }
      e.preventDefault();
      e.stopPropagation();
      capturingPromptPopupHotkey = false;
      if (editPromptPopupHotkeySetBtn) {
        editPromptPopupHotkeySetBtn.textContent = "Set";
      }
      const duplicate = findPromptKindWithPopupHotkey(combo, kind);
      if (duplicate) {
        toast(
          `${combo} is already assigned to ${getPromptUserName(duplicate)}.`,
          "error",
        );
        refreshEditPromptPopupHotkeyDisplay();
        return;
      }
      setPromptPopupHotkey(kind, combo);
      refreshEditPromptPopupHotkeyDisplay();
      void syncPromptPopupHotkeysToMain().then((r) => {
        const worksInBackground =
          r &&
          !r.failed?.includes(combo) &&
          (r.shortcutCombos?.includes(combo) || r.listenerCombos?.includes(combo));
        if (worksInBackground) {
          toast(
            `Popup hotkey set to ${combo} (works in the background; existing window is shown again).`,
            "success",
            5000,
          );
        } else {
          toast(
            `Popup hotkey set to ${combo}. Press it while this app is focused to open the window.`,
            "success",
            5000,
          );
        }
      });
    },
    true,
  );

  document.addEventListener(
    "keydown",
    (e) => {
      if (shouldIgnorePromptPopupHotkeyTrigger(e)) {
        return;
      }
      for (const p of listAllMainPrompts()) {
        const combo = getPromptPopupHotkey(p.id);
        if (!combo || !keyboardEventMatchesHotkeyCombo(e, combo)) {
          continue;
        }
        e.preventDefault();
        e.stopPropagation();
        triggerPromptPopupHotkeyForKind(p.id);
        return;
      }
    },
    true,
  );
}

function setupDedicatedCompanyInput() {
  dedicatedCompanyInput?.addEventListener("input", () => {
    syncDedicatedCompanyShell();
    if (dedicatedPromptKind && dedicatedActiveWorkspacePage !== "prompt") {
      saveDedicatedBrowserTabFormState(dedicatedActiveWorkspacePage);
    } else {
      persistDedicatedCompanyName();
    }
  });
}

function setupDedicatedWorkspaceTabs() {
  tabDedicatedPrompt?.addEventListener("click", activateDedicatedPromptTab);

  tabDedicatedPrompt?.addEventListener("keydown", (e) => {
    if (e.key === "ArrowRight") {
      e.preventDefault();
      activateDedicatedBrowserTab();
    }
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      activateDedicatedPromptTab();
    }
  });

  dedicatedAddBrowserTabBtn?.addEventListener("click", () => {
    addDedicatedBrowserTab({ activate: true });
  });
}

function setupWorkspaceCtrlTabShortcut() {
  document.addEventListener(
    "keydown",
    (e) => {
      if (!e.ctrlKey || e.key !== "Tab" || e.shiftKey) {
        return;
      }
      e.preventDefault();
      if (dedicatedPromptKind) {
        cycleDedicatedWorkspaceTabs();
      }
    },
    true,
  );
}

setupWorkspaceCtrlTabShortcut();

document.querySelector(".edit-nav")?.addEventListener("click", (e) => {
  const btn = e.target.closest(".edit-file-btn[data-edit]");
  if (!(btn instanceof HTMLElement) || btn.classList.contains("edit-file-btn--add")) {
    return;
  }
  const key = btn.getAttribute("data-edit");
  if (isBuiltinPromptKind(key)) {
    void openEditKey(/** @type {"other" | "justin"} */ (key));
  } else if (isCustomPromptKind(key)) {
    void openEditCustomPrompt(key);
  }
});

/** @param {string} kind */
async function openEditCustomPrompt(kind) {
  const entry = getCustomPromptByKind(kind);
  if (!entry) {
    toast("Prompt not found.", "error");
    return;
  }
  await loadPromptIntoEditor(kind);
}

/** @param {string} kind */
function setEditFileNavActiveCustom(kind) {
  setEditFileNavActive(kind);
}

async function loadCustomPromptTemplate(id) {
  const entry = customPrompts.find((p) => p.id === id);
  if (!entry) {
    return "";
  }
  try {
    const r = await tryReadPromptFileAtPath(getPromptPathForKind(id));
    if (isPromptReadSuccess(r)) {
      customTemplateStrById[id] = r.content;
      return r.content;
    }
    delete customTemplateStrById[id];
  } catch {
    delete customTemplateStrById[id];
  }
  return "";
}

async function loadAllCustomPromptTemplates() {
  await Promise.all(customPrompts.map((p) => loadCustomPromptTemplate(p.id)));
}

function renderEditPromptNav() {
  const nav = document.querySelector(".edit-nav");
  const addBtn = addPromptNavBtn;
  if (!nav || !addBtn) {
    return;
  }
  nav.querySelectorAll(".edit-file-btn[data-edit]").forEach((el) => {
    if (!el.classList.contains("edit-file-btn--add")) {
      el.remove();
    }
  });
  if (editCustomPromptNav) {
    editCustomPromptNav.hidden = true;
    editCustomPromptNav.textContent = "";
  }
  const activeKind = getEditPromptSettingsKind();
  for (const p of listAllMainPrompts()) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "edit-file-btn";
    if (isCustomPromptKind(p.id)) {
      btn.classList.add("edit-file-btn-custom");
    }
    btn.setAttribute("role", "tab");
    btn.setAttribute("aria-selected", p.id === activeKind ? "true" : "false");
    if (p.id === activeKind) {
      btn.classList.add("edit-file-active");
    }
    btn.dataset.edit = p.id;
    const hotkey = getPromptPopupHotkey(p.id);
    btn.innerHTML = `${userAvatarHtml(p.label, true)}<span class="edit-file-btn-text"><span class="edit-file-label">${escapeHtml(p.label)}</span>${hotkey ? `<span class="edit-file-hotkey">${escapeHtml(hotkey)}</span>` : ""}</span>`;
    btn.title = "Edit prompt and settings";
    nav.insertBefore(btn, addBtn);
  }
}

function renderCustomPromptUi() {
  renderEditPromptNav();
  renderBuildPromptCards();
  applyBuiltinPromptVisibilityToUi();
  syncEditLayoutEmptyState();
  syncMainWindowChatgptPanes();
}

async function refreshCustomPrompts() {
  const api = window.promptApi;
  if (!api?.getCustomPrompts) {
    return;
  }
  const r = await api.getCustomPrompts();
  const raw = Array.isArray(r?.prompts) ? r.prompts : [];
  /** @type {Map<string, CustomPromptEntry>} */
  const byId = new Map();
  for (const p of raw) {
    if (p?.id && !byId.has(p.id)) {
      byId.set(p.id, p);
    }
  }
  customPrompts = [...byId.values()];
  for (const id of Object.keys(customTemplateStrById)) {
    if (!customPrompts.some((p) => p.id === id)) {
      delete customTemplateStrById[id];
    }
  }
  await loadAllCustomPromptTemplates();
  renderCustomPromptUi();
}

function resetAddPromptForm() {
  addPromptForm?.reset();
  if (addPromptPathHidden) {
    addPromptPathHidden.value = "";
  }
  if (addPromptStructureHidden) {
    addPromptStructureHidden.value = "";
  }
  if (addPromptEditId) {
    addPromptEditId.value = "";
  }
  if (addPromptPathDisplay) {
    addPromptPathDisplay.value = "";
  }
  if (addPromptStructureDisplay) {
    addPromptStructureDisplay.value = "";
  }
  if (addPromptModalTitle) {
    addPromptModalTitle.textContent = "Add prompt";
  }
  if (addPromptDeleteBtn) {
    addPromptDeleteBtn.hidden = true;
  }
}

/** @param {CustomPromptEntry} [entry] */
function openAddPromptModal(entry) {
  if (!addPromptModal) {
    return;
  }
  resetAddPromptForm();
  if (entry) {
    if (addPromptModalTitle) {
      addPromptModalTitle.textContent = "Edit prompt";
    }
    if (addPromptEditId) {
      addPromptEditId.value = entry.id;
    }
    if (addPromptLabelInput) {
      addPromptLabelInput.value = entry.label;
    }
    if (addPromptPathHidden) {
      addPromptPathHidden.value = entry.promptPath;
    }
    if (addPromptStructureHidden) {
      addPromptStructureHidden.value = entry.structurePath;
    }
    if (addPromptPathDisplay) {
      addPromptPathDisplay.value = entry.promptPath;
    }
    if (addPromptStructureDisplay) {
      addPromptStructureDisplay.value = entry.structurePath;
    }
    if (addPromptDeleteBtn) {
      addPromptDeleteBtn.hidden = false;
    }
  }
  addPromptModal.hidden = false;
  document.body.classList.add("add-prompt-modal-open");
  addPromptLabelInput?.focus();
}

function closeAddPromptModal() {
  if (addPromptModal) {
    addPromptModal.hidden = true;
  }
  document.body.classList.remove("add-prompt-modal-open");
  resetAddPromptForm();
}

function setupAddPromptModal() {
  editPromptDeleteBtn?.addEventListener("click", () => {
    const kind = getEditPromptSettingsKind();
    if (!kind || !isPromptKindAvailable(kind)) {
      return;
    }
    void deletePromptForKind(kind).then((deleted) => {
      if (deleted && isUsersManageModalOpen()) {
        refreshEditUsersNav();
      }
    });
  });
  addPromptDeleteBtn?.addEventListener("click", () => {
    const id = addPromptEditId?.value?.trim() || "";
    if (!id || !isCustomPromptKind(id)) {
      return;
    }
    void deleteCustomPromptForKind(id).then((deleted) => {
      if (deleted) {
        closeAddPromptModal();
      }
    });
  });
  addPromptNavBtn?.addEventListener("click", () => openAddPromptModal());
  addPromptModalClose?.addEventListener("click", closeAddPromptModal);
  addPromptCancelBtn?.addEventListener("click", closeAddPromptModal);
  addPromptModal?.addEventListener("click", (e) => {
    if (e.target === addPromptModal) {
      closeAddPromptModal();
    }
  });

  addPromptBrowseTxtBtn?.addEventListener("click", async () => {
    const api = window.promptApi;
    if (!api?.pickPromptFile) {
      return;
    }
    try {
      const r = await api.pickPromptFile();
      if (r?.canceled) {
        return;
      }
      if (addPromptPathHidden) {
        addPromptPathHidden.value = r.path;
      }
      if (addPromptPathDisplay) {
        addPromptPathDisplay.value = r.path;
      }
    } catch (_e) {
      toast("Could not choose prompt file.", "error");
    }
  });

  addPromptBrowseDocxBtn?.addEventListener("click", async () => {
    const api = window.promptApi;
    if (!api?.pickStructureDocx) {
      return;
    }
    try {
      const r = await api.pickStructureDocx();
      if (r?.canceled) {
        return;
      }
      if (addPromptStructureHidden) {
        addPromptStructureHidden.value = r.path;
      }
      if (addPromptStructureDisplay) {
        addPromptStructureDisplay.value = r.path;
      }
    } catch (_e) {
      toast("Could not choose structure file.", "error");
    }
  });

  addPromptForm?.addEventListener("submit", async (e) => {
    e.preventDefault();
    const api = window.promptApi;
    if (!api?.saveCustomPrompt) {
      toast("Save is not available.", "error");
      return;
    }
    const label = addPromptLabelInput?.value?.trim() || "";
    const promptPath = addPromptPathHidden?.value?.trim() || "";
    const structurePath = addPromptStructureHidden?.value?.trim() || "";
    const id = addPromptEditId?.value?.trim() || "";
    try {
      const r = await api.saveCustomPrompt({
        id: id || undefined,
        label,
        promptPath,
        structurePath,
        chatgptProfile: "olu",
      });
      if (!r?.ok) {
        toast(String(r?.error || "Could not save prompt."), "error");
        return;
      }
      closeAddPromptModal();
      await refreshCustomPrompts();
      refreshBuildPreviewIfCardSelected();
      toast(id ? "Prompt updated." : "Prompt added.");
    } catch (err) {
      toast(String(err && err.message ? err.message : err), "error");
    }
  });
}

editFileBrowseBtn?.addEventListener("click", async () => {
  const api = window.promptApi;
  if (!api?.pickPromptFile) {
    return;
  }
  try {
    const r = await api.pickPromptFile();
    if (r?.canceled) {
      return;
    }
    applyExternalEditor(r.path, r.fileUrl, r.content);
    toast("File opened.");
  } catch (_e) {
    toast("Could not open file.", "error");
  }
});

editFileLoadBtn?.addEventListener("click", async () => {
  const raw = editFileUrlInput?.value?.trim();
  if (!raw) {
    toast("Enter a file path or file:// URL.", "error");
    return;
  }
  try {
    const r = await tryReadPromptFileAtPath(raw);
    if (!isPromptReadSuccess(r)) {
      toast("Load failed — file not found or unreadable.", "error");
      return;
    }
    applyExternalEditor(r.path, r.fileUrl || "", r.content);
    toast("File loaded.");
  } catch (_e) {
    toast("Load failed — check path or URL.", "error");
  }
});

editSaveBtn.addEventListener("click", async () => {
  if (editMode === "external") {
    const target = externalAbsPath || editFileUrlInput?.value?.trim();
    if (!target) {
      toast("No file path to save to.", "error");
      return;
    }
    try {
      await window.promptApi.savePromptExternal(target, editTemplate.value);
      externalAbsPath = target;
      const kind = resolveEditPromptKindForPath(target) || getEditPromptSettingsKind();
      const customEntry = customPrompts.find(
        (p) => normalizePathForCompare(p.promptPath) === normalizePathForCompare(target),
      );
      if (customEntry) {
        customTemplateStrById[customEntry.id] = editTemplate.value;
        refreshBuildPreviewIfCardSelected();
      } else if (isBuiltinPromptKind(kind)) {
        setStrForEdit(kind, editTemplate.value);
        refreshBuildPreviewIfCardSelected();
      }
      toast("Saved to file.");
    } catch (_e) {
      toast("Save failed.", "error");
    }
    return;
  }
  syncEditorToMemory();
  const id = getIdForEdit(editSelection);
  const body = getStrForEdit(editSelection);
  try {
    await window.promptApi.savePrompt(id, body);
    toast(`Saved: ${EDIT_LABELS[editSelection]}`);
  } catch (_e) {
    toast("Save failed.", "error");
  }
});

editReloadBtn.addEventListener("click", async () => {
  const kind = getEditPromptSettingsKind();
  if (kind) {
    const ok = await loadPromptIntoEditor(kind);
    if (ok) {
      toast("Reloaded from disk.");
    }
    return;
  }
  const raw = externalAbsPath || editFileUrlInput?.value?.trim();
  if (editMode === "external" && raw) {
    try {
      const r = await tryReadPromptFileAtPath(raw);
      if (!isPromptReadSuccess(r)) {
        toast("Reload failed — file not found or unreadable.", "error");
        return;
      }
      setPromptEditorContent(r.content);
      externalAbsPath = r.path;
      applyExternalEditor(r.path, r.fileUrl || "", r.content);
      toast("Reloaded from disk.");
    } catch (_e) {
      toast("Reload failed.", "error");
    }
    return;
  }
  if (!isBuiltinPromptKind(editSelection)) {
    toast("No file to reload.", "error");
    return;
  }
  const id = getIdForEdit(editSelection);
  try {
    const text = await window.promptApi.readPrompt(id);
    setStrForEdit(editSelection, text);
    setPromptEditorContent(text);
    toast(`Reloaded: ${EDIT_LABELS[editSelection]}`);
  } catch (_e) {
    toast("Reload failed.", "error");
  }
});

async function loadAllTemplates() {
  const api = window.promptApi;
  if (!api || typeof api.readPrompt !== "function") {
    throw new Error("promptApi unavailable");
  }
  await reloadPromptTemplateForKind("other", { silent: true });
  if (!otherTemplateStr) {
    otherTemplateStr = await api.readPrompt(OTHER_ID);
  }
  await reloadPromptTemplateForKind("justin", { silent: true });
  if (!justinTemplateStr) {
    justinTemplateStr = await api.readPrompt(JUSTIN_ID);
  }
  await refreshCustomPrompts();
}

async function reloadAllTemplatesFromDisk() {
  await loadAllTemplates();
  if (isUsersManageModalOpen()) {
    const kind = getEditPromptSettingsKind();
    if (isBuiltinPromptKind(kind)) {
      await openEditKey(kind);
    } else if (isCustomPromptKind(kind)) {
      await openEditCustomPrompt(kind);
    } else if (isBuiltinPromptKind(editSelection)) {
      await openEditKey(editSelection);
    }
  }
  toast("All templates reloaded from disk.");
  refreshBuildPreviewIfCardSelected();
}

function selectDefaultPromptCard() {
  const firstKind = listAllMainPrompts()[0]?.id;
  if (!firstKind) {
    return;
  }
  const first = document.querySelector(
    `.prompt-card[data-prompt="${CSS.escape(firstKind)}"]`,
  );
  if (!first) {
    return;
  }
  if (!dedicatedPromptKind) {
    loadMainCompanyName();
    void loadMainPreviewDirectories(firstKind);
  }
  document.querySelectorAll(".prompt-card").forEach((c) => {
    c.classList.remove("is-highlighted");
  });
  first.classList.add("is-highlighted");
  applyPreviewForCard(first, { showToast: false });
}

/** @type {string | null} */
let dedicatedPromptKind = null;

/** @returns {string | null} */
function getDedicatedPromptKind() {
  return dedicatedPromptKind;
}

/** @returns {string | null} */
function parseLaunchPromptKind() {
  /** @type {string | null} */
  let raw = null;
  try {
    if (window.location.search) {
      raw = new URLSearchParams(window.location.search).get("prompt");
    }
    if (!raw && window.location.href.includes("?")) {
      const q = window.location.href.split("?")[1]?.split("#")[0] ?? "";
      raw = new URLSearchParams(q).get("prompt");
    }
    if (!raw && window.location.hash) {
      const h = window.location.hash.replace(/^#/, "");
      if (h.includes("=")) {
        raw = new URLSearchParams(h.replace(/^\?/, "")).get("prompt");
      } else if (h === "other" || h === "justin") {
        raw = h;
      }
    }
  } catch (_e) {
    raw = null;
  }
  if (raw === "justin") {
    return "justin";
  }
  if (raw === "other") {
    return "other";
  }
  if (raw && raw.startsWith("custom:")) {
    return raw;
  }
  return null;
}

/**
 * @param {string} kind
 * @returns {boolean}
 */
function applyDedicatedPromptWindowLayout(kind) {
  if (!isBuiltinPromptKind(kind) && !isCustomPromptKind(kind)) {
    return false;
  }
  dedicatedPromptKind = kind;

  document.body.classList.add("dedicated-prompt-window");
  applyHeaderColorForScope(kind, getStoredHeaderColor(kind));
  document.body.dataset.promptWindow = kind;
  syncDedicatedWindowTitle(kind);

  document.querySelectorAll(".prompt-card[data-prompt]").forEach((card) => {
    const k = card.getAttribute("data-prompt");
    const isActive = k === kind;
    card.hidden = !isActive;
    card.setAttribute("aria-hidden", isActive ? "false" : "true");
    card.classList.toggle("is-highlighted", isActive);
  });

  document.querySelectorAll(".prompt-new-window-btn").forEach((btn) => {
    btn.hidden = true;
  });

  if (addPromptNavBtn) {
    addPromptNavBtn.hidden = true;
  }
  if (editCustomPromptNav) {
    editCustomPromptNav.hidden = true;
  }
  applyDedicatedDownloadDirVisibility();

  if (tabsListMain) {
    tabsListMain.hidden = true;
  }
  if (tabsListDedicated) {
    tabsListDedicated.hidden = false;
  }

  initDedicatedBrowserTabs();
  parkDedicatedBrowserOnlyPanels();
  mountDedicatedPromptTabPanels();
  document.body.classList.add("dedicated-workspace-page-prompt");
  document.body.classList.remove("dedicated-workspace-page-browser");

  document.querySelectorAll(".edit-file-btn[data-edit]").forEach((btn) => {
    const k = btn.getAttribute("data-edit");
    btn.hidden = k !== kind;
  });
  if (isBuiltinPromptKind(kind)) {
    editSelection = /** @type {"other" | "justin"} */ (kind);
  }

  ensureDedicatedJdToolbar();
  showDedicatedPage("prompt");
  return true;
}

/** @param {string} kind */
async function finishDedicatedPromptWindowSetup(kind) {
  if (!isPromptKindAvailable(kind)) {
    toast("That user was removed from Resume Builder.", "error", 4500);
    return;
  }
  if (!applyDedicatedPromptWindowLayout(kind)) {
    return;
  }
  if (isCustomPromptKind(kind)) {
    await loadCustomPromptTemplate(kind);
  } else if (isBuiltinPromptKind(kind)) {
    await reloadPromptTemplateForKind(kind, { silent: true });
  }
  await refreshDedicatedPromptTabPreview();
  updateGptResultReadyUi();
}

/** @returns {boolean} */
function applyInitialPromptFromQuery() {
  const prompt = parseLaunchPromptKind();
  if (!prompt) {
    return false;
  }
  finishDedicatedPromptWindowSetup(prompt);
  return true;
}

/** @type {string | null} */
let pendingDedicatedWindowPrompt = null;
let dedicatedWindowInitDone = false;

function normalizeDedicatedWindowPrompt(kind) {
  if (kind === "justin" || kind === "other") {
    return kind;
  }
  if (typeof kind === "string" && kind.startsWith("custom:")) {
    return kind;
  }
  return null;
}

function setupDedicatedWindowChannel() {
  const api = window.promptApi;
  if (!api || typeof api.onWindowPromptKind !== "function") {
    return;
  }
  api.onWindowPromptKind((kind) => {
    const prompt = normalizeDedicatedWindowPrompt(kind);
    if (!prompt) {
      return;
    }
    if (!dedicatedWindowInitDone) {
      pendingDedicatedWindowPrompt = prompt;
      return;
    }
    finishDedicatedPromptWindowSetup(prompt);
  });
}

async function init() {
  syncAppVersionLabel();
  ensureMainWindowTabLayout();
  if (!dedicatedPromptKind) {
    document.body.classList.add("workspace-tab-build");
  }
  const launchKind = parseLaunchPromptKind();
  if (launchKind) {
    applyDedicatedPromptWindowLayout(launchKind);
  }
  try {
    await loadAllTemplates();
  } catch (e) {
    const msg = e && typeof e === "object" && "message" in e ? String(e.message) : String(e);
    toast(
      msg.includes("promptApi") || msg.includes("unavailable")
        ? "App bridge failed to load. Close fully and restart, or reinstall."
        : `Could not load templates: ${msg}`,
      "error",
    );
    return;
  }
  try {
    await removeDefaultSampleUsers();
    mountMainEditLayout();
    syncAllBuildInputShellStates();
    setupChatgptShowPanesHandlers();
    setupMainChatgptPanesHost();
    renderEditPromptNav();
    syncEditLayoutEmptyState();
    renderBuildPromptCards();
    applyBuiltinPromptVisibilityToUi();
    syncMainWindowChatgptPanes();
    renderChatgptShowPanesBar();
    refreshPromptDisplayNamesAcrossUi();
    void syncGlobalResumeSaveRootDisplays();
    void syncPromptPopupHotkeysToMain();
    if (!dedicatedPromptKind) {
      loadMainCompanyName();
    }
    if (!applyInitialPromptFromQuery()) {
      selectDefaultPromptCard();
    }
    if (editTemplate && !dedicatedPromptKind) {
      const editKind = isPromptKindAvailable(editSelection)
        ? editSelection
        : listAllMainPrompts()[0]?.id;
      if (isBuiltinPromptKind(editKind)) {
        void openEditKey(/** @type {"other" | "justin"} */ (editKind));
      } else if (isCustomPromptKind(editKind)) {
        void openEditCustomPrompt(editKind);
      }
    }
    updateGptResultReadyUi();
    dedicatedWindowInitDone = true;
    if (pendingDedicatedWindowPrompt) {
      finishDedicatedPromptWindowSetup(pendingDedicatedWindowPrompt);
      pendingDedicatedWindowPrompt = null;
    } else if (dedicatedPromptKind) {
      refreshDedicatedPromptTabPreview();
    }
  } catch (e) {
    console.error("Resume Builder init UI", e);
    toast("Templates loaded, but preview UI failed to start. Try Reload.", "error");
  }
}

reloadAllBtn?.addEventListener("click", async () => {
  try {
    await reloadAllTemplatesFromDisk();
  } catch (_e) {
    toast("Reload failed.", "error");
  }
});

exportResumeJsonBtn?.addEventListener("click", async () => {
  const api = window.promptApi;
  if (!api || typeof api.exportResumeJsonToDocx !== "function") {
    toast("Export is not available.", "error");
    return;
  }
  try {
    const r = await api.exportResumeJsonToDocx();
    if (r?.canceled) {
      return;
    }
    if (r?.ok) {
      toast("Word document saved.");
    } else {
      toast(String(r?.error || "Export failed."), "error");
    }
  } catch (e) {
    toast(String(e && e.message ? e.message : e), "error");
  }
});

reloadBuildBtn?.addEventListener("click", async () => {
  try {
    await reloadAllTemplatesFromDisk();
  } catch (_e) {
    toast("Reload failed.", "error");
  }
});

/**
 * @param {string} kind
 * @param {{ originOnly?: boolean }} [opts]
 */
function previewContentForPrompt(kind, opts = {}) {
  const originOnly = opts.originOnly === true;
  if (!jdInput) {
    return { ok: false, message: "Build inputs are not available." };
  }
  const jdRaw = jdInput.value;
  const jdFilled = jdRaw.trim().length > 0;

  if (kind === "other") {
    const tplRaw = otherTemplateStr;
    if (!tplRaw.trim()) {
      return { ok: false, message: "Olu prompt template is empty." };
    }
    if (originOnly || !jdFilled) {
      return {
        ok: true,
        text: tplRaw,
        metaHtml: "",
      };
    }
    return {
      ok: true,
      text: applyPlaceholders(tplRaw, jdRaw, ""),
      metaHtml: "",
    };
  }

  if (kind === "justin") {
    const tplRaw = justinTemplateStr;
    if (!tplRaw.trim()) {
      return { ok: false, message: "Justin prompt template is empty." };
    }
    if (originOnly || !jdFilled) {
      return {
        ok: true,
        text: tplRaw,
        metaHtml: "",
      };
    }
    return {
      ok: true,
      text: applyPlaceholders(tplRaw, jdRaw, ""),
      metaHtml: "",
    };
  }

  if (isCustomPromptKind(kind)) {
    const entry = getCustomPromptByKind(kind);
    if (!entry) {
      return { ok: false, message: "Prompt not found." };
    }
    const tplRaw = customTemplateStrById[kind] || "";
    if (!tplRaw.trim()) {
      return { ok: false, message: `${entry.label} prompt file is empty or missing.` };
    }
    if (originOnly || !jdFilled) {
      return {
        ok: true,
        text: tplRaw,
        metaHtml: "",
      };
    }
    return {
      ok: true,
      text: applyPlaceholders(tplRaw, jdRaw, ""),
      metaHtml: "",
    };
  }

  return { ok: false, message: "Unknown prompt." };
}

function getSelectedPromptKind() {
  const activeEdit = document.querySelector(
    ".edit-nav .edit-file-btn.edit-file-active[data-edit]",
  );
  const editKind = activeEdit?.getAttribute("data-edit") ?? null;
  if (isBuiltinPromptKind(editKind) || isCustomPromptKind(editKind)) {
    return editKind;
  }
  const card = document.querySelector(".prompt-card.is-highlighted");
  const k = card?.getAttribute("data-prompt") ?? null;
  if (isBuiltinPromptKind(k) || isCustomPromptKind(k)) {
    return k;
  }
  return null;
}

/** @param {string} key */
function syncPromptCardHighlightFromKind(key) {
  document.querySelectorAll(".prompt-card[data-prompt]").forEach((card) => {
    card.classList.toggle("is-highlighted", card.getAttribute("data-prompt") === key);
  });
}

/**
 * @param {string} kind
 */
function isPasteCompleteForPromptKind(kind) {
  if (dedicatedPromptKind && isBuiltinPromptKind(kind)) {
    if (dedicatedActiveWorkspacePage === "prompt") {
      return true;
    }
    return jdInput.value.trim().length > 0;
  }
  return jdInput.value.trim().length > 0;
}

function getDedicatedBuildTopRow() {
  return (
    document.querySelector("#pageBuild .build-top-row") ??
    document.querySelector(".build-layout .build-top-row")
  );
}

function getDedicatedBuildPreviewBlock() {
  return (
    document.getElementById("buildPreviewBlock") ??
    document.querySelector(".build-preview-block")
  );
}

function getDedicatedBuildPreviewColumn() {
  return document.querySelector("#pageBuild .build-preview-column");
}

/** Hidden holder so JD / company are not in the Prompt tab layout. */
function getDedicatedWorkspaceParking() {
  let el = document.getElementById("dedicatedWorkspaceParking");
  if (!el) {
    el = document.createElement("div");
    el.id = "dedicatedWorkspaceParking";
    el.hidden = true;
    el.setAttribute("aria-hidden", "true");
    el.style.display = "none";
    document.body.appendChild(el);
  }
  return el;
}

function parkDedicatedBrowserOnlyPanels() {
  const parking = getDedicatedWorkspaceParking();
  const topRow = getDedicatedBuildTopRow();
  if (topRow && topRow.parentElement !== parking) {
    parking.appendChild(topRow);
  }
  if (
    dedicatedCompanySection &&
    dedicatedCompanySection.parentElement !== parking
  ) {
    parking.appendChild(dedicatedCompanySection);
    dedicatedCompanySection.hidden = true;
  }
  parkDedicatedPromptDirsPanel();
}

/** Keep download-directory settings in the preview block when not on the browser tab. */
function parkDedicatedPromptDirsPanel() {
  const block = getDedicatedBuildPreviewBlock();
  if (!dedicatedPromptDirsPanel || !block) {
    return;
  }
  if (dedicatedPromptDirsPanel.parentElement !== block) {
    block.appendChild(dedicatedPromptDirsPanel);
  }
  dedicatedPromptDirsPanel.hidden = true;
  parkDedicatedJdDownloadStatusHost();
}

/** @param {HTMLElement} parent */
function mountDedicatedPromptDirsPanelIn(parent) {
  if (!dedicatedPromptDirsPanel || !parent) {
    return;
  }
  if (dedicatedPromptDirsPanel.parentElement !== parent) {
    if (
      dedicatedCompanySection &&
      dedicatedCompanySection.parentElement === parent
    ) {
      dedicatedCompanySection.insertAdjacentElement(
        "afterend",
        dedicatedPromptDirsPanel,
      );
    } else {
      const topRow = getDedicatedBuildTopRow();
      if (topRow && topRow.parentElement === parent) {
        topRow.insertAdjacentElement("afterend", dedicatedPromptDirsPanel);
      } else {
        parent.appendChild(dedicatedPromptDirsPanel);
      }
    }
  }
  dedicatedPromptDirsPanel.hidden = false;
  mountDedicatedJdDownloadStatusHostIn(parent);
}

async function refreshDedicatedPromptTabPreview() {
  const kind = dedicatedPromptKind;
  if (!kind) {
    return;
  }
  if (isCustomPromptKind(kind) && !String(customTemplateStrById[kind] || "").trim()) {
    await loadCustomPromptTemplate(kind);
  } else if (isBuiltinPromptKind(kind)) {
    const tpl =
      kind === "justin" ? justinTemplateStr : otherTemplateStr;
    if (!String(tpl || "").trim()) {
      await reloadPromptTemplateForKind(kind, { silent: true });
    }
  }
  applyPreviewForPromptKind(kind, { showToast: false });
}

/** @param {string} tabId @param {"jd" | "company"} field */
function dedicatedBrowserTabStorageKey(tabId, field) {
  const kind = dedicatedPromptKind || "other";
  return `${DEDICATED_BROWSER_TAB_STORAGE_PREFIX}${kind}.${tabId}.${field}`;
}

/** @param {string} tabId @param {"jd" | "company"} field */
function legacyDedicatedBrowserTabStorageKey(tabId, field) {
  const legacyKind = dedicatedPromptKind === "justin" ? "justin" : "other";
  return `${DEDICATED_BROWSER_TAB_STORAGE_PREFIX}${legacyKind}.${tabId}.${field}`;
}

/** @param {string} tabId */
function clearDedicatedBrowserTabStorage(tabId) {
  const id = String(tabId || "").trim();
  if (!id) {
    return;
  }
  const fields = ["jd", "company"];
  const kinds = new Set([
    dedicatedPromptKind || "other",
    "other",
    "justin",
  ]);
  for (const p of customPrompts) {
    if (p?.id) {
      kinds.add(p.id);
    }
  }
  try {
    for (const kind of kinds) {
      for (const field of fields) {
        localStorage.removeItem(`${DEDICATED_BROWSER_TAB_STORAGE_PREFIX}${kind}.${id}.${field}`);
      }
    }
    const suffixes = [`.${id}.jd`, `.${id}.company`];
    const keysToRemove = [];
    for (let i = 0; i < localStorage.length; i++) {
      const key = localStorage.key(i);
      if (!key || !key.startsWith(DEDICATED_BROWSER_TAB_STORAGE_PREFIX)) {
        continue;
      }
      if (suffixes.some((suffix) => key.endsWith(suffix))) {
        keysToRemove.push(key);
      }
    }
    for (const key of keysToRemove) {
      localStorage.removeItem(key);
    }
  } catch (_e) {
    /* ignore */
  }
}

/** @param {DedicatedBrowserTabEntry} entry */
function loadDedicatedBrowserTabFormStateFromStorage(entry) {
  void entry;
}

/** @param {string} tabId */
function saveDedicatedBrowserTabFormState(tabId) {
  const entry = dedicatedBrowserTabs.find((t) => t.id === tabId);
  if (!entry) {
    return;
  }
  entry.jd = jdInput?.value ?? "";
  entry.company = dedicatedCompanyInput?.value?.trim() ?? "";
  refreshDedicatedBrowserTabLabels();
}

/** @param {string} tabId */
function loadDedicatedBrowserTabFormState(tabId) {
  const entry = dedicatedBrowserTabs.find((t) => t.id === tabId);
  if (!entry) {
    return;
  }
  loadDedicatedBrowserTabFormStateFromStorage(entry);
  if (jdInput) {
    jdInput.value = entry.jd;
    syncBuildInputShellState(jdInput);
  }
  if (dedicatedCompanyInput) {
    dedicatedCompanyInput.value = entry.company ?? "";
    syncDedicatedCompanyShell();
  }
  if (entry.paneEl) {
    updateChatgptPaneToolbarUi(entry.paneEl);
  }
  refreshBuildPreviewIfCardSelected();
  refreshDedicatedBrowserTabLabels();
}

/** @returns {DedicatedBrowserTabEntry | null} */
function getActiveDedicatedBrowserTabEntry() {
  if (dedicatedActiveWorkspacePage === "prompt") {
    return null;
  }
  return dedicatedBrowserTabs.find((t) => t.id === dedicatedActiveWorkspacePage) ?? null;
}

function getDedicatedCompanyName() {
  const entry = getActiveDedicatedBrowserTabEntry();
  if (entry?.company) {
    return entry.company;
  }
  return dedicatedCompanyInput?.value?.trim() || "";
}

function syncDedicatedCompanyShell() {
  if (!dedicatedCompanyInput) {
    return;
  }
  const shell = dedicatedCompanyInput.closest(".build-input-shell");
  if (!shell) {
    return;
  }
  shell.classList.toggle("is-empty", !dedicatedCompanyInput.value.trim());
  shell.classList.toggle("is-filled", !!dedicatedCompanyInput.value.trim());
}

function persistDedicatedCompanyName() {
  if (!dedicatedPromptKind || !dedicatedCompanyInput) {
    return;
  }
  if (dedicatedActiveWorkspacePage !== "prompt") {
    saveDedicatedBrowserTabFormState(dedicatedActiveWorkspacePage);
  }
}

function applyDedicatedDownloadDirVisibility() {
  if (!dedicatedPromptKind) {
    if (dedicatedPromptDirsPanel) {
      dedicatedPromptDirsPanel.hidden = true;
    }
    return;
  }
  const onBrowserTab = dedicatedActiveWorkspacePage !== "prompt";
  if (dedicatedCompanySection) {
    dedicatedCompanySection.hidden = !onBrowserTab;
  }
  if (dedicatedPromptDirsPanel) {
    dedicatedPromptDirsPanel.hidden = !onBrowserTab;
  }
  void loadMainPreviewDirectories(dedicatedPromptKind);
}

/** @returns {string} */
function downloadKeyForContext() {
  if (dedicatedPromptKind) {
    return dedicatedPromptKind;
  }
  return getSelectedPromptKind() || "other";
}

function mountDedicatedPromptTabPanels() {
  parkDedicatedBrowserOnlyPanels();
  const legacy = document.getElementById("buildLegacyPanel");
  if (legacy instanceof HTMLElement) {
    legacy.hidden = false;
    legacy.removeAttribute("aria-hidden");
  }
  const editSlot = document.getElementById("editLayoutSlot");
  if (editSlot instanceof HTMLElement) {
    editSlot.hidden = true;
    editSlot.setAttribute("aria-hidden", "true");
  }
  const column = getDedicatedBuildPreviewColumn();
  const block = getDedicatedBuildPreviewBlock();
  if (!column || !block) {
    return;
  }
  if (block.parentElement !== column) {
    column.appendChild(block);
  }
  const heading = block.querySelector(".preview-heading");
  if (heading instanceof HTMLElement) {
    heading.textContent = "Prompt content";
  }
  parkDedicatedPromptDirsPanel();
  void refreshDedicatedPromptTabPreview();
  applyDedicatedDownloadDirVisibility();
  ensureDedicatedJdToolbar();
}

function mountDedicatedBrowserTabPanels() {
  const parking = getDedicatedWorkspaceParking();
  const topRow =
    getDedicatedBuildTopRow() ?? parking.querySelector(".build-top-row");
  const block = getDedicatedBuildPreviewBlock();
  const column = getDedicatedBuildPreviewColumn();
  if (!dedicatedBrowserSideInner) {
    return;
  }
  if (topRow && topRow.parentElement !== dedicatedBrowserSideInner) {
    dedicatedBrowserSideInner.appendChild(topRow);
  }
  if (dedicatedCompanySection) {
    if (dedicatedCompanySection.parentElement !== dedicatedBrowserSideInner) {
      dedicatedBrowserSideInner.appendChild(dedicatedCompanySection);
    }
    dedicatedCompanySection.hidden = false;
  }
  if (block && column && block.parentElement === dedicatedBrowserSideInner) {
    column.appendChild(block);
  }
  mountDedicatedPromptDirsPanelIn(dedicatedBrowserSideInner);
  refreshBuildPreviewIfCardSelected();
  applyDedicatedDownloadDirVisibility();
  ensureDedicatedJdToolbar();
  syncDedicatedJdToolbarUi();
}

/** @returns {HTMLElement | null} */
function getActiveDedicatedBrowserWebview() {
  if (!dedicatedPromptKind) {
    return null;
  }
  const tabId =
    dedicatedActiveWorkspacePage !== "prompt"
      ? dedicatedActiveWorkspacePage
      : dedicatedBrowserTabs[0]?.id;
  if (!tabId) {
    return getChatgptWebviewForProfile(getDedicatedGptProfile());
  }
  const entry = dedicatedBrowserTabs.find((t) => t.id === tabId);
  const wv = entry?.paneEl?.querySelector(".chatgpt-webview");
  return wv && typeof wv.executeJavaScript === "function"
    ? /** @type {HTMLElement} */ (wv)
    : getChatgptWebviewForProfile(getDedicatedGptProfile());
}

/**
 * Full prompt text for copy / GPT. Never trims payload; JD-only prompts merge JD.
 * @param {"other" | "justin"} kind
 */
function getFullPromptTextForKind(kind) {
  const built = previewContentForPrompt(kind);
  if (built.ok && built.text) {
    return built.text;
  }
  return previewOutput?.value ?? "";
}

function getPreviewCopyControls() {
  return document.querySelectorAll(
    ".build-preview-block .preview-copy-control.preview-copy-only",
  );
}

function updatePreviewCopyUi() {
  const controls = getPreviewCopyControls();
  const kind = getSelectedPromptKind();
  let disabled = true;
  let title = "Copy preview";
  if (!kind) {
    title = "Select a prompt card";
    } else if (!dedicatedPromptKind && !isPasteCompleteForPromptKind(kind)) {
      title = "Paste a job description (JD) to enable copy";
  } else if (!previewOutput.value.length) {
    title = "Nothing to copy yet";
  } else {
    disabled = false;
    title = "Copy preview";
  }
  controls.forEach((btn) => {
    btn.disabled = disabled;
    btn.title = title;
    btn.setAttribute("aria-label", title);
  });

  updateAllChatgptPaneToolbarsUi();
}

/**
 * @param {string} text
 * @param {{ emptyMsg?: string; successMsg?: string }} [opts]
 */
async function copyStringToClipboard(text, opts = {}) {
  const emptyMsg = opts.emptyMsg ?? "Nothing to copy.";
  const successMsg = opts.successMsg ?? "Copied.";
  if (!String(text ?? "").length) {
    toast(emptyMsg, "error");
    return false;
  }
  try {
    await navigator.clipboard.writeText(text);
    toast(successMsg);
    return true;
  } catch (_e) {
    toast("Copy failed.", "error");
    return false;
  }
}

/**
 * @param {string} kind
 * @param {{ showToast?: boolean }} [opts]
 */
function applyPreviewForPromptKind(kind, opts = {}) {
  const { showToast = true } = opts;
  if (!previewOutput || !previewMeta) {
    return;
  }
  const originOnly =
    !!dedicatedPromptKind && dedicatedActiveWorkspacePage === "prompt";
  const result = previewContentForPrompt(kind, { originOnly });
  if (!result.ok) {
    if (showToast) {
      toast(result.message, "error");
    }
    previewOutput.value = dedicatedPromptKind ? String(result.message || "") : "";
    previewMeta.textContent = "";
    updatePreviewCopyUi();
    return;
  }
  previewMeta.textContent = "";
  previewOutput.value = result.text;
  syncBuildFieldHighlightsForKind(kind);
  if (!dedicatedPromptKind) {
    void loadMainPreviewDirectories(kind);
    void refreshChatgptPaneWordDirectoriesForKind(kind);
  }
  updatePreviewCopyUi();
}

/**
 * @param {string | null | undefined} kind
 * @returns {"olu" | "justin"}
 */
function profileForPromptKind(kind) {
  return gptResumeProfileForPromptKind(kind);
}

/**
 * @param {HTMLElement} card
 * @param {{ showToast?: boolean }} [opts]
 */
function applyPreviewForCard(card, opts = {}) {
  const kind = card.getAttribute("data-prompt");
  if (!kind) {
    updatePreviewCopyUi();
    return;
  }
  applyPreviewForPromptKind(kind, opts);
}

function refreshBuildPreviewIfCardSelected() {
  const selected = document.querySelector(".prompt-card.is-highlighted");
  if (!selected) {
    clearBuildFieldHighlights();
    updatePreviewCopyUi();
    return;
  }
  applyPreviewForCard(selected, { showToast: false });
}

if (jdInput) {
  jdInput.addEventListener("input", onBuildInputsChanged);
}

function setupPreviewCopyHandlers() {
  const run = async () => {
    const ctrl = getPreviewCopyControls()[0];
    if (!ctrl || ctrl.disabled) {
      return;
    }
    const kind = getSelectedPromptKind();
    const text = kind ? getFullPromptTextForKind(kind) : previewOutput.value;
    await copyStringToClipboard(text, {
      emptyMsg: "Nothing in preview to copy.",
      successMsg: "Preview copied.",
    });
  };
  getPreviewCopyControls().forEach((btn) => {
    btn.addEventListener("click", run);
  });
}

/**
 * Guest: inject full composer text without trimming. Returns "ok" | "no-input" | "empty" | "clipboard-failed".
 * @param {string} textInitJs JavaScript statements that assign normalized string to `t`
 */
function buildChatgptInjectComposerScript(textInitJs) {
  return `(async function () {
    ${textInitJs}
    if (!t.length) return "empty";
    function fireInput(el) {
      try {
        el.dispatchEvent(new InputEvent("input", { bubbles: true, inputType: "insertText" }));
      } catch (e) {
        el.dispatchEvent(new Event("input", { bubbles: true }));
      }
    }
    function readText(el) {
      if (el.tagName === "TEXTAREA") return String(el.value || "").replace(/\\r\\n/g, "\\n").replace(/\\r/g, "\\n");
      return String(el.innerText || el.textContent || "").replace(/\\r\\n/g, "\\n").replace(/\\r/g, "\\n");
    }
    function hasFull(el, expected) {
      var got = readText(el);
      return got.length >= expected.length * 0.98 || got === expected;
    }
    function setTextareaValue(el, val) {
      var proto = window.HTMLTextAreaElement.prototype;
      var d = Object.getOwnPropertyDescriptor(proto, "value");
      if (d && d.set) d.set.call(el, val);
      else el.value = val;
      fireInput(el);
    }
    function resolveEditable(root) {
      if (!root) return null;
      if (root.tagName === "TEXTAREA") return root;
      var pm = root.querySelector && root.querySelector(".ProseMirror");
      if (!pm) pm = document.querySelector("#prompt-textarea .ProseMirror, .ProseMirror");
      if (pm && (pm.isContentEditable || pm.getAttribute("contenteditable") === "true")) return pm;
      var ta = root.querySelector && root.querySelector("textarea");
      if (ta) return ta;
      if (root.isContentEditable) return root;
      return root.querySelector("textarea, [contenteditable='true']");
    }
    function setParagraphs(el, val) {
      var lines = String(val).split("\\n");
      var frag = document.createDocumentFragment();
      for (var i = 0; i < lines.length; i++) {
        var p = document.createElement("p");
        if (lines[i].length) p.textContent = lines[i];
        else p.appendChild(document.createElement("br"));
        frag.appendChild(p);
      }
      el.replaceChildren(frag);
      fireInput(el);
    }
    function setPreWrap(el, val) {
      var block = document.createElement("div");
      block.style.whiteSpace = "pre-wrap";
      block.textContent = val;
      el.replaceChildren(block);
      fireInput(el);
    }
    function findComposerRoot() {
      var sels = [
        "textarea#prompt-textarea",
        "motion-textarea textarea",
        'textarea[data-id="root"]',
        "#prompt-textarea .ProseMirror",
        'div#prompt-textarea[contenteditable="true"]',
        "#prompt-textarea"
      ];
      for (var si = 0; si < sels.length; si++) {
        var hit = document.querySelector(sels[si]);
        if (hit && hit.offsetParent !== null) return hit;
      }
      var cand = document.querySelectorAll("textarea, [contenteditable='true']");
      for (var j = 0; j < cand.length; j++) {
        var node = cand[j];
        if (node.offsetParent === null) continue;
        var ph = (node.getAttribute && node.getAttribute("data-placeholder")) || "";
        var id = node.id || "";
        if (id === "prompt-textarea" || ph.toLowerCase().indexOf("message") >= 0) return node;
      }
      return null;
    }
    var root = findComposerRoot();
    if (!root) return "no-input";
    var el = resolveEditable(root);
    if (!el) return "no-input";
    if (el.tagName === "TEXTAREA") {
      setTextareaValue(el, t);
      return hasFull(el, t) ? "ok" : "no-input";
    }
    if (!el.isContentEditable) return "no-input";
    el.focus();
    el.replaceChildren();
    setParagraphs(el, t);
    if (hasFull(el, t)) return "ok";
    el.replaceChildren();
    setPreWrap(el, t);
    if (hasFull(el, t)) return "ok";
    el.replaceChildren();
    el.innerText = t;
    fireInput(el);
    return hasFull(el, t) ? "ok" : "no-input";
  })()`;
}

/** Clipboard fallback when webview IPC inject is unavailable. */
const CHATGPT_INJECT_FROM_CLIPBOARD_SCRIPT = buildChatgptInjectComposerScript(
  `var t = "";
try { t = await navigator.clipboard.readText(); } catch (e) { return "clipboard-failed"; }
t = String(t).replace(/\\r\\n/g, "\\n").replace(/\\r/g, "\\n");`,
);

/** @param {string} escaped JSON.stringify(previewText) */
function buildChatgptInjectComposerScriptFromText(escaped) {
  return buildChatgptInjectComposerScript(
    `var t = ${escaped};
t = String(t).replace(/\\r\\n/g, "\\n").replace(/\\r/g, "\\n");`,
  );
}

/**
 * Inject full prompt into ChatGPT via webview preload IPC (no inline string limits).
 * @param {HTMLElement} wv
 * @param {string} text
 * @returns {Promise<string>}
 */
function injectTextIntoChatGptWebview(wv, text) {
  if (typeof wv.send === "function") {
    return new Promise((resolve) => {
      let settled = false;
      const timeout = setTimeout(() => {
        if (!settled) {
          settled = true;
          resolve("timeout");
        }
      }, 15000);
      const onMsg = (event) => {
        if (event.channel !== "pb-inject-result") {
          return;
        }
        wv.removeEventListener("ipc-message", onMsg);
        if (!settled) {
          settled = true;
          clearTimeout(timeout);
          const args = Array.isArray(event.args) ? event.args : [];
          resolve(String(args[0] ?? "error"));
        }
      };
      wv.addEventListener("ipc-message", onMsg);
      try {
        wv.send("pb-inject-composer", text);
      } catch (_e) {
        wv.removeEventListener("ipc-message", onMsg);
        clearTimeout(timeout);
        settled = true;
        resolve("error");
      }
    });
  }
  if (typeof wv.executeJavaScript !== "function") {
    return Promise.resolve("error");
  }
  return wv.executeJavaScript(CHATGPT_INJECT_FROM_CLIPBOARD_SCRIPT, true).catch(() => "error");
}

/** After text is in the composer: wait if upload UI shows, then Enter to send. */
const CHATGPT_WAIT_UPLOAD_AND_ENTER_SCRIPT = `(async function () {
  function sleep(ms) {
    return new Promise(function (r) {
      setTimeout(r, ms);
    });
  }
  function findComposer() {
    var el =
      document.querySelector("#prompt-textarea") ||
      document.querySelector('textarea[data-id="root"]') ||
      document.querySelector("textarea#prompt-textarea") ||
      document.querySelector('div#prompt-textarea[contenteditable="true"]');
    if (!el) {
      var cand = document.querySelectorAll("textarea, div[contenteditable='true']");
      for (var j = 0; j < cand.length; j++) {
        var node = cand[j];
        if (node.offsetParent === null) continue;
        var ph = (node.getAttribute && node.getAttribute("data-placeholder")) || "";
        var id = node.id || "";
        if (id === "prompt-textarea" || ph.toLowerCase().indexOf("message") >= 0) {
          el = node;
          break;
        }
      }
    }
    return el;
  }
  function composerScope() {
    var el = findComposer();
    if (!el) return null;
    var root = el;
    for (var i = 0; i < 10 && root && root.parentElement; i++) {
      root = root.parentElement;
    }
    return root || el;
  }
  function isLikelyUploading() {
    var scope = composerScope();
    if (!scope) return false;
    var bars = scope.querySelectorAll('[role="progressbar"]');
    for (var i = 0; i < bars.length; i++) {
      if (bars[i].offsetParent !== null) return true;
    }
    var st = scope.querySelectorAll('[data-state="uploading"], [data-state="loading"]');
    for (var u = 0; u < st.length; u++) {
      if (st[u].offsetParent !== null) return true;
    }
    var btns = scope.querySelectorAll("button");
    for (var b = 0; b < btns.length; b++) {
      var btn = btns[b];
      if (btn.offsetParent === null) continue;
      var al = ((btn.getAttribute && btn.getAttribute("aria-label")) || "").toLowerCase();
      if (al.indexOf("cancel") >= 0 && al.indexOf("upload") >= 0) return true;
    }
    return false;
  }
  function fireEnter(el) {
    var o = {
      key: "Enter",
      code: "Enter",
      keyCode: 13,
      which: 13,
      bubbles: true,
      cancelable: true,
    };
    try {
      el.dispatchEvent(new KeyboardEvent("keydown", o));
      el.dispatchEvent(new KeyboardEvent("keypress", o));
      el.dispatchEvent(new KeyboardEvent("keyup", o));
    } catch (e2) {
      el.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", bubbles: true }));
      el.dispatchEvent(new KeyboardEvent("keyup", { key: "Enter", bubbles: true }));
    }
  }

  var el = findComposer();
  if (!el) return "no-composer";
  await sleep(120);

  var maxMs = 120000;
  var t0 = Date.now();
  if (isLikelyUploading()) {
    while (Date.now() - t0 < maxMs) {
      await sleep(220);
      if (!isLikelyUploading()) {
        await sleep(280);
        if (!isLikelyUploading()) break;
      }
    }
  } else {
    await sleep(100);
  }

  el = findComposer() || el;
  el.focus();
  await sleep(120);

  function findSendButton() {
    var sels = [
      'button[data-testid="send-button"]',
      'button[data-testid*="send-button"]',
      'button[aria-label="Send prompt"]',
      'button[aria-label="Send message"]',
      'button[aria-label*="Send"]'
    ];
    for (var si = 0; si < sels.length; si++) {
      var hits = document.querySelectorAll(sels[si]);
      for (var h = 0; h < hits.length; h++) {
        if (hits[h].offsetParent !== null && !hits[h].disabled) return hits[h];
      }
    }
    return null;
  }
  var sendBtn = null;
  var sendDeadline = Date.now() + 8000;
  while (Date.now() < sendDeadline) {
    sendBtn = findSendButton();
    if (sendBtn) break;
    await sleep(150);
  }
  if (sendBtn) {
    var rc = sendBtn.getBoundingClientRect();
    return JSON.stringify({ x: Math.round(rc.left + rc.width / 2), y: Math.round(rc.top + rc.height / 2) });
  }
  return "no-send-button";
})()`;

/** True once ChatGPT has accepted the prompt (stop button shown, or composer emptied). */
const CHATGPT_SUBMITTED_SCRIPT = `(function () {
  if (document.querySelector('[data-testid="stop-button"], button[aria-label*="Stop"]')) return true;
  var c = document.querySelector("#prompt-textarea");
  var txt = c ? String(c.value !== undefined && c.tagName === "TEXTAREA" ? c.value : c.innerText || "") : "";
  return txt.trim().length < 20;
})()`;

/** Submit with trusted input events (DOM .click()/synthetic Enter are ignored by ChatGPT). */
async function submitChatgptComposerTrusted(wv) {
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const submitted = async () => {
    for (let i = 0; i < 10; i += 1) {
      await sleep(300);
      try {
        if (await wv.executeJavaScript(CHATGPT_SUBMITTED_SCRIPT, true)) return true;
      } catch (_e) {
        /* retry */
      }
    }
    return false;
  };
  let pos = null;
  try {
    const raw = String(await wv.executeJavaScript(CHATGPT_WAIT_UPLOAD_AND_ENTER_SCRIPT, true));
    pos = JSON.parse(raw);
  } catch (_e) {
    pos = null;
  }
  if (typeof wv.sendInputEvent !== "function") return false;
  try {
    wv.focus();
  } catch (_e) {
    /* ignore */
  }
  if (pos && Number.isFinite(pos.x)) {
    const base = { x: pos.x, y: pos.y, button: "left", clickCount: 1 };
    try {
      wv.sendInputEvent({ type: "mouseMove", x: pos.x, y: pos.y });
      wv.sendInputEvent({ type: "mouseDown", ...base });
      await sleep(60);
      wv.sendInputEvent({ type: "mouseUp", ...base });
    } catch (_e) {
      /* fall through to Enter */
    }
    if (await submitted()) return true;
  }
  try {
    wv.sendInputEvent({ type: "keyDown", keyCode: "Enter" });
    wv.sendInputEvent({ type: "char", keyCode: "\r" });
    wv.sendInputEvent({ type: "keyUp", keyCode: "Enter" });
  } catch (_e) {
    return false;
  }
  return submitted();
}

/* ------------------------------------------------------------------------
 * GPT reply reader (send → wait → JSON → Word).
 * Primary: ChatGPT's own conversation API, called from inside the logged-in page.
 *          message.status === "finished_successfully" is the definitive "done" signal.
 * Fallback: scan the page text for a complete resume JSON object that stops changing.
 * No dependency on ChatGPT's DOM class names / data attributes.
 * ---------------------------------------------------------------------- */

/** Append a line to <userData>/gpt-debug.log (send → reply → Word build trace). */
function gptDbg(msg) {
  try {
    void window.promptApi?.gptDebugLog?.(msg);
  } catch (_e) {
    /* ignore */
  }
}

/**
 * In-page script. Returns JSON:
 * { url, api: { status, endTurn, text } | null, apiError, stop, bodyText }
 * @param {number} sinceSec only consider assistant messages created at/after this unix time (0 = any)
 */
function buildReadGptReplyScript(sinceSec) {
  return `(async function (sinceSec) {
  var out = { url: location.href, api: null, apiError: "", stop: false, bodyText: "" };
  try {
    var parts = location.pathname.split("/");
    var ci = parts.indexOf("c");
    var convId = ci >= 0 ? parts[ci + 1] || "" : "";
    if (!convId) {
      out.apiError = "no conversation id in URL";
    } else {
      if (!window.__pbTok || Date.now() - (window.__pbTokAt || 0) > 600000) {
        var sres = await fetch("/api/auth/session", { credentials: "include" });
        var sess = sres.ok ? await sres.json() : null;
        window.__pbTok = (sess && sess.accessToken) || "";
        window.__pbTokAt = Date.now();
      }
      var headers = window.__pbTok ? { Authorization: "Bearer " + window.__pbTok } : {};
      var res = await fetch("/backend-api/conversation/" + convId, { credentials: "include", headers: headers });
      if (!res.ok) {
        out.apiError = "conversation http " + res.status;
        if (res.status === 401) window.__pbTok = "";
      } else {
        var conv = await res.json();
        var map = (conv && conv.mapping) || {};
        var best = null;
        for (var k in map) {
          var m = map[k] && map[k].message;
          if (!m || !m.author || m.author.role !== "assistant") continue;
          if (m.recipient && m.recipient !== "all") continue;
          var c = m.content || {};
          if (c.content_type && c.content_type !== "text") continue;
          var ct = m.create_time || 0;
          if (ct < sinceSec) continue;
          if (!best || ct >= (best.create_time || 0)) best = m;
        }
        if (best) {
          var bc = best.content || {};
          var txt = Array.isArray(bc.parts)
            ? bc.parts.filter(function (p) { return typeof p === "string"; }).join("")
            : String(bc.text || "");
          out.api = { status: String(best.status || ""), endTurn: best.end_turn, text: txt };
        } else {
          out.api = { status: "none", endTurn: null, text: "" };
        }
      }
    }
  } catch (e) {
    out.apiError = String((e && e.message) || e);
  }
  try {
    if (document.querySelector('[data-testid="stop-button"]')) out.stop = true;
    var btns = document.querySelectorAll("button");
    for (var i = 0; i < btns.length && !out.stop; i++) {
      var al = (btns[i].getAttribute("aria-label") || "").toLowerCase();
      if (al.indexOf("stop") >= 0 && al.indexOf("dictat") < 0 && al.indexOf("voice") < 0) out.stop = true;
    }
    out.bodyText = String((document.querySelector("main") || document.body).innerText || "");
  } catch (e2) {
    /* ignore */
  }
  return JSON.stringify(out);
})(${Number(sinceSec) || 0})`;
}

/**
 * Last complete resume JSON object in `text` (has non-empty work_experience), or null.
 * Scans balanced {...} blocks, string/escape aware, so surrounding prose is ignored.
 * @param {string} text
 * @returns {Record<string, unknown> | null}
 */
function extractResumeJsonFromText(text) {
  const s = String(text || "");
  let found = null;
  let depth = 0;
  let start = -1;
  let inStr = false;
  let esc = false;
  for (let i = 0; i < s.length; i += 1) {
    const ch = s[i];
    if (inStr) {
      if (esc) esc = false;
      else if (ch === "\\") esc = true;
      else if (ch === '"') inStr = false;
      continue;
    }
    if (ch === '"') {
      if (depth > 0) inStr = true;
    } else if (ch === "{") {
      if (depth === 0) start = i;
      depth += 1;
    } else if (ch === "}" && depth > 0) {
      depth -= 1;
      if (depth === 0 && start >= 0) {
        try {
          const o = JSON.parse(s.slice(start, i + 1));
          const wx = o && (o.work_experience || o.work_xperience);
          if (o && typeof o === "object" && !Array.isArray(o) && Array.isArray(wx) && wx.length > 0) {
            found = o;
          }
        } catch (_e) {
          /* not JSON — keep scanning */
        }
        start = -1;
      }
    }
  }
  return found;
}

/**
 * @param {any} wv
 * @param {number} sinceSec
 * @returns {Promise<{ url?: string, api: { status: string, endTurn: any, text: string } | null, apiError: string, stop: boolean, bodyText: string }>}
 */
async function readGptReply(wv, sinceSec) {
  const empty = { api: null, apiError: "", stop: false, bodyText: "" };
  if (!wv || typeof wv.executeJavaScript !== "function") {
    return { ...empty, apiError: "webview unavailable" };
  }
  try {
    const raw = await wv.executeJavaScript(buildReadGptReplyScript(sinceSec), true);
    return { ...empty, ...JSON.parse(String(raw || "{}")) };
  } catch (e) {
    return { ...empty, apiError: `executeJavaScript: ${e?.message || e}` };
  }
}

/**
 * Wait until ChatGPT finishes, then return the resume JSON.
 * @param {any} wv
 * @param {{ sinceSec: number, baselineJson: string, timeoutMs: number }} opts
 * @returns {Promise<{ ok: true, json: Record<string, unknown> } | { ok: false, error: string }>}
 */
async function waitForGptResumeJson(wv, { sinceSec, baselineJson, timeoutMs }) {
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const deadline = Date.now() + timeoutMs;
  let domCandidate = "";
  let domStable = 0;
  let iter = 0;
  gptDbg(`wait start sinceSec=${sinceSec} timeoutMs=${timeoutMs}`);
  await sleep(1500);
  while (Date.now() < deadline) {
    iter += 1;
    const st = await readGptReply(wv, sinceSec);
    const api = st.api;
    if (iter % 5 === 1) {
      gptDbg(
        `wait #${iter} url=${st.url || ""} api=${api ? `${api.status}/${api.text.length}` : "null"} apiError=${st.apiError} stop=${st.stop} bodyLen=${st.bodyText.length}`,
      );
    }

    /* 1) API: definitive status. */
    if (api && api.status === "finished_successfully" && api.endTurn !== false) {
      const json = extractResumeJsonFromText(api.text);
      if (json) {
        gptDbg(`wait done via api len=${api.text.length}`);
        return { ok: true, json };
      }
      gptDbg(`api finished but no resume JSON: ${JSON.stringify(api.text.slice(0, 200))}`);
      return {
        ok: false,
        error: "ChatGPT finished, but the reply has no complete resume JSON (with work_experience).",
      };
    }
    const apiStreaming = !!api && api.status === "in_progress";

    /* 2) Page-text fallback: complete JSON, no Stop button, unchanged for 3 polls. */
    if (!st.stop && !apiStreaming) {
      const json = extractResumeJsonFromText(st.bodyText);
      const key = json ? JSON.stringify(json) : "";
      if (key && key !== baselineJson) {
        if (key === domCandidate) {
          domStable += 1;
        } else {
          domCandidate = key;
          domStable = 0;
        }
        if (domStable >= 2) {
          gptDbg(`wait done via page text len=${key.length}`);
          return { ok: true, json };
        }
      } else {
        domCandidate = "";
        domStable = 0;
      }
    } else {
      domCandidate = "";
      domStable = 0;
    }
    await sleep(2000);
  }
  gptDbg("wait timeout");
  return { ok: false, error: "Timed out waiting for ChatGPT to finish the resume JSON." };
}

/**
 * Used by View / refresh: current reply as text (pretty JSON when found).
 * @param {any} wv
 * @returns {Promise<{ text: string, streaming: boolean }>}
 */
async function readChatgptAssistantState(wv) {
  const st = await readGptReply(wv, 0);
  const api = st.api;
  const streaming = api && !st.apiError ? api.status === "in_progress" : st.stop;
  const json = extractResumeJsonFromText(api?.text || "") || extractResumeJsonFromText(st.bodyText);
  if (json) {
    return { text: JSON.stringify(json, null, 2), streaming };
  }
  return { text: String(api?.text || "").trim(), streaming };
}

/**
 * @param {HTMLElement} pane
 * @param {{ silent?: boolean }} [opts]
 * @returns {Promise<{ ok: boolean, text?: string }>}
 */
async function fetchGptResultForPane(pane, opts = {}) {
  const { silent = false } = opts;
  const ctx = getChatgptPaneContextFromPane(pane);
  if (!ctx.wv) {
    if (!silent) {
      toast("ChatGPT pane is not available.", "error");
    }
    return { ok: false };
  }
  if (isGptBusyForSendTarget(ctx.profile, { browserTabId: ctx.browserTabId })) {
    if (!silent) {
      toast("Wait until ChatGPT finishes generating.", "error");
    }
    return { ok: false };
  }

  const resultLabel = contextLabelForGptUi({
    promptKind: ctx.kind,
    browserTabId: ctx.browserTabId,
    profile: ctx.profile,
  });

  const modalFetchBtn =
    gptResultModal && !gptResultModal.hidden ? gptResultModalFetchBtn : null;
  const modalDownloadBtn =
    gptResultModal && !gptResultModal.hidden ? gptResultModalDownloadBtn : null;
  if (modalFetchBtn instanceof HTMLButtonElement) {
    modalFetchBtn.disabled = true;
  }
  if (modalDownloadBtn instanceof HTMLButtonElement) {
    modalDownloadBtn.disabled = true;
  }

  try {
    const { text, streaming } = await readChatgptAssistantState(ctx.wv);
    if (streaming) {
      if (!silent) {
        toast(`${resultLabel} is still generating — try again in a moment.`, "error");
      }
      return { ok: false };
    }
    if (!text) {
      if (!silent) {
        toast(
          "No assistant reply found in ChatGPT. Make sure a response is visible, then try again.",
          "error",
        );
      }
      return { ok: false };
    }
    storeGptResult(ctx.profile, text, false, { browserTabId: ctx.browserTabId });
    updateChatgptPaneToolbarUi(pane);
    if (!silent) {
      toast(`${resultLabel}: result updated.`, "success", 4500);
    }
    return { ok: true, text };
  } finally {
    if (modalFetchBtn instanceof HTMLButtonElement) {
      modalFetchBtn.disabled = false;
    }
    if (modalDownloadBtn instanceof HTMLButtonElement) {
      modalDownloadBtn.disabled = false;
    }
    updateGptResultReadyUi();
  }
}

/** @returns {Promise<{ ok: boolean, text?: string }>} */
async function fetchGptResultForModalContext() {
  const pane = getChatgptPaneForModalContext();
  if (!pane) {
    toast("ChatGPT pane is not available.", "error");
    return { ok: false };
  }
  const result = await fetchGptResultForPane(pane, { silent: true });
  if (!result.ok || !result.text) {
    return result;
  }
  showGptResultModalDone(result.text, false, { viewOnly: true });
  return result;
}

/** Re-read latest assistant reply from the active ChatGPT webview and store it. */
async function fetchGptResultFromActiveWebview() {
  /** @type {HTMLElement | null} */
  let pane = null;
  if (dedicatedPromptKind) {
    const tabId = getActiveDedicatedBrowserTabId();
    pane = dedicatedBrowserTabs.find((t) => t.id === tabId)?.paneEl ?? null;
  } else {
    const kind = getSelectedPromptKind();
    if (kind) {
      pane = document.querySelector(
        `#chatgptPanesHost .chatgpt-pane[data-prompt="${CSS.escape(kind)}"]`,
      );
    }
  }
  if (!pane) {
    toast("ChatGPT pane is not available.", "error");
    return;
  }
  await fetchGptResultForPane(pane);
}

/**
 * @param {HTMLElement} pane
 * @returns {Promise<boolean>}
 */
async function ensureParseableGptResultForPane(pane) {
  const ctx = getChatgptPaneContextFromPane(pane);
  const api = window.promptApi;
  if (!api?.checkGptDocxExport) {
    return false;
  }
  const resultLabel = contextLabelForGptUi({
    promptKind: ctx.kind,
    browserTabId: ctx.browserTabId,
    profile: ctx.profile,
  });

  let stored = getStoredGptResultForPaneContext(ctx);
  if (stored?.text?.trim() && !stored.isError) {
    const check = await api.checkGptDocxExport(stored.text);
    if (check?.ok) {
      return true;
    }
  }

  if (!ctx.wv) {
    toast("ChatGPT pane is not available.", "error");
    return false;
  }

  toast(`${resultLabel}: refreshing from ChatGPT…`, "success", 2500);
  const fetched = await fetchGptResultForPane(pane, { silent: true });
  if (!fetched.ok || !fetched.text) {
    return false;
  }

  const exportCheck = await api.checkGptDocxExport(fetched.text);
  if (!exportCheck?.ok) {
    toast(
      exportCheck?.error ||
        "ChatGPT reply must be complete resume_structure JSON (with work_experience). Open View, refresh, then try again.",
      "error",
      7000,
    );
    updateGptResultReadyUi();
    return false;
  }
  return true;
}

/** @param {HTMLElement} pane */
async function downloadGptDocxForPane(pane) {
  const ctx = getChatgptPaneContextFromPane(pane);
  const downloadKey = ctx.kind || ctx.promptId || downloadKeyForContext();
  const stored = getStoredGptResultForPaneContext(ctx);
  if (!stored || stored.isError || !stored.text.trim()) {
    toast("No result to download yet.", "error");
    return;
  }

  const api = window.promptApi;
  const alreadyParseable =
    api?.checkGptDocxExport &&
    (await api.checkGptDocxExport(stored.text))?.ok;

  gptDbg(`download: storedLen=${stored.text.length} parseable=${!!alreadyParseable}`);
  if (!alreadyParseable) {
    const ok = await ensureParseableGptResultForPane(pane);
    gptDbg(`download: ensureParseable=${ok}`);
    if (!ok) {
      return;
    }
  }

  gptDbg("download: building Word from JSON + skeleton");
  await downloadGptDocxForProfile(ctx.profile, {
    useDefaultDirectory: true,
    downloadKey,
    browserTabId: ctx.browserTabId,
    paneEl: pane,
  });
}

function sanitizeDocxNameSegment(s) {
  return String(s || "")
    .replace(/[\\/:*?"<>|]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
}

/**
 * Suggest a filename like "Oluwaseun Adeniji - ML Engineer" from reply, EXP/JD when possible.
 * @param {string} [replyText] optional ChatGPT reply — first "Name - Role" line wins
 * @returns {string} stem without .docx
 */
/**
 * @param {string} text
 * @returns {string}
 */
function roleTitleFromResumeJsonText(text) {
  const s = String(text || "").trim();
  if (!s) {
    return "";
  }
  try {
    const o = JSON.parse(s);
    if (o && typeof o === "object" && !Array.isArray(o)) {
      const rt = o.role_title ?? o.role;
      if (rt) {
        return String(rt).trim();
      }
    }
  } catch {
    /* continue */
  }
  const fenceRe = /```(?:json)?\s*\n?([\s\S]*?)```/gi;
  let m;
  while ((m = fenceRe.exec(s)) !== null) {
    try {
      const o = JSON.parse(m[1].trim());
      if (o && typeof o === "object" && !Array.isArray(o)) {
        const rt = o.role_title ?? o.role;
        if (rt) {
          return String(rt).trim();
        }
      }
    } catch {
      /* try next fence */
    }
  }
  const i = s.indexOf("{");
  const j = s.lastIndexOf("}");
  if (i >= 0 && j > i) {
    try {
      const o = JSON.parse(s.slice(i, j + 1));
      if (o && typeof o === "object" && !Array.isArray(o)) {
        const rt = o.role_title ?? o.role;
        if (rt) {
          return String(rt).trim();
        }
      }
    } catch {
      /* ignore */
    }
  }
  return "";
}

function suggestGptDocxDefaultStem(
  replyText,
  profile = activeGptResumeProfile,
  companyName = getCompanyFolderForDownload(),
) {
  const company = sanitizeDocxNameSegment(companyName);
  if (company) {
    return company;
  }
  const reply = String(replyText || "").trim();
  const jsonRole = sanitizeDocxNameSegment(roleTitleFromResumeJsonText(reply));
  if (jsonRole) {
    return jsonRole;
  }
  if (reply) {
    const replyLine =
      reply
        .split(/\r?\n/)
        .map((l) => l.trim())
        .find((l) => l.length > 0) || "";
    if (replyLine && /\s-\s/.test(replyLine)) {
      const stem = sanitizeDocxNameSegment(replyLine);
      if (stem) {
        return stem;
      }
    }
  }
  const jdLine =
    (jdInput?.value?.trim() || "")
      .split(/\r?\n/)
      .find((l) => l.trim())?.trim() || "";
  const jdRole = sanitizeDocxNameSegment(jdLine);
  if (jdRole) {
    return jdRole;
  }

  return "Resume";
}

/** @param {"olu" | "justin"} profile */
function getLastGptResult(profile) {
  return lastGptResults[profile] || null;
}

/** @returns {string | null} */
function getActiveDedicatedBrowserTabId() {
  if (!dedicatedPromptKind || dedicatedActiveWorkspacePage === "prompt") {
    return null;
  }
  return dedicatedActiveWorkspacePage;
}

/** @returns {GptStoredResult | null} */
function getLastGptResultForContext() {
  const browserTabId = getActiveDedicatedBrowserTabId();
  if (browserTabId) {
    return lastGptResultsByBrowserTab[browserTabId] ?? null;
  }
  const profile = profileForSelectedPrompt();
  return getLastGptResult(profile);
}

const TAB_BROWSER_STATUS_GPT_READY_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.25" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="9"/><polyline points="8 12 11 15 16 9"/></svg>';

const TAB_BROWSER_STATUS_DOWNLOADED_SVG =
  '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="3" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><polyline points="20 6 9 17 4 12"/></svg>';

/** @param {string} tabId */
function updateDedicatedBrowserTabStatus(tabId) {
  const entry = dedicatedBrowserTabs.find((t) => t.id === tabId);
  if (!entry) {
    return;
  }
  const busy = isGptBrowserTabBusy(tabId);
  const downloaded = !!gptBrowserTabDownloadDone[tabId];
  const stored = lastGptResultsByBrowserTab[tabId];
  const ready = !!stored && !stored.isError && !!stored.text.trim();
  const gptDone = !busy && !!stored && !!stored.text.trim();

  entry.tabEl.classList.toggle("is-gpt-loading", busy);
  entry.tabEl.classList.toggle("has-gpt-result", ready);
  entry.tabEl.classList.toggle("has-gpt-answered", gptDone && !downloaded);
  entry.tabEl.classList.toggle("has-download-done", downloaded && !busy);

  const labelNode = entry.tabEl.querySelector(".tab-dedicated-browser-label");
  if (labelNode) {
    const label = dedicatedBrowserTabLabel(tabId);
    labelNode.textContent = label;
    labelNode.title = label;
  }

  let statusEl = entry.tabEl.querySelector(".tab-dedicated-browser-status");
  if (!statusEl) {
    statusEl = document.createElement("span");
    statusEl.className = "tab-dedicated-browser-status";
    const closeBtn = entry.tabEl.querySelector(".tab-dedicated-browser-close");
    if (closeBtn) {
      entry.tabEl.insertBefore(statusEl, closeBtn);
    } else {
      entry.tabEl.appendChild(statusEl);
    }
  }

  statusEl.className = "tab-dedicated-browser-status";
  statusEl.textContent = "";
  statusEl.removeAttribute("aria-label");
  statusEl.setAttribute("aria-hidden", "true");

  if (busy) {
    statusEl.classList.add("tab-dedicated-browser-status--loading");
    statusEl.setAttribute("aria-label", "ChatGPT is generating");
    statusEl.removeAttribute("aria-hidden");
  } else if (downloaded) {
    statusEl.classList.add("tab-dedicated-browser-status--downloaded");
    statusEl.innerHTML = TAB_BROWSER_STATUS_DOWNLOADED_SVG;
    statusEl.setAttribute("aria-label", "Downloaded successfully");
    statusEl.removeAttribute("aria-hidden");
  } else if (gptDone) {
    statusEl.classList.add("tab-dedicated-browser-status--gpt-ready");
    statusEl.innerHTML = TAB_BROWSER_STATUS_GPT_READY_SVG;
    if (stored.isError) {
      statusEl.setAttribute("aria-label", "ChatGPT finished with issues");
    } else {
      statusEl.setAttribute("aria-label", "ChatGPT finished");
    }
    statusEl.removeAttribute("aria-hidden");
  }
}

/** @param {string} tabId */
function refreshDedicatedBrowserTabResultBadge(tabId) {
  updateDedicatedBrowserTabStatus(tabId);
}

function refreshAllDedicatedBrowserTabResultBadges() {
  dedicatedBrowserTabs.forEach((entry) => {
    refreshDedicatedBrowserTabResultBadge(entry.id);
  });
}

/** @param {"olu" | "justin"} profile */
function profileForSelectedPrompt() {
  const kind = getSelectedPromptKind();
  return kind ? gptResumeProfileForPromptKind(kind) : activeGptResumeProfile;
}

function updateGptResultReadyUi() {
  const profile = profileForSelectedPrompt();
  const browserTabId = getActiveDedicatedBrowserTabId();
  const stored = getLastGptResultForContext();
  const busy = isGptBusyForSendTarget(profile, { browserTabId });
  const resultLabel = contextLabelForGptUi({
    promptKind: getSelectedPromptKind(),
    browserTabId,
    profile,
  });

  if (gptResultReadyBar) {
    const onDedicatedBrowser = dedicatedPromptKind && !!browserTabId;
    const showBar =
      !!stored &&
      !busy &&
      (!dedicatedPromptKind || onDedicatedBrowser);
    gptResultReadyBar.hidden = !showBar;
  }
  if (gptResultReadyLabel) {
    if (!stored) {
      gptResultReadyLabel.textContent = "";
    } else if (stored.isError) {
      gptResultReadyLabel.textContent = `${resultLabel}: generation finished with an issue.`;
    } else {
      gptResultReadyLabel.textContent = `${resultLabel}: reply ready.`;
    }
  }
  if (viewGptResultBtn) {
    viewGptResultBtn.disabled = !stored;
  }
  if (downloadGptResultBtn) {
    downloadGptResultBtn.disabled = true;
  }
  if (gptResultModalFetchBtn instanceof HTMLButtonElement) {
    gptResultModalFetchBtn.disabled = busy;
  }
  if (gptResultModalDownloadBtn instanceof HTMLButtonElement) {
    gptResultModalDownloadBtn.disabled = busy;
  }
  updateAllChatgptPaneToolbarsUi();
  if (
    (gptResultModalDownloadWord || gptResultModalDownloadBtn) &&
    !gptResultModal?.hidden &&
    stored?.text
  ) {
    void refreshGptResultModalExportUi(stored.text, stored.isError);
  }
}

/** @param {string} text @param {boolean} isError */
async function refreshGptResultModalExportUi(text, isError) {
  const modalDownloadBtns = [gptResultModalDownloadWord, gptResultModalDownloadBtn].filter(
    (el) => el instanceof HTMLButtonElement,
  );
  if (!modalDownloadBtns.length) {
    return;
  }
  const hasBody = !!String(text ?? "").trim() && !isError;
  const disabledTitle = "Reply must be complete resume_structure JSON";
  const enabledTitle = "Download Word";
  if (!hasBody) {
    for (const btn of modalDownloadBtns) {
      btn.disabled = true;
      btn.title = disabledTitle;
    }
    return;
  }
  for (const btn of modalDownloadBtns) {
    btn.disabled = true;
  }
  const exportable = await isGptResultExportable(text);
  for (const btn of modalDownloadBtns) {
    btn.disabled = !exportable;
    btn.title = exportable ? enabledTitle : disabledTitle;
  }
}

/**
 * @param {"olu" | "justin"} profile
 * @param {string} text
 * @param {boolean} isError
 * @param {{ browserTabId?: string | null }} [opts]
 */
function storeGptResult(profile, text, isError, opts = {}) {
  const browserTabId = opts.browserTabId ?? null;
  const result = { text, isError };
  if (browserTabId) {
    lastGptResultsByBrowserTab[browserTabId] = result;
    refreshDedicatedBrowserTabResultBadge(browserTabId);
  } else {
    lastGptResults[profile] = result;
  }
  activeGptResumeProfile = profile;
  updateGptResultReadyUi();
}

function setGptResultModalContextFromPane(pane) {
  const ctx = getChatgptPaneContextFromPane(pane);
  activeGptResumeProfile = ctx.profile;
  gptResultModalContext = {
    profile: ctx.profile,
    browserTabId: ctx.browserTabId,
    downloadKey: ctx.kind || ctx.promptId || downloadKeyForContext(),
  };
}

/** @param {HTMLElement} pane */
function openGptResultModalForPane(pane) {
  const ctx = getChatgptPaneContextFromPane(pane);
  const stored = getStoredGptResultForPaneContext(ctx);
  if (!stored) {
    toast("No result yet. Use Get result again first.", "error");
    return;
  }
  setGptResultModalContextFromPane(pane);
  showGptResultModalDone(stored.text, stored.isError, { viewOnly: true });
}

function openGptResultModalForCurrentContext() {
  const stored = getLastGptResultForContext();
  if (!stored) {
    toast("No result to show yet.", "error");
    return;
  }
  const browserTabId = getActiveDedicatedBrowserTabId();
  const kind = dedicatedPromptKind || getSelectedPromptKind();
  gptResultModalContext = {
    profile: profileForSelectedPrompt(),
    browserTabId,
    downloadKey: kind || downloadKeyForContext(),
  };
  showGptResultModalDone(stored.text, stored.isError, { viewOnly: true });
}

/** @param {"olu" | "justin"} profile */
function openGptResultModalForProfile(profile) {
  const browserTabId = getActiveDedicatedBrowserTabId();
  const stored = browserTabId
    ? lastGptResultsByBrowserTab[browserTabId]
    : getLastGptResult(profile);
  if (!stored) {
    toast("No result to show yet.", "error");
    return;
  }
  const kind = dedicatedPromptKind || getSelectedPromptKind();
  gptResultModalContext = {
    profile,
    browserTabId,
    downloadKey: kind || downloadKeyForContext(),
  };
  activeGptResumeProfile = profile;
  showGptResultModalDone(stored.text, stored.isError, { viewOnly: true });
}

/**
 * @param {"olu" | "justin"} profile
 * @param {{ useDefaultDirectory?: boolean, showSaveDialog?: boolean, paneEl?: HTMLElement | null, browserTabId?: string | null, downloadKey?: string, companyFolder?: string }} [opts]
 */
async function downloadGptDocxForProfile(profile, opts = {}) {
  const api = window.promptApi;
  if (!api || typeof api.saveGptDocx !== "function") {
    toast("Word export is not available.", "error");
    return;
  }
  const browserTabId =
    opts.browserTabId !== undefined
      ? opts.browserTabId
      : getActiveDedicatedBrowserTabId();
  const stored = browserTabId
    ? lastGptResultsByBrowserTab[browserTabId]
    : getLastGptResult(profile);
  if (!stored || stored.isError || !stored.text.trim()) {
    toast("Nothing to download.", "error");
    return;
  }
  const companyFolder =
    opts.companyFolder !== undefined
      ? String(opts.companyFolder).trim()
      : getCompanyFolderForDownload(browserTabId);
  if (!companyFolder) {
    toast(
      dedicatedPromptKind
        ? "Enter company name on the Browser tab first."
        : "Enter company name above before downloading.",
      "error",
    );
    return;
  }
  const downloadKey =
    opts.downloadKey !== undefined && opts.downloadKey !== null
      ? String(opts.downloadKey)
      : downloadKeyForContext();
  try {
    const kind =
      opts.downloadKey !== undefined && opts.downloadKey !== null
        ? String(opts.downloadKey)
        : getSelectedPromptKind();
    const structureTemplatePath = kind
      ? getStructureTemplatePathForKind(kind)
      : null;
    const r = await api.saveGptDocx({
      text: stored.text,
      profile,
      downloadKey,
      profileName: getPromptUserName(downloadKey),
      defaultFileName: suggestGptDocxDefaultStem(stored.text, profile, companyFolder),
      useDefaultDirectory: opts.useDefaultDirectory,
      showSaveDialog: opts.showSaveDialog,
      companyFolder,
      structureTemplatePath: structureTemplatePath || undefined,
      jdText: getJdTextForDownload(browserTabId),
    });
    if (r?.canceled) {
      return;
    }
    if (r?.ok) {
      const who = userDisplayNameForPromptKind(downloadKey);
      if (browserTabId) {
        gptBrowserTabDownloadDone[browserTabId] = true;
        updateDedicatedBrowserTabStatus(browserTabId);
      }
      const savedName = fileNameFromPath(r.filePath || "");
      const folderPath = String(r.folderPath || "").trim();
      const jdNote = r.jdSaved
        ? ` (+ ${fileNameFromPath(String(r.jdPath || "JD.txt"))})`
        : "";
      const paneEl =
        opts.paneEl ??
        (browserTabId
          ? dedicatedBrowserTabs.find((t) => t.id === browserTabId)?.paneEl ?? null
          : null);
      showGptDownloadStatus(r, who, {
        paneEl: opts.paneEl ?? null,
        browserTabId,
        downloadKey,
      });
      toast(
        folderPath
          ? `${who}: saved ${savedName}${jdNote}\n${folderPath}`
          : `${who}: saved ${savedName}${jdNote}`,
        "success",
        folderPath ? 8000 : 5000,
      );
    } else if (r?.needsDirectory) {
      toast(String(r.error || "Set global resume download folder in Settings."), "error", 5500);
    } else {
      toast(String(r?.error || "Save failed."), "error");
    }
  } catch (e) {
    toast(String(e && e.message ? e.message : e), "error");
  }
}

function hideGptResultModal() {
  if (gptResultModal) {
    gptResultModal.hidden = true;
  }
  clearGptModalWorkExperience();
  updatePreviewCopyUi();
}

function clearGptModalWorkExperience() {
  if (gptResultModalWxWrap) {
    gptResultModalWxWrap.hidden = true;
  }
  if (gptResultModalWorkExperience) {
    gptResultModalWorkExperience.textContent = "";
  }
  if (gptResultModalWxWordPreview) {
    gptResultModalWxWordPreview.textContent = "";
    gptResultModalWxWordPreview.hidden = true;
  }
}

/**
 * @param {unknown} rows
 */
function renderGptModalWordPreviewRows(rows) {
  if (!gptResultModalWxWordPreview) {
    return;
  }
  const host = gptResultModalWxWordPreview;
  host.textContent = "";
  if (!Array.isArray(rows) || rows.length === 0) {
    host.hidden = true;
    return;
  }
  renderWordPreviewRows(host, rows);
  host.hidden = host.childElementCount === 0;
}

/**
 * When the reply contains resume_structure JSON, show parsed work_experience in the modal.
 * @param {string} replyText
 */
async function refreshGptModalWorkExperiencePanel(replyText) {
  if (!gptResultModalWxWrap || !gptResultModalWorkExperience) {
    return;
  }
  const api = window.promptApi;
  if (!api || typeof api.resolveWorkExperienceFromGptText !== "function") {
    clearGptModalWorkExperience();
    return;
  }
  try {
    const r = await api.resolveWorkExperienceFromGptText(String(replyText || ""));
    if (r && r.ok && Object.prototype.hasOwnProperty.call(r, "work_experience")) {
      gptResultModalWorkExperience.textContent = JSON.stringify(
        r.work_experience,
        null,
        2,
      );
      renderGptModalWordPreviewRows(r.word_preview_rows);
      gptResultModalWxWrap.hidden = false;
      return;
    }
  } catch (_e) {
    /* ignore */
  }
  clearGptModalWorkExperience();
}

/**
 * @param {string} text
 * @param {boolean} isError
 * @param {{ viewOnly?: boolean }} [opts]
 */
function showGptResultModalDone(text, isError, opts = {}) {
  const { viewOnly = false } = opts;
  if (
    !gptResultModal ||
    !gptResultModalStatus ||
    !gptResultModalBody ||
    !gptResultModalCopy ||
    !gptResultModalDownloadWord
  ) {
    return;
  }
  if (!viewOnly && !isError && String(text).trim()) {
    storeGptResult(gptResultModalContext.profile, text, false, {
      browserTabId: gptResultModalContext.browserTabId,
    });
  }
  gptResultModal.hidden = false;
  if (isError) {
    gptResultModalStatus.textContent = text;
    gptResultModalStatus.classList.add("gpt-result-status--error");
    gptResultModalBody.value = "";
    gptResultModalBody.classList.remove("gpt-result-body--waiting");
    gptResultModalBody.placeholder = "Response will appear here…";
    gptResultModalCopy.disabled = true;
    gptResultModalDownloadWord.disabled = true;
    clearGptModalWorkExperience();
  } else {
    gptResultModalStatus.textContent = viewOnly
      ? "Retrieved assistant reply:"
      : "Latest assistant reply:";
    gptResultModalStatus.classList.remove("gpt-result-status--error");
    gptResultModalBody.value = text;
    gptResultModalBody.classList.remove("gpt-result-body--waiting");
    gptResultModalBody.placeholder = "Response will appear here…";
    const hasBody = !!String(text).trim();
    gptResultModalCopy.disabled = !hasBody;
    gptResultModalDownloadWord.disabled = true;
    void refreshGptModalWorkExperiencePanel(text);
    void refreshGptResultModalExportUi(text, false);
  }
}

function setupGptResultModal() {
  if (
    !gptResultModal ||
    !gptResultModalClose ||
    !gptResultModalCopy ||
    !gptResultModalBody ||
    !gptResultModalDownloadWord
  ) {
    return;
  }
  if (gptResultModalFetchBtn) {
    gptResultModalFetchBtn.innerHTML = CHATGPT_PANE_FETCH_ICON_SVG;
    gptResultModalFetchBtn.addEventListener("click", () => {
      void fetchGptResultForModalContext();
    });
  }
  if (gptResultModalDownloadBtn) {
    gptResultModalDownloadBtn.innerHTML = CHATGPT_PANE_DOWNLOAD_ICON_SVG;
    gptResultModalDownloadBtn.addEventListener("click", async () => {
      const pane = getChatgptPaneForModalContext();
      if (pane) {
        await downloadGptDocxForPane(pane);
        return;
      }
      toast("ChatGPT pane is not available.", "error");
    });
  }
  gptResultModalClose.addEventListener("click", () => {
    hideGptResultModal();
  });
  gptResultModal.addEventListener("click", (e) => {
    if (e.target === gptResultModal) {
      hideGptResultModal();
    }
  });
  gptResultModalCopy.addEventListener("click", async () => {
    await copyStringToClipboard(gptResultModalBody.value, {
      emptyMsg: "Nothing to copy.",
      successMsg: "Reply copied.",
    });
  });
  gptResultModalDownloadWord.addEventListener("click", async () => {
    const pane = getChatgptPaneForModalContext();
    if (pane) {
      await downloadGptDocxForPane(pane);
      return;
    }
    toast("ChatGPT pane is not available.", "error");
  });
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape") {
      return;
    }
    if (!gptResultModal || gptResultModal.hidden) {
      return;
    }
    hideGptResultModal();
  });
}

/**
 * Workflow: paste prompt+JD into ChatGPT → click Send → wait until GPT finishes →
 * take the resume JSON → build Word from JSON + skeleton.
 * @param {HTMLElement} pane
 */
async function sendPreviewToChatGptForPane(pane) {
  const ctx = getChatgptPaneContextFromPane(pane);
  const kind = ctx.browserTabId ? dedicatedPromptKind : ctx.promptId;
  const blocked = inputGptDisabledReasonForPaneContext(ctx);
  if (blocked) {
    toast(blocked, "error");
    return;
  }
  if (!kind) {
    toast("Select a prompt first.", "error");
    return;
  }

  const text = getFullPromptTextForKind(kind);
  if (!String(text || "").trim()) {
    toast("Prompt template is empty or missing.", "error");
    return;
  }
  activeGptResumeProfile = ctx.profile;
  const profile = ctx.profile;
  const browserTabId = ctx.browserTabId;
  const wv = ctx.wv;
  if (!wv || typeof wv.executeJavaScript !== "function") {
    toast("ChatGPT pane is not available.", "error");
    return;
  }

  if (browserTabId && dedicatedActiveWorkspacePage !== browserTabId) {
    activateDedicatedBrowserTabById(browserTabId);
    await new Promise((r) => setTimeout(r, 80));
  } else if (dedicatedPromptKind && dedicatedActiveWorkspacePage === "prompt") {
    activateDedicatedBrowserTab();
    await new Promise((r) => setTimeout(r, 80));
  }

  const browserEntry = browserTabId
    ? dedicatedBrowserTabs.find((t) => t.id === browserTabId)
    : null;
  const busyOpts = { browserTabId, paneEl: browserEntry?.paneEl ?? null };
  const tabLabel = contextLabelForGptUi({ promptKind: kind, browserTabId, profile });

  let resumeJson = null;
  setGptProfileBusy(profile, true, busyOpts);
  try {
    /* 1) Paste. */
    try {
      await navigator.clipboard.writeText(text);
    } catch (_e) {
      /* clipboard is only a fallback path */
    }
    try {
      wv.focus();
    } catch (_e) {
      /* ignore */
    }
    let injected = "error";
    try {
      injected = await injectTextIntoChatGptWebview(wv, text);
    } catch (_e) {
      injected = "error";
    }
    if (injected !== "ok") {
      try {
        injected = await wv.executeJavaScript(
          buildChatgptInjectComposerScriptFromText(JSON.stringify(text)),
          true,
        );
      } catch (_e) {
        injected = "error";
      }
    }
    if (injected !== "ok") {
      gptDbg(`paste failed: ${injected}`);
      toast("Could not paste into the ChatGPT message box.", "error");
      return;
    }

    /* Baseline: any resume JSON already on the page must not be mistaken for the new reply. */
    const before = await readGptReply(wv, 0);
    const baselineObj = extractResumeJsonFromText(before.bodyText);
    const baselineJson = baselineObj ? JSON.stringify(baselineObj) : "";
    const sinceSec = Math.floor(Date.now() / 1000) - 60; /* tolerate clock skew */

    /* 2) Send. */
    await new Promise((r) => setTimeout(r, 150));
    const sent = await submitChatgptComposerTrusted(wv);
    gptDbg(`send: injected=${injected} sent=${sent}`);
    if (!sent) {
      toast("Could not click ChatGPT Send automatically.", "error");
      return;
    }
    hideGptResultModal();
    toast(`Sent to ${tabLabel} ChatGPT — waiting for it to finish…`, "success");

    /* 3) Wait for GPT to finish and take the JSON. */
    const res = await waitForGptResumeJson(wv, { sinceSec, baselineJson, timeoutMs: 600000 });
    if (!res.ok) {
      storeGptResult(profile, res.error, true, { browserTabId });
      toast(`${tabLabel}: ${res.error}`, "error", 8000);
      return;
    }
    resumeJson = res.json;
    storeGptResult(profile, JSON.stringify(resumeJson, null, 2), false, { browserTabId });
  } finally {
    setGptProfileBusy(profile, false, busyOpts);
  }

  /* 4) Build Word from JSON + skeleton. */
  toast(`${tabLabel} finished — building Word resume…`, "success", 4500);
  const downloadPane = browserEntry?.paneEl ?? pane;
  gptDbg("build: JSON ready, building Word from skeleton");
  await downloadGptDocxForPane(downloadPane);
}

/**
 * Copy preview to clipboard, inject text into the ChatGPT composer (guest JS), wait if uploading, then Enter.
 */
async function sendPreviewToChatGpt() {
  /** @type {HTMLElement | null} */
  let pane = null;
  if (dedicatedPromptKind) {
    const tabId = getActiveDedicatedBrowserTabId();
    pane = dedicatedBrowserTabs.find((t) => t.id === tabId)?.paneEl ?? null;
  } else {
    const kind = getSelectedPromptKind();
    if (kind) {
      pane = document.querySelector(
        `#chatgptPanesHost .chatgpt-pane[data-prompt="${CSS.escape(kind)}"]`,
      );
    }
  }
  if (!pane) {
    toast("ChatGPT pane is not available.", "error");
    return;
  }
  await sendPreviewToChatGptForPane(pane);
}

function setupInputGptHandler() {
  if (!inputGptBtn) {
    return;
  }
  inputGptBtn.addEventListener("click", () => {
    void sendPreviewToChatGpt();
  });
}

/**
 * @param {string} downloadKey
 * @param {HTMLElement | null} labelEl
 * @param {string} [prefixOverride]
 */
/**
 * @returns {Promise<boolean>}
 */
async function pickGlobalResumeSaveRoot() {
  const api = window.promptApi;
  if (!api || typeof api.pickDownloadDirectory !== "function") {
    toast("Folder picker is not available.", "error");
    return false;
  }
  try {
    const r = await api.pickDownloadDirectory("global");
    if (r?.canceled) {
      return false;
    }
    if (r?.ok && r.directory) {
      await syncGlobalResumeSaveRootDisplays();
      toast("Global resume save folder updated.", "success", 4500);
      return true;
    }
    toast(String(r?.error || "Could not set folder."), "error");
  } catch (e) {
    toast(String(e && e.message ? e.message : e), "error");
  }
  return false;
}

/** @param {string} kind */
async function loadDedicatedPromptStructureDirectory(kind) {
  await loadMainPromptStructureDirectory(kind);
  if (dedicatedPromptStructureDisplay) {
    dedicatedPromptStructureDisplay.value = getPromptStructurePathForKind(kind);
    dedicatedPromptStructureDisplay.title = dedicatedPromptStructureDisplay.value;
  }
}

const BROWSER_PANEL_HIDDEN_KEY = "resume-builder.browser-panel-hidden";
const LEGACY_PROMPT_PANEL_HIDDEN_KEY = "prompt-builder.prompt-panel-hidden";

function isBrowserPanelHidden() {
  return workspaceBrowserSplit?.classList.contains("browser-panel-hidden") ?? false;
}

/** @param {boolean} hidden */
function setBrowserPanelHidden(hidden) {
  if (!workspaceBrowserSplit) {
    return;
  }
  workspaceBrowserSplit.classList.toggle("browser-panel-hidden", hidden);
  document.body.classList.toggle("browser-panel-collapsed", hidden);
  const showLabel = "Show ChatGPT panel";
  const hideLabel = "Hide ChatGPT panel";
  if (togglePromptPanelBtn) {
    togglePromptPanelBtn.setAttribute("aria-pressed", hidden ? "true" : "false");
    togglePromptPanelBtn.title = hidden ? showLabel : hideLabel;
    togglePromptPanelBtn.setAttribute("aria-label", hidden ? showLabel : hideLabel);
    togglePromptPanelBtn.classList.toggle("browser-panel-toggle-collapsed", hidden);
  }
  if (hideBrowserPanelBtn) {
    hideBrowserPanelBtn.hidden = hidden;
  }
  try {
    localStorage.setItem(BROWSER_PANEL_HIDDEN_KEY, hidden ? "1" : "0");
  } catch (_e) {
    /* ignore */
  }
}

function toggleBrowserPanel() {
  setBrowserPanelHidden(!isBrowserPanelHidden());
}

function restoreBrowserPanelHiddenState() {
  let hidden = false;
  try {
    if (localStorage.getItem(BROWSER_PANEL_HIDDEN_KEY) === "1") {
      hidden = true;
    } else if (
      localStorage.getItem(BROWSER_PANEL_HIDDEN_KEY) === null &&
      localStorage.getItem(LEGACY_PROMPT_PANEL_HIDDEN_KEY) === "1"
    ) {
      hidden = false;
    }
  } catch (_e) {
    hidden = false;
  }
  setBrowserPanelHidden(hidden);
}

function setupChatgptPaneToolbarHandlers() {
  if (document.body.dataset.chatgptPaneToolbarBound) {
    return;
  }
  document.body.dataset.chatgptPaneToolbarBound = "1";

  document.addEventListener("click", (e) => {
    const inputBtn = e.target.closest(".chatgpt-pane-input-gpt");
    if (inputBtn) {
      const pane = resolveChatgptPaneForToolbarClick(inputBtn);
      if (pane) {
        void sendPreviewToChatGptForPane(pane);
      }
      return;
    }
    const viewBtn = e.target.closest(".chatgpt-pane-view");
    if (viewBtn) {
      const pane = resolveChatgptPaneForToolbarClick(viewBtn);
      if (pane) {
        openGptResultModalForPane(pane);
      }
      return;
    }
    const downloadBtn = e.target.closest(".chatgpt-pane-download");
    if (downloadBtn) {
      const pane = resolveChatgptPaneForToolbarClick(downloadBtn);
      if (pane) {
        void downloadGptDocxForPane(pane);
      }
      return;
    }
    const fetchBtn = e.target.closest(".chatgpt-pane-fetch");
    if (fetchBtn) {
      const pane = resolveChatgptPaneForToolbarClick(fetchBtn);
      if (pane) {
        void fetchGptResultForPane(pane);
      }
    }
  });
}

function setupGptResultReadyBar() {
  fetchGptResultAgainBtn?.addEventListener("click", () => {
    void fetchGptResultFromActiveWebview();
  });
  viewGptResultBtn?.addEventListener("click", () => {
    openGptResultModalForCurrentContext();
  });
  downloadGptResultBtn?.addEventListener("click", () => {
    void downloadGptDocxForProfile(profileForSelectedPrompt(), {
      useDefaultDirectory: true,
    });
  });
}

function setupPromptPanelToggle() {
  restoreBrowserPanelHiddenState();
  togglePromptPanelBtn?.addEventListener("click", () => {
    toggleBrowserPanel();
  });
  hideBrowserPanelBtn?.addEventListener("click", () => {
    setBrowserPanelHidden(true);
  });
}

function setupPromptSectionHighlights() {
  const promptHost = document.querySelector(".build-prompts-column");
  if (promptHost && !promptHost.dataset.promptCardsBound) {
    promptHost.dataset.promptCardsBound = "1";
    promptHost.addEventListener("click", (e) => {
      if (e.target.closest(".prompt-new-window-btn")) {
        return;
      }
      const card = e.target.closest(".prompt-card[data-prompt]");
      if (!card || !promptHost.contains(card)) {
        return;
      }
      const kind = card.getAttribute("data-prompt");
      if (!kind) {
        return;
      }
      document.querySelectorAll(".prompt-card").forEach((c) => {
        c.classList.remove("is-highlighted");
      });
      card.classList.add("is-highlighted");
      applyPreviewForCard(/** @type {HTMLElement} */ (card));
      syncMainChatgptPaneHighlights();
      updateGptResultReadyUi();
    });
    promptHost.addEventListener("keydown", (e) => {
      const card = e.target.closest?.(".prompt-card[data-prompt]");
      if (!card || !promptHost.contains(card)) {
        return;
      }
      if (e.key !== "Enter" && e.key !== " ") {
        return;
      }
      if (e.target !== card) {
        return;
      }
      e.preventDefault();
      document.querySelectorAll(".prompt-card").forEach((c) => {
        c.classList.remove("is-highlighted");
      });
      card.classList.add("is-highlighted");
      applyPreviewForCard(/** @type {HTMLElement} */ (card));
      syncMainChatgptPaneHighlights();
      updateGptResultReadyUi();
    });
  }

  if (!document.body.dataset.promptNewWindowBound) {
    document.body.dataset.promptNewWindowBound = "1";
    document.addEventListener("click", (e) => {
      const btn = e.target.closest(".prompt-new-window-btn");
      if (!btn || dedicatedPromptKind) {
        return;
      }
      e.preventDefault();
      e.stopPropagation();
      const promptId = btn.getAttribute("data-prompt-new-window");
      if (!promptId) {
        return;
      }
      void openPromptWindowForKind(promptId);
    });
  }
}

const LEGACY_DEDICATED_HEADER_COLOR_STORAGE = "pb_dedicated_window_header_color";
const DEFAULT_HEADER_COLOR = "#fceee3";
const DEFAULT_MAIN_HEADER_COLOR = "#d1fae5";

/** @type {{ id: string, label: string, hex: string }[]} */
const HEADER_COLOR_PALETTE = [
  { id: "peach", label: "Peach", hex: "#fceee3" },
  { id: "sky", label: "Sky", hex: "#dbeafe" },
  { id: "mint", label: "Mint", hex: "#d1fae5" },
  { id: "lavender", label: "Lavender", hex: "#ede9fe" },
  { id: "rose", label: "Rose", hex: "#ffe4e6" },
  { id: "sand", label: "Sand", hex: "#fef3c7" },
  { id: "silver", label: "Silver", hex: "#e2e8f0" },
];

/** @returns {string} main | prompt kind */
function getActiveHeaderColorScope() {
  const kind = dedicatedPromptKind || parseLaunchPromptKind();
  return kind || "main";
}

/** @param {string} scope */
function getHeaderColorStorageKey(scope) {
  return scope === "main" ? "pb_header_color_main" : `pb_header_color_${scope}`;
}

/** @param {string} raw */
function resolveHeaderColorHex(raw) {
  const s = String(raw || "").trim();
  const byHex = HEADER_COLOR_PALETTE.find(
    (c) => c.hex.toLowerCase() === s.toLowerCase(),
  );
  if (byHex) {
    return byHex.hex;
  }
  const byId = HEADER_COLOR_PALETTE.find((c) => c.id === s);
  if (byId) {
    return byId.hex;
  }
  return DEFAULT_HEADER_COLOR;
}

/** @param {string} [scope] */
function getStoredHeaderColor(scope) {
  const s = scope || getActiveHeaderColorScope();
  try {
    let raw = String(localStorage.getItem(getHeaderColorStorageKey(s)) || "").trim();
    if (!raw && s !== "main") {
      raw = String(
        localStorage.getItem(LEGACY_DEDICATED_HEADER_COLOR_STORAGE) || "",
      ).trim();
    }
    if (raw) {
      return resolveHeaderColorHex(raw);
    }
  } catch (_e) {
    /* ignore */
  }
  return s === "main" ? DEFAULT_MAIN_HEADER_COLOR : DEFAULT_HEADER_COLOR;
}

/** @param {string} hex @param {number} whiteRatio */
function mixHexWithWhite(hex, whiteRatio) {
  const n = String(hex || "")
    .trim()
    .replace(/^#/, "");
  if (!/^[0-9a-f]{6}$/i.test(n)) {
    return null;
  }
  const r = parseInt(n.slice(0, 2), 16);
  const g = parseInt(n.slice(2, 4), 16);
  const b = parseInt(n.slice(4, 6), 16);
  const t = Math.min(1, Math.max(0, whiteRatio));
  const wr = Math.round(r + (255 - r) * t);
  const wg = Math.round(g + (255 - g) * t);
  const wb = Math.round(b + (255 - b) * t);
  return `rgba(${wr}, ${wg}, ${wb}, 0.92)`;
}

/** @param {string} scope @param {string} hex */
function applyHeaderColorForScope(scope, hex) {
  const match =
    HEADER_COLOR_PALETTE.find(
      (c) => c.hex.toLowerCase() === String(hex || "").trim().toLowerCase(),
    ) ?? HEADER_COLOR_PALETTE[0];
  if (scope === "main") {
    const titlebarBg =
      mixHexWithWhite(match.hex, 0.08) ?? "rgba(209, 250, 229, 0.95)";
    const panelBg = mixHexWithWhite(match.hex, 0.82) ?? "#ffffff";
    const appBg = mixHexWithWhite(match.hex, 0.92) ?? "#f4f6f9";
    document.body.style.setProperty("--pb-main-titlebar-bg", titlebarBg);
    document.body.style.setProperty("--pb-main-panel-bg", panelBg);
    document.body.style.setProperty("--pb-app-bg", appBg.replace("0.92)", "1)"));
  } else {
    document.body.style.setProperty("--dedicated-tabstrip-bg", match.hex);
  }
  return match.hex;
}

/** @param {string} scope @param {string} hex */
function saveHeaderColorForScope(scope, hex) {
  const applied = applyHeaderColorForScope(scope, hex);
  try {
    localStorage.setItem(getHeaderColorStorageKey(scope), applied);
  } catch (_e) {
    /* ignore */
  }
  return applied;
}

/** @param {string} selectedHex */
function renderDedicatedHeaderColorGrid(selectedHex) {
  const grid = document.getElementById("dedicatedHeaderColorGrid");
  if (!(grid instanceof HTMLElement)) {
    return;
  }
  const scope = getActiveHeaderColorScope();
  const selected = getStoredHeaderColor(scope);
  const active = selectedHex || selected;
  grid.replaceChildren();
  for (const color of HEADER_COLOR_PALETTE) {
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "dedicated-header-color-btn";
    btn.dataset.colorHex = color.hex;
    btn.setAttribute("role", "option");
    btn.setAttribute("aria-selected", color.hex === active ? "true" : "false");
    btn.title = color.label;
    if (color.hex === active) {
      btn.classList.add("is-selected");
    }
    const swatch = document.createElement("span");
    swatch.className = "dedicated-header-color-swatch";
    swatch.style.backgroundColor = color.hex;
    swatch.setAttribute("aria-hidden", "true");
    const label = document.createElement("span");
    label.className = "dedicated-header-color-label";
    label.textContent = color.label;
    btn.append(swatch, label);
    btn.addEventListener("click", () => {
      saveHeaderColorForScope(scope, color.hex);
      renderDedicatedHeaderColorGrid(color.hex);
      toast(`This window: ${color.label}`, "success");
    });
    grid.appendChild(btn);
  }
}

function openWindowSettingsModal() {
  const modal = document.getElementById("windowSettingsModal");
  if (!(modal instanceof HTMLElement)) {
    return;
  }
  const scope = getActiveHeaderColorScope();
  renderDedicatedHeaderColorGrid(getStoredHeaderColor(scope));
  modal.hidden = false;
  const dialog = modal.querySelector(".window-settings-dialog");
  if (dialog instanceof HTMLElement) {
    dialog.focus();
  }
}

function closeWindowSettingsModal() {
  const modal = document.getElementById("windowSettingsModal");
  if (modal instanceof HTMLElement) {
    modal.hidden = true;
  }
}

function isUsersManageModalOpen() {
  const modal = document.getElementById("usersManageModal");
  return modal instanceof HTMLElement && !modal.hidden;
}

function getEditLayoutElement() {
  return document.querySelector(".edit-layout");
}

function restoreEditLayoutToEditPage() {
  const slot = document.getElementById("editLayoutSlot");
  const layout = getEditLayoutElement();
  if (slot instanceof HTMLElement && layout instanceof HTMLElement && layout.parentElement !== slot) {
    slot.appendChild(layout);
  }
}

function mountMainEditLayout() {
  if (dedicatedPromptKind) {
    return;
  }
  restoreEditLayoutToEditPage();
}

function syncEditTemplateFromEditor() {
  if (editMode === "builtin" && editTemplate) {
    setStrForEdit(editSelection, editTemplate.value);
  }
}

function closeUsersManageModal() {
  /* Main window uses inline edit layout; nothing to close. */
}

async function loadUsersManageEditorForKind(kind) {
  if (!isPromptKindAvailable(kind)) {
    await openFirstAvailablePromptInEditor();
    return;
  }
  if (isBuiltinPromptKind(kind)) {
    await openEditKey(/** @type {"other" | "justin"} */ (kind));
    return;
  }
  if (isCustomPromptKind(kind)) {
    await openEditCustomPrompt(kind);
  }
}

function resolveUsersManageModalKind() {
  const candidates = [
    getSelectedPromptKind(),
    getEditPromptSettingsKind(),
    editSelection,
  ].filter(Boolean);
  for (const k of candidates) {
    if (isPromptKindAvailable(k)) {
      return k;
    }
  }
  return listAllMainPrompts()[0]?.id ?? null;
}

async function openUsersManageModal() {
  if (dedicatedPromptKind) {
    toast("Manage users from the main Resume Builder window.", "info", 4500);
    return;
  }
  mountMainEditLayout();
  const kind = resolveUsersManageModalKind();
  refreshEditUsersNav();
  if (kind) {
    await loadUsersManageEditorForKind(kind);
  } else {
    await openFirstAvailablePromptInEditor();
  }
  const host = document.getElementById("editLayoutSlot");
  if (host instanceof HTMLElement) {
    host.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }
}

function setupUsersManageModal() {
  const openBtn = document.getElementById("openUsersManageBtn");
  const openFromSettingsBtn = document.getElementById("openUsersManageFromSettingsBtn");
  const addFromBuildBtn = document.getElementById("addUserFromBuildBtn");
  const editNavAddBtn = document.getElementById("editNavAddUserBtn");
  const editNavReloadBtn = document.getElementById("editNavReloadBtn");
  const editUserOpenBtn = document.getElementById("editUserHeaderOpenBtn");
  const mainSaveRootBrowseBtn = document.getElementById("mainResumeSaveRootBrowseBtn");
  const editSaveRootBrowseBtn = document.getElementById("editResumeSaveRootBrowseBtn");
  const modal = document.getElementById("usersManageModal");
  const closeBtn = document.getElementById("usersManageModalClose");
  if (!(modal instanceof HTMLElement)) {
    return;
  }
  const openManage = () => {
    void openUsersManageModal();
  };
  openBtn?.addEventListener("click", openManage);
  openFromSettingsBtn?.addEventListener("click", openManage);
  addFromBuildBtn?.addEventListener("click", () => openAddPromptModal());
  editNavAddBtn?.addEventListener("click", () => openAddPromptModal());
  editNavReloadBtn?.addEventListener("click", () => {
    void reloadBuildBtn?.click();
  });
  editUserOpenBtn?.addEventListener("click", () => {
    const kind = getEditPromptSettingsKind();
    if (kind) {
      void openPromptWindowForKind(kind);
    }
  });
  mainSaveRootBrowseBtn?.addEventListener("click", () => {
    void pickGlobalResumeSaveRoot();
  });
  editSaveRootBrowseBtn?.addEventListener("click", () => {
    void pickGlobalResumeSaveRoot();
  });
  closeBtn?.addEventListener("click", () => {
    closeUsersManageModal();
  });
  modal.addEventListener("click", (e) => {
    if (e.target === modal) {
      closeUsersManageModal();
    }
  });
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape" || modal.hidden) {
      return;
    }
    closeUsersManageModal();
  });
}

function setupWindowSettingsModal() {
  const openBtn = document.getElementById("openWindowSettingsBtn");
  const modal = document.getElementById("windowSettingsModal");
  const closeBtn = document.getElementById("windowSettingsModalClose");
  if (!(openBtn instanceof HTMLButtonElement) || !(modal instanceof HTMLElement)) {
    return;
  }
  openBtn.addEventListener("click", () => {
    openWindowSettingsModal();
  });
  closeBtn?.addEventListener("click", () => {
    closeWindowSettingsModal();
  });
  modal.addEventListener("click", (e) => {
    if (e.target === modal) {
      closeWindowSettingsModal();
    }
  });
  document.addEventListener("keydown", (e) => {
    if (e.key !== "Escape" || modal.hidden) {
      return;
    }
    closeWindowSettingsModal();
  });
  const scope = getActiveHeaderColorScope();
  applyHeaderColorForScope(scope, getStoredHeaderColor(scope));
  renderDedicatedHeaderColorGrid(getStoredHeaderColor(scope));
}

function setupWindowAppearanceControls() {
  const topBtn = document.getElementById("alwaysOnTopBtn");
  const api = typeof windowControls !== "undefined" ? windowControls : null;

  const STORAGE_TOP = "pb_always_on_top";

  if (!api || !topBtn) {
    return;
  }

  function syncPinButton(on) {
    topBtn.setAttribute("aria-pressed", on ? "true" : "false");
    topBtn.classList.toggle("is-active", on);
  }

  let topStored = false;
  try {
    topStored = localStorage.getItem(STORAGE_TOP) === "1";
  } catch (_e) {
    /* ignore */
  }

  syncPinButton(topStored);
  void api.setAlwaysOnTop(topStored);

  topBtn.addEventListener("click", () => {
    const on = topBtn.getAttribute("aria-pressed") !== "true";
    syncPinButton(on);
    try {
      localStorage.setItem(STORAGE_TOP, on ? "1" : "0");
    } catch (_e) {
      /* ignore */
    }
    void api.setAlwaysOnTop(on);
  });
}

function setupTitleBarWindowButtons() {
  const api = typeof windowControls !== "undefined" ? windowControls : null;
  const minBtn = document.getElementById("winMinBtn");
  const maxBtn = document.getElementById("winMaxBtn");
  const closeBtn = document.getElementById("winCloseBtn");
  const drag = document.querySelector(".titlebar-drag");
  if (!api || !minBtn || !maxBtn || !closeBtn) {
    return;
  }

  const maxIcon = maxBtn.querySelector(".win-max-icon");
  const restIcon = maxBtn.querySelector(".win-restore-icon");

  function syncMaxButton(isMax) {
    if (isMax) {
      maxBtn.setAttribute("aria-label", "Restore down");
      maxBtn.title = "Restore down";
      maxIcon?.setAttribute("hidden", "");
      restIcon?.removeAttribute("hidden");
    } else {
      maxBtn.setAttribute("aria-label", "Maximize");
      maxBtn.title = "Maximize";
      maxIcon?.removeAttribute("hidden");
      restIcon?.setAttribute("hidden", "");
    }
  }

  async function toggleMax() {
    if (typeof api.maximizeToggle !== "function") return;
    const m = await api.maximizeToggle();
    syncMaxButton(m);
  }

  minBtn.addEventListener("click", () => {
    void api.minimize?.();
  });
  maxBtn.addEventListener("click", () => {
    void toggleMax();
  });
  closeBtn.addEventListener("click", () => {
    if (!dedicatedPromptKind) {
      clearMainWindowSessionInputs();
    }
    void api.close?.();
  });

  drag?.addEventListener("dblclick", () => {
    void toggleMax();
  });

  if (typeof api.isMaximized === "function") {
    void api.isMaximized().then((m) => syncMaxButton(!!m));
  }
  if (typeof api.onMaximized === "function") {
    api.onMaximized((m) => syncMaxButton(!!m));
  }
}

function clearMainWindowSessionInputs() {
  if (dedicatedPromptKind) {
    return;
  }
  if (jdInput) {
    jdInput.value = "";
    syncBuildInputShellState(jdInput);
  }
  if (mainCompanyInput) {
    mainCompanyInput.value = "";
    syncMainCompanyShell();
  }
  purgePersistedInputHistory();
}

window.addEventListener("beforeunload", () => {
  if (!dedicatedPromptKind) {
    clearMainWindowSessionInputs();
  }
});

/** @param {HTMLElement} wv */
function setupChatgptWebviewLoading(wv) {
  if (!wv || wv.dataset.pbLoadingUi === "1") {
    return;
  }
  wv.dataset.pbLoadingUi = "1";

  let host = wv.parentElement;
  if (!host?.classList.contains("chatgpt-webview-host")) {
    host = document.createElement("div");
    host.className = "chatgpt-webview-host";
    const parent = wv.parentNode;
    if (parent) {
      parent.insertBefore(host, wv);
      host.appendChild(wv);
    }
  }

  let overlay = host.querySelector(".chatgpt-webview-loading");
  if (!overlay) {
    overlay = document.createElement("div");
    overlay.className = "chatgpt-webview-loading";
    overlay.setAttribute("role", "status");
    overlay.setAttribute("aria-live", "polite");
    overlay.setAttribute("hidden", "");
    overlay.innerHTML =
      '<span class="chatgpt-webview-loading-label">Loading ChatGPT…</span>';
    host.appendChild(overlay);
  }

  let loadCount = 0;
  /** @type {ReturnType<typeof setTimeout> | null} */
  let hideDelay = null;

  const setVisible = (/** @type {boolean} */ on) => {
    if (on) {
      host.classList.add("is-loading");
      overlay.removeAttribute("hidden");
      overlay.setAttribute("aria-busy", "true");
    } else {
      host.classList.remove("is-loading");
      overlay.setAttribute("hidden", "");
      overlay.setAttribute("aria-busy", "false");
    }
  };

  const bumpStart = () => {
    if (hideDelay) {
      clearTimeout(hideDelay);
      hideDelay = null;
    }
    loadCount += 1;
  };

  const bumpStop = () => {
    loadCount = Math.max(0, loadCount - 1);
    if (loadCount > 0) {
      return;
    }
    hideDelay = setTimeout(() => {
      hideDelay = null;
      if (loadCount === 0) {
        setVisible(false);
      }
    }, 120);
  };

  wv.addEventListener("did-start-loading", bumpStart);
  wv.addEventListener("did-stop-loading", bumpStop);
  wv.addEventListener("did-fail-load", (e) => {
    const ev = /** @type {Event & { errorCode?: number, errorDescription?: string, isMainFrame?: boolean, validatedURL?: string }} */ (
      e
    );
    if (ev.isMainFrame === false) {
      return;
    }
    const failedUrl = String(ev.validatedURL || wv.getAttribute("src") || "");
    if (ev.errorCode === -20 && isBlockedFederatedAuthUrl(failedUrl)) {
      navigateChatgptWebviewToEmailLogin(/** @type {HTMLElement} */ (wv));
      return;
    }
    // Navigation replaced or cancelled — not a real failure.
    if (ev.errorCode === -3 || wv.dataset.pbIntentionalNav === "1") {
      return;
    }
    bumpStop();
    const label = overlay.querySelector(".chatgpt-webview-loading-label");
    const retries = Number(wv.dataset.pbLoadRetry || "0");
    let nextUrl = CHATGPT_HOME_URL;
    if (retries >= 1) {
      nextUrl = CHATGPT_EMAIL_LOGIN_URL;
    }
    if (retries >= 2) {
      if (label) {
        label.textContent =
          "ChatGPT could not load. Use Email login, or open ChatGPT in your system browser.";
      }
      setVisible(true);
      return;
    }
    wv.dataset.pbLoadRetry = String(retries + 1);
    if (label) {
      label.textContent = "Retrying ChatGPT…";
    }
    setVisible(true);
    setTimeout(() => {
      if (failedUrl !== nextUrl) {
        navigateChatgptWebviewTo(wv, nextUrl);
      } else if (typeof wv.reload === "function") {
        wv.reload();
      }
    }, 400);
  });
  wv.addEventListener(
    "dom-ready",
    () => {
      loadCount = 0;
      if (hideDelay) {
        clearTimeout(hideDelay);
        hideDelay = null;
      }
      setVisible(false);
    },
    { once: true },
  );

  if (wv.getAttribute("src")) {
    let alreadyLoading = true;
    try {
      if (typeof wv.isLoading === "function") {
        alreadyLoading = wv.isLoading();
      }
    } catch (_e) {
      /* treat as loading until events say otherwise */
    }
    if (alreadyLoading) {
      bumpStart();
    }
  }
}

function setupChatgptZoomForWebview(wv) {
  if (!wv || typeof wv.getWebContents !== "function") {
    return;
  }

  const STORAGE = "pb_chatgpt_zoom_factor";
  const ZOOM_MIN = 0.5;
  const ZOOM_MAX = 2;
  const ZOOM_STEP = 0.1;

  let factor = 1;
  try {
    const raw = localStorage.getItem(STORAGE);
    if (raw != null) {
      const n = parseFloat(raw, 10);
      if (Number.isFinite(n)) {
        factor = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, n));
      }
    }
  } catch (_e) {
    /* ignore */
  }

  function applyZoom() {
    try {
      const wc = wv.getWebContents();
      if (wc && typeof wc.setZoomFactor === "function") {
        wc.setZoomFactor(factor);
      }
    } catch (_e) {
      /* ignore */
    }
  }

  function saveZoom() {
    try {
      localStorage.setItem(STORAGE, String(factor));
    } catch (_e) {
      /* ignore */
    }
  }

  function bump(delta) {
    factor = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, Math.round((factor + delta) * 100) / 100));
    saveZoom();
    for (const pane of getAllChatgptWebviews()) {
      try {
        const wc = pane.getWebContents();
        if (wc && typeof wc.setZoomFactor === "function") {
          wc.setZoomFactor(factor);
        }
      } catch (_e) {
        /* ignore */
      }
    }
  }

  function resetZoom() {
    factor = 1;
    saveZoom();
    for (const pane of getAllChatgptWebviews()) {
      try {
        const wc = pane.getWebContents();
        if (wc && typeof wc.setZoomFactor === "function") {
          wc.setZoomFactor(factor);
        }
      } catch (_e) {
        /* ignore */
      }
    }
  }

  function pointOverThisWebview(clientX, clientY) {
    const r = wv.getBoundingClientRect();
    return clientX >= r.left && clientX <= r.right && clientY >= r.top && clientY <= r.bottom;
  }

  wv.addEventListener(
    "wheel",
    (e) => {
      if (!(e.ctrlKey || e.metaKey)) {
        return;
      }
      if (!pointOverThisWebview(e.clientX, e.clientY)) {
        return;
      }
      e.preventDefault();
      e.stopPropagation();
      bump(e.deltaY > 0 ? -ZOOM_STEP : ZOOM_STEP);
    },
    { capture: true, passive: false },
  );

  let guestZoomBeforeInputHandler = null;

  wv.addEventListener("dom-ready", () => {
    try {
      applyZoom();
    } catch (_e) {
      /* ignore */
    }
    try {
      const wc = wv.getWebContents();
      if (!wc || typeof wc.on !== "function") {
        return;
      }
      if (guestZoomBeforeInputHandler) {
        try {
          wc.removeListener("before-input-event", guestZoomBeforeInputHandler);
        } catch (_e2) {
          /* ignore */
        }
      }
      guestZoomBeforeInputHandler = (event, input) => {
        if (input.type !== "keyDown" || !(input.control || input.meta)) {
          return;
        }
        const code = input.code || "";
        const key = String(input.key || "");
        const isPlus =
          code === "NumpadAdd" ||
          code === "Equal" ||
          key === "+" ||
          (code === "Equal" && (key === "=" || key === "+"));
        const isMinus =
          code === "NumpadSubtract" ||
          code === "Minus" ||
          key === "-" ||
          key === "−" ||
          key === "–";
        const isZero = code === "Digit0" || code === "Numpad0";
        if (isPlus) {
          event.preventDefault();
          bump(ZOOM_STEP);
        } else if (isMinus) {
          event.preventDefault();
          bump(-ZOOM_STEP);
        } else if (isZero) {
          event.preventDefault();
          resetZoom();
        }
      };
      wc.on("before-input-event", guestZoomBeforeInputHandler);
    } catch (_e) {
      /* ignore */
    }
  });

  wv.addEventListener("ipc-message", (event) => {
    const args = Array.isArray(event.args) ? event.args : [];
    if (event.channel === "gpt-zoom-wheel") {
      const dy = Number(args[0]);
      if (Number.isFinite(dy)) {
        bump(dy > 0 ? -ZOOM_STEP : ZOOM_STEP);
      }
    }
  });

  applyZoom();
}

function setupChatgptZoom() {
  const webviews = getAllChatgptWebviews();
  if (!webviews.length) {
    return;
  }
  webviews.forEach((wv) => {
    setupChatgptWebviewLoading(wv);
    setupChatgptZoomForWebview(wv);
  });
}

function setupChatgptBrowserLogin() {
  if (document.body.dataset.chatgptLoginBound) {
    return;
  }
  document.body.dataset.chatgptLoginBound = "1";

  if (typeof window.promptApi?.onChatgptLoginComplete === "function") {
    window.promptApi.onChatgptLoginComplete((payload) => {
      const part =
        payload && typeof payload.partition === "string" ? payload.partition : "";
      if (part) {
        reloadChatgptWebviewsForPartition(part);
      }
      toast("ChatGPT signed in. The pane will refresh.", "success", 5000);
    });
  }

  document.addEventListener("click", async (e) => {
    const emailBtn = e.target.closest(".chatgpt-email-login");
    if (emailBtn) {
      await openChatgptSignInForPaneButton(emailBtn);
      return;
    }

    const resetBtn = e.target.closest(".chatgpt-reset-login");
    if (resetBtn) {
      const partition = getChatgptPartitionForPaneButton(resetBtn);
      try {
        if (typeof window.promptApi?.clearChatgptSession === "function") {
          await window.promptApi.clearChatgptSession({ partition });
        }
        reloadChatgptWebviewsForPartition(partition);
        await openChatgptSignInForPaneButton(resetBtn);
        toast("Login cleared. Sign in again with email and password.", "success", 8000);
      } catch (_err) {
        toast("Could not reset login.", "error");
      }
      return;
    }

    const browserBtn = e.target.closest(".chatgpt-open-browser");
    if (!browserBtn) {
      return;
    }
    const profile = browserBtn.getAttribute("data-chatgpt-profile");
    const label = profile === "justin" ? "Justin" : "Oluwaseun";
    try {
      if (typeof window.promptApi?.openChatgptExternal === "function") {
        await window.promptApi.openChatgptExternal(profile);
      } else {
        window.open("https://chatgpt.com/", "_blank");
      }
      toast(
        `${label}: login page opened in your browser. That login is separate from the app — use Sign in in the pane to log in inside Resume Builder.`,
        "success",
        9000,
      );
    } catch (_err) {
      toast("Could not open ChatGPT in browser.", "error");
    }
  });
}

const WORKSPACE_SPLIT_STORAGE_KEY = "prompt-builder.workspace-left-width";
const WORKSPACE_SPLIT_DEFAULT = 38;

function setupWorkspaceSplitResize() {
  if (!workspaceBrowserSplit || !workspaceLeft || !workspaceSplitResizer) {
    return;
  }

  const stored = localStorage.getItem(WORKSPACE_SPLIT_STORAGE_KEY);
  const initial = stored ? Number.parseFloat(stored) : WORKSPACE_SPLIT_DEFAULT;
  if (Number.isFinite(initial) && initial > 10 && initial < 90) {
    workspaceBrowserSplit.style.setProperty("--workspace-left-width", `${initial}%`);
  }

  let dragging = false;

  const applyWidthFromPointer = (clientX) => {
    const rect = workspaceBrowserSplit.getBoundingClientRect();
    if (rect.width <= 0) {
      return;
    }
    const pct = ((clientX - rect.left) / rect.width) * 100;
    const clamped = Math.min(78, Math.max(22, pct));
    workspaceBrowserSplit.style.setProperty(
      "--workspace-left-width",
      `${clamped}%`,
    );
    localStorage.setItem(WORKSPACE_SPLIT_STORAGE_KEY, String(clamped));
  };

  const stopDrag = () => {
    if (!dragging) {
      return;
    }
    dragging = false;
    workspaceSplitResizer.classList.remove("is-dragging");
    document.body.classList.remove("workspace-split-dragging");
    document.removeEventListener("pointermove", onPointerMove);
    document.removeEventListener("pointerup", stopDrag);
    document.removeEventListener("pointercancel", stopDrag);
  };

  /** @param {PointerEvent} e */
  const onPointerMove = (e) => {
    if (!dragging) {
      return;
    }
    applyWidthFromPointer(e.clientX);
  };

  workspaceSplitResizer.addEventListener("pointerdown", (e) => {
    if (e.button !== 0) {
      return;
    }
    dragging = true;
    workspaceSplitResizer.classList.add("is-dragging");
    document.body.classList.add("workspace-split-dragging");
    workspaceSplitResizer.setPointerCapture(e.pointerId);
    applyWidthFromPointer(e.clientX);
    document.addEventListener("pointermove", onPointerMove);
    document.addEventListener("pointerup", stopDrag);
    document.addEventListener("pointercancel", stopDrag);
    e.preventDefault();
  });

  workspaceSplitResizer.addEventListener("keydown", (e) => {
    const rect = workspaceBrowserSplit.getBoundingClientRect();
    const current =
      Number.parseFloat(
        getComputedStyle(workspaceBrowserSplit)
          .getPropertyValue("--workspace-left-width")
          .replace("%", ""),
      ) || WORKSPACE_SPLIT_DEFAULT;
    let next = current;
    if (e.key === "ArrowLeft") {
      next = current - (e.shiftKey ? 5 : 2);
    } else if (e.key === "ArrowRight") {
      next = current + (e.shiftKey ? 5 : 2);
    } else {
      return;
    }
    e.preventDefault();
    const clamped = Math.min(78, Math.max(22, next));
    workspaceBrowserSplit.style.setProperty(
      "--workspace-left-width",
      `${clamped}%`,
    );
    localStorage.setItem(WORKSPACE_SPLIT_STORAGE_KEY, String(clamped));
    applyWidthFromPointer(rect.left + (rect.width * clamped) / 100);
  });
}

const DEDICATED_JD_SPLIT_STORAGE_KEY = "resume-builder.dedicated-jd-width";
const DEDICATED_JD_SPLIT_DEFAULT = 38;

function setupDedicatedBrowserSplitResize() {
  if (!dedicatedBrowserLayout || !dedicatedBrowserSide || !dedicatedBrowserSplitResizer) {
    return;
  }
  if (dedicatedBrowserSplitResizer.dataset.splitBound === "1") {
    return;
  }
  dedicatedBrowserSplitResizer.dataset.splitBound = "1";

  const stored = localStorage.getItem(DEDICATED_JD_SPLIT_STORAGE_KEY);
  const initial = stored ? Number.parseFloat(stored) : DEDICATED_JD_SPLIT_DEFAULT;
  if (Number.isFinite(initial) && initial > 10 && initial < 90) {
    dedicatedBrowserLayout.style.setProperty("--dedicated-jd-width", `${initial}%`);
  }

  let dragging = false;

  const applyWidthFromPointer = (clientX) => {
    const rect = dedicatedBrowserLayout.getBoundingClientRect();
    if (rect.width <= 0) {
      return;
    }
    const pct = ((clientX - rect.left) / rect.width) * 100;
    const clamped = Math.min(72, Math.max(22, pct));
    dedicatedBrowserLayout.style.setProperty("--dedicated-jd-width", `${clamped}%`);
    localStorage.setItem(DEDICATED_JD_SPLIT_STORAGE_KEY, String(clamped));
  };

  const stopDrag = () => {
    if (!dragging) {
      return;
    }
    dragging = false;
    dedicatedBrowserSplitResizer.classList.remove("is-dragging");
    document.body.classList.remove("dedicated-browser-split-dragging");
    document.removeEventListener("pointermove", onPointerMove);
    document.removeEventListener("pointerup", stopDrag);
    document.removeEventListener("pointercancel", stopDrag);
  };

  /** @param {PointerEvent} e */
  const onPointerMove = (e) => {
    if (!dragging) {
      return;
    }
    applyWidthFromPointer(e.clientX);
  };

  dedicatedBrowserSplitResizer.addEventListener("pointerdown", (e) => {
    if (e.button !== 0) {
      return;
    }
    dragging = true;
    dedicatedBrowserSplitResizer.classList.add("is-dragging");
    document.body.classList.add("dedicated-browser-split-dragging");
    dedicatedBrowserSplitResizer.setPointerCapture(e.pointerId);
    applyWidthFromPointer(e.clientX);
    document.addEventListener("pointermove", onPointerMove);
    document.addEventListener("pointerup", stopDrag);
    document.addEventListener("pointercancel", stopDrag);
    e.preventDefault();
  });

  dedicatedBrowserSplitResizer.addEventListener("keydown", (e) => {
    const current =
      Number.parseFloat(
        getComputedStyle(dedicatedBrowserLayout)
          .getPropertyValue("--dedicated-jd-width")
          .replace("%", ""),
      ) || DEDICATED_JD_SPLIT_DEFAULT;
    let next = current;
    if (e.key === "ArrowLeft") {
      next = current - (e.shiftKey ? 5 : 2);
    } else if (e.key === "ArrowRight") {
      next = current + (e.shiftKey ? 5 : 2);
    } else {
      return;
    }
    e.preventDefault();
    const clamped = Math.min(72, Math.max(22, next));
    dedicatedBrowserLayout.style.setProperty("--dedicated-jd-width", `${clamped}%`);
    localStorage.setItem(DEDICATED_JD_SPLIT_STORAGE_KEY, String(clamped));
  });
}

function setupChatgptRefresh() {
  document.addEventListener("click", (e) => {
    const btn = e.target instanceof Element ? e.target.closest(".chatgpt-pane-refresh") : null;
    if (!btn) {
      return;
    }
    const profile = btn.getAttribute("data-chatgpt-profile");
    const promptId = btn.getAttribute("data-prompt");
    const tabId = btn.getAttribute("data-dedicated-browser-tab");
    let wv = null;
    if (tabId) {
      const entry = dedicatedBrowserTabs.find((t) => t.id === tabId);
      wv = entry?.paneEl?.querySelector(".chatgpt-webview") ?? null;
    } else if (promptId) {
      wv = getChatgptWebviewForKind(promptId);
    } else {
      wv = getChatgptWebviewForProfile(profile === "justin" ? "justin" : "olu");
    }
    if (!wv) {
      return;
    }
    try {
      delete wv.dataset.pbLoadRetry;
      loadChatgptWebviewIfNeeded(/** @type {HTMLElement} */ (wv), { force: true });
      const label = dedicatedPromptKind
        ? promptLabelForKind(dedicatedPromptKind)
        : promptId
          ? promptLabelForKind(promptId)
          : profile === "justin"
            ? "Justin"
            : "Oluwaseun";
      toast(`${label} ChatGPT refreshed.`);
    } catch (_err) {
      toast("Could not refresh ChatGPT.", "error");
    }
  });
}

purgePersistedInputHistory();
setupBuildPasteToasts();
setupDedicatedWindowChannel();
setupDedicatedWorkspaceTabs();
setupMainCompanyInput();
setupMainPromptWordDirectory();
setupEditPromptSettings();
setupPromptPopupHotkeys();
setupDedicatedCompanyInput();
setupPromptSectionHighlights();
setupAddPromptModal();
setupPreviewCopyHandlers();
setupGptResultModal();
setupChatgptPaneToolbarHandlers();
setupChatgptShowPanesHandlers();
setupGptResultReadyBar();
setupPromptPanelToggle();
try {
  setupChatgptZoom();
} catch (_e) {
  /* avoid breaking init() and the rest of the UI */
}
setupChatgptBrowserLogin();
setupChatgptRefresh();
setupWorkspaceSplitResize();
setupDedicatedBrowserSplitResize();
setupInputGptHandler();
setupWindowSettingsModal();
setupUsersManageModal();
setupDocxPreviewModal();
setupWindowAppearanceControls();
setupTitleBarWindowButtons();
init();
