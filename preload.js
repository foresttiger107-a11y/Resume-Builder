const { contextBridge, ipcRenderer } = require("electron");
const path = require("path");
const { pathToFileURL } = require("url");

/** Never throw from preload — a failed path would break the whole window API. */
let chatgptZoomPreloadPath = "";
try {
  chatgptZoomPreloadPath = pathToFileURL(
    path.join(__dirname, "chatgpt-webview-preload.js"),
  ).href;
} catch (_e) {
  chatgptZoomPreloadPath = "";
}

contextBridge.exposeInMainWorld("pbPaths", {
  /** Absolute file URL so the ChatGPT <webview> preload resolves (empty if unavailable). */
  chatgptWebviewZoomPreload: chatgptZoomPreloadPath,
});

contextBridge.exposeInMainWorld("promptApi", {
  gptDebugLog: (msg) => ipcRenderer.invoke("gpt-debug-log", String(msg)),
  getAppVersion: () => ipcRenderer.invoke("get-app-version"),
  readPrompt: (relPath) => ipcRenderer.invoke("read-prompt", relPath),
  savePrompt: (relPath, content) =>
    ipcRenderer.invoke("save-prompt", relPath, content),
  pickPromptFile: () => ipcRenderer.invoke("pick-prompt-file"),
  pickStructureDocx: () => ipcRenderer.invoke("pick-structure-docx"),
  getCustomPrompts: () => ipcRenderer.invoke("get-custom-prompts"),
  saveCustomPrompt: (entry) => ipcRenderer.invoke("save-custom-prompt", entry),
  updateCustomPromptPath: (entry) =>
    ipcRenderer.invoke("update-custom-prompt-path", entry),
  deleteCustomPrompt: (id) => ipcRenderer.invoke("delete-custom-prompt", { id }),
  readPromptFilePath: (pathOrFileUrl) =>
    ipcRenderer.invoke("read-prompt-file-path", pathOrFileUrl),
  savePromptExternal: (pathOrFileUrl, content) =>
    ipcRenderer.invoke("save-prompt-external", pathOrFileUrl, content),
  saveGptDocx: (payload) => ipcRenderer.invoke("save-gpt-docx", payload),
  getDownloadDirectory: (downloadKey) =>
    ipcRenderer.invoke("get-download-directory", {
      downloadKey: String(downloadKey ?? "other"),
    }),
  pickDownloadDirectory: (downloadKey) =>
    ipcRenderer.invoke("pick-download-directory", {
      downloadKey: String(downloadKey ?? "other"),
    }),
  exportResumeJsonToDocx: () => ipcRenderer.invoke("export-resume-json-to-docx"),
  resolveWorkExperienceFromGptText: (text) =>
    ipcRenderer.invoke("resolve-work-experience-preview", text),
  checkGptDocxExport: (text) =>
    ipcRenderer.invoke("check-gpt-docx-export", text),
  openGptDownloadFolder: (payload) =>
    ipcRenderer.invoke("open-gpt-download-folder", payload),
  openGptDownloadFile: (payload) =>
    ipcRenderer.invoke("open-gpt-download-file", payload),
  readGptDocxFile: (filePath) =>
    ipcRenderer.invoke("read-gpt-docx-file", filePath),
  copyResumeFileToClipboard: (filePath) =>
    ipcRenderer.invoke("copy-resume-file-to-clipboard", filePath),
  showPathInFolder: (filePath) =>
    ipcRenderer.invoke("show-path-in-folder", filePath),
  openChatgptExternal: (profile) =>
    ipcRenderer.invoke("open-chatgpt-external", profile),
  openChatgptLoginWindow: (payload) =>
    ipcRenderer.invoke("open-chatgpt-login-window", payload),
  clearChatgptSession: (payload) =>
    ipcRenderer.invoke("clear-chatgpt-session", payload),
  onChatgptLoginComplete: (fn) => {
    ipcRenderer.on("chatgpt-login-complete", (_event, payload) => {
      fn(payload);
    });
  },
  openPromptWindow: (promptId) =>
    ipcRenderer.invoke("open-prompt-window", {
      prompt: String(promptId ?? "other"),
    }),
  closePromptWindow: (promptId) =>
    ipcRenderer.invoke("close-prompt-window", {
      id: String(promptId ?? ""),
    }),
  syncPromptPopupHotkeys: (payload) =>
    ipcRenderer.invoke("sync-prompt-popup-hotkeys", payload),
  onWindowPromptKind: (fn) => {
    ipcRenderer.on("window-prompt-kind", (_event, kind) => {
      fn(kind);
    });
  },
});

contextBridge.exposeInMainWorld("windowControls", {
  setTitle: (title) =>
    ipcRenderer.invoke("window-set-title", String(title ?? "")),
  setAlwaysOnTop: (flag) =>
    ipcRenderer.invoke("window-set-always-on-top", !!flag),
  minimize: () => ipcRenderer.invoke("window-minimize"),
  maximizeToggle: () => ipcRenderer.invoke("window-maximize-toggle"),
  isMaximized: () => ipcRenderer.invoke("window-is-maximized"),
  close: () => ipcRenderer.invoke("window-close"),
  onMaximized: (fn) => {
    ipcRenderer.on("window-max-state", (_event, value) => {
      fn(!!value);
    });
  },
});
