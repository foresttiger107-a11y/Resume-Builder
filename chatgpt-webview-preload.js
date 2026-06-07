/**
 * Runs in the ChatGPT <webview> guest. Forwards Ctrl/Cmd + wheel to the embedder
 * so the host can call webContents.setZoomFactor. Injects full composer text via IPC
 * without trimming (paragraph DOM, innerText, then chunked insertText).
 */
const { ipcRenderer } = require("electron");

/** Block passkeys only — do not break password managers or email/password login. */
(function disablePasskeysInWebview() {
  const passkeyError = () =>
    Promise.reject(
      new DOMException(
        "Passkeys are disabled in Resume Builder. Use Sign in → email and password.",
        "NotAllowedError",
      ),
    );
  try {
    if (navigator.credentials) {
      const origGet = navigator.credentials.get?.bind(navigator.credentials);
      const origCreate = navigator.credentials.create?.bind(navigator.credentials);
      if (origGet) {
        navigator.credentials.get = (options) =>
          options && "publicKey" in options ? passkeyError() : origGet(options);
      }
      if (origCreate) {
        navigator.credentials.create = (options) =>
          options && "publicKey" in options ? passkeyError() : origCreate(options);
      }
    }
    if (typeof window.PublicKeyCredential === "function") {
      PublicKeyCredential.isUserVerifyingPlatformAuthenticatorAvailable = () =>
        Promise.resolve(false);
      PublicKeyCredential.isConditionalMediationAvailable = () => Promise.resolve(false);
    }
  } catch (_e) {
    /* ignore */
  }
})();

function zoomChord(e) {
  return e.ctrlKey || e.metaKey;
}

window.addEventListener(
  "wheel",
  (e) => {
    if (!zoomChord(e)) return;
    e.preventDefault();
    e.stopPropagation();
    ipcRenderer.sendToHost("gpt-zoom-wheel", e.deltaY);
  },
  { capture: true, passive: false },
);

function normalizeNewlines(val) {
  return String(val ?? "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
}

/** Max chars per insertText call (avoids browser limits on huge strings). */
const INSERT_TEXT_CHUNK_SIZE = 500;

/** Fraction of expected length required before we accept an injection attempt. */
const INJECT_LENGTH_RATIO = 0.98;

function sleep(ms) {
  return new Promise((resolve) => {
    setTimeout(resolve, ms);
  });
}

function fireInput(el, inputType = "insertText") {
  try {
    el.dispatchEvent(
      new InputEvent("input", { bubbles: true, inputType }),
    );
  } catch (_e) {
    el.dispatchEvent(new Event("input", { bubbles: true }));
  }
  try {
    el.dispatchEvent(new Event("change", { bubbles: true }));
  } catch (_e) {
    /* ignore */
  }
}

function setTextareaValue(el, val) {
  const t = normalizeNewlines(val);
  const proto = window.HTMLTextAreaElement.prototype;
  const d = Object.getOwnPropertyDescriptor(proto, "value");
  if (d && d.set) {
    d.set.call(el, t);
  } else {
    el.value = t;
  }
  fireInput(el);
}

function readComposerText(el) {
  if (!el) {
    return "";
  }
  if (el.tagName === "TEXTAREA") {
    return normalizeNewlines(el.value || "");
  }
  return normalizeNewlines(el.innerText || el.textContent || "");
}

function hasFullComposerText(el, expected) {
  const got = readComposerText(el);
  if (!expected.length) {
    return !got.length;
  }
  if (got.length >= expected.length * INJECT_LENGTH_RATIO) {
    return true;
  }
  if (got === expected) {
    return true;
  }
  return got.length >= expected.length - 8 && expected.startsWith(got);
}

function findComposerRoot() {
  const selectors = [
    "textarea#prompt-textarea",
    'motion-textarea textarea',
    'textarea[data-id="root"]',
    '[data-testid="composer-textarea"]',
    "#prompt-textarea .ProseMirror",
    '.ProseMirror[contenteditable="true"]',
    'div#prompt-textarea[contenteditable="true"]',
    '[contenteditable="true"][id="prompt-textarea"]',
    "#prompt-textarea",
    "motion-textarea",
    "motion-div#prompt-textarea",
    "motion-div #prompt-textarea",
    "motion-div[contenteditable='true']",
  ];
  for (const sel of selectors) {
    const el = document.querySelector(sel);
    if (el && el.offsetParent !== null) {
      return el;
    }
  }
  const cand = document.querySelectorAll("textarea, [contenteditable='true']");
  for (let j = 0; j < cand.length; j++) {
    const node = cand[j];
    if (node.offsetParent === null) continue;
    const ph = (node.getAttribute && node.getAttribute("data-placeholder")) || "";
    const id = node.id || "";
    if (id === "prompt-textarea" || ph.toLowerCase().indexOf("message") >= 0) {
      return node;
    }
  }
  return null;
}

/** Prefer the real typing surface (textarea or ProseMirror), not the outer wrapper. */
function resolveComposerEditable(root) {
  if (!root) {
    return null;
  }
  if (root.tagName === "TEXTAREA") {
    return root;
  }
  const pm =
    root.querySelector?.(".ProseMirror") ||
    document.querySelector("#prompt-textarea .ProseMirror, .ProseMirror");
  if (pm && (pm.isContentEditable || pm.getAttribute("contenteditable") === "true")) {
    return pm;
  }
  const ta = root.querySelector?.("textarea");
  if (ta) {
    return ta;
  }
  if (root.isContentEditable) {
    return root;
  }
  const inner = root.querySelector?.("textarea, [contenteditable='true']");
  return inner || null;
}

function clearEditable(el) {
  el.focus();
  if (el.tagName === "TEXTAREA") {
    setTextareaValue(el, "");
    return;
  }
  if (!el.isContentEditable) {
    return;
  }
  try {
    const range = document.createRange();
    range.selectNodeContents(el);
    const sel = window.getSelection();
    if (sel) {
      sel.removeAllRanges();
      sel.addRange(range);
      document.execCommand("delete", false, null);
      return;
    }
  } catch (_e) {
    /* fall through */
  }
  el.replaceChildren();
}

function insertNewlineInEditable() {
  try {
    if (document.execCommand("insertLineBreak", false, null)) {
      return true;
    }
    if (document.execCommand("insertParagraph", false, null)) {
      return true;
    }
  } catch (_e) {
    /* ignore */
  }
  return false;
}

function insertTextInEditable(text) {
  if (!text.length) {
    return;
  }
  for (let i = 0; i < text.length; i += INSERT_TEXT_CHUNK_SIZE) {
    const chunk = text.slice(i, i + INSERT_TEXT_CHUNK_SIZE);
    let inserted = false;
    try {
      inserted = document.execCommand("insertText", false, chunk);
    } catch (_e) {
      inserted = false;
    }
    if (inserted) {
      continue;
    }
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      const range = sel.getRangeAt(0);
      const node = document.createTextNode(chunk);
      range.insertNode(node);
      range.setStartAfter(node);
      range.collapse(true);
      sel.removeAllRanges();
      sel.addRange(range);
      continue;
    }
    const active = document.activeElement;
    if (active && active.isContentEditable) {
      active.appendChild(document.createTextNode(chunk));
    }
  }
}

/** One <p> per source line — full document in one DOM update (no mid-stream trim). */
function setContentEditableParagraphs(el, val) {
  const lines = normalizeNewlines(val).split("\n");
  el.focus();
  const frag = document.createDocumentFragment();
  for (let i = 0; i < lines.length; i++) {
    const p = document.createElement("p");
    if (lines[i].length) {
      p.textContent = lines[i];
    } else {
      p.appendChild(document.createElement("br"));
    }
    frag.appendChild(p);
  }
  el.replaceChildren(frag);
  fireInput(el);
}

/** Preserve exact newlines/spaces in a single pre-wrap block. */
function setContentEditablePreWrap(el, val) {
  const t = normalizeNewlines(val);
  el.focus();
  const block = document.createElement("div");
  block.style.whiteSpace = "pre-wrap";
  block.textContent = t;
  el.replaceChildren(block);
  fireInput(el);
}

function setContentEditableInnerText(el, val) {
  const t = normalizeNewlines(val);
  el.focus();
  el.innerText = t;
  fireInput(el);
}

/** Chunked insertText fallback when DOM bulk set is rejected by ProseMirror. */
async function setContentEditableLineByLineAsync(el, val) {
  const lines = normalizeNewlines(val).split("\n");
  clearEditable(el);
  el.focus();

  for (let i = 0; i < lines.length; i++) {
    if (i > 0) {
      if (!insertNewlineInEditable()) {
        const br = document.createElement("br");
        const sel = window.getSelection();
        if (sel && sel.rangeCount > 0) {
          const range = sel.getRangeAt(0);
          range.insertNode(br);
          range.setStartAfter(br);
          range.collapse(true);
          sel.removeAllRanges();
          sel.addRange(range);
        } else {
          el.appendChild(br);
        }
      }
    }
    insertTextInEditable(lines[i]);
    if (i % 4 === 3) {
      await sleep(0);
    }
  }
  fireInput(el);
}

/**
 * @param {HTMLElement} el
 * @param {string} expected
 * @returns {Promise<boolean>}
 */
async function injectIntoEditable(el, expected) {
  const t = normalizeNewlines(expected);
  if (el.tagName === "TEXTAREA") {
    setTextareaValue(el, t);
    return hasFullComposerText(el, t);
  }
  if (!el.isContentEditable) {
    return false;
  }

  clearEditable(el);
  setContentEditableParagraphs(el, t);
  if (hasFullComposerText(el, t)) {
    return true;
  }

  clearEditable(el);
  setContentEditablePreWrap(el, t);
  if (hasFullComposerText(el, t)) {
    return true;
  }

  clearEditable(el);
  setContentEditableInnerText(el, t);
  if (hasFullComposerText(el, t)) {
    return true;
  }

  clearEditable(el);
  await setContentEditableLineByLineAsync(el, t);
  return hasFullComposerText(el, t);
}

function findComposer() {
  return findComposerRoot();
}

/**
 * @param {string} val
 * @returns {Promise<"ok" | "empty" | "no-input">}
 */
async function injectComposerText(val) {
  const t = normalizeNewlines(val);
  if (!t.length) {
    return "empty";
  }
  const root = findComposerRoot();
  if (!root) {
    return "no-input";
  }
  const el = resolveComposerEditable(root);
  if (!el) {
    return "no-input";
  }

  const ok = await injectIntoEditable(el, t);
  return ok ? "ok" : "no-input";
}

ipcRenderer.on("pb-inject-composer", (_event, text) => {
  Promise.resolve(injectComposerText(text))
    .then((result) => {
      ipcRenderer.sendToHost("pb-inject-result", result);
    })
    .catch(() => {
      ipcRenderer.sendToHost("pb-inject-result", "error");
    });
});
