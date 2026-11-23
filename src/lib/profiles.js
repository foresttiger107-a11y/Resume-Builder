/** @typedef {"olu" | "justin"} ChatgptProfile */

/**
 * @typedef {{
 *   id: string,
 *   label: string,
 *   chatgptProfile: ChatgptProfile,
 *   promptPath: string,
 *   structurePath: string,
 * }} ProfileSummary
 */

/** @param {unknown} raw */
function normalizeProfileId(raw) {
  const id = String(raw ?? "").trim();
  return id.startsWith("custom:") ? id : "";
}

/**
 * @param {{ id: string, label: string, promptPath: string, structurePath: string, chatgptProfile: ChatgptProfile }} entry
 * @returns {ProfileSummary | null}
 */
function profileFromEntry(entry) {
  if (!entry?.id?.startsWith("custom:")) {
    return null;
  }
  if (!entry.label?.trim() || !entry.promptPath?.trim() || !entry.structurePath?.trim()) {
    return null;
  }
  return {
    id: entry.id,
    label: entry.label,
    chatgptProfile: entry.chatgptProfile === "justin" ? "justin" : "olu",
    promptPath: entry.promptPath,
    structurePath: entry.structurePath,
  };
}

/**
 * User-managed profiles shown in the app (Manage users / build cards).
 * @param {Array<{ id: string, label: string, promptPath: string, structurePath: string, chatgptProfile: ChatgptProfile }>} customEntries
 * @returns {ProfileSummary[]}
 */
function listProfiles(customEntries = []) {
  /** @type {ProfileSummary[]} */
  const profiles = [];
  const seen = new Set();
  for (const entry of customEntries) {
    const profile = profileFromEntry(entry);
    if (!profile || seen.has(profile.id)) {
      continue;
    }
    seen.add(profile.id);
    profiles.push(profile);
  }
  return profiles;
}

/**
 * @param {unknown} rawId
 * @param {Array<{ id: string, label: string, promptPath: string, structurePath: string, chatgptProfile: ChatgptProfile }>} customEntries
 * @returns {ProfileSummary | null}
 */
function findProfile(rawId, customEntries = []) {
  const id = normalizeProfileId(rawId);
  if (!id) {
    return null;
  }
  const entry = customEntries.find((p) => p.id === id);
  return entry ? profileFromEntry(entry) : null;
}

/** @param {unknown} raw */
function normalizeProfileName(raw) {
  return String(raw ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

/** @param {string} a @param {string} b */
function levenshtein(a, b) {
  const m = a.length;
  const n = b.length;
  if (m === 0) {
    return n;
  }
  if (n === 0) {
    return m;
  }
  /** @type {number[]} */
  let prev = Array.from({ length: n + 1 }, (_, i) => i);
  /** @type {number[]} */
  let cur = new Array(n + 1);
  for (let i = 1; i <= m; i += 1) {
    cur[0] = i;
    for (let j = 1; j <= n; j += 1) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      cur[j] = Math.min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + cost);
    }
    prev = cur;
    cur = new Array(n + 1);
  }
  return prev[n];
}

/** @param {string} needle @param {string} hay */
function usernameMatchDistance(needle, hay) {
  if (!needle || !hay) {
    return Number.POSITIVE_INFINITY;
  }
  if (needle === hay) {
    return 0;
  }
  if (hay.includes(needle) || needle.includes(hay)) {
    return 1;
  }
  const maxLen = Math.max(needle.length, hay.length);
  const maxDistance = Math.max(2, Math.floor(maxLen * 0.18));
  return levenshtein(needle, hay) <= maxDistance
    ? levenshtein(needle, hay)
    : Number.POSITIVE_INFINITY;
}

/**
 * @param {unknown} username
 * @param {Array<{ id: string, label: string, promptPath: string, structurePath: string, chatgptProfile: ChatgptProfile }>} customEntries
 * @returns {ProfileSummary[]}
 */
function rankProfilesByUsername(username, customEntries = []) {
  const needle = normalizeProfileName(username);
  if (!needle) {
    return [];
  }
  return listProfiles(customEntries)
    .map((profile) => ({
      profile,
      distance: usernameMatchDistance(needle, normalizeProfileName(profile.label)),
    }))
    .filter((item) => Number.isFinite(item.distance))
    .sort((a, b) => a.distance - b.distance || a.profile.label.localeCompare(b.profile.label));
}

/**
 * @param {unknown} username
 * @param {Array<{ id: string, label: string, promptPath: string, structurePath: string, chatgptProfile: ChatgptProfile }>} customEntries
 * @returns {ProfileSummary | null}
 */
function findProfileByUsername(username, customEntries = []) {
  const ranked = rankProfilesByUsername(username, customEntries);
  if (ranked.length === 0) {
    return null;
  }
  if (ranked.length === 1) {
    return ranked[0].profile;
  }
  if (ranked[0].distance < ranked[1].distance) {
    return ranked[0].profile;
  }
  return null;
}

/**
 * @param {unknown} username
 * @param {Array<{ id: string, label: string, promptPath: string, structurePath: string, chatgptProfile: ChatgptProfile }>} customEntries
 * @returns {ProfileSummary[]}
 */
function findProfileCandidatesByUsername(username, customEntries = []) {
  const ranked = rankProfilesByUsername(username, customEntries);
  if (ranked.length === 0) {
    return [];
  }
  const bestDistance = ranked[0].distance;
  return ranked
    .filter((item) => item.distance === bestDistance)
    .map((item) => item.profile);
}

module.exports = {
  normalizeProfileId,
  normalizeProfileName,
  listProfiles,
  findProfile,
  findProfileByUsername,
  findProfileCandidatesByUsername,
};
