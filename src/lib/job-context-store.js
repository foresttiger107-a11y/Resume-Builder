/** @typedef {{ jd: string, company: string, updatedAt: string }} ProfileJobContext */

/**
 * @param {Record<string, unknown>} settings
 * @returns {Record<string, ProfileJobContext>}
 */
function readProfileJobContextMap(settings) {
  /** @type {Record<string, ProfileJobContext>} */
  const out = {};
  const raw = settings?.profileJobContext;
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) {
    return out;
  }
  for (const [key, value] of Object.entries(raw)) {
    if (!value || typeof value !== "object" || Array.isArray(value)) {
      continue;
    }
    const rec = /** @type {Record<string, unknown>} */ (value);
    out[String(key)] = {
      jd: String(rec.jd ?? ""),
      company: String(rec.company ?? ""),
      updatedAt: String(rec.updatedAt ?? ""),
    };
  }
  return out;
}

/**
 * @param {ProfileJobContext | null | undefined} current
 * @param {{ jd?: unknown, company?: unknown }} patch
 * @returns {ProfileJobContext}
 */
function mergeProfileJobContext(current, patch) {
  const jdProvided = patch && Object.prototype.hasOwnProperty.call(patch, "jd");
  const companyProvided =
    patch && Object.prototype.hasOwnProperty.call(patch, "company");

  const jd = jdProvided ? String(patch.jd ?? "") : String(current?.jd ?? "");
  const company = companyProvided
    ? String(patch.company ?? "")
    : String(current?.company ?? "");

  return {
    jd,
    company,
    updatedAt: new Date().toISOString(),
  };
}

/**
 * @param {Record<string, unknown>} settings
 * @param {string} profileId
 * @param {ProfileJobContext} context
 * @returns {Record<string, unknown>}
 */
function withProfileJobContext(settings, profileId, context) {
  const map = readProfileJobContextMap(settings);
  map[profileId] = context;
  return {
    ...settings,
    profileJobContext: map,
  };
}

module.exports = {
  readProfileJobContextMap,
  mergeProfileJobContext,
  withProfileJobContext,
};
