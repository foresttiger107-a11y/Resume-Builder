/**
 * GPT resume JSON (see resume_structure.json) → same parsed shape as
 * parsePlainResumeToJustinFields for Justin McGhee .docx (template + programmatic).
 */

/**
 * @param {unknown} obj
 * @returns {boolean}
 */
function looksLikeResumeStructure(obj) {
  if (!obj || typeof obj !== "object" || Array.isArray(obj)) {
    return false;
  }
  const o = /** @type {Record<string, unknown>} */ (obj);
  if (Array.isArray(o.work_experience) && o.work_experience.length > 0) {
    return true;
  }
  if (Array.isArray(o.work_xperience) && o.work_xperience.length > 0) {
    return true;
  }
  if (typeof o.role_title === "string" && o.role_title.trim().length > 2) {
    return true;
  }
  if (typeof o.career_summary === "string" && o.career_summary.trim().length > 20) {
    return true;
  }
  if (Array.isArray(o.skills) && o.skills.length > 0) {
    return true;
  }
  return false;
}

/**
 * Download / Word export: parsed resume_structure JSON with at least one job.
 * @param {unknown} obj
 * @returns {boolean}
 */
function isCompleteResumeStructureJsonForExport(obj) {
  if (!obj || typeof obj !== "object" || Array.isArray(obj)) {
    return false;
  }
  if (!looksLikeResumeStructure(obj)) {
    return false;
  }
  const rec = /** @type {Record<string, unknown>} */ (obj);
  const wx = Array.isArray(rec.work_experience)
    ? rec.work_experience
    : Array.isArray(rec.work_xperience)
      ? rec.work_xperience
      : [];
  return wx.length > 0;
}

/**
 * @param {string} text
 * @returns {unknown | null}
 */
function tryParseJsonObject(text) {
  const t = String(text || "").trim();
  if (!t) {
    return null;
  }
  try {
    return JSON.parse(t);
  } catch {
    /* continue */
  }
  const i = t.indexOf("{");
  const j = t.lastIndexOf("}");
  if (i >= 0 && j > i) {
    try {
      return JSON.parse(t.slice(i, j + 1));
    } catch {
      return null;
    }
  }
  return null;
}

/**
 * Extract resume_structure-style JSON from a ChatGPT reply (fenced ```json or raw object).
 * @param {string} raw
 * @returns {object | null}
 */
function extractResumeJsonObjectFromGptText(raw) {
  const s0 = String(raw || "").replace(/\r\n/g, "\n");
  let best = null;
  let bestScore = 0;
  const fenceRe = /```(?:json)?\s*\n?([\s\S]*?)```/gi;
  let m;
  while ((m = fenceRe.exec(s0)) !== null) {
    const inner = m[1].trim();
    const o = tryParseJsonObject(inner);
    if (o && typeof o === "object" && !Array.isArray(o) && looksLikeResumeStructure(o)) {
      const score = inner.length;
      if (score > bestScore) {
        best = /** @type {object} */ (o);
        bestScore = score;
      }
    }
  }
  if (best) {
    return best;
  }
  const o2 = tryParseJsonObject(s0);
  if (o2 && typeof o2 === "object" && !Array.isArray(o2) && looksLikeResumeStructure(o2)) {
    return /** @type {object} */ (o2);
  }
  return null;
}

/** @param {unknown} v */
function str(v) {
  return String(v ?? "").trim();
}

/** @param {unknown} arr */
function asStringArray(arr) {
  if (!Array.isArray(arr)) {
    return [];
  }
  return arr.map((x) => str(x)).filter(Boolean);
}

/**
 * One-line job header (left of the date tab), matching common resume layout:
 * Role [– Specialization] | Company | City, Region (Remote|Hybrid|Onsite).
 * GPT may omit optional fields; still produces a valid three-part line when possible.
 * @param {Record<string, unknown>} w one work_experience object
 */
function buildWorkExperienceHeaderLeft(w) {
  const rt = str(w.role_title);
  const spec = str(
    w.specialization ??
      w.role_specialization ??
      w.specialty ??
      w.focus_area ??
      w.focus,
  );
  const company = str(w.company);
  const geo = getWorkExperienceGeo(w);
  /** City/geo line only — not the same as work mode (Remote/Hybrid). */
  const lt = str(w.location_type ?? w.locationType);

  const roleSeg = spec ? `${rt} \u2013 ${spec}` : rt;
  let placeSeg = "";
  if (geo && lt) {
    placeSeg = `${geo} (${lt})`;
  } else if (geo) {
    placeSeg = geo;
  } else if (lt) {
    placeSeg = `(${lt})`;
  }
  const parts = [roleSeg, company, placeSeg].filter((p) => p && p.trim());
  return parts.join(" | ");
}

/**
 * City / region from GPT JSON (`location`, `office_location`, etc.).
 * @param {Record<string, unknown>} w
 */
function getWorkExperienceGeo(w) {
  return str(
    w.location ??
      w.office_location ??
      w.officeLocation ??
      w.city_state_country ??
      w.city_region ??
      w.city_state ??
      w.location_detail,
  );
}

/**
 * `Middletown, DW, US (Remote)` or partial when fields are missing.
 * @param {string} geo
 * @param {string} locationType
 */
function formatWorkExperiencePlace(geo, locationType) {
  const g = String(geo || "").trim();
  const lt = String(locationType || "").trim();
  if (g && lt) {
    return `${g} (${lt})`;
  }
  if (g) {
    return g;
  }
  if (lt) {
    return `(${lt})`;
  }
  return "";
}

/**
 * Suffix after bold `{role_title} | {company}`: ` | City (Remote)` or ` | (Hybrid)`.
 * @param {string} geo
 * @param {string} locationType
 */
function formatWorkExperienceLocationLine(geo, locationType) {
  const g = String(geo || "").trim();
  const lt = String(locationType || "").trim();
  if (g && lt) {
    return ` | ${g} (${lt})`;
  }
  if (g) {
    return ` | ${g}`;
  }
  if (lt) {
    return ` | (${lt})`;
  }
  return "";
}

/**
 * Split `buildWorkExperienceHeaderLeft` output for Word: bold first segment,
 * regular remainder ` | company | (location)…`, right tab + duration.
 * @param {string} experienceLineLeft
 * @returns {{ line_role: string; header_regular_suffix: string }}
 */
function splitExperienceLineForWord(experienceLineLeft) {
  const segs = String(experienceLineLeft || "")
    .split(/\s*\|\s*/)
    .map((x) => x.trim())
    .filter(Boolean);
  return {
    line_role: segs[0] || "",
    header_regular_suffix: segs.length > 1 ? ` | ${segs.slice(1).join(" | ")}` : "",
  };
}

/**
 * Rows for UI / previews: same segments as the Justin Word template job line.
 * @param {unknown} wx
 * @returns {{ line_role: string; header_regular_suffix: string; experience_line_left: string; duration: string }[]}
 */
function buildWorkExperienceWordPreviewRows(wx) {
  if (!Array.isArray(wx)) {
    return [];
  }
  /** @type {{ company: string; location: string; location_type: string; place: string; role_title: string; duration: string }[]} */
  const out = [];
  for (const entry of wx) {
    if (!entry || typeof entry !== "object") {
      continue;
    }
    const w = /** @type {Record<string, unknown>} */ (entry);
    const location = getWorkExperienceGeo(w);
    const location_type = str(w.location_type ?? w.locationType);
    out.push({
      company: str(w.company),
      location,
      location_type,
      place: formatWorkExperiencePlace(location, location_type),
      location_line: formatWorkExperienceLocationLine(location, location_type),
      role_title: str(w.role_title),
      duration: str(w.duration),
    });
  }
  return out;
}

/**
 * GPT may send bullet_points as string[] or a single newline-separated string.
 * Strips common list markers from each line.
 * @param {unknown} raw
 * @returns {string[]}
 */
function normalizeBulletLines(raw) {
  if (raw == null) {
    return [];
  }
  if (Array.isArray(raw)) {
    /** @type {string[]} */
    const out = [];
    for (const item of raw) {
      if (item == null) {
        continue;
      }
      if (typeof item === "string") {
        const t = stripOneBulletLine(item);
        if (t) {
          out.push(t);
        }
        continue;
      }
      if (typeof item === "object" && item !== null && "text" in item) {
        const t = stripOneBulletLine(str(/** @type {Record<string, unknown>} */ (item).text));
        if (t) {
          out.push(t);
        }
        continue;
      }
      const t = stripOneBulletLine(String(item));
      if (t) {
        out.push(t);
      }
    }
    return out;
  }
  if (typeof raw === "string") {
    return raw
      .split(/\r?\n/)
      .map((ln) => stripOneBulletLine(ln))
      .filter(Boolean);
  }
  return [];
}

/**
 * @param {string} line
 * @returns {string}
 */
function stripOneBulletLine(line) {
  return String(line || "")
    .replace(/^\s*(?:[-*+•·▪]|\d+[.)])\s+/, "")
    .replace(/^\s*[•·▪]\s*/, "")
    .trim();
}

/**
 * Map resume_structure.json object to parsePlainResumeToJustinFields output.
 * @param {object} obj
 */
function resumeJsonObjectToJustinFields(obj) {
  const nameRaw = str(
    /** @type {Record<string, unknown>} */ (obj).full_name ??
      /** @type {Record<string, unknown>} */ (obj).fullName ??
      /** @type {Record<string, unknown>} */ (obj).name,
  );
  const roleTitle = str(/** @type {Record<string, unknown>} */ (obj).role_title);
  const displayName = (nameRaw || roleTitle || "Resume").toUpperCase();

  const contactLine = str(
    /** @type {Record<string, unknown>} */ (obj).contact_line ??
      /** @type {Record<string, unknown>} */ (obj).contactLine ??
      /** @type {Record<string, unknown>} */ (obj).contact,
  );

  const summary = str(/** @type {Record<string, unknown>} */ (obj).career_summary);

  /** @type {{ header: string; detail: string }[]} */
  const skillsBlocks = [];
  const skillsArr = /** @type {unknown} */ (obj).skills;
  if (Array.isArray(skillsArr)) {
    for (const block of skillsArr) {
      if (!block || typeof block !== "object") {
        continue;
      }
      const b = /** @type {Record<string, unknown>} */ (block);
      const header = str(b.name);
      const items = asStringArray(b.items);
      const detail = items.join(", ");
      if (header || detail) {
        skillsBlocks.push({ header, detail });
      }
    }
  }

  const skillsRaw = skillsBlocks
    .map((b) => {
      if (b.header && b.detail) {
        return `${b.header}\n${b.detail}`;
      }
      return b.header || b.detail;
    })
    .filter(Boolean)
    .join("\n\n");

  /** @type {{ titlePart: string; dateRight: string; bullets: string[] }[]} */
  const jobs = [];
  const rec = /** @type {Record<string, unknown>} */ (obj);
  const wx =
    Array.isArray(rec.work_experience) ?
      rec.work_experience
    : Array.isArray(rec.work_xperience) ?
      rec.work_xperience
    : null;
  if (Array.isArray(wx)) {
    for (const entry of wx) {
      if (!entry || typeof entry !== "object") {
        continue;
      }
      const w = /** @type {Record<string, unknown>} */ (entry);
      const titlePart = buildWorkExperienceHeaderLeft(w);
      const duration = str(w.duration);
      const rawBullets =
        w.bullet_points ??
        w.bulletPoints ??
        w.bullets ??
        w.achievements ??
        w.highlights;
      const bullets = normalizeBulletLines(rawBullets);
      if (titlePart || bullets.length) {
        jobs.push({
          titlePart,
          dateRight: duration,
          bullets,
          expCompany: str(w.company),
          expLocation: getWorkExperienceGeo(w),
          expLocationType: str(w.location_type ?? w.locationType),
          expRoleTitle: str(w.role_title),
        });
      }
    }
  }

  const eduArr = /** @type {unknown} */ (obj).education;
  /** @type {string[]} */
  const eduLines = [];
  if (Array.isArray(eduArr)) {
    for (const e of eduArr) {
      if (!e || typeof e !== "object") {
        continue;
      }
      const ed = /** @type {Record<string, unknown>} */ (e);
      const degree = str(ed.degree);
      const inst = str(ed.institution);
      const dur = str(ed.duration);
      const parts = [degree, inst, dur].filter(Boolean);
      if (parts.length) {
        eduLines.push(parts.join(" — "));
      }
    }
  }
  const education = eduLines.join("\n");

  const rawExperienceBody = jobs
    .map((j) => {
      const head = [j.titlePart, j.dateRight].filter(Boolean).join("\n");
      const bs = j.bullets.map((b) => `• ${b}`).join("\n");
      return [head, bs].filter(Boolean).join("\n");
    })
    .join("\n\n");

  return {
    fullName: displayName,
    role: roleTitle,
    contactLine,
    summary,
    skills: skillsRaw,
    skillsRaw,
    skillsBlocks,
    education,
    jobs,
    rawExperienceBody,
  };
}

/**
 * Parse a raw JSON string as resume_structure (no GPT fences required).
 * @param {string} text
 * @returns {object | null}
 */
function parseResumeStructureJson(text) {
  const o = tryParseJsonObject(String(text || "").trim());
  if (!o || typeof o !== "object" || Array.isArray(o)) {
    return null;
  }
  if (!looksLikeResumeStructure(o)) {
    return null;
  }
  return /** @type {object} */ (o);
}

const JUSTIN_MCGHEE_DOCX_NAME = "Justin Mcghee";
const OLUWASEUN_ADENIJI_DOCX_NAME = "Oluwaseun Adeniji";

/**
 * @param {string} personName
 * @param {object | null | undefined} obj
 * @returns {string}
 */
function suggestResumeDocxFileStem(personName, obj) {
  if (!obj || typeof obj !== "object" || Array.isArray(obj)) {
    return `${personName} - Resume`;
  }
  const rec = /** @type {Record<string, unknown>} */ (obj);
  const role = String(rec.role_title ?? rec.role ?? rec.roleTitle ?? "")
    .replace(/[\\/:*?"<>|]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 120);
  if (role) {
    return `${personName} - ${role}`;
  }
  return `${personName} - Resume`;
}

/**
 * Default save-as filename stem: "Justin Mcghee - {role_title}".
 * @param {object | null | undefined} obj resume_structure JSON
 * @returns {string} filename stem without .docx
 */
function suggestJustinMcgheeDocxFileStem(obj) {
  return suggestResumeDocxFileStem(JUSTIN_MCGHEE_DOCX_NAME, obj);
}

/**
 * Default save-as filename stem: "Oluwaseun Adeniji - {role_title}".
 * @param {object | null | undefined} obj resume_structure JSON
 * @returns {string} filename stem without .docx
 */
function suggestOluwaseunAdenijiDocxFileStem(obj) {
  return suggestResumeDocxFileStem(OLUWASEUN_ADENIJI_DOCX_NAME, obj);
}

/**
 * @param {"justin" | "olu"} profile
 * @param {object | null | undefined} obj
 * @returns {string}
 */
function suggestResumeDocxFileStemForProfile(profile, obj) {
  return profile === "olu"
    ? suggestOluwaseunAdenijiDocxFileStem(obj)
    : suggestJustinMcgheeDocxFileStem(obj);
}

module.exports = {
  looksLikeResumeStructure,
  isCompleteResumeStructureJsonForExport,
  extractResumeJsonObjectFromGptText,
  parseResumeStructureJson,
  resumeJsonObjectToJustinFields,
  normalizeBulletLines,
  buildWorkExperienceHeaderLeft,
  buildWorkExperienceWordPreviewRows,
  splitExperienceLineForWord,
  getWorkExperienceGeo,
  formatWorkExperiencePlace,
  formatWorkExperienceLocationLine,
  suggestJustinMcgheeDocxFileStem,
  suggestOluwaseunAdenijiDocxFileStem,
  suggestResumeDocxFileStemForProfile,
  OLUWASEUN_ADENIJI_DOCX_NAME,
  JUSTIN_MCGHEE_DOCX_NAME,
};
