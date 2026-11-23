/**
 * Fills "resume structure.docx" via Docxtemplater (Word styles come from the template).
 * Resolution order: project root, then assets/ (bundled default).
 */

const fs = require("fs");
const path = require("path");
const PizZip = require("pizzip");
const Docxtemplater = require("docxtemplater");
const {
  parsePlainResumeToFields,
  stripMd,
  stripInlineMdBold,
} = require("./resume-docx.js");
const {
  extractResumeJsonObjectFromGptText,
  parseResumeStructureJson,
  resumeJsonObjectToFields,
  looksLikeResumeStructure,
  formatWorkExperiencePlace,
  formatWorkExperienceLocationLine,
} = require("./resume-json-to-fields.js");

/**
 * Human-readable message from docxtemplater (often throws "Multi error" with details in .properties.errors).
 * @param {unknown} err
 * @returns {string}
 */
function formatDocxtemplaterError(err) {
  if (!err || typeof err !== "object") {
    return String(err ?? "Word template error.");
  }
  const e = /** @type {{ message?: string; properties?: { errors?: unknown[] } }} */ (
    err
  );
  /** @type {string[]} */
  const parts = [];
  const subs = e.properties?.errors;
  if (Array.isArray(subs) && subs.length > 0) {
    for (const sub of subs.slice(0, 8)) {
      const item = /** @type {{ message?: string; properties?: { explanation?: string } }} */ (
        sub
      );
      const line =
        item.properties?.explanation ||
        item.message ||
        (typeof sub === "string" ? sub : "");
      if (line && String(line).trim()) {
        parts.push(String(line).trim());
      }
    }
    if (subs.length > 8) {
      parts.push(`(+${subs.length - 8} more template issues)`);
    }
  }
  if (parts.length === 0 && e.message && e.message !== "Multi error") {
    parts.push(String(e.message));
  }
  if (parts.length === 0) {
    return "Word template error. Check structure .docx placeholders in Manage users, or use the default template.";
  }
  return `Word template: ${parts.join(" — ")}`;
}

/** Bundled default structure template filename(s). */
const RESUME_TEMPLATE_FILENAMES = ["resume structure.docx"];

/**
 * @param {string} appRoot
 * @param {unknown} [_profile] unused; kept for call-site compatibility
 * @returns {string | null}
 */
function findResumeStructureTemplatePath(appRoot, _profile) {
  for (const name of RESUME_TEMPLATE_FILENAMES) {
    for (const dir of [appRoot, path.join(appRoot, "assets")]) {
      const candidate = path.join(dir, name);
      if (fs.existsSync(candidate)) {
        return candidate;
      }
    }
  }
  return null;
}

/** @param {string} appRoot */
function findDefaultStructureTemplatePath(appRoot) {
  return findResumeStructureTemplatePath(appRoot);
}

/**
 * @param {string} titlePart
 * @param {{ roleTitle?: string; company?: string; location?: string }} [j]
 */
function splitJobHeaderParts(titlePart, j) {
  const jr = j || {};
  if (jr.roleTitle || jr.company || jr.location) {
    return {
      role: String(jr.roleTitle || "").trim(),
      company: String(jr.company || "").trim(),
      location: String(jr.location || "").trim(),
    };
  }
  const parts = String(titlePart || "")
    .split(/\s*\|\s*/)
    .map((x) => x.trim())
    .filter(Boolean);
  return {
    role: parts[0] || "",
    company: parts[1] || "",
    location: parts.slice(2).join(" | ") || "",
  };
}

/**
 * Multiline string for a single `{bullet_points}` placeholder in Word (linebreaks: true → w:br).
 * @param {string[]} lines
 */
function formatExperienceBulletsDocx(lines) {
  return (lines || [])
    .map((t) => String(t || "").trim())
    .filter(Boolean)
    .map((t) => `\u2022\t${t}`)
    .join("\n");
}

/** @param {object} parsed */
function buildTemplateData(parsed) {
  const p = parsed;
  const jobRows = (p.jobs || []).map((j) => {
    const cleanBullets = (j.bullets || [])
      .map((b) => stripInlineMdBold(stripMd(String(b ?? "").trim())))
      .filter(Boolean);
    const parts = splitJobHeaderParts(String(j.titlePart || ""), j);
    const role = stripInlineMdBold(stripMd(parts.role));
    const company = stripInlineMdBold(stripMd(parts.company));
    const location = stripInlineMdBold(stripMd(parts.location));
    const rawTitle = stripInlineMdBold(stripMd(String(j.titlePart || "").trim()))
      .replace(/\s*\n\s*/g, " ")
      .replace(/\s+/g, " ")
      .trim();
    const titleJoined = [role, company, location].filter(Boolean).join(" | ");
    const title = titleJoined || rawTitle;
    const experienceLineLeft = rawTitle || titleJoined;
    const dateStr = stripInlineMdBold(stripMd(String(j.dateRight || "").trim()));
    const jr = /** @type {Record<string, unknown>} */ (j);
    const expCompany = stripInlineMdBold(
      stripMd(String(jr.expCompany ?? jr.exp_company ?? "").trim()),
    );
    const expLocation = stripInlineMdBold(
      stripMd(String(jr.expLocation ?? jr.exp_location ?? "").trim()),
    );
    const expLocationType = stripInlineMdBold(
      stripMd(String(jr.expLocationType ?? jr.exp_location_type ?? "").trim()),
    );
    const expRoleTitle = stripInlineMdBold(
      stripMd(String(jr.expRoleTitle ?? jr.exp_role_title ?? "").trim()),
    );
    const tmplCompany = expCompany || company;
    const tmplLocation = expLocation;
    const tmplLocationType = expLocationType;
    const tmplRoleTitle = expRoleTitle || role;
    const tmplPlace = formatWorkExperiencePlace(tmplLocation, tmplLocationType);
    const tmplLocationLine = formatWorkExperienceLocationLine(
      tmplLocation,
      tmplLocationType,
    );
    return {
      cleanBullets,
      role,
      company,
      location,
      title,
      dateStr,
      experienceLineLeft,
      tmplCompany,
      tmplLocation,
      tmplLocationType,
      tmplPlace,
      tmplLocationLine,
      tmplRoleTitle,
    };
  });

  const jobs = jobRows.map((r) => ({
    title: r.title,
    role: r.role,
    company: r.company,
    location: r.location,
    date: r.dateStr,
    bullets: r.cleanBullets.map((bullet) => ({ bullet })),
    /**
     * Legacy `{#jobs}` templates: `{#bullet_points}•\t{point}{/bullet_points}`.
     */
    bullet_points: r.cleanBullets.map((t) => ({
      point: stripInlineMdBold(stripMd(String(t || "").trim())),
    })),
    bullet_points_text: formatExperienceBulletsDocx(r.cleanBullets),
  }));

  /**
   * `{#work_experience}` / `{#work_xperience}`: `exp_algo` header line
   * Bold `{role_title} | {company}` + normal `{location_line}` + tab `{duration}`; then `{#bullet_points}…{bullet}…`.
   * Same array on `work_experience` and `work_xperience`. Do not use `{work_experience}` (scalar).
   * Optional debug string: `{work_experience_json}` (pretty-printed array).
   */
  const work_experience = jobRows.map((r) => ({
    company: r.tmplCompany,
    location: r.tmplLocation,
    location_type: r.tmplLocationType,
    place: r.tmplPlace,
    location_line: r.tmplLocationLine,
    role_title: r.tmplRoleTitle,
    duration: r.dateStr,
    experience_line_left: r.experienceLineLeft,
    bullet_points: r.cleanBullets.map((t) => ({
      bullet: stripInlineMdBold(stripMd(String(t || "").trim())),
    })),
  }));
  const work_xperience = work_experience;

  /** Plain string for `{work_experience_json}` — never use `{work_experience}` on an array. */
  let work_experience_json = "";
  try {
    work_experience_json = JSON.stringify(
      work_experience.map((row) => ({
        company: row.company,
        location: row.location,
        location_type: row.location_type,
        place: row.place,
        role_title: row.role_title,
        duration: row.duration,
        bullet_points: (row.bullet_points || []).map((b) => b.bullet),
      })),
      null,
      2,
    );
  } catch {
    work_experience_json = "";
  }

  const skillBlocks = (p.skillsBlocks || [])
    .filter((b) => (b.header || "").trim() || (b.detail || "").trim())
    .map((b) => ({
      header: stripInlineMdBold(stripMd((b.header || "").trim())),
      detail: stripInlineMdBold(stripMd((b.detail || "").trim())),
    }));

  const summaryText = stripInlineMdBold(stripMd(String(p.summary || "").trim()));
  const roleText = stripInlineMdBold(stripMd(String(p.role || "").trim()));
  const fullNameText = String(p.fullName || "");
  const contactText = stripInlineMdBold(stripMd(String(p.contactLine || "").trim()));
  const skillsText = stripInlineMdBold(stripMd(String(p.skillsRaw || "").trim()));
  const educationText = stripInlineMdBold(stripMd(String(p.education || "").trim()));

  const data = {
    fullName: fullNameText,
    full_name: fullNameText,
    role: roleText,
    role_title: roleText,
    contactLine: contactText,
    contact_line: contactText,
    summary: summaryText,
    career_summary: summaryText,
    careerSummary: summaryText,
    skill: skillsText,
    skills: skillsText,
    skillBlocks,
    education: educationText,
    jobs,
    work_xperience,
    work_experience,
    work_experience_json,
  };

  for (let i = 0; i < 10; i++) {
    const job = jobs[i];
    data[`job${i + 1}_title`] = job ? job.title : "";
    data[`job${i + 1}_date`] = job ? job.date : "";
    data[`job${i + 1}_bullets`] = job ? job.bullets : [];
  }

  return data;
}

/**
 * Build a .docx buffer from resume_structure JSON (GPT file or pasted string).
 * Uses resume structure.docx when present; otherwise the programmatic default layout.
 * @param {string | object} jsonInput raw JSON string, GPT reply, or parsed object
 * @param {string} appRoot
 * @returns {Promise<Buffer>}
 */
async function buildResumeFromResumeStructureJson(
  jsonInput,
  appRoot,
  profile = "primary",
  templatePathOverride = null,
) {
  /** @type {object | null} */
  let obj = null;
  if (jsonInput && typeof jsonInput === "object" && !Array.isArray(jsonInput)) {
    obj = /** @type {object} */ (jsonInput);
  } else {
    const s = String(jsonInput ?? "");
    obj =
      extractResumeJsonObjectFromGptText(s) || parseResumeStructureJson(s);
  }
  if (!obj || !looksLikeResumeStructure(obj)) {
    throw new Error(
      "Not valid resume_structure JSON. Expected career_summary, skills, work_experience (or work_xperience), or role_title.",
    );
  }
  const parsed = resumeJsonObjectToFields(obj);
  let buf = null;
  try {
    buf = await buildResumeFromStructureDocxParsed(
      parsed,
      appRoot,
      profile,
      templatePathOverride,
    );
  } catch (e) {
    console.error(
      "[resume structure.docx] template render failed:",
      formatDocxtemplaterError(e),
    );
  }
  if (!buf) {
    const { buildDefaultResumeDocxBufferParsed } = require("./resume-docx.js");
    buf = await buildDefaultResumeDocxBufferParsed(parsed);
  }
  return buf;
}

/**
 * @param {object} parsed same shape as parsePlainResumeToFields
 * @param {string} appRoot
 * @param {"primary" | "secondary"} [profile]
 * @returns {Promise<Buffer | null>}
 */
async function buildResumeFromStructureDocxParsed(
  parsed,
  appRoot,
  profile = "primary",
  templatePathOverride = null,
) {
  const override =
    typeof templatePathOverride === "string" && templatePathOverride.trim()
      ? path.normalize(templatePathOverride.trim())
      : "";
  const templatePath =
    override && fs.existsSync(override)
      ? override
      : findResumeStructureTemplatePath(appRoot, profile);
  if (!templatePath) {
    return null;
  }

  try {
    const content = fs.readFileSync(templatePath);
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, {
      paragraphLoop: true,
      linebreaks: true,
      delimiters: { start: "{", end: "}" },
      nullGetter() {
        return "";
      },
    });

    doc.render(buildTemplateData(parsed));

    return Buffer.from(
      doc.getZip().generate({
        type: "uint8array",
        compression: "DEFLATE",
      }),
    );
  } catch (e) {
    console.error(
      `[docxtemplater] ${path.basename(templatePath)}:`,
      formatDocxtemplaterError(e),
    );
    return null;
  }
}

/**
 * @param {string} plainText resume plain text (same extraction as save-gpt-docx)
 * @param {string} appRoot __dirname of main process
 * @returns {Promise<Buffer | null>} null if template file missing
 */
async function buildResumeFromStructureDocx(
  plainText,
  appRoot,
  profile = "primary",
  templatePathOverride = null,
) {
  const parsed = parsePlainResumeToFields(plainText);
  return buildResumeFromStructureDocxParsed(
    parsed,
    appRoot,
    profile,
    templatePathOverride,
  );
}

module.exports = {
  findDefaultStructureTemplatePath,
  findResumeStructureTemplatePath,
  formatDocxtemplaterError,
  buildResumeFromStructureDocx,
  buildResumeFromStructureDocxParsed,
  buildResumeFromResumeStructureJson,
  buildTemplateData,
  RESUME_TEMPLATE_FILENAMES,
};
