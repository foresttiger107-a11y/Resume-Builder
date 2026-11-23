/**
 * Parses plain-text / markdown-ish GPT resume output and builds a .docx matching
 * the default style: name (caps), role, contact row, rule, sections,
 * experience rows with right-aligned dates and bullets.
 */

const {
  AlignmentType,
  BorderStyle,
  convertInchesToTwip,
  Document,
  LevelFormat,
  Packer,
  Paragraph,
  Tab,
  TabStopPosition,
  TabStopType,
  TextRun,
} = require("docx");
const {
  formatWorkExperienceLocationLine,
} = require("./resume-json-to-fields.js");

const BULLET_REF = "resume-bullets";

function stripMd(s) {
  return String(s || "")
    .replace(/^#{1,6}\s+/, "")
    .replace(/\*+/g, "")
    .trim();
}

/** Email, phone, URL, LinkedIn — whole line or pipe-separated contact row. */
function isContactLine(s) {
  const L = String(s || "").trim();
  if (!L) {
    return false;
  }
  return (
    /@/.test(L) ||
    /linkedin\.com/i.test(L) ||
    /\+?\d[\d\s().\-–—]{8,}\d/.test(L) ||
    /\bhttps?:\/\//i.test(L)
  );
}

/**
 * First matching major CV section on this line (after stripMd), or null.
 * @param {string} rawLine
 * @returns {"summary"|"skills"|"experience"|"education"|null}
 */
function sectionHeaderKey(rawLine) {
  const t = stripMd(String(rawLine || ""));
  if (!t) {
    return null;
  }
  if (
    /^(career\s+summary|professional\s+summary|executive\s+summary|summary|profile|objective|professional\s+profile|overview|highlights?)\s*:?$/i.test(
      t,
    )
  ) {
    return "summary";
  }
  if (
    /^(technical\s+skills|core\s+skills|key\s+skills|skills|core\s+competencies|competencies|technical\s+expertise|relevant\s+skills|expertise|technical\s+competencies|tools\s+and\s+technologies)\s*:?$/i.test(
      t,
    )
  ) {
    return "skills";
  }
  if (
    /^(professional\s+experience|work\s+experience|relevant\s+experience|employment(\s+history)?|work\s+history|selected\s+experience|career\s+history|career\s+experience|job\s+history|professional\s+background|key\s+experience|experience)\s*:?$/i.test(
      t,
    )
  ) {
    return "experience";
  }
  if (/^education(\s+and\s+certifications)?\s*:?$/i.test(t)) {
    return "education";
  }
  return null;
}

/**
 * @param {string[]} lines
 * @returns {{ key: string; idx: number }[]}
 */
function collectSectionMarkers(lines) {
  /** @type {{ key: string; idx: number }[]} */
  const markers = [];
  const seen = new Set();
  for (let i = 0; i < lines.length; i++) {
    const k = sectionHeaderKey(lines[i]);
    if (k && !seen.has(k)) {
      markers.push({ key: k, idx: i });
      seen.add(k);
    }
  }
  markers.sort((a, b) => a.idx - b.idx);
  return markers;
}

/**
 * @param {{ key: string; idx: number }[]} markers
 * @param {string} key
 * @param {number} linesLen
 */
function sectionEndExclusive(markers, key, linesLen) {
  const pos = markers.findIndex((m) => m.key === key);
  if (pos < 0) {
    return -1;
  }
  const next = markers[pos + 1];
  return next ? next.idx : linesLen;
}

function stripBulletPrefix(s) {
  return String(s || "")
    .replace(/^\s*[-*+•·▪▸○◦](?:[\s\u00a0]+|(?=\S))/, "")
    .replace(/^\s*(?:[1-9]\d?\.|\d+\))(?:[\s\u00a0]+|(?=\S))/, "")
    .replace(/^\s*[•·▪][\s\u00a0]*/, "")
    .replace(/^\s*\*\s+/, "")
    .replace(/^\s*\*(?=\S)/, "")
    .replace(/^[\u2022\u2023\u2043\u25E6\u25AA\u25AB]\s*/, "")
    .trim();
}

/** GPT bullets: `- `, `• `, `-Word`, `1. `, `1) `, unicode bullets. */
function isExperienceBulletLine(L) {
  const t = String(L).trim();
  if (!t) {
    return false;
  }
  if (/^\s*[-*+•·▪▸○◦](?:[\s\u00a0]+|(?=\S))/.test(t)) {
    return true;
  }
  if (/^\s*(?:[1-9]\d?\.|\d+\))(?:[\s\u00a0]+|(?=\S))/.test(t)) {
    return true;
  }
  if (/^\s*[•·▪]\s?\S/.test(t)) {
    return true;
  }
  if (/^\s*\*(?=\S)/.test(t)) {
    return true;
  }
  if (/^[\u2022\u2023\u2043\u25E6\u25AA\u25AB\u2610\u263B]\s*\S/.test(t)) {
    return true;
  }
  return false;
}

/** @param {string[]} lines */
function joinNonEmpty(lines) {
  return lines
    .map((l) => l.trim())
    .filter(Boolean)
    .join("\n")
    .trim();
}

/** Join lines keeping blank lines; only trims leading/trailing of whole string. */
function joinLinesPreservingBlanks(lines) {
  return lines.map((l) => String(l).replace(/\r$/, "")).join("\n").replace(/^\s+|\s+$/g, "");
}

/**
 * @param {string} text
 * @param {{ bold?: boolean; size?: number }} [runOpts]
 * @param {{ spacingAfter?: number }} [paraOpts]
 */
function paragraphFromMultilineText(text, runOpts = {}, paraOpts = {}) {
  const raw = String(text ?? "");
  const parts = raw.split(/\n/);
  const { bold = false, size = 22 } = runOpts;
  const { spacingAfter = 120 } = paraOpts;
  /** @type {import("docx").TextRun[]} */
  const runs = [];
  for (let i = 0; i < parts.length; i++) {
    runs.push(new TextRun({ text: parts[i], font: "Calibri", size, bold: !!bold }));
    if (i < parts.length - 1) {
      runs.push(
        new TextRun({
          break: 1,
          font: "Calibri",
          size,
          bold: false,
        }),
      );
    }
  }
  if (runs.length === 0) {
    runs.push(new TextRun({ text: "", font: "Calibri", size, bold: !!bold }));
  }
  return new Paragraph({
    spacing: { after: spacingAfter },
    style: bold ? undefined : "Normal",
    children: runs,
  });
}

/**
 * Split skills body into groups (new group when a **…** header line appears after content).
 * @param {string[]} rawLines
 * @returns {string[][]}
 */
function splitSkillLinesIntoGroups(rawLines) {
  const lines = rawLines.map((r) => String(r).replace(/\r$/, ""));
  const groups = [];
  let cur = [];
  for (const line of lines) {
    const t = line.trim();
    const isMdHeaderLine = /^\*\*.+\*\*/.test(t);
    if (isMdHeaderLine && cur.length) {
      groups.push(cur);
      cur = [];
    }
    if (!t && cur.length === 0) {
      continue;
    }
    cur.push(line);
  }
  if (cur.length) {
    groups.push(cur);
  }
  return groups;
}

/**
 * @param {string[]} group
 * @returns {{ header: string; detail: string } | null}
 */
function skillGroupToBlock(group) {
  if (!group.length) {
    return null;
  }
  const firstTrim = group[0].trim();
  const md = firstTrim.match(/^\*\*(.+?)\*\*(.*)$/);
  if (md) {
    const header = stripInlineMdBold(md[1].trim());
    const tail = md[2].replace(/\r$/, "");
    const rest = group.slice(1).map((l) => String(l).replace(/\r$/, ""));
    const pieces = [];
    if (tail.trim()) {
      pieces.push(tail);
    }
    pieces.push(...rest);
    const detail = stripInlineMdBold(pieces.join("\n").replace(/^\s+|\s+$/g, ""));
    return { header, detail };
  }
  const first = firstTrim;
  const second = (group[1] || "").trim();
  const fc = (first.match(/,/g) || []).length;
  const sc = (second.match(/,/g) || []).length;
  if (group.length >= 2 && fc <= 2 && first.length <= 140 && (sc >= 3 || second.length > 90)) {
    const detail = stripInlineMdBold(
      group
        .slice(1)
        .map((l) => String(l).replace(/\r$/, ""))
        .join("\n")
        .replace(/^\s+|\s+$/g, ""),
    );
    return { header: stripInlineMdBold(first), detail };
  }
  const detail = stripInlineMdBold(
    group
      .map((l) => String(l).replace(/\r$/, ""))
      .join("\n")
      .replace(/^\s+|\s+$/g, ""),
  );
  return { header: "", detail };
}

function parseSkillsIntoBlocks(rawLines) {
  const groups = splitSkillLinesIntoGroups(rawLines);
  const blocks = [];
  for (const g of groups) {
    const b = skillGroupToBlock(g);
    if (b && (b.header || b.detail)) {
      blocks.push(b);
    }
  }
  return blocks;
}

/**
 * Lines from start+1 until next section header (or end).
 * @param {string[]} lines
 * @param {number} startIdx header row index
 * @param {number} endExclusive
 */
function sliceSectionBody(lines, startIdx, endExclusive) {
  const body = [];
  for (let i = startIdx + 1; i < endExclusive && i < lines.length; i++) {
    const raw = lines[i];
    if (sectionHeaderKey(raw)) {
      break;
    }
    body.push(raw);
  }
  return body;
}

const DATE_TAIL =
  /\s+((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{4}\s*[–—\-]+\s*(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{4}|(?:\d{1,2}\/\d{1,2}\/\d{2,4})\s*[–—\-]+\s*(?:\d{1,2}\/\d{1,2}\/\d{2,4}))\s*$/i;

/** Date range as the entire line (e.g. "Feb 2023 – Apr 2026"). */
const DATE_LINE_ONLY =
  /^((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{4}\s*[–—\-]+\s*(?:(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{4}|Present|Current|Ongoing))\s*$/i;

/** Index of first non-empty line at or after `start`, or lines.length if none. */
function nextNonEmptyLineIndex(lines, start) {
  let j = Math.max(0, start);
  while (j < lines.length && !String(lines[j] || "").trim()) {
    j += 1;
  }
  return j;
}

/** Short meta line (company / location), not a sentence or date line. */
function looksLikeCompanyOrLocationLine(s) {
  const t = String(s || "").trim();
  if (!t || t.length > 100) {
    return false;
  }
  if (DATE_LINE_ONLY.test(t) || DATE_TAIL.test(t)) {
    return false;
  }
  if (/\.\s*$/.test(t) && t.length > 42) {
    return false;
  }
  return true;
}

/** Long sentence-like line (usually a wrapped bullet), not a job title row. */
function looksLikeBulletContinuationSentence(s) {
  const t = String(s || "").trim();
  if (t.length < 48) {
    return false;
  }
  return /\.\s*$/.test(t);
}

/**
 * True when a new experience job appears to start at flat[i] (non-bullet line),
 * so wrapped bullet lines are not mistaken for job headers.
 */
function startsStructuredNewJob(flat, i) {
  const L0 = String(flat[i] ?? "").trim();
  if (!L0 || isExperienceBulletLine(flat[i])) {
    return false;
  }
  if (looksLikeBulletContinuationSentence(L0)) {
    return false;
  }
  const dm0 = L0.match(DATE_TAIL);
  if (dm0 && dm0.index != null && dm0.index > 0) {
    const before = L0.slice(0, dm0.index).trim();
    if (before.length > 0 && before.length <= 200) {
      return true;
    }
  }
  const j1 = nextNonEmptyLineIndex(flat, i + 1);
  if (j1 >= flat.length || isExperienceBulletLine(flat[j1])) {
    return false;
  }
  const L1 = flat[j1].trim();
  if (DATE_LINE_ONLY.test(L1)) {
    return true;
  }
  if (DATE_TAIL.test(L1)) {
    return true;
  }
  const j2 = nextNonEmptyLineIndex(flat, j1 + 1);
  if (j2 >= flat.length || isExperienceBulletLine(flat[j2])) {
    return false;
  }
  const L2 = flat[j2].trim();
  if (DATE_LINE_ONLY.test(L2)) {
    return true;
  }
  const j3 = nextNonEmptyLineIndex(flat, j2 + 1);
  if (j3 >= flat.length || isExperienceBulletLine(flat[j3])) {
    return false;
  }
  const L3 = flat[j3].trim();
  if (
    DATE_LINE_ONLY.test(L3) &&
    looksLikeCompanyOrLocationLine(L1) &&
    looksLikeCompanyOrLocationLine(L2)
  ) {
    return true;
  }
  return false;
}

/**
 * Append a non-marker experience line: merge wraps, or start a new plain bullet after a finished sentence.
 * @param {string[]} bullets
 * @param {string} t trimmed line
 */
function appendExperienceBulletBodyLine(bullets, t) {
  const text = String(t || "").trim();
  if (!text) {
    return;
  }
  if (!bullets.length) {
    bullets.push(text);
    return;
  }
  const prev = bullets[bullets.length - 1];
  const startsNewPlain =
    /\.\s*$/.test(prev) &&
    /^[A-Z(0-9]/.test(text) &&
    text.length >= 28;
  if (startsNewPlain) {
    bullets.push(text);
  } else {
    bullets[bullets.length - 1] = `${prev} ${text}`.replace(/\s{2,}/g, " ").trim();
  }
}

/**
 * @param {string[]} lines experience body only
 * @returns {{ titlePart: string; dateRight: string; bullets: string[] }[]}
 */
function parseExperienceJobs(lines) {
  const jobs = [];
  const flat = lines.map((l) => l.replace(/\r$/, ""));
  let i = 0;
  /** Bullets that appear before the first job title in this section. */
  let leadingBullets = [];
  while (i < flat.length) {
    const L0 = flat[i].trim();
    if (!L0) {
      i += 1;
      continue;
    }
    if (isExperienceBulletLine(flat[i])) {
      const b = stripBulletPrefix(flat[i]);
      if (jobs.length > 0) {
        jobs[jobs.length - 1].bullets.push(b);
      } else {
        leadingBullets.push(b);
      }
      i += 1;
      continue;
    }

    let titlePart = L0;
    let dateRight = "";
    let lastUsedIdx = i;

    const dm0 = L0.match(DATE_TAIL);
    if (dm0) {
      dateRight = dm0[1].trim();
      titlePart = L0.slice(0, dm0.index).trim();
    } else {
      const j1 = nextNonEmptyLineIndex(flat, i + 1);
      if (j1 < flat.length && !isExperienceBulletLine(flat[j1])) {
        const L1 = flat[j1].trim();
        const d1only = L1.match(DATE_LINE_ONLY);
        if (d1only) {
          dateRight = d1only[1].trim();
          titlePart = L0;
          lastUsedIdx = j1;
        } else {
          const d1tail = L1.match(DATE_TAIL);
          if (d1tail) {
            dateRight = d1tail[1].trim();
            titlePart = (L0 + "\n" + L1.slice(0, d1tail.index).trim()).trim();
            lastUsedIdx = j1;
          } else {
            const j2 = nextNonEmptyLineIndex(flat, j1 + 1);
            if (j2 < flat.length && !isExperienceBulletLine(flat[j2])) {
              const L2 = flat[j2].trim();
              const d2only = L2.match(DATE_LINE_ONLY);
              if (d2only) {
                dateRight = d2only[1].trim();
                titlePart = `${L0}\n${L1}`;
                lastUsedIdx = j2;
              } else {
                const j3 = nextNonEmptyLineIndex(flat, j2 + 1);
                if (
                  j3 < flat.length &&
                  !isExperienceBulletLine(flat[j3])
                ) {
                  const L3 = flat[j3].trim();
                  const d3only = L3.match(DATE_LINE_ONLY);
                  if (
                    d3only &&
                    looksLikeCompanyOrLocationLine(L1) &&
                    looksLikeCompanyOrLocationLine(L2) &&
                    L0.length >= 10 &&
                    !L0.match(DATE_LINE_ONLY)
                  ) {
                    dateRight = d3only[1].trim();
                    titlePart = `${L0} | ${L1} | ${L2}`;
                    lastUsedIdx = j3;
                  }
                }
              }
            }
          }
        }
      }
    }
    const consumed = lastUsedIdx - i + 1;

    const bullets = [];
    i += consumed;
    while (i < flat.length) {
      const L = flat[i];
      const t = L.trim();
      if (!t) {
        i += 1;
        continue;
      }
      if (isExperienceBulletLine(L)) {
        bullets.push(stripBulletPrefix(L));
        i += 1;
      } else if (!startsStructuredNewJob(flat, i)) {
        appendExperienceBulletBodyLine(bullets, t);
        i += 1;
      } else {
        break;
      }
    }
    if (leadingBullets.length) {
      bullets.unshift(...leadingBullets);
      leadingBullets = [];
    }
    if (titlePart || bullets.length) {
      jobs.push({ titlePart, dateRight, bullets });
    }
  }
  if (leadingBullets.length) {
    jobs.push({ titlePart: "", dateRight: "", bullets: [...leadingBullets] });
  }
  return jobs;
}

/**
 * @param {string} plain
 */
function parsePlainResumeToFields(plain) {
  const lines = String(plain || "")
    .replace(/\r\n/g, "\n")
    .split("\n");

  const markers = collectSectionMarkers(lines);
  const sectionStarts = markers.map((m) => m.idx);
  const headerEnd = sectionStarts.length
    ? Math.min(...sectionStarts)
    : Math.min(12, lines.length);

  const headerBlock =
    headerEnd > 0 ? lines.slice(0, headerEnd) : lines.slice(0, Math.min(12, lines.length));

  let fullName = "";
  let role = "";
  const contactBits = [];
  const headerSummaryLines = [];

  for (let h = 0; h < headerBlock.length; h++) {
    const raw = headerBlock[h];
    const L = stripMd(raw);
    if (!L) {
      continue;
    }

    if (isContactLine(L)) {
      if (/\|/.test(L) && /@/.test(L)) {
        const parts = L.split(/\s*\|\s*/).map((p) => p.trim());
        const contactParts = parts.filter((p) => isContactLine(p));
        const roleParts = parts.filter((p) => p && !isContactLine(p));
        if (roleParts.length) {
          const merged = roleParts.join(" | ").trim();
          if (merged.length < 220 && !role) {
            role = merged;
          }
        }
        if (contactParts.length) {
          contactBits.push(contactParts.join(" | "));
        } else {
          contactBits.push(L.replace(/\s*\|\s*/g, " | "));
        }
      } else {
        contactBits.push(L.replace(/\s*\|\s*/g, " | "));
      }
      continue;
    }

    if (!fullName && L.length > 2) {
      if (/\s-\s/.test(L) && L.length < 160) {
        const parts = L.split(/\s-\s/);
        fullName = parts[0].trim();
        role = parts.slice(1).join(" - ").trim();
      } else {
        fullName = L;
      }
      continue;
    }

    if (fullName && !role && !isContactLine(L)) {
      if (L.length <= 180) {
        role = L;
      } else {
        headerSummaryLines.push(String(raw).replace(/\r$/, ""));
      }
      continue;
    }

    if (fullName && role && !isContactLine(L)) {
      headerSummaryLines.push(String(raw).replace(/\r$/, ""));
    }
  }

  if (!fullName) {
    fullName = stripMd(lines[0] || "Resume") || "Resume";
  }

  if (!role) {
    for (let j = headerEnd; j < Math.min(headerEnd + 8, lines.length); j++) {
      const probe = stripMd(lines[j] || "");
      if (!probe) {
        continue;
      }
      if (sectionHeaderKey(lines[j])) {
        break;
      }
      if (isContactLine(probe)) {
        break;
      }
      if (probe.length <= 200) {
        role = probe;
        break;
      }
    }
  }

  const contactLine = contactBits.join(" | ");

  const idxSummary = markers.find((m) => m.key === "summary")?.idx ?? -1;
  const idxSkills = markers.find((m) => m.key === "skills")?.idx ?? -1;
  const idxExp = markers.find((m) => m.key === "experience")?.idx ?? -1;
  const idxEdu = markers.find((m) => m.key === "education")?.idx ?? -1;

  const endSum = idxSummary >= 0 ? sectionEndExclusive(markers, "summary", lines.length) : lines.length;
  const endSkills = idxSkills >= 0 ? sectionEndExclusive(markers, "skills", lines.length) : lines.length;
  const endExp = idxExp >= 0 ? sectionEndExclusive(markers, "experience", lines.length) : lines.length;

  let summary =
    idxSummary >= 0 ? joinNonEmpty(sliceSectionBody(lines, idxSummary, endSum)) : "";

  const headerSummary = joinLinesPreservingBlanks(headerSummaryLines);
  if (!summary.trim() && headerSummary.trim()) {
    summary = headerSummary;
  }

  const skillLines =
    idxSkills >= 0 ? sliceSectionBody(lines, idxSkills, endSkills > idxSkills ? endSkills : lines.length) : [];
  const skillsBlocks = parseSkillsIntoBlocks(skillLines);
  const skillsRaw = joinLinesPreservingBlanks(skillLines);

  const expLines =
    idxExp >= 0 ? sliceSectionBody(lines, idxExp, endExp > idxExp ? endExp : lines.length) : [];
  const jobs = parseExperienceJobs(expLines);

  const education =
    idxEdu >= 0 ? joinNonEmpty(sliceSectionBody(lines, idxEdu, lines.length)) : "";

  const displayName = fullName.toUpperCase();

  return {
    fullName: displayName,
    role: role || "",
    contactLine,
    summary: summary.trim() || "",
    skills: skillsRaw,
    skillsRaw,
    skillsBlocks,
    education,
    jobs,
    rawExperienceBody: joinLinesPreservingBlanks(expLines),
  };
}

function bulletNumbering() {
  return {
    config: [
      {
        reference: BULLET_REF,
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: {
              paragraph: {
                indent: {
                  left: convertInchesToTwip(0.35),
                  hanging: convertInchesToTwip(0.18),
                },
                spacing: { after: 80 },
              },
              run: { font: "Calibri", size: 22, bold: false },
            },
          },
        ],
      },
    ],
  };
}

function sectionHeaderParagraph(text) {
  return new Paragraph({
    spacing: { before: convertInchesToTwip(0.12), after: 140 },
    children: [
      new TextRun({
        text: String(text || "").toUpperCase(),
        font: "Calibri",
        size: 24,
        bold: true,
        color: "111827",
      }),
    ],
  });
}

function bodyParagraph(text, opts = {}) {
  const { spacingAfter = 120 } = opts;
  return new Paragraph({
    spacing: { after: spacingAfter },
    style: "Normal",
    children: [
      new TextRun({
        text: String(text || ""),
        font: "Calibri",
        size: 22,
        bold: false,
      }),
    ],
  });
}

function stripInlineMdBold(s) {
  let t = String(s || "");
  let prev;
  do {
    prev = t;
    t = t
      .replace(/\*\*([^*]+)\*\*/g, "$1")
      .replace(/\*([^*]+)\*/g, "$1")
      .replace(/__([^_]+)__/g, "$1");
  } while (t !== prev);
  return t.replace(/\*\*/g, "");
}

/** Thin rule under a section title (matches resume structure.docx). */
function sectionTitleRuleParagraph() {
  return new Paragraph({
    spacing: { after: 120 },
    border: {
      bottom: {
        color: "9CA3AF",
        space: 1,
        style: BorderStyle.SINGLE,
        size: 8,
      },
    },
    children: [new TextRun({ text: "\u00a0", size: 2 })],
  });
}

/** Rule after each job block (matches resume structure.docx). */
function jobBlockDividerParagraph() {
  return new Paragraph({
    spacing: { before: 48, after: 72 },
    border: {
      bottom: {
        color: "D1D5DB",
        space: 1,
        style: BorderStyle.SINGLE,
        size: 6,
      },
    },
    children: [new TextRun({ text: "\u00a0", font: "Arial", size: 4 })],
  });
}

/**
 * Job header: bold `{role_title} | {company}`, normal ` | {location} ({location_type})`, tab + `{duration}`.
 * @param {{ titlePart?: string; dateRight?: string; expCompany?: string; expLocation?: string; expLocationType?: string; expRoleTitle?: string }} job
 * @param {boolean} isFirstJob
 */
function jobExperienceHeaderParagraph(job, isFirstJob) {
  const date = stripInlineMdBold(stripMd(String(job.dateRight || "").trim()));
  const jr = /** @type {Record<string, unknown>} */ (job);
  let c = stripInlineMdBold(stripMd(String(jr.expCompany ?? "").trim()));
  let geo = stripInlineMdBold(stripMd(String(jr.expLocation ?? "").trim()));
  let lt = stripInlineMdBold(stripMd(String(jr.expLocationType ?? "").trim()));
  let rt = stripInlineMdBold(stripMd(String(jr.expRoleTitle ?? "").trim()));
  if (!c && !geo && !lt && !rt) {
    const parts = String(job.titlePart || "")
      .split(/\s*\|\s*/)
      .map((x) => x.trim())
      .filter(Boolean);
    rt = stripInlineMdBold(stripMd(parts[0] || ""));
    c = stripInlineMdBold(stripMd(parts[1] || ""));
    geo = stripInlineMdBold(stripMd(parts.slice(2).join(" | ") || ""));
  }
  const locationLine = formatWorkExperienceLocationLine(geo, lt);
  const A = { font: "Arial", size: 22 };
  const runs = [
    new TextRun({
      text: [rt, c].filter(Boolean).join(" | ") || "\u00a0",
      ...A,
      bold: true,
    }),
    new TextRun({ text: locationLine, ...A, bold: false }),
  ];
  if (date) {
    runs.push(new Tab(), new TextRun({ text: date, ...A, bold: false }));
  }
  return new Paragraph({
    spacing: { before: isFirstJob ? 80 : 120, after: date ? 100 : 80 },
    style: "Normal",
    tabStops: date ? [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }] : undefined,
    children: runs,
  });
}

/**
 * Skill detail: bold subheader lines only; comma lists normal. Handles one-line
 * "Languages & Software Engineering, Python, Bash, ..." and multiline GPT blocks.
 * @param {import("docx").Paragraph[]} children
 * @param {string} detailRaw
 * @param {{ spacingAfter: number; blockIndex: number; knownHeader?: string }} opts
 */
function appendFormattedSkillDetail(children, detailRaw, opts) {
  const { spacingAfter, blockIndex, knownHeader = "" } = opts;
  const cleaned = stripMd(stripInlineMdBold(String(detailRaw || ""))).trim();
  if (!cleaned) {
    return;
  }
  const known = stripMd(stripInlineMdBold(String(knownHeader || ""))).trim();
  const norm = (s) => s.toLowerCase().replace(/\s+/g, " ").trim();

  const tryCommaSplit = () => {
    if (cleaned.includes("\n")) {
      return false;
    }
    const i = cleaned.indexOf(",");
    if (i <= 0 || i > 120) {
      return false;
    }
    const head = cleaned.slice(0, i).trim();
    const tail = cleaned.slice(i + 1).trim();
    if (!tail || (head.match(/,/g) || []).length > 0) {
      return false;
    }
    if (head.length < 10 || head.length > 110) {
      return false;
    }
    if (tail.length < 12) {
      return false;
    }
    if (known && norm(head) === norm(known)) {
      return false;
    }
    children.push(
      new Paragraph({
        spacing: {
          before: blockIndex > 0 ? convertInchesToTwip(0.04) : 0,
          after: spacingAfter,
        },
        style: "Normal",
        children: [
          new TextRun({ text: head, font: "Calibri", size: 22, bold: true }),
          new TextRun({ text: ", ", font: "Calibri", size: 22, bold: false }),
          new TextRun({ text: tail, font: "Calibri", size: 22, bold: false }),
        ],
      }),
    );
    return true;
  };

  if (tryCommaSplit()) {
    return;
  }

  const lines = cleaned.split(/\n/).map((l) => l.trim()).filter(Boolean);
  const looksSub = (L, next) => {
    if (!next) {
      return false;
    }
    if (known && norm(L) === norm(known)) {
      return false;
    }
    const fc = (L.match(/,/g) || []).length;
    const nc = (next.match(/,/g) || []).length;
    if (fc > 2 || L.length > 130) {
      return false;
    }
    return nc >= 3 || next.length > 85;
  };

  let buf = [];
  const flush = (after) => {
    if (!buf.length) {
      return;
    }
    const body = buf.join("\n").trim();
    buf = [];
    if (!body) {
      return;
    }
    children.push(paragraphFromMultilineText(body, { bold: false, size: 22 }, { spacingAfter: after }));
  };

  const subBefore = blockIndex > 0 ? convertInchesToTwip(0.04) : 0;
  for (let i = 0; i < lines.length; i++) {
    const L = lines[i];
    if (known && norm(L) === norm(known)) {
      continue;
    }
    const next = lines[i + 1] || "";
    if (looksSub(L, next)) {
      flush(100);
      children.push(
        new Paragraph({
          spacing: { before: subBefore, after: 90 },
          style: "Normal",
          children: [new TextRun({ text: L, font: "Calibri", size: 22, bold: true })],
        }),
      );
    } else {
      buf.push(L);
    }
  }
  flush(spacingAfter);
}

/**
 * Build programmatic default-style .docx from an already-parsed resume object
 * (same shape as returned by parsePlainResumeToFields).
 * @param {object} p
 * @returns {Promise<Buffer>}
 */
async function buildDefaultResumeDocxBufferParsed(p) {
  const numbering = bulletNumbering();

  /** @type {import("docx").Paragraph[]} */
  const children = [];

  children.push(
    new Paragraph({
      spacing: { after: 80 },
      children: [
        new TextRun({
          text: p.fullName,
          font: "Calibri",
          size: 56,
          bold: true,
        }),
      ],
    }),
  );

  if (p.role) {
    children.push(
      new Paragraph({
        spacing: { after: 100 },
        children: [
          new TextRun({
            text: stripInlineMdBold(stripMd(p.role)),
            font: "Calibri",
            size: 24,
          }),
        ],
      }),
    );
  }

  if (p.contactLine) {
    children.push(
      new Paragraph({
        spacing: { after: 120 },
        children: [new TextRun({ text: p.contactLine, font: "Calibri", size: 20, color: "374151" })],
      }),
    );
  }

  children.push(
    new Paragraph({
      spacing: { after: 220 },
      border: {
        bottom: {
          color: "9CA3AF",
          space: 1,
          style: BorderStyle.SINGLE,
          size: 8,
        },
      },
      children: [new TextRun({ text: "\u00a0", size: 2 })],
    }),
  );

  children.push(sectionHeaderParagraph("Career summary"));
  children.push(bodyParagraph(stripInlineMdBold(stripMd(p.summary)), { spacingAfter: 200 }));

  children.push(sectionHeaderParagraph("Technical skills"));
  if (p.skillsBlocks.length) {
    p.skillsBlocks.forEach((b, bi) => {
      const isLast = bi === p.skillsBlocks.length - 1;
      const afterDetail = isLast ? 200 : 180;
      if (b.header) {
        children.push(
          new Paragraph({
            spacing: {
              before: bi > 0 ? convertInchesToTwip(0.1) : 0,
              after: 72,
            },
            children: [
              new TextRun({
                text: stripMd(stripInlineMdBold(b.header)),
                font: "Calibri",
                size: 22,
                bold: true,
              }),
            ],
          }),
        );
      }
      if (b.detail) {
        appendFormattedSkillDetail(children, b.detail, {
          spacingAfter: afterDetail,
          blockIndex: bi,
          knownHeader: b.header || "",
        });
      }
    });
  } else if (p.skillsRaw) {
    appendFormattedSkillDetail(children, p.skillsRaw, {
      spacingAfter: 200,
      blockIndex: 0,
      knownHeader: "",
    });
  } else {
    children.push(bodyParagraph("—", { spacingAfter: 200 }));
  }

  children.push(sectionHeaderParagraph("Professional experience"));
  children.push(sectionTitleRuleParagraph());
  if (p.jobs.length === 0) {
    children.push(
      bodyParagraph((p.rawExperienceBody || "").trim() || "—", { spacingAfter: 160 }),
    );
  } else {
    for (let ji = 0; ji < p.jobs.length; ji++) {
      const job = p.jobs[ji];
      children.push(
        jobExperienceHeaderParagraph(job, ji === 0),
      );
      for (const b of job.bullets) {
        children.push(
          new Paragraph({
            numbering: { reference: BULLET_REF, level: 0 },
            spacing: { after: 80 },
            children: [
              new TextRun({
                text: stripInlineMdBold(stripMd(String(b || ""))),
                font: "Arial",
                size: 22,
                bold: false,
              }),
            ],
          }),
        );
      }
      children.push(jobBlockDividerParagraph());
    }
  }

  if (p.education) {
    children.push(sectionHeaderParagraph("Education"));
    children.push(bodyParagraph(p.education, { spacingAfter: 120 }));
  }

  const doc = new Document({
    creator: "Resume Builder",
    title: "Resume",
    numbering,
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertInchesToTwip(0.7),
              right: convertInchesToTwip(0.75),
              bottom: convertInchesToTwip(0.7),
              left: convertInchesToTwip(0.75),
            },
          },
        },
        children,
      },
    ],
  });
  return Packer.toBuffer(doc);
}

/**
 * @param {string} plain extracted resume plain text
 * @returns {Promise<Buffer>}
 */
async function buildDefaultResumeDocxBuffer(plain) {
  const p = parsePlainResumeToFields(plain);
  return buildDefaultResumeDocxBufferParsed(p);
}

module.exports = {
  parsePlainResumeToFields,
  buildDefaultResumeDocxBuffer,
  buildDefaultResumeDocxBufferParsed,
  stripMd,
  stripInlineMdBold,
};
