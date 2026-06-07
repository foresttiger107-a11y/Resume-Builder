/**
 * One-time / dev script: writes assets/Justin Mcghee structure.docx — a Word template
 * with Docxtemplater placeholders matching the Justin McGhee layout (Calibri, margins,
 * section headers, `{#work_experience}` + `exp_algo` line, `{#skillBlocks}` (bold `{header}`, normal `{detail}`), etc.
 */
"use strict";

const fs = require("fs");
const path = require("path");
const {
  BorderStyle,
  convertInchesToTwip,
  Document,
  Packer,
  Paragraph,
  Tab,
  TabStopPosition,
  TabStopType,
  TextRun,
} = require("docx");

async function main() {
  const outDir = path.join(__dirname, "..", "assets");
  fs.mkdirSync(outDir, { recursive: true });
  const outPath = path.join(outDir, "Justin Mcghee structure.docx");

  const children = [
    new Paragraph({
      spacing: { after: 80 },
      children: [
        new TextRun({
          text: "{fullName}",
          font: "Calibri",
          size: 56,
          bold: true,
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 100 },
      children: [new TextRun({ text: "{role}", font: "Calibri", size: 24 })],
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [
        new TextRun({
          text: "{contactLine}",
          font: "Calibri",
          size: 20,
          color: "374151",
        }),
      ],
    }),
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
    new Paragraph({
      spacing: { before: convertInchesToTwip(0.12), after: 140 },
      children: [
        new TextRun({
          text: "CAREER SUMMARY",
          font: "Calibri",
          size: 24,
          bold: true,
          color: "111827",
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 200 },
      children: [new TextRun({ text: "{summary}", font: "Calibri", size: 22, bold: false })],
    }),
    new Paragraph({
      spacing: { before: convertInchesToTwip(0.12), after: 140 },
      children: [
        new TextRun({
          text: "TECHNICAL SKILLS",
          font: "Calibri",
          size: 24,
          bold: true,
          color: "111827",
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 40 },
      children: [
        new TextRun({ text: "{#skillBlocks}", font: "Calibri", size: 2, color: "FFFFFF" }),
      ],
    }),
    new Paragraph({
      spacing: { before: convertInchesToTwip(0.1), after: 72 },
      children: [
        new TextRun({
          text: "{header}",
          font: "Calibri",
          size: 22,
          bold: true,
          color: "111827",
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: convertInchesToTwip(0.14) },
      children: [
        new TextRun({
          text: "{detail}",
          font: "Calibri",
          size: 22,
          bold: false,
          color: "111827",
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 40 },
      children: [
        new TextRun({ text: "{/skillBlocks}", font: "Calibri", size: 2, color: "FFFFFF" }),
      ],
    }),
    new Paragraph({
      spacing: { after: 40 },
      children: [
        new TextRun({ text: "{^skillBlocks}", font: "Calibri", size: 2, color: "FFFFFF" }),
      ],
    }),
    new Paragraph({
      spacing: { after: 200 },
      children: [new TextRun({ text: "{skill}", font: "Calibri", size: 22, bold: false })],
    }),
    new Paragraph({
      spacing: { after: 40 },
      children: [
        new TextRun({ text: "{/skillBlocks}", font: "Calibri", size: 2, color: "FFFFFF" }),
      ],
    }),
    new Paragraph({
      spacing: { before: convertInchesToTwip(0.12), after: 140 },
      children: [
        new TextRun({
          text: "PROFESSIONAL EXPERIENCE",
          font: "Calibri",
          size: 24,
          bold: true,
          color: "111827",
        }),
      ],
    }),
    new Paragraph({
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
    }),
    new Paragraph({
      spacing: { after: 40 },
      children: [
        new TextRun({ text: "{#work_experience}", font: "Calibri", size: 2, color: "FFFFFF" }),
      ],
    }),
    new Paragraph({
      spacing: { before: 120, after: 100 },
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      children: [
        new TextRun({
          text: "{role_title} | {company}",
          font: "Arial",
          size: 22,
          bold: true,
        }),
        new TextRun({
          text: "{location_line}",
          font: "Arial",
          size: 22,
          bold: false,
        }),
        new Tab(),
        new TextRun({
          text: "{duration}",
          font: "Arial",
          size: 22,
          bold: false,
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 40 },
      children: [
        new TextRun({ text: "{#bullet_points}", font: "Calibri", size: 2, color: "FFFFFF" }),
      ],
    }),
    new Paragraph({
      spacing: { after: 80 },
      indent: { left: convertInchesToTwip(0.35), hanging: convertInchesToTwip(0.18) },
      children: [
        new TextRun({ text: "\u2022\t{bullet}", font: "Arial", size: 22, bold: false }),
      ],
    }),
    new Paragraph({
      spacing: { after: 40 },
      children: [
        new TextRun({ text: "{/bullet_points}", font: "Calibri", size: 2, color: "FFFFFF" }),
      ],
    }),
    new Paragraph({
      spacing: { before: 48, after: 72 },
      border: {
        bottom: {
          color: "D1D5DB",
          space: 1,
          style: BorderStyle.SINGLE,
          size: 6,
        },
      },
      children: [
        new TextRun({
          text: "\u00a0",
          font: "Arial",
          size: 4,
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 40 },
      children: [
        new TextRun({ text: "{/work_experience}", font: "Calibri", size: 2, color: "FFFFFF" }),
      ],
    }),
    new Paragraph({
      spacing: { before: convertInchesToTwip(0.12), after: 140 },
      children: [
        new TextRun({
          text: "EDUCATION",
          font: "Calibri",
          size: 24,
          bold: true,
          color: "111827",
        }),
      ],
    }),
    new Paragraph({
      spacing: { after: 120 },
      children: [new TextRun({ text: "{education}", font: "Calibri", size: 22, bold: false })],
    }),
  ];

  const doc = new Document({
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

  const buf = await Packer.toBuffer(doc);
  fs.writeFileSync(outPath, buf);
  console.log("Wrote", outPath);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
