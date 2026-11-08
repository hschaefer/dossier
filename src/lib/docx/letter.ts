import { Paragraph, TableRow, type FileChild } from "docx";

import { cell, fixedTable } from "./layout";
import { contactRuns, INK, lineBox, MUTED, runsFor, textWidth, twipsIn, type DocContext } from "./runs";

import type { PreparedLetter } from "@/lib/document-prep";

// A cover letter: always one column, whatever the resume theme's layout. The
// theme still supplies the faces, accent, banner and density; the fixed gaps
// are the CSS letter's own (sender 30, date 18, recipient 22, greeting 12,
// paragraph 14, sign-off 26, signature 4).

/** The letter's line height, from document.module.css (`.letter`). */
const LETTER_LINE_HEIGHT = 1.62;

/** A font size in page units at the letter's density. */
const body = (ctx: DocContext, factor = 1) => ctx.resolved.density.fontSize * factor;

/** A paragraph at the letter's 1.62 line height, for text `units` tall. */
function paragraph(
  ctx: DocContext,
  children: Paragraph["options"]["children"],
  before: number,
  after: number,
  keepNext = false,
  units = ctx.resolved.density.fontSize,
): Paragraph {
  return new Paragraph({
    bidirectional: ctx.dir === "rtl" || undefined,
    keepNext: keepNext || undefined,
    spacing: { before, after, ...lineBox(ctx, units, LETTER_LINE_HEIGHT) },
    children,
  });
}

/** The sender's name and contact lines, shaded as a banner when the theme asks for one. */
function senderBlocks(ctx: DocContext, prepared: PreparedLetter): FileChild[] {
  const banner = ctx.resolved.banner;
  const ink = banner ? ctx.resolved.onAccent : INK;
  const muted = banner ? ctx.resolved.onAccent : MUTED;
  // The banner's table is followed by the date's own 30-unit gap; without one,
  // the sender's last line carries the CSS margin-bottom instead.
  const bottom = banner ? 0 : twipsIn(ctx.paper, 30);
  const name = prepared.senderName
    ? [
        new Paragraph({
          bidirectional: ctx.dir === "rtl" || undefined,
          spacing: { before: 0, after: prepared.contactLines.length === 0 ? bottom : 0, ...lineBox(ctx, 20, 1.2) },
          children: runsFor(ctx, prepared.senderName, {
            role: "head",
            weight: ctx.resolved.name.weight,
            units: 20,
            color: ink,
          }),
        }),
      ]
    : [];
  const contacts = prepared.contactLines.map((items, index) =>
    paragraph(
      ctx,
      contactRuns(ctx, items, { units: body(ctx, 0.92), color: muted }),
      index === 0 && name.length > 0 ? twipsIn(ctx.paper, 3) : 0,
      index === prepared.contactLines.length - 1 ? bottom : 0,
      false,
      body(ctx, 0.92),
    ),
  );
  const lines = [...name, ...contacts];

  if (!banner) return lines;
  return [
    fixedTable(
      ctx,
      [
        new TableRow({
          children: [
            cell(textWidth(ctx), lines, {
              shading: ctx.resolved.accent,
              margins: {
                top: twipsIn(ctx.paper, 22),
                bottom: twipsIn(ctx.paper, 22),
                left: twipsIn(ctx.paper, 24),
                right: twipsIn(ctx.paper, 24),
              },
            }),
          ],
        }),
      ],
      [textWidth(ctx)],
    ),
  ];
}

/** The letter's children in reading order: sender, date, recipient, greeting, body, sign-off. */
export function letterChildren(ctx: DocContext, prepared: PreparedLetter): FileChild[] {
  const children: FileChild[] = senderBlocks(ctx, prepared);
  // A banner is a table, which carries no spacing of its own, so the sender's
  // bottom gap moves to whichever part starts the rest of the letter.
  let beforeNext = ctx.resolved.banner ? twipsIn(ctx.paper, 30) : 0;
  const takeBefore = () => {
    const before = beforeNext;
    beforeNext = 0;
    return before;
  };

  if (prepared.date) {
    children.push(paragraph(ctx, runsFor(ctx, prepared.date, { units: body(ctx) }), takeBefore(), twipsIn(ctx.paper, 18)));
  }
  prepared.recipient.forEach((line, index) => {
    children.push(
      paragraph(
        ctx,
        runsFor(ctx, line, { units: body(ctx) }),
        index === 0 ? takeBefore() : 0,
        index === prepared.recipient.length - 1 ? twipsIn(ctx.paper, 22) : 0,
      ),
    );
  });
  if (prepared.greeting) {
    children.push(paragraph(ctx, runsFor(ctx, prepared.greeting, { units: body(ctx) }), takeBefore(), twipsIn(ctx.paper, 12)));
  }
  prepared.paragraphs.forEach((unit, index) => {
    children.push(
      paragraph(
        ctx,
        runsFor(ctx, unit.text, { units: body(ctx) }),
        0,
        index === prepared.paragraphs.length - 1 ? 0 : twipsIn(ctx.paper, 14),
      ),
    );
  });
  if (prepared.closing) {
    children.push(paragraph(ctx, runsFor(ctx, prepared.closing, { units: body(ctx) }), twipsIn(ctx.paper, 26), twipsIn(ctx.paper, 4), true));
  }
  if (prepared.signature) {
    children.push(paragraph(ctx, runsFor(ctx, prepared.signature, { weight: 600, color: INK, units: body(ctx) }), 0, 0));
  }
  return children;
}
