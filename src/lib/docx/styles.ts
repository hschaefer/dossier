import { BorderStyle, LineRuleType, Paragraph, ShadingType, type IStylesOptions } from "docx";

import { unitsToHalfPoints, unitsToPoints } from "./geometry";
import { bidi, hex, lineBox, mixHex, runsFor, TEXT, twipsIn, type DocContext, type RunStyle } from "./runs";

// The document's defaults and the five section-heading treatments. Word's own
// Normal style adds space after every paragraph and a theme font; both are
// cleared here so only the theme's measurements decide the layout.

/** Default body text: the body face, size, line height, ink and language. */
export function documentStyles(ctx: DocContext): IStylesOptions {
  const family = ctx.fonts.face("body", 400).family;
  return {
    default: {
      document: {
        run: {
          font: family,
          size: unitsToHalfPoints(ctx.resolved.density.fontSize, ctx.paper),
          color: hex(TEXT),
          language: ctx.language,
        },
        paragraph: {
          spacing: { before: 0, after: 0, ...lineBox(ctx, ctx.resolved.density.fontSize) },
        },
      },
    },
  };
}

/** The character spacing a heading's tracking asks for, in twentieths of a point. */
export function headingTracking(ctx: DocContext): number {
  const { tracking, size } = ctx.resolved.heading;
  // Joined scripts must not be pulled apart (document.module.css does the same).
  if (ctx.dir === "rtl" || ctx.lang?.split("-")[0] === "hi") return 0;
  const em = /^(-?(?:\d+|\d*\.\d+))em$/.exec(tracking);
  if (!em) return 0;
  return Math.round(Number(em[1]) * unitsToPoints(size, ctx.paper) * 20);
}

/** A section heading's run formatting: head face, accent, caps and tracking. */
export function headingTextStyle(ctx: DocContext): RunStyle {
  const { heading } = ctx.resolved;
  return {
    role: "head",
    weight: heading.weight,
    optical: "display",
    units: heading.size,
    color: ctx.resolved.accent,
    allCaps: heading.uppercase,
    characterSpacing: headingTracking(ctx),
  };
}

/**
 * A section heading paragraph. The rule is a bottom border, the bar a left
 * border and the band paragraph shading. Word draws those around the
 * paragraph's spacing too, so the heading carries no spacing of its own: the
 * gap above it is its own empty paragraph and the gap below it belongs to the
 * section's first item (see `sectionTitle`). Its bar is a 3-unit border with
 * 7 units of space, drawn at the paragraph's indent, so the indent is both.
 */
export function sectionHeading(ctx: DocContext, title: string): Paragraph {
  const { heading } = ctx.resolved;
  const accent = hex(ctx.resolved.accent);
  const rule = heading.rule
    ? { bottom: { style: BorderStyle.SINGLE, color: accent, size: 8, space: unitsToPoints(heading.padding, ctx.paper) } }
    : undefined;
  const barBorder = {
    style: BorderStyle.SINGLE,
    color: accent,
    size: Math.max(8, Math.round(unitsToPoints(3, ctx.paper) * 8)),
    space: unitsToPoints(7, ctx.paper),
  } as const;
  // The bar is on the inline start: the left in LTR, the right in RTL. The
  // border sits 7 units before the text, so the 3-unit bar itself lands on the
  // text edge when the paragraph is indented by both.
  const bar = heading.marker === "bar" ? (ctx.dir === "rtl" ? { right: barBorder } : { left: barBorder }) : undefined;
  const band = heading.marker === "band";

  return new Paragraph({
    keepNext: true,
    bidirectional: bidi(ctx),
    border: rule || bar ? { ...rule, ...bar } : undefined,
    shading: band ? { type: ShadingType.CLEAR, fill: hex(mixHex(ctx.resolved.accent, 0.1)) } : undefined,
    indent: band
      ? { start: twipsIn(ctx.paper, 7), end: twipsIn(ctx.paper, 7) }
      : heading.marker === "bar"
        ? ctx.dir === "rtl"
          ? { end: twipsIn(ctx.paper, 10) }
          : { start: twipsIn(ctx.paper, 10) }
        : undefined,
    spacing: { before: 0, after: 0, ...lineBox(ctx, heading.size) },
    children: runsFor(ctx, title, headingTextStyle(ctx)),
  });
}

/**
 * The gap above a section heading, as its own empty paragraph of that exact
 * height. Word includes a paragraph's space before and after in the border and
 * shading it draws, so a bar or band with the gap would stretch over it.
 *
 * The gap deliberately has no `keepNext`: chaining it to the heading also
 * chains the previous section's last entry to this section's first entry, so
 * Word moves the whole run to the next page and leaves a hole on the previous
 * one. The heading itself keeps with its first item.
 */
export function sectionGapParagraph(ctx: DocContext, before: number): Paragraph {
  return new Paragraph({
    bidirectional: bidi(ctx),
    spacing: { before: 0, after: 0, line: before, lineRule: LineRuleType.EXACT },
  });
}

/** A section heading, with the gap above it when there is one. */
export function sectionTitle(ctx: DocContext, title: string, before: number): Paragraph[] {
  return [...(before > 0 ? [sectionGapParagraph(ctx, before)] : []), sectionHeading(ctx, title)];
}
