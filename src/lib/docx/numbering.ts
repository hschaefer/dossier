import { AlignmentType, LevelFormat, type INumberingOptions } from "docx";

import { unitsToTwips } from "./geometry";
import { bodyHalfPoints, hex, TEXT, type DocContext } from "./runs";

// The resume's bullet list. docx's built-in bullet is a filled circle (●) at the
// document's size, which reads as a large disc, and its default indent pushes
// the text half an inch in. A resume's bullets follow the print stylesheet
// instead: a small disc and a 1.1em inset (document.module.css `.bullets`).

/** The numbering reference a bullet paragraph names. */
export const BULLET_REFERENCE = "resume-bullets";

/** The bullet list's text inset from the entry's inline start, in ems. */
export const BULLET_INSET_EM = 1.1;

/** How far the bullet itself sits before the text, in ems. */
export const BULLET_HANGING_EM = 0.6;

/** The list's inset and hanging widths in twips for the theme's body size. */
export function bulletIndent(ctx: DocContext): { inset: number; hanging: number } {
  const em = ctx.resolved.density.fontSize;
  return {
    inset: unitsToTwips(BULLET_INSET_EM * em, ctx.paper),
    hanging: unitsToTwips(BULLET_HANGING_EM * em, ctx.paper),
  };
}

/** The resume's bullet numbering: one level, a small disc in the body face. */
export function bulletNumbering(ctx: DocContext): INumberingOptions {
  return {
    config: [
      {
        reference: BULLET_REFERENCE,
        levels: [
          {
            level: 0,
            format: LevelFormat.BULLET,
            text: "•",
            alignment: AlignmentType.LEFT,
            style: {
              run: {
                font: ctx.fonts.face("body", 400).family,
                size: bodyHalfPoints(ctx),
                color: hex(TEXT),
              },
            },
          },
        ],
      },
    ],
  };
}
