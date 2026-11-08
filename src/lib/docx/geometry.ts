import type { PaperSize } from "@/generated/prisma/enums";
import { PAGE_WIDTH, PAPER_SIZES } from "@/lib/themes";

// DOCX lengths from the page units documents use. A page is 640 units wide
// whatever its paper (src/lib/themes.ts); Word works in twips (1/1440 inch)
// and half-points. Keep values unrounded until a DOCX API needs a number.

const MM_PER_INCH = 25.4;
const POINTS_PER_INCH = 72;
const TWIPS_PER_INCH = 1440;
const PIXELS_PER_INCH = 96;

/** How many millimetres one page unit is on `paper`. */
export function unitsToMm(units: number, paper: PaperSize): number {
  return (units * PAPER_SIZES[paper].widthMm) / PAGE_WIDTH;
}

/** A length in page units as points. */
export function unitsToPoints(units: number, paper: PaperSize): number {
  return (unitsToMm(units, paper) / MM_PER_INCH) * POINTS_PER_INCH;
}

/** A length in page units as twips, as Word stores indents, margins and spacing. */
export function unitsToTwips(units: number, paper: PaperSize): number {
  return Math.round((unitsToMm(units, paper) / MM_PER_INCH) * TWIPS_PER_INCH);
}

/** A length in page units as pixels at 96 dpi, as docx sizes images. */
export function unitsToPixels(units: number, paper: PaperSize): number {
  return Math.round((unitsToMm(units, paper) / MM_PER_INCH) * PIXELS_PER_INCH);
}

export function mmToTwips(mm: number): number {
  return Math.round((mm / MM_PER_INCH) * TWIPS_PER_INCH);
}

/** A font size in page units as DOCX half-points. */
export function unitsToHalfPoints(units: number, paper: PaperSize): number {
  return Math.round(unitsToPoints(units, paper) * 2);
}

/**
 * One CSS line box in twips, from its line-height and font size. Word's own
 * "multiple" spacing is a multiple of the font's height, not of the size
 * (`1.5 lines` of Geist is 1.95 em, a third more than the CSS 1.5 em), so the
 * renderer writes the line height itself and only lets Word grow it when the
 * font needs more room (see `lineBox`).
 */
export function lineBoxTwips(lineHeight: number, points: number): number {
  return Math.round(lineHeight * points * 20);
}

/** The width Word leaves for text on a page with the theme's margins. */
export function contentWidthTwips(paper: PaperSize, paddingX: number): number {
  return mmToTwips(PAPER_SIZES[paper].widthMm) - 2 * unitsToTwips(paddingX, paper);
}
