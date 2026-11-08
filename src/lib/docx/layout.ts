import {
  BorderStyle,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  WidthType,
  type ITableCellOptions,
  type ISectionPropertiesOptions,
  type TableVerticalAlign,
} from "docx";

import { PAPER_SIZES } from "@/lib/themes";

import { mmToTwips, unitsToTwips } from "./geometry";
import { mixHex, sidebarGap, sidebarWidth, textWidth, type DocContext } from "./runs";

// Word's page and table geometry. Layout tables carry no borders, no default
// cell padding and fixed column widths, so a theme's sidebar is exactly as
// wide as its measurements say and nothing pushes content off the paper.

/** A4 or Letter, with the theme's page padding as Word page margins. */
export function pageProperties(ctx: DocContext): ISectionPropertiesOptions {
  const { widthMm, heightMm } = PAPER_SIZES[ctx.paper];
  const { paddingY, paddingX } = ctx.resolved.density;
  return {
    page: {
      size: { width: mmToTwips(widthMm), height: mmToTwips(heightMm) },
      margin: {
        top: unitsToTwips(paddingY, ctx.paper),
        bottom: unitsToTwips(paddingY, ctx.paper),
        left: unitsToTwips(paddingX, ctx.paper),
        right: unitsToTwips(paddingX, ctx.paper),
      },
    },
  };
}

/**
 * A layout table drawn as nothing. Word adds inside borders unless they are
 * named too, which showed as stray vertical lines between the reference,
 * gap and sidebar cells.
 */
export const noBorders = {
  top: { style: BorderStyle.NONE, size: 0, color: "auto" },
  bottom: { style: BorderStyle.NONE, size: 0, color: "auto" },
  left: { style: BorderStyle.NONE, size: 0, color: "auto" },
  right: { style: BorderStyle.NONE, size: 0, color: "auto" },
  insideHorizontal: { style: BorderStyle.NONE, size: 0, color: "auto" },
  insideVertical: { style: BorderStyle.NONE, size: 0, color: "auto" },
} as const;

/** Cell padding, in twips; docx doesn't export the option type itself. */
type CellMargins = NonNullable<ITableCellOptions["margins"]>;

const ZERO_MARGINS: CellMargins = { top: 0, bottom: 0, left: 0, right: 0 };

/** A borderless table with fixed columns: a layout container, not a data table. */
export function fixedTable(
  ctx: DocContext,
  rows: TableRow[],
  widths: number[],
  options: {
    /**
     * Reverses the columns without changing cell order. A sidebar at the end of
     * the line is written sidebar-first and mirrored, which Word and
     * LibreOffice keep on the first page where a main-first table does not.
     */
    rightToLeft?: boolean;
  } = {},
): Table {
  return new Table({
    rows,
    width: { size: widths.reduce((total, width) => total + width, 0), type: WidthType.DXA },
    columnWidths: widths,
    layout: TableLayoutType.FIXED,
    borders: noBorders,
    margins: ZERO_MARGINS,
    // Mirrors the columns for a right-to-left document but keeps cell order.
    visuallyRightToLeft: options.rightToLeft ?? ctx.dir === "rtl",
  });
}

export function cell(
  width: number,
  children: readonly (Paragraph | Table)[],
  options: { shading?: string; margins?: CellMargins; verticalAlign?: TableVerticalAlign } = {},
): TableCell {
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: options.shading ? { fill: options.shading.replace("#", ""), type: "clear" } : undefined,
    margins: options.margins ?? ZERO_MARGINS,
    verticalAlign: options.verticalAlign,
    children,
  });
}

/** One group of a column's blocks; columnTable flattens these into the cell. */
export type ColumnRow = readonly (Paragraph | Table)[];

/**
 * The sidebar and main column as a three-column table: content, a real gap
 * column, more content. Both columns live in one row, so their paragraphs flow
 * independently: a tall sidebar photo does not push the main column down, and
 * page breaks follow each column's own content. Word splits the row across
 * pages as long as nothing sets `cantSplit` or a fixed row height. The tinted
 * sidebar cell repeats its shading on every page fragment.
 */
export function columnTable(
  ctx: DocContext,
  input: {
    sidebar: readonly ColumnRow[];
    main: readonly ColumnRow[];
    sidebarEnd: boolean;
    tinted: boolean;
  },
): Table {
  const width = textWidth(ctx);
  const side = sidebarWidth(ctx);
  const gap = sidebarGap(ctx);
  const mainWidth = width - side - gap;
  const shaded = input.tinted ? mixHex(ctx.resolved.accent, 0.07) : undefined;
  // Pad the panel around its content; the cell now spans every page it reaches.
  const margins: CellMargins = input.tinted
    ? {
        top: unitsToTwips(14, ctx.paper),
        bottom: unitsToTwips(14, ctx.paper),
        left: unitsToTwips(12, ctx.paper),
        right: unitsToTwips(12, ctx.paper),
      }
    : ZERO_MARGINS;

  const sidebarCell = cell(side, input.sidebar.flat(), { shading: shaded, margins });
  const gapCell = cell(gap, [new Paragraph({})]);
  const mainCell = cell(mainWidth, input.main.flat());
  const row = new TableRow({ children: [sidebarCell, gapCell, mainCell] });

  // Sidebar first in cell order, mirrored when it belongs at the end of the
  // line; see fixedTable. This keeps a right sidebar on the first page.
  const mirrored = ctx.dir === "rtl" ? !input.sidebarEnd : input.sidebarEnd;
  return fixedTable(ctx, [row], [side, gap, mainWidth], { rightToLeft: mirrored });
}

/** One row of a side-headings layout: the heading on its section's first row, then content beside it. */
export type SideHeadingRow = {
  heading?: Paragraph;
  content: readonly (Paragraph | Table)[];
  /** The gap above this row, in twips; a table cell has no spacing of its own. */
  before: number;
};

/**
 * The side-headings layout as one table with a row per entry. The heading
 * column is empty on the rows after a section's first, so a long section can
 * break between entries instead of moving the whole section to another page.
 */
export function sideHeadingTable(ctx: DocContext, rows: readonly SideHeadingRow[]): Table {
  const width = textWidth(ctx);
  const headingWidth = unitsToTwips(112, ctx.paper);
  const gap = unitsToTwips(16, ctx.paper);
  return fixedTable(
    ctx,
    rows.map((row) => {
      const margins: CellMargins = { top: row.before, bottom: 0, left: 0, right: 0 };
      return new TableRow({
        children: [
          cell(headingWidth, row.heading ? [row.heading] : [new Paragraph({})], { margins: { ...margins, right: gap } }),
          cell(width - headingWidth, [...row.content], { margins }),
        ],
      });
    }),
    [headingWidth, width - headingWidth],
  );
}
