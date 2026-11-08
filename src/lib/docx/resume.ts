import {
  AlignmentType,
  ExternalHyperlink,
  LineRuleType,
  Paragraph,
  Tab,
  Table,
  TableRow,
  TextRun,
  VerticalAlign,
  type FileChild,
} from "docx";

import type { PreparedResume } from "@/lib/document-prep";
import { PhotoPosition } from "@/generated/prisma/enums";
import { citationSpans } from "@/lib/citation";
import type { ResumeContent } from "@/lib/resume-content";
import { customKey, inSidebar, sectionHeading as resumeSectionHeading, shownSections } from "@/lib/resume-sections";
import { skillLine } from "@/lib/skill-groups";
import { unitsToPoints, unitsToTwips } from "./geometry";

import { cell, columnTable, fixedTable, sideHeadingTable, type SideHeadingRow } from "./layout";
import { BULLET_REFERENCE, bulletIndent } from "./numbering";
import { photoImage, photoParagraph, photoWidth } from "./photo";
import {
  bidi,
  contactRuns,
  INK,
  lineBox,
  MUTED,
  runsFor,
  sidebarGap,
  sidebarWidth,
  textWidth,
  twipsIn,
  type DocContext,
  type RunStyle,
} from "./runs";
import { sectionHeading, sectionTitle } from "./styles";

// A resume laid out from the prepared content and the resolved theme: header,
// sections, a real sidebar table or side-heading tables, and timeline or
// inline dates. Nothing here knows a built-in theme by name.

type Block = Paragraph | Table;

/**
 * A section's content as items: one item is one entry, one skill line or one
 * contact. In a column layout the items are flattened into the column's cell,
 * so each column keeps its own flow across pages.
 */
type Section = { title: string; items: Block[][] };

/** A section's blocks in order: its heading, then every item. */
const flat = (section: Section): Block[] => section.items.flat();

/** The body size in page units, scaled by `factor`. */
const BODY = (ctx: DocContext, factor = 1) => ctx.resolved.density.fontSize * factor;

/** Text width inside the main column, accounting for a sidebar or side headings. */
function bodyWidth(ctx: DocContext): number {
  if (ctx.resolved.split) return textWidth(ctx) - sidebarWidth(ctx) - sidebarGap(ctx);
  if (ctx.resolved.sideHeadings) return textWidth(ctx) - unitsToTwips(112, ctx.paper);
  return textWidth(ctx);
}

const muted = (ctx: DocContext, factor = 0.9): RunStyle => ({ units: BODY(ctx, factor), color: MUTED });

/**
 * The gap between a section's children (`.section` in document.module.css), in
 * page units. The heading carries it as its space after; every item after the
 * first carries it too, on top of its own margin (`.entry + .entry` and the
 * like), which is how the print layout stacks a section.
 */
const sectionGap = (ctx: DocContext) => twipsIn(ctx.paper, 6);

/**
 * The gap above a section's first item, which the heading no longer carries.
 * The side-heading layout puts the title beside the content, so its first item
 * starts on the title's top rather than below it.
 */
const firstItemGap = (ctx: DocContext) => (ctx.resolved.sideHeadings ? 0 : sectionGap(ctx));

function headerText(ctx: DocContext, prepared: PreparedResume, color: { name: string; headline: string; contact: string }): Paragraph[] {
  const { content, resolved } = prepared;
  const alignment = resolved.centered ? AlignmentType.CENTER : undefined;
  const paragraphs: Paragraph[] = [
    new Paragraph({
      alignment,
      bidirectional: bidi(ctx),
      spacing: { before: 0, after: twipsIn(ctx.paper, 4), ...lineBox(ctx, resolved.name.size, 1.1) },
      children: runsFor(ctx, content.name, {
        role: "head",
        weight: resolved.name.weight,
        units: resolved.name.size,
        color: color.name,
        characterSpacing: Math.round(-0.01 * unitsToPoints(resolved.name.size, ctx.paper) * 20),
      }),
    }),
  ];
  if (content.headline) {
    paragraphs.push(
      new Paragraph({
        alignment,
        bidirectional: bidi(ctx),
        // The header's 4-unit gap plus the contact line's 4-unit margin.
        spacing: { before: 0, after: twipsIn(ctx.paper, 8), ...lineBox(ctx, BODY(ctx, 1.15)) },
        children: runsFor(ctx, content.headline, { units: BODY(ctx, 1.15), color: color.headline }),
      }),
    );
  }
  // A sidebar layout prints the contacts in its Contact section instead.
  if (content.contacts.length > 0 && !resolved.split) {
    paragraphs.push(
      new Paragraph({
        alignment,
        bidirectional: bidi(ctx),
        spacing: { before: 0, after: 0, ...lineBox(ctx, BODY(ctx, 0.9)) },
        children: contactRuns(ctx, content.contacts, { units: BODY(ctx, 0.9), color: color.contact }),
      }),
    );
  }
  return paragraphs;
}

/**
 * The header, with the photo above the name or in a two-column row beside it.
 * In a banner theme the whole header sits in a shaded one-cell table.
 */
async function headerBlocks(ctx: DocContext, prepared: PreparedResume): Promise<Block[]> {
  const { content, resolved } = prepared;
  const banner = resolved.banner;
  const colors = banner
    ? { name: resolved.onAccent, headline: resolved.onAccent, contact: resolved.onAccent }
    : { name: INK, headline: resolved.accent, contact: MUTED };
  const width = banner ? textWidth(ctx) - 2 * twipsIn(ctx.paper, 24) : textWidth(ctx);
  const image = await photoImage(ctx, content.photo, content.name, "header");
  const position = resolved.photo.position;
  const text = headerText(ctx, prepared, colors);

  let blocks: Block[];
  if (image && position === PhotoPosition.ABOVE_NAME) {
    blocks = [photoParagraph(ctx, image), ...text];
  } else if (image && (position === PhotoPosition.HEADER_START || position === PhotoPosition.HEADER_END)) {
    const gap = unitsToTwips(20, ctx.paper);
    const photo = photoWidth(ctx, "header");
    const photoCell = cell(photo, [new Paragraph({ spacing: { before: 0, after: 0 }, children: [image] })], {
      verticalAlign: VerticalAlign.CENTER,
    });
    const textCell = cell(width - photo - gap, text, { verticalAlign: VerticalAlign.CENTER });
    const gapCell = cell(gap, [new Paragraph({})]);
    const cells = position === PhotoPosition.HEADER_START ? [photoCell, gapCell, textCell] : [textCell, gapCell, photoCell];
    const widths = position === PhotoPosition.HEADER_START ? [photo, gap, width - photo - gap] : [width - photo - gap, gap, photo];
    blocks = [fixedTable(ctx, [new TableRow({ children: cells })], widths)];
  } else {
    blocks = text;
  }

  if (!banner) return blocks;
  return [
    fixedTable(
      ctx,
      [
        new TableRow({
          children: [
            cell(textWidth(ctx), blocks, {
              shading: resolved.accent,
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

type EntryHeadInput = { title: string; at?: string; dates: string };

/** The date column's width in a timeline. */
const dateColumn = (ctx: DocContext) => twipsIn(ctx.paper, ctx.resolved.split ? 78 : 108);

/** The left inset every part of a timeline entry starts at. */
const timelineIndent = (ctx: DocContext) =>
  ctx.resolved.timeline ? (ctx.dir === "rtl" ? { start: dateColumn(ctx) } : { left: dateColumn(ctx) }) : undefined;

/** A title/company line with its dates, or a timeline date column before it. */
function entryHead(ctx: DocContext, input: EntryHeadInput, width: number, before: number): Paragraph {
  const { title, at, dates } = input;
  const dateWidth = dateColumn(ctx);
  const titleRuns = runsFor(ctx, title, { weight: 600, color: INK });
  const atRuns = at ? [...runsFor(ctx, title ? ", " : "", { weight: 400 }), ...runsFor(ctx, at, { weight: 400 })] : [];
  const dateRuns = dates ? runsFor(ctx, dates, muted(ctx)) : [];
  const tab = new TextRun({ children: [new Tab()] });

  if (ctx.resolved.timeline) {
    const indent = ctx.dir === "rtl" ? { start: dateWidth, hanging: dateWidth } : { left: dateWidth, hanging: dateWidth };
    return new Paragraph({
      keepNext: true,
      bidirectional: bidi(ctx),
      indent,
      tabStops: [{ type: ctx.dir === "rtl" ? "start" : "left", position: dateWidth }],
      spacing: { before, after: 0 },
      children: [...dateRuns, tab, ...titleRuns, ...atRuns],
    });
  }
  return new Paragraph({
    keepNext: true,
    bidirectional: bidi(ctx),
    tabStops: [{ type: ctx.dir === "rtl" ? "end" : "right", position: width }],
    spacing: { before, after: 0 },
    children: [...titleRuns, ...atRuns, ...(dates ? [tab, ...dateRuns] : [])],
  });
}

function descriptor(ctx: DocContext, text: string): Paragraph {
  return new Paragraph({
    bidirectional: bidi(ctx),
    indent: timelineIndent(ctx),
    spacing: { before: twipsIn(ctx.paper, 2), after: 0, ...lineBox(ctx, BODY(ctx, 0.9)) },
    children: runsFor(ctx, text, { ...muted(ctx), italic: true }),
  });
}

function bullets(ctx: DocContext, items: readonly string[]): Paragraph[] {
  const { inset, hanging } = bulletIndent(ctx);
  const offset = ctx.resolved.timeline ? dateColumn(ctx) : 0;
  const indent =
    ctx.dir === "rtl" ? { start: offset + inset, hanging } : { left: offset + inset, hanging };
  return items.map(
    (item, index) =>
      new Paragraph({
        // A real Word list, so the bullets stay bullets when edited. `custom`
        // keeps docx from adding its built-in ListParagraph style, which is not
        // part of the theme.
        numbering: { reference: BULLET_REFERENCE, level: 0, custom: true },
        keepLines: true,
        bidirectional: bidi(ctx),
        indent,
        spacing: { before: index === 0 ? twipsIn(ctx.paper, 3) : 0, after: 0 },
        children: runsFor(ctx, item, { units: BODY(ctx) }),
      }),
  );
}

function entrySkills(ctx: DocContext, skills: readonly string[], label: string): Paragraph {
  return new Paragraph({
    bidirectional: bidi(ctx),
    indent: timelineIndent(ctx),
    spacing: { before: twipsIn(ctx.paper, 3), after: 0, ...lineBox(ctx, BODY(ctx, 0.9)) },
    children: [
      ...runsFor(ctx, `${label} · `, { ...muted(ctx), color: ctx.resolved.accent }),
      ...runsFor(ctx, skills.join(", "), muted(ctx)),
    ],
  });
}

function experienceSections(ctx: DocContext, prepared: PreparedResume, width: number, key: "experience" | "volunteering" = "experience"): Section[] {
  const { content, labels, resolved } = prepared;
  if (content[key].length === 0) return [];
  const itemGap = twipsIn(ctx.paper, resolved.density.itemGap) + sectionGap(ctx);
  const roles = content[key].map((role, index) => {
    const blocks: Block[] = [entryHead(ctx, { title: role.role, at: role.company, dates: role.dates }, width, index === 0 ? firstItemGap(ctx) : itemGap)];
    if (role.descriptor) blocks.push(descriptor(ctx, role.descriptor));
    blocks.push(...bullets(ctx, role.bullets));
    if (role.skills.length > 0) blocks.push(entrySkills(ctx, role.skills, labels.skills));
    return blocks;
  });
  return [{ title: labels[key], items: roles }];
}

/** A verified credential, keeping its details and web link editable. */
function certificationSections(ctx: DocContext, prepared: PreparedResume, width: number): Section[] {
  const { content, labels, resolved } = prepared;
  if (content.certifications.length === 0) return [];
  const itemGap = twipsIn(ctx.paper, resolved.density.itemGap) + sectionGap(ctx);
  return [{ title: labels.certifications, items: content.certifications.map((entry, index) => {
    const blocks: Block[] = [entryHead(ctx, { title: entry.name, at: entry.issuer, dates: entry.date }, width, index === 0 ? firstItemGap(ctx) : itemGap)];
    if (entry.details.length > 0 || entry.link) {
      blocks.push(new Paragraph({
        bidirectional: bidi(ctx),
        indent: timelineIndent(ctx),
        spacing: { before: 0, after: 0 },
        children: [
          ...runsFor(ctx, entry.details.join(" · "), muted(ctx, 1)),
          ...(entry.link ? [...runsFor(ctx, entry.details.length ? " · " : "", muted(ctx, 1)), ...documentLink(ctx, entry.link, entry.href)] : []),
        ],
      }));
    }
    return blocks;
  }) }];
}

function documentLink(ctx: DocContext, text: string, href: string | null): (TextRun | ExternalHyperlink)[] {
  const children = runsFor(ctx, text, muted(ctx));
  return href ? [new ExternalHyperlink({ link: href, children })] : children;
}

/** Custom sections use the same three formats as the preview. */
function customSection(ctx: DocContext, section: ResumeContent["customSections"][number], width: number): Section {
  if (section.format === "INLINE") {
    return { title: section.title, items: [[new Paragraph({
      bidirectional: bidi(ctx),
      spacing: { before: firstItemGap(ctx), after: 0 },
      children: runsFor(ctx, section.items.map((item) => item.title).join(" · ")),
    })]] };
  }
  return { title: section.title, items: section.items.map((item, index) => {
    const before = index === 0 ? firstItemGap(ctx) : sectionGap(ctx) + twipsIn(ctx.paper, section.format === "LINES" ? 5 : ctx.resolved.density.itemGap);
    if (section.format === "LINES") {
      return [new Paragraph({
        bidirectional: bidi(ctx),
        spacing: { before, after: 0 },
        children: [
          ...citationSpans(item.title).flatMap((span) => runsFor(ctx, span.text, { weight: span.bold ? 700 : 400, italic: span.italic ?? false })),
          ...(item.link ? [...runsFor(ctx, " "), ...documentLink(ctx, item.link, item.href)] : []),
        ],
      })];
    }
    const blocks: Block[] = [entryHead(ctx, { title: item.title, at: item.subtitle, dates: item.dates }, width, before), ...bullets(ctx, item.bullets)];
    if (item.link) blocks.push(new Paragraph({
      bidirectional: bidi(ctx),
      indent: timelineIndent(ctx),
      spacing: { before: 0, after: 0 },
      children: documentLink(ctx, item.link, item.href),
    }));
    return blocks;
  }) };
}

function projectSections(ctx: DocContext, prepared: PreparedResume, width: number): Section[] {
  const { content, labels, resolved } = prepared;
  if (content.projects.length === 0) return [];
  const itemGap = twipsIn(ctx.paper, resolved.density.itemGap) + sectionGap(ctx);
  const projects = content.projects.map((project, index) => {
    const blocks: Block[] = [entryHead(ctx, { title: project.name, at: project.link, dates: project.dates }, width, index === 0 ? firstItemGap(ctx) : itemGap)];
    blocks.push(...bullets(ctx, project.bullets));
    if (project.skills.length > 0) blocks.push(entrySkills(ctx, project.skills, labels.skills));
    return blocks;
  });
  return [{ title: labels.projects, items: projects }];
}

function educationSections(ctx: DocContext, prepared: PreparedResume, width: number): Section[] {
  const { content, labels, resolved } = prepared;
  if (content.education.length === 0) return [];
  const itemGap = twipsIn(ctx.paper, resolved.density.itemGap) + sectionGap(ctx);
  const entries = content.education.map((entry, index) => {
    const blocks: Block[] = [entryHead(ctx, { title: entry.degree, at: entry.school, dates: entry.dates }, width, index === 0 ? firstItemGap(ctx) : itemGap)];
    if (entry.notes) {
      blocks.push(
        new Paragraph({
          bidirectional: bidi(ctx),
          indent: timelineIndent(ctx),
          spacing: { before: 0, after: 0 },
          children: runsFor(ctx, entry.notes, { ...muted(ctx), units: BODY(ctx) }),
        }),
      );
    }
    return blocks;
  });
  return [{ title: labels.education, items: entries }];
}

function publicationSections(ctx: DocContext, prepared: PreparedResume): Section[] {
  const { content, labels } = prepared;
  if (content.publications.length === 0) return [];
  const paragraphs = content.publications.map((entry, index) => {
    const children: (TextRun | ExternalHyperlink)[] = citationSpans(entry.citation).flatMap((span) =>
      runsFor(ctx, span.text, {
        weight: span.bold ? 700 : 400,
        italic: span.italic ?? false,
      }),
    );
    if (entry.link) {
      children.push(...runsFor(ctx, " ", muted(ctx)));
      const linkStyle: RunStyle = { ...muted(ctx), italic: false };
      if (entry.href) children.push(new ExternalHyperlink({ link: entry.href, children: runsFor(ctx, entry.link, linkStyle) }));
      else children.push(...runsFor(ctx, entry.link, linkStyle));
    }
    return new Paragraph({
      bidirectional: bidi(ctx),
      spacing: { before: index === 0 ? firstItemGap(ctx) : twipsIn(ctx.paper, 5) + sectionGap(ctx), after: 0 },
      children,
    });
  });
  return [{ title: labels.publications, items: paragraphs.map((paragraph) => [paragraph]) }];
}

function referenceSections(ctx: DocContext, prepared: PreparedResume, width: number): Section[] {
  const { content, labels } = prepared;
  if (content.references.length === 0) return [];
  const gap = twipsIn(ctx.paper, 10);
  const column = Math.floor((width - gap) / 2);
  const rows: TableRow[] = [];
  for (let index = 0; index < content.references.length; index += 2) {
    const pair = content.references.slice(index, index + 2);
    const margin = index === 0 ? firstItemGap(ctx) : twipsIn(ctx.paper, 10);
    const referenceCell = (reference: (typeof content.references)[number] | undefined) => {
      const children: Paragraph[] = reference
        ? [
            new Paragraph({ bidirectional: bidi(ctx), spacing: { before: 0, after: 0 }, children: runsFor(ctx, reference.name, { weight: 600, color: INK }) }),
            ...(reference.role
              ? [
                  new Paragraph({
                    bidirectional: bidi(ctx),
                    spacing: { before: 0, after: 0, ...lineBox(ctx, BODY(ctx, 0.9)) },
                    children: runsFor(ctx, reference.role, muted(ctx)),
                  }),
                ]
              : []),
            ...(reference.contact
              ? [new Paragraph({ bidirectional: bidi(ctx), spacing: { before: 0, after: 0 }, children: runsFor(ctx, reference.contact, { units: BODY(ctx) }) })]
              : []),
          ]
        : [new Paragraph({})];
      return cell(column, children, { margins: { top: margin, bottom: 0, left: 0, right: 0 } });
    };
    rows.push(new TableRow({ children: [referenceCell(pair[0]), cell(gap, [new Paragraph({})]), referenceCell(pair[1])] }));
  }
  return [{ title: labels.references, items: [[fixedTable(ctx, rows, [column, gap, column])]] }];
}

function skillsSection(ctx: DocContext, prepared: PreparedResume): Section[] {
  const { content, labels } = prepared;
  if (!content.skills) return [];
  const lines = content.skills
    .split("\n")
    .map((line) => line.trim())
    .filter(Boolean);
  if (lines.length === 0) return [];
  const paragraphs = lines.map((line, index) => {
    const { label, text } = skillLine(line);
    return new Paragraph({
      bidirectional: bidi(ctx),
      spacing: { before: index === 0 ? firstItemGap(ctx) : twipsIn(ctx.paper, 3), after: 0 },
      children: [
        ...(label ? runsFor(ctx, `${label}: `, { weight: 600, color: INK }) : []),
        ...runsFor(ctx, text, { units: BODY(ctx) }),
      ],
    });
  });
  return [{ title: labels.skills, items: paragraphs.map((paragraph) => [paragraph]) }];
}

function languagesSection(ctx: DocContext, prepared: PreparedResume): Section[] {
  if (!prepared.content.languages) return [];
  return [
    {
      title: prepared.labels.languages,
      items: [[new Paragraph({ bidirectional: bidi(ctx), spacing: { before: firstItemGap(ctx), after: 0 }, children: runsFor(ctx, prepared.content.languages, { units: BODY(ctx) }) })]],
    },
  ];
}

function contactSection(ctx: DocContext, prepared: PreparedResume): Section[] {
  if (prepared.content.contacts.length === 0) return [];
  return [
    {
      title: prepared.labels.contact,
      items: prepared.content.contacts.map((item, index) => [
        new Paragraph({
          bidirectional: bidi(ctx),
          spacing: { before: index === 0 ? firstItemGap(ctx) : twipsIn(ctx.paper, 2), after: 0, ...lineBox(ctx, BODY(ctx, 0.92)) },
          children: contactRuns(ctx, [item], { units: BODY(ctx, 0.92) }),
        }),
      ]),
    },
  ];
}

function summarySections(ctx: DocContext, prepared: PreparedResume): Section[] {
  if (!prepared.content.summary) return [];
  return [
    {
      title: prepared.labels.profile,
      items: [[new Paragraph({ bidirectional: bidi(ctx), spacing: { before: firstItemGap(ctx), after: 0 }, children: runsFor(ctx, prepared.content.summary, { units: BODY(ctx) }) })]],
    },
  ];
}

/** Section headings at the theme's gap, each followed by its content. */
function renderSections(ctx: DocContext, sections: readonly Section[]): Block[] {
  return sections.flatMap((section) => [...sectionTitle(ctx, section.title, twipsIn(ctx.paper, ctx.resolved.density.gap)), ...flat(section)]);
}

/** One group per entry, with its section heading on the first. */
function sectionRows(ctx: DocContext, sections: readonly Section[], gap: number): Block[][] {
  return sections.flatMap((section) =>
    section.items.map((item, index) => [...(index === 0 ? sectionTitle(ctx, section.title, gap) : []), ...item]),
  );
}

/**
 * One group per section, all its items together: a sidebar's few short
 * sections read as one continuous column.
 */
function sectionGroupRows(ctx: DocContext, sections: readonly Section[], gap: number): Block[][] {
  return sections.map((section) => [...sectionTitle(ctx, section.title, gap), ...flat(section)]);
}

/**
 * Word keeps a paragraph after a table at the end of the body. If the table
 * reaches the bottom margin, that paragraph lands on a page of its own. Give
 * it one explicitly, with no height, so a full page does not gain a blank one.
 */
function tableEnd(ctx: DocContext): Paragraph {
  return new Paragraph({
    bidirectional: bidi(ctx),
    spacing: { before: 0, after: 0, line: 1, lineRule: LineRuleType.EXACT },
  });
}

/** Side-heading rows: a heading beside each section's first entry, then its entries. */
function sideHeadingRows(ctx: DocContext, sections: readonly Section[], gap: number): SideHeadingRow[] {
  return sections.flatMap((section) =>
    section.items.map((item, index) => ({
      heading: index === 0 ? sectionHeading(ctx, section.title) : undefined,
      content: item,
      before: index === 0 ? gap : 0,
    })),
  );
}

/** A resume's body: header, then sidebar/main or side-heading sections. */
export async function resumeChildren(ctx: DocContext, prepared: PreparedResume): Promise<FileChild[]> {
  const { content, resolved } = prepared;
  const width = bodyWidth(ctx);
  // Render only visible sections: constructing runs also registers fonts for embedding.
  const builders: Record<string, () => Section | undefined> = {
    summary: () => summarySections(ctx, prepared)[0],
    experience: () => experienceSections(ctx, prepared, width)[0],
    volunteering: () => experienceSections(ctx, prepared, width, "volunteering")[0],
    projects: () => projectSections(ctx, prepared, width)[0],
    education: () => educationSections(ctx, prepared, width)[0],
    certifications: () => certificationSections(ctx, prepared, width)[0],
    publications: () => publicationSections(ctx, prepared)[0],
    references: () => referenceSections(ctx, prepared, width)[0],
    skills: () => skillsSection(ctx, prepared)[0],
    languages: () => languagesSection(ctx, prepared)[0],
  };
  for (const section of content.customSections) builders[customKey(section.id)] = () => customSection(ctx, section, width);
  const sections = shownSections(prepared.sections, Object.keys(builders)).flatMap((key) => {
    const section = builders[key]();
    return section ? [{ ...section, key, title: resumeSectionHeading(prepared.sections, key, section.title) }] : [];
  });

  const header = await headerBlocks(ctx, prepared);

  if (resolved.split) {
    const main = sections.filter((section) => !inSidebar(section.key, content.customSections));
    const sidebar = [...contactSection(ctx, prepared), ...sections.filter((section) => inSidebar(section.key, content.customSections))];
    const gap = twipsIn(ctx.paper, resolved.density.gap);
    const photo = resolved.photo.position === PhotoPosition.SIDEBAR ? await photoImage(ctx, content.photo, content.name, "sidebar") : null;
    const sidebarRows = sectionGroupRows(ctx, sidebar, gap);
    if (photo) {
      if (sidebarRows.length > 0) sidebarRows[0] = [photoParagraph(ctx, photo, 0), ...sidebarRows[0]];
      else sidebarRows.push([photoParagraph(ctx, photo, 0)]);
    }
    return [
      ...header,
      columnTable(ctx, {
        sidebar: sidebarRows,
        main: sectionRows(ctx, main, gap),
        sidebarEnd: resolved.sidebarEnd,
        tinted: resolved.tintedSidebar && sidebarRows.length > 0,
      }),
      tableEnd(ctx),
    ];
  }

  if (resolved.sideHeadings) {
    // Every side-heading row, the first included, carries the gap after the
    // header; the header itself is not a table and has no margin of its own.
    const gap = twipsIn(ctx.paper, resolved.density.gap);
    return [...header, sideHeadingTable(ctx, sideHeadingRows(ctx, sections, gap)), tableEnd(ctx)];
  }

  return [...header, ...renderSections(ctx, sections)];
}
