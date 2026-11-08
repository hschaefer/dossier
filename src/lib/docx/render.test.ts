import { deflateSync, strFromU8, unzipSync } from "fflate";
import sharp from "sharp";
import { describe, expect, it } from "vitest";

import { DateFormat, LanguageLevel, PaperSize, ThemeLayout, Tone } from "@/generated/prisma/enums";
import type { Unit } from "@/lib/application-data";
import { prepareLetter, prepareResume } from "@/lib/document-prep";
import { mixHex } from "@/lib/docx/runs";
import type { Basics, ProfileData } from "@/lib/profile";
import { BUILT_IN_SECTIONS, customKey } from "@/lib/resume-sections";
import { BUILT_IN_THEMES, DENSITIES, unitsToMm } from "@/lib/themes";

import { unitsToHalfPoints, unitsToPoints, unitsToTwips } from "./geometry";
import { renderLetterDocx, renderResumeDocx } from "./render";

// Synthetic documents only. The tests read the OOXML package itself: names,
// relationships, real numbering, embedded font parts and escaped user text.

const theme = (id: string) => {
  const found = BUILT_IN_THEMES.find((entry) => entry.id === id);
  if (!found) throw new Error(`No built-in theme ${id}`);
  return found;
};

const basics: Basics = {
  name: "Test Person",
  headline: "Product designer",
  email: "test@example.com",
  phone: "+31 6 0000 0000",
  dob: "",
  website: "test.example",
  address: "",
  city: "Amsterdam",
  country: "Netherlands",
  linkedin: "",
  summary: "Designs calm interfaces with <care> and an ampersand & a pipe |.",
  showAge: false,
};

const profile: ProfileData = {
  basics,
  photo: null,
  skills: [
    { id: "skill-1", name: "Figma", group: "Design" },
    { id: "skill-2", name: "TypeScript", group: "" },
  ],
  languages: [{ id: "lang-1", name: "Dutch", level: LanguageLevel.NATIVE }],
  experience: [
    {
      id: "role-1",
      role: "Senior <Designer>",
      company: "Northwind & Co",
      location: "Amsterdam",
      descriptor: "Design systems lead",
      bullets: "Built a <design> system\nRan research with 20 users",
      skillIds: ["skill-1"],
      startYear: 2019,
      startMonth: 1,
      endYear: 2024,
      endMonth: 6,
      current: false,
      volunteering: false,
    },
  ],
  projects: [
    {
      id: "project-1",
      name: "Tidewater",
      link: "github.com/test/tidewater",
      bullets: "A small project",
      skillIds: [],
      startYear: 2020,
      startMonth: 1,
      endYear: null,
      endMonth: null,
      current: true,
    },
  ],
  education: [
    {
      id: "school-1",
      degree: "MSc Design",
      school: "Test University",
      notes: "With honours",
      startYear: 2014,
      startMonth: 9,
      endYear: 2016,
      endMonth: 6,
      current: false,
    },
  ],
  certifications: [],
  customSections: [],
  publications: [
    {
      id: "paper-1",
      citation: "Doe, J. (2020). **Trust** in *systems*.",
      link: "https://example.com/paper",
      onResumes: true,
    },
  ],
  references: [
    { id: "reference-1", name: "Referent Person", relation: "Manager", company: "Northwind", email: "ref@example.com", phone: "" },
  ],
  facts: [],
};

const resume = (id: string) => prepareResume({ theme: theme(id), profile, paper: PaperSize.A4, includeReferences: true });

const expandedProfile: ProfileData = {
  ...profile,
  experience: [...profile.experience, { ...profile.experience[0], id: "volunteer", role: "Community mentor", company: "Code club", volunteering: true, bullets: "Taught local pupils" }],
  certifications: [{ id: "credential", name: "Certified designer", issuer: "Design institute", year: 2024, month: 3, expiryYear: 2027, expiryMonth: 3, credentialId: "CERT-42", link: "https://example.org/verify", onResumes: true }],
  customSections: (["ENTRIES", "LINES", "INLINE"] as const).map((format) => ({
    id: format,
    title: `${format} heading`,
    format,
    items: [{ id: `${format}-item`, title: format === "LINES" ? "**Award** for *service*" : `${format} item`, subtitle: "Award foundation", startYear: 2022, startMonth: null, endYear: null, endMonth: null, current: false, details: "Won the regional award", link: "https://example.org/award", onResumes: true }],
  })),
};

const units: Unit[] = [
  { id: "paragraph:1", label: "First paragraph", text: "I would like to apply for the role.\nThis is a second line." },
  { id: "paragraph:2", label: "Second paragraph", text: "Thank you for considering it. <script> is not markup here." },
];

const letter = (id: string) =>
  prepareLetter({
    theme: theme(id),
    paper: PaperSize.A4,
    basics,
    units,
    company: "Northwind Logistics",
    location: "Amsterdam",
    tone: Tone.WARM,
    date: new Date(Date.UTC(2026, 9, 3)),
    dateFormat: DateFormat.DAY_MONTH_YEAR,
  });

/** A 1×1 PNG: enough for Word's image part and for isPhotoDataUrl's signature check. */
const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Y9Zl1sAAAAASUVORK5CYII=";
/** A real 48×64 (3:4) PNG, so the circle and rounded shapes can be masked. */
const PHOTO = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAADAAAABACAIAAADTQmMRAAAACXBIWXMAAAPoAAAD6AG1e1JrAAAAdklEQVRoge3YQQ3AQBBC0XrCE55wWxE9TA8vWQETAh+yT7pfvef8gjhoFCoPRcp2jp8A4yhUHoqU7RxCMdBGofJQpGw4NGCs6ohynfnR+2kWA20UKg9FyoZDA8aqjijXmR+952EMtFGoPBQp2zmE4id/FOpXD72Yi1jTt+7dxwAAAABJRU5ErkJggg==";
/** RIFF…WEBP… bytes; docx has no WebP drawing, so this must be left out, not crash. */
const WEBP = `data:image/webp;base64,${Buffer.from("RIFF0000WEBPVP8 0000000000", "binary").toString("base64")}`;

const photoProfile = (src: string): ProfileData => ({ ...profile, photo: { src, onResumes: true } });

const unzip = (bytes: Uint8Array) => unzipSync(bytes);
const xml = (files: Record<string, Uint8Array>, name: string) => strFromU8(files[name] ?? new Uint8Array());

/** The document's visible text, with tags removed. */
const plain = (documentXml: string) =>
  documentXml
    .replace(/<w:tab\/>/g, "\t")
    .replace(/<[^>]+>/g, "")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");

describe("renderResumeDocx", () => {
  it("renders volunteering, certification details and every custom format with real links", async () => {
    const files = unzip(await renderResumeDocx(prepareResume({ theme: theme("editorial"), profile: expandedProfile, paper: PaperSize.A4 })));
    const document = xml(files, "word/document.xml");
    const text = plain(document);
    for (const part of ["Volunteering", "Community mentor", "Taught local pupils", "Certifications", "Certified designer", "Design institute", "Mar 2024", "Valid until Mar 2027", "Credential ID CERT-42", "ENTRIES item", "Award foundation", "Won the regional award", "Award for service", "INLINE item"]) expect(text).toContain(part);
    expect(text).not.toContain("**Award**");
    expect(document).toContain('w:ascii="Geist Bold"');
    expect(document).toContain("<w:i/>");
    const links = xml(files, "word/_rels/document.xml.rels");
    expect(links).toContain("https://example.org/verify");
    expect(links).toContain("https://example.org/award");
  });

  it.each(["editorial", "folio"])("honors ordering, hiding and renamed headings in %s", async (id) => {
    const prepared = prepareResume({
      theme: theme(id), profile: expandedProfile, paper: PaperSize.A4,
      sections: {
        order: ["languages", customKey("ENTRIES"), "certifications", ...BUILT_IN_SECTIONS.filter((key) => key !== "languages" && key !== "certifications"), customKey("INLINE"), customKey("LINES")],
        off: ["experience", "projects", customKey("LINES")],
        titles: { certifications: "Credentials", [customKey("ENTRIES")]: "Recognition", summary: "About me" },
      },
    });
    const text = plain(xml(unzip(await renderResumeDocx(prepared)), "word/document.xml"));
    const order = ["Languages", "Recognition", "Credentials", "About me", "Volunteering", "Education", "Skills", "INLINE heading"];
    let previous = -1;
    for (const part of order) {
      const at = text.indexOf(part, previous + 1);
      expect(at, part).toBeGreaterThan(previous);
      previous = at;
    }
    for (const absent of ["Senior <Designer>", "Tidewater", "LINES heading", "Award for service", "Certifications", "ENTRIES heading"]) expect(text).not.toContain(absent);
  });

  it("keeps inline custom sections in the sidebar, but follows global order for a single-column override", async () => {
    const input = {
      theme: theme("modern"), profile: expandedProfile, paper: PaperSize.A4,
      sections: { order: [customKey("ENTRIES"), customKey("INLINE"), "summary", ...BUILT_IN_SECTIONS], off: ["skills", "languages"] },
    };
    const sidebar = plain(xml(unzip(await renderResumeDocx(prepareResume(input))), "word/document.xml"));
    expect(sidebar.indexOf("Contact")).toBeLessThan(sidebar.indexOf("INLINE heading"));
    expect(sidebar.indexOf("INLINE heading")).toBeLessThan(sidebar.indexOf("ENTRIES heading"));
    const single = plain(xml(unzip(await renderResumeDocx(prepareResume({ ...input, layout: ThemeLayout.SINGLE }))), "word/document.xml"));
    expect(single.indexOf("ENTRIES heading")).toBeLessThan(single.indexOf("INLINE heading"));
    expect(single).not.toContain("Contact");
  });

  it("exports tailored new sections rather than their original profile text", async () => {
    const prepared = prepareResume({
      theme: theme("editorial"), profile: expandedProfile, paper: PaperSize.A4,
      units: [
        { id: "experience:volunteer", label: "Volunteering", text: "Tailored mentoring" },
        { id: "certification:credential", label: "Certification", text: "Tailored credential" },
        { id: "custom-entry:ENTRIES-item", label: "Award", text: "Tailored award details" },
        { id: "custom-line:LINES-item", label: "Line", text: "Tailored **line**" },
        { id: "custom-inline:INLINE", label: "Interests", text: "Chess\nClimbing" },
      ],
    });
    const text = plain(xml(unzip(await renderResumeDocx(prepared)), "word/document.xml"));
    for (const part of ["Tailored mentoring", "Tailored credential", "Tailored award details", "Tailored line", "Chess · Climbing"]) expect(text).toContain(part);
    for (const absent of ["Taught local pupils", "Certified designer", "Won the regional award", "Award for service", "INLINE item", "Senior <Designer>"]) expect(text).not.toContain(absent);
  });

  it("uses the shared one-page fit typography in the Word styles", async () => {
    const prepared = prepareResume({ theme: theme("editorial"), profile, paper: PaperSize.A4, fit: 4 });
    const styles = xml(unzip(await renderResumeDocx(prepared)), "word/styles.xml");
    const density = prepared.resolved.density;
    expect(density.fontSize).toBeLessThan(DENSITIES[theme("editorial").density].fontSize);
    expect(styles).toContain(`w:line="${Math.round(density.lineHeight * unitsToPoints(density.fontSize, PaperSize.A4) * 20)}"`);
  });

  it("writes a valid A4 package with theme fonts, embedded faces and escaped text", async () => {
    const files = unzip(await renderResumeDocx(resume("editorial")));
    const document = xml(files, "word/document.xml");
    const fontTable = xml(files, "word/fontTable.xml");
    const styles = xml(files, "word/styles.xml");
    const contentTypes = xml(files, "[Content_Types].xml");
    const relationships = xml(files, "word/_rels/document.xml.rels");
    const numbering = xml(files, "word/numbering.xml");
    const settings = xml(files, "word/settings.xml");

    expect(document).toContain("<w:document");
    // A4, with the Regular density's margins as the page margins.
    expect(document).toContain('<w:pgSz w:w="11906" w:h="16838"');
    const top = Math.round((50 * 210 * 1440) / (640 * 25.4));
    const left = Math.round((54 * 210 * 1440) / (640 * 25.4));
    expect(document).toContain(`w:top="${top}"`);
    expect(document).toContain(`w:left="${left}"`);
    expect(unitsToMm(54, PaperSize.A4)).toBeCloseTo(17.719, 2);
    expect(DENSITIES[theme("editorial").density].fontSize).toBe(10.5);

    // The theme font is named by runs, and every named face has a font part.
    expect(document).toContain('w:ascii="Newsreader');
    expect(fontTable).toContain("Newsreader");
    expect(fontTable).toContain("<w:embedRegular");
    expect(Object.keys(files).some((name) => name.startsWith("word/fonts/") && name.endsWith(".odttf"))).toBe(true);
    expect(contentTypes).toContain("obfuscatedFont");
    expect(styles).toContain("w:docDefaults");
    expect(numbering).toContain("w:abstractNum");
    // Word keeps the faces when the document is saved again, and keeps all
    // their glyphs rather than a subset of what this text used.
    expect(settings).toContain("<w:embedTrueTypeFonts/>");
    expect(settings).toContain('<w:saveSubsetFonts w:val="false"/>');

    const text = plain(document);
    expect(text).toContain("Test Person");
    expect(text).toContain("Product designer");
    expect(text).toContain("Profile");
    expect(text).toContain("Experience");
    expect(text).toContain("Projects");
    expect(text).toContain("Education");
    expect(text).toContain("Publications");
    expect(text).toContain("References");

    // Real list numbering, not a "• " string.
    expect(document).toContain("<w:numPr>");
    expect(document).toContain("Ran research with 20 users");
    // User text is escaped XML text.
    expect(document).toContain("&lt;Designer&gt;");
    expect(document).toContain("&lt;design&gt;");
    expect(document).not.toContain("<Designer>");

    expect(relationships).toContain("relationships/hyperlink");
    expect(relationships).toContain("https://example.com/paper");
    // The visible text is the shortened link; the relationship holds the URL.
    expect(plain(document)).toContain("example.com/paper");
  });

  it("puts a sidebar theme's two columns in one row so they flow independently", async () => {
    const noReferences = prepareResume({
      theme: theme("modern"),
      profile: { ...profile, references: [] },
      paper: PaperSize.A4,
      includeReferences: false,
    });
    const files = unzip(await renderResumeDocx(noReferences));
    const document = xml(files, "word/document.xml");
    const text = plain(document);

    expect(document).toContain("<w:tbl>");
    // One outer row with sidebar, gap and main cells. No nested table in this
    // profile, so the count is exact.
    expect(document.match(/<w:tr>/g)).toHaveLength(1);
    expect(document.match(/<w:tc>/g)).toHaveLength(3);
    // A row taller than the page has to be allowed to split. A fixed height or
    // cantSplit would make Word move the whole resume to page two.
    expect(document).not.toContain("w:cantSplit");
    expect(document).not.toContain("w:trHeight");
    // Cell order is the whole sidebar, then the whole main column. The sidebar
    // is written first when it opens the page.
    expect(text.indexOf("Contact")).toBeGreaterThan(-1);
    expect(text.indexOf("Contact")).toBeLessThan(text.indexOf("Skills"));
    expect(text.indexOf("Skills")).toBeLessThan(text.indexOf("Profile"));
  });

  it("ends a layout table with a zero-height paragraph so Word does not add a blank page", async () => {
    const files = unzip(await renderResumeDocx(resume("technical")));
    const document = xml(files, "word/document.xml");
    const tail = document.slice(document.lastIndexOf("</w:tbl>"));

    // Word keeps a paragraph after a table; without one it adds a normal-sized
    // paragraph, which spills to page two when the table fills the page.
    expect(tail).toMatch(/^<\/w:tbl><w:p>/);
    expect(tail).toContain('w:line="1"');
    expect(tail).toContain('w:lineRule="exact"');
    expect(tail.indexOf('w:line="1"')).toBeLessThan(tail.indexOf("<w:sectPr"));
  });

  it("shades Studio's tinted sidebar at 7% of its accent", async () => {
    const files = unzip(await renderResumeDocx(resume("studio")));
    const document = xml(files, "word/document.xml");
    const fill = mixHex(theme("studio").accent, 0.07).replace("#", "");

    expect(document).toContain(`w:fill="${fill}"`);
    // The sidebar at the end is written first and mirrored: it stays on the
    // first page, while the table's cell order is sidebar, main.
    expect(document).toContain("<w:bidiVisual/>");
    const text = plain(document);
    expect(text.indexOf("Contact")).toBeLessThan(text.indexOf("Profile"));
  });

  it("leaves out sections with no content, as the preview does", async () => {
    const empty: ProfileData = {
      ...profile,
      experience: [],
      projects: [],
      education: [],
      publications: [],
      references: [],
      skills: [],
      languages: [],
    };
    const prepared = prepareResume({ theme: theme("editorial"), profile: empty, paper: PaperSize.A4, includeReferences: true });
    const text = plain(xml(unzip(await renderResumeDocx(prepared)), "word/document.xml"));

    expect(text).toContain("Profile");
    for (const heading of ["Experience", "Projects", "Education", "Publications", "References", "Skills", "Languages"]) {
      expect(text).not.toContain(heading);
    }
  });

  it("writes timeline dates in an indent column and keeps headings with their content", async () => {
    const files = unzip(await renderResumeDocx(resume("chronicle")));
    const document = xml(files, "word/document.xml");

    expect(document).toContain("<w:tabs>");
    expect(document).toMatch(/<w:tab w:val="left"/);
    expect(document).toContain('w:hanging="');
    expect(document).toContain("<w:keepNext");
    expect(plain(document)).toContain("Jan 2019 – Jun 2024");
  });

  it("writes the CSS typography as Word line boxes, character scale and kerning", async () => {
    const files = unzip(await renderResumeDocx(resume("editorial")));
    const document = xml(files, "word/document.xml");
    const styles = xml(files, "word/styles.xml");
    const density = DENSITIES[theme("editorial").density];

    // The body line box is 1.5 × 9.77pt, not Word's "1.5 lines" (1.5 × the
    // font's own 1.3em height). Word's multiple made every line a third taller.
    const bodyLine = Math.round(density.lineHeight * unitsToPoints(density.fontSize, PaperSize.A4) * 20);
    expect(bodyLine).toBe(293);
    expect(styles).toContain(`w:line="${bodyLine}"`);
    expect(styles).toContain('w:lineRule="atLeast"');

    // Half-points round 9.77pt up to 10pt, so runs scale to 98% to keep the
    // advances (and the PDF's line breaks); kerning makes the letters the
    // same distance apart as the browser draws them.
    expect(document).toContain('<w:w w:val="98"/>');
    expect(document).toContain('<w:kern w:val="2"/>');

    // The 30-unit name keeps its 1.1 line height in points.
    const nameLine = Math.round(1.1 * unitsToPoints(30, PaperSize.A4) * 20);
    expect(document).toContain(`w:line="${nameLine}"`);
  });

  it("draws a bar heading at the text edge, one line tall, with its gaps outside the border", async () => {
    const document = xml(unzip(await renderResumeDocx(resume("signal"))), "word/document.xml");
    const paragraph = (needle: string) =>
      document.match(new RegExp(`<w:p>(?:(?!<\\/w:p>)[\\s\\S])*?${needle}[\\s\\S]*?<\\/w:p>`))?.[0] ?? "";
    const heading = paragraph(">Profile</w:t>");

    // The bar is a 3-unit left border with 7 units of space, on a paragraph
    // indented by both so the bar lands on the text edge like the CSS one.
    expect(heading).toContain('<w:left w:val="single"');
    expect(heading).toContain(`<w:ind w:start="${unitsToTwips(10, PaperSize.A4)}"/>`);
    // No spacing of its own: Word would draw the border around it as well.
    expect(heading).toContain('w:before="0"');
    expect(heading).toContain('w:after="0"');
    expect(heading).toContain(`w:line="${Math.round(1.5 * unitsToPoints(9.5, PaperSize.A4) * 20)}"`);

    // The section gap above it is an empty paragraph of exactly that height,
    // and the gap below moved to the section's first item.
    expect(document).toContain(`w:line="${unitsToTwips(20, PaperSize.A4)}" w:lineRule="exact"`);
    expect(paragraph("Designs calm interfaces")).toContain(`w:before="${unitsToTwips(6, PaperSize.A4)}"`);
  });

  it("styles bullets like the print list: a small disc at a 1.1em inset", async () => {
    const fontSize = DENSITIES[theme("editorial").density].fontSize;
    const inset = unitsToTwips(1.1 * fontSize, PaperSize.A4);
    const hanging = unitsToTwips(0.6 * fontSize, PaperSize.A4);
    const firstBullet = (document: string) =>
      document.match(/<w:p>(?:(?!<\/w:p>)[\s\S])*?<w:numPr>(?:(?!<\/w:p>)[\s\S])*?<\/w:p>/)?.[0] ?? "";

    const files = unzip(await renderResumeDocx(resume("editorial")));
    const document = xml(files, "word/document.xml");
    const numbering = xml(files, "word/numbering.xml");
    const bullet = firstBullet(document);

    // A small disc in the body face and size, not docx's default ● at the
    // paragraph's inherited size.
    expect(numbering).toContain('<w:lvlText w:val="•"/>');
    expect(numbering).toContain('<w:rFonts w:ascii="Geist"');
    expect(numbering).toContain(`<w:sz w:val="${unitsToHalfPoints(fontSize, PaperSize.A4)}"/>`);
    // Real list numbering, with the print inset rather than Word's 0.5".
    expect(bullet).toContain("<w:numPr>");
    expect(bullet).not.toContain('w:pStyle w:val="ListParagraph"');
    expect(bullet).toContain(`<w:ind w:left="${inset}" w:hanging="${hanging}"/>`);

    // A timeline entry's bullets sit in the date column, then at the list inset.
    const chronicle = xml(unzip(await renderResumeDocx(resume("chronicle"))), "word/document.xml");
    const timelineInset = unitsToTwips(108, PaperSize.A4) + inset;
    expect(firstBullet(chronicle)).toContain(`<w:ind w:left="${timelineInset}" w:hanging="${hanging}"/>`);
  });
});

describe("photos and scripts", () => {
  it("bakes the circle and rounded shapes into the PNG's alpha channel", async () => {
    for (const id of ["classic", "modern", "signal"] as const) {
      const prepared = prepareResume({ theme: theme(id), profile: photoProfile(PHOTO), paper: PaperSize.A4 });
      const files = unzip(await renderResumeDocx(prepared));
      const document = xml(files, "word/document.xml");
      const media = Object.keys(files).filter((name) => name.startsWith("word/media/"));

      expect(media.length, id).toBe(1);
      expect(media[0], id).toMatch(/\.png$/);
      // The mask replaces Word's square crop: the image is already square.
      expect(document.match(/<a:srcRect[^>]*>/)?.[0] ?? "", id).toBe("<a:srcRect/>");

      const image = sharp(Buffer.from(files[media[0]])).ensureAlpha();
      const { data, info } = await image.raw().toBuffer({ resolveWithObject: true });
      const alpha = (x: number, y: number) => data[(y * info.width + x) * 4 + 3];
      expect(alpha(0, 0), `${id} corner`).toBe(0);
      expect(alpha(info.width - 1, info.height - 1), `${id} corner`).toBe(0);
      expect(alpha(info.width >> 1, info.height >> 1), `${id} centre`).toBe(255);
      // Stored at print resolution: never larger than the source (48px here),
      // and capped at the 300 dpi size for the shape's place on the paper.
      expect(info.width, id).toBeLessThanOrEqual(420);
      expect(info.width, id).toBeGreaterThan(0);
    }
  });

  it("crops the square shape in Word and leaves the portrait as it is", async () => {
    for (const [id, cropped] of [
      ["minimal", true],
      ["chronicle", false],
    ] as const) {
      const prepared = prepareResume({ theme: theme(id), profile: photoProfile(PNG), paper: PaperSize.A4 });
      const files = unzip(await renderResumeDocx(prepared));
      const document = xml(files, "word/document.xml");
      const media = Object.keys(files).filter((name) => name.startsWith("word/media/"));

      expect(media.length, id).toBe(1);
      expect(media[0]).toMatch(/\.png$/);
      expect(xml(files, "[Content_Types].xml")).toContain("image/png");
      expect(document).toContain('descr="Test Person"');
      // An uncropped image has an empty srcRect; a square crop trims 12.5% top and bottom.
      const rect = document.match(/<a:srcRect[^>]*>/)?.[0] ?? "";
      expect(rect).toContain(cropped ? 't="12500"' : "<a:srcRect/>");
    }
  });

  it("omits a WebP photo without crashing the package", async () => {
    const prepared = prepareResume({ theme: theme("modern"), profile: photoProfile(WEBP), paper: PaperSize.A4 });
    const files = unzip(await renderResumeDocx(prepared));
    const document = xml(files, "word/document.xml");

    expect(Object.keys(files).some((name) => name.startsWith("word/media/"))).toBe(false);
    expect(document).not.toContain("<a:blip");
    expect(plain(document)).toContain("Test Person");
  });

  it("selects the script faces and marks Arabic runs right to left", async () => {
    const prepared = prepareResume({
      theme: theme("classic"),
      profile: { ...profile, basics: { ...basics, summary: "ليلى حسن تعمل في التصميم" } },
      paper: PaperSize.A4,
      language: "Arabic",
    });
    const document = xml(unzip(await renderResumeDocx(prepared)), "word/document.xml");

    expect(document).toContain("Dossier Arabic");
    expect(document).toMatch(/<w:rtl/);
    expect(document).toContain("<w:bidi");
  });

  it.each([
    ["नमस्ते दुनिया", "Hindi", "Dossier Devanagari"],
    ["こんにちは世界", "Japanese", "Dossier Japanese"],
    ["안녕하세요", "Korean", "Dossier Korean"],
    ["你好世界", "Chinese (Simplified)", "Dossier Chinese"],
    // Classic is a serif theme: Greek takes EB Garamond, not the sans fallback.
    ["Καλημέρα κόσμε", "Greek", "EB Garamond"],
  ])("embeds the face for %s (%s)", async (summary, language, family) => {
    const prepared = prepareResume({
      theme: theme("classic"),
      profile: { ...profile, basics: { ...basics, summary } },
      paper: PaperSize.A4,
      language,
    });
    expect(xml(unzip(await renderResumeDocx(prepared)), "word/document.xml")).toContain(family);
  });

  it("reports the embedded font bytes for a Studio A4 resume", async () => {
    const files = unzip(await renderResumeDocx(resume("studio")));
    const fonts = Object.entries(files).filter(([name]) => name.startsWith("word/fonts/") && name.endsWith(".odttf"));
    const bytes = fonts.reduce((total, [, data]) => total + data.byteLength, 0);
    const deflated = fonts.reduce((total, [, data]) => total + deflateSync(data).byteLength, 0);

    console.info(`Studio A4 embedded fonts: ${fonts.length} faces, ${bytes} bytes uncompressed, ${deflated} bytes deflated`);
    expect(bytes).toBeGreaterThan(0);
  });
});

describe("renderLetterDocx", () => {
  it("keeps the letter parts in reading order and escapes its text", async () => {
    const files = unzip(await renderLetterDocx(letter("editorial")));
    const document = xml(files, "word/document.xml");
    const text = plain(document);

    const order = ["Test Person", "3 October 2026", "Northwind Logistics · Hiring team", "Amsterdam", "Dear Northwind Logistics team,"];
    let at = -1;
    for (const part of order) {
      const index = text.indexOf(part, at + 1);
      expect(index, part).toBeGreaterThan(at);
      at = index;
    }
    const firstParagraph = text.indexOf("I would like to apply for the role.");
    const secondParagraph = text.indexOf("Thank you for considering it.");
    expect(firstParagraph).toBeGreaterThan(at);
    expect(secondParagraph).toBeGreaterThan(firstParagraph);
    expect(text.indexOf("Warm regards,")).toBeGreaterThan(secondParagraph);
    expect(text.lastIndexOf("Test Person")).toBeGreaterThan(text.indexOf("Warm regards,"));

    // A line break, not a new paragraph, inside the first letter paragraph.
    expect(document).toContain("<w:br/>");
    expect(document).toContain("&lt;script&gt;");
    expect(document).not.toContain("<script>");
  });
});
