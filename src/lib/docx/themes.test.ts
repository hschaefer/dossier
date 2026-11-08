import { strFromU8, unzipSync } from "fflate";
import { describe, expect, it } from "vitest";

import { demoProfile } from "@/lib/demo-profile";
import { prepareLetter, prepareResume } from "@/lib/document-prep";
import { BUILT_IN_THEMES } from "@/lib/themes";

import { renderLetterDocx, renderResumeDocx } from "./render";

// Every built-in theme and every layout setting goes through the same
// renderer. This is a structural smoke test, not visual acceptance: it catches
// a theme value the renderer cannot express before a person opens Word.

const profile = demoProfile();
const today = new Date(2026, 8, 28);

const documentXml = (bytes: Uint8Array) => strFromU8(unzipSync(bytes)["word/document.xml"]);

describe("every built-in resume theme", () => {
  it.each(BUILT_IN_THEMES.map((theme) => [theme.name, theme] as const))("renders %s on A4", async (_, theme) => {
    const bytes = await renderResumeDocx(prepareResume({ theme, profile, paper: "A4", today, includeReferences: true }));
    const xml = documentXml(bytes);
    expect(xml).toContain("<w:document");
    expect(xml).toContain(profile.basics.name.slice(0, 4));
    expect(xml).toContain("Experience");
    expect(xml).toContain("<w:sectPr");
  });

  it.each([
    ["SINGLE", "single column"],
    ["SIDE_HEADINGS", "side headings"],
    ["SPLIT", "sidebar left"],
    ["SPLIT_END", "sidebar right"],
  ] as const)("renders a %s layout on Letter", async (layout, label) => {
    const bytes = await renderResumeDocx(
      prepareResume({ theme: BUILT_IN_THEMES[0], profile, paper: "LETTER", layout, today }),
    );
    const xml = documentXml(bytes);
    expect(label).toBeTruthy();
    // Letter is 215.9 x 279.4 mm: 12240 x 15840 twips, rounded.
    expect(xml).toContain('<w:pgSz w:w="12240" w:h="15840"');
    if (layout === "SPLIT" || layout === "SPLIT_END") expect(xml).toContain("<w:tbl>");
  });
});

describe("every built-in theme as a cover letter", () => {
  it.each(BUILT_IN_THEMES.map((theme) => [theme.name, theme] as const))("renders %s", async (_, theme) => {
    const bytes = await renderLetterDocx(
      prepareLetter({
        theme,
        paper: "A4",
        basics: profile.basics,
        units: [{ id: "paragraph:1", label: "First paragraph", text: "Hello there." }],
        company: "Ledgerly",
        location: "Utrecht",
        tone: "WARM",
        date: today,
        dateFormat: "DAY_MONTH_YEAR",
      }),
    );
    const xml = documentXml(bytes);
    expect(xml).toContain("Hello there.");
    expect(xml).toContain(profile.basics.name.slice(0, 4));
  });
});
