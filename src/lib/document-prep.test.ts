import { describe, expect, it } from "vitest";

import { demoProfile, demoResumeUnits } from "@/lib/demo-profile";
import { prepareLetter, prepareResume } from "@/lib/document-prep";
import { DEFAULT_SECTION_LAYOUT } from "@/lib/resume-sections";
import { BUILT_IN_THEMES } from "@/lib/themes";

const theme = (id: string) => BUILT_IN_THEMES.find((entry) => entry.id === id)!;
const profile = demoProfile();
const today = new Date(2026, 8, 28);

describe("prepareResume", () => {
  it("prepares the base resume from the profile and settings", () => {
    const prepared = prepareResume({ theme: theme("editorial"), profile, paper: "A4", today, includeReferences: true });
    expect(prepared.resolved.split).toBe(false);
    expect(prepared.labels.experience).toBe("Experience");
    expect(prepared.content.name).toBe(profile.basics.name);
    expect(prepared.content.references.length).toBeGreaterThan(0);
    expect(prepared.lang).toBeUndefined();
    expect(prepared.dir).toBe("ltr");
    expect(prepared.sections).toEqual(DEFAULT_SECTION_LAYOUT);
  });

  it("shares section settings and fitted theme measurements with the preview", () => {
    const sections = { ...DEFAULT_SECTION_LAYOUT, off: ["experience"], titles: { summary: "About me" } };
    const regular = prepareResume({ theme: theme("editorial"), profile, paper: "A4" });
    const prepared = prepareResume({ theme: theme("editorial"), profile, paper: "A4", sections, fit: 2 });
    expect(prepared.sections).toBe(sections);
    expect(prepared.resolved.density.paddingY).toBeLessThan(regular.resolved.density.paddingY);
    // Visibility is applied by each renderer, not by deleting the source content.
    expect(prepared.content.experience).toEqual(regular.content.experience);
    expect(prepared.content.certifications).toEqual(regular.content.certifications);
  });

  it("uses tailored units in their own order, not the profile's raw text", () => {
    const prepared = prepareResume({ theme: theme("modern"), profile, paper: "A4", units: demoResumeUnits(), today });
    expect(prepared.content.summary).not.toBe(profile.basics.summary);
    expect(prepared.content.experience.map((entry) => entry.id)).toEqual(
      demoResumeUnits()
        .filter((unit) => unit.id.startsWith("experience:"))
        .map((unit) => unit.id.slice("experience:".length)),
    );
  });

  it("folds a tailored resume's sidebar into a single column when asked", () => {
    expect(prepareResume({ theme: theme("modern"), profile, paper: "A4", today }).resolved.split).toBe(true);
    expect(prepareResume({ theme: theme("modern"), profile, paper: "A4", layout: "SINGLE", today }).resolved.split).toBe(false);
  });

  it("keeps the photo policy and reads the document's language", () => {
    const photo = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
    const withPhoto = { ...profile, photo: { src: photo, onResumes: true } };
    const shown = prepareResume({ theme: theme("editorial"), profile: withPhoto, paper: "A4", language: "German", today });
    expect(shown.labels.experience).toBe("Berufserfahrung");
    expect(shown.content.photo).toBe(photo);
    expect(prepareResume({ theme: theme("editorial"), profile: withPhoto, paper: "A4", showPhoto: false, today }).content.photo).toBeNull();
    expect(prepareResume({ theme: theme("editorial"), profile, paper: "A4", today }).content.photo).toBeNull();
  });
});

describe("prepareLetter", () => {
  const basics = profile.basics;
  const letter = () =>
    prepareLetter({
      theme: theme("editorial"),
      paper: "A4",
      basics,
      units: [
        { id: "recipient", label: "Recipient", text: "Ledgerly · Hiring team\nUtrecht" },
        { id: "greeting", label: "Greeting", text: "Dear Ledgerly team," },
        { id: "paragraph:1", label: "First paragraph", text: "First line.\nSecond line." },
        { id: "paragraph:2", label: "Second paragraph", text: "Another paragraph." },
        { id: "sign-off", label: "Sign-off", text: "Kind regards," },
      ],
      company: "Ledgerly",
      location: "Utrecht",
      tone: "WARM",
      date: new Date(Date.UTC(2026, 8, 28)),
      dateFormat: "DAY_MONTH_YEAR",
      language: "English",
    });

  it("resolves the frame from the profile and the letter's own parts", () => {
    const prepared = letter();
    expect(prepared.senderName).toBe(basics.name);
    expect(prepared.contactLines[0]).toContain(basics.email);
    expect(prepared.date).toBe("28 September 2026");
    expect(prepared.recipient).toEqual(["Ledgerly · Hiring team", "Utrecht"]);
    expect(prepared.greeting).toBe("Dear Ledgerly team,");
    expect(prepared.paragraphs.map((unit) => unit.id)).toEqual(["paragraph:1", "paragraph:2"]);
    expect(prepared.paragraphs[0].text).toBe("First line.\nSecond line.");
    expect(prepared.closing).toBe("Kind regards,");
    expect(prepared.signature).toBe(basics.name);
    expect(prepared.resolved.split).toBe(false);
  });

  it("falls back to the tone's sign-off and the letter date in its language", () => {
    const prepared = prepareLetter({
      theme: theme("editorial"),
      paper: "LETTER",
      basics,
      units: [],
      company: "",
      location: "",
      tone: "DIRECT",
      date: new Date(Date.UTC(2026, 9, 3)),
      dateFormat: "DAY_MONTH_YEAR",
      language: "German",
    });
    expect(prepared.greeting).toBe("Sehr geehrte Damen und Herren,");
    expect(prepared.date).toBe("3. Oktober 2026");
    expect(prepared.closing).toBe("Best,");
    expect(prepared.paper).toBe("LETTER");
  });

  it("honours a sign-off override and a right-to-left language", () => {
    const prepared = prepareLetter({
      theme: theme("editorial"),
      paper: "A4",
      basics,
      units: [],
      company: "Ledgerly",
      location: "",
      tone: "WARM",
      date: new Date(Date.UTC(2026, 8, 28)),
      dateFormat: "DAY_MONTH_YEAR",
      signOff: "Best,",
      language: "Persian",
    });
    expect(prepared.closing).toBe("Best,");
    expect(prepared.dir).toBe("rtl");
    expect(prepared.lang).toBe("fa");
  });
});
