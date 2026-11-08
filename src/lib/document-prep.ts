import { ThemeLayout, type DateFormat, type PaperSize, type Tone } from "@/generated/prisma/enums";
import { SIGN_OFF_UNIT, type Unit } from "@/lib/application-data";
import { documentLabels, type DocumentLabels } from "@/lib/document-labels";
import { documentLanguage } from "@/lib/document-language";
import { defaultFrame, isFrameUnit, withFrame, type FrameUnitId } from "@/lib/letter-frame";
import type { Basics, ProfileData } from "@/lib/profile";
import { resumeContent, type ResumeContent, type ResumeOptions } from "@/lib/resume-content";
import { DEFAULT_SECTION_LAYOUT, type SectionLayout } from "@/lib/resume-sections";
import { resolveTheme, type ResolvedTheme, type ThemeData } from "@/lib/themes";
import { TONES } from "@/lib/tones";

// What a document shows, worked out once for every renderer: the preview and
// the PDF draw it as React, the Word export draws it as DOCX. Nothing here
// touches the database or React, so synthetic fixtures can test it.

/** Enough of a resume's own words to tell its script and direction. */
export function resumeSample(content: ResumeContent): string {
  return [
    content.name,
    content.headline,
    content.summary,
    ...[...content.experience, ...content.volunteering].flatMap((role) => [role.role, ...role.bullets]),
    ...content.projects.flatMap((project) => [project.name, ...project.bullets]),
    ...content.customSections.flatMap((section) => section.items.flatMap((item) => [item.title, ...item.bullets])),
  ].join(" ");
}

export type ResumeInput = ResumeOptions & {
  theme: ThemeData;
  profile: ProfileData;
  paper: PaperSize;
  /** Overrides the theme's layout, e.g. a tailored resume in a single column. */
  layout?: ThemeLayout;
  sections?: SectionLayout;
  fit?: number;
  language?: string | null;
};

export type PreparedResume = {
  theme: ThemeData;
  resolved: ResolvedTheme;
  paper: PaperSize;
  labels: DocumentLabels;
  content: ResumeContent;
  sections: SectionLayout;
  lang?: string;
  dir: "ltr" | "rtl";
};

/** The resume as both renderers need it. */
export function prepareResume(input: ResumeInput): PreparedResume {
  const { theme, profile, paper, layout, language, sections = DEFAULT_SECTION_LAYOUT, fit = 0, ...options } = input;
  const resolved = resolveTheme(theme, layout, fit);
  const labels = documentLabels(language);
  const content = resumeContent(profile, options, labels);
  const { lang, dir } = documentLanguage(language, resumeSample(content));
  return { theme, resolved, paper, labels, content, sections, lang, dir };
}

/** The contact items in a typed contact line ("a · b · c"). */
export const contactsIn = (line: string): string[] =>
  line
    .split(/\s+·\s+/)
    .map((item) => item.trim())
    .filter(Boolean);

export type LetterInput = {
  theme: ThemeData;
  paper: PaperSize;
  /** The sender, for a letter without its own sender and signature units. */
  basics: Basics;
  /** The letter's parts: frame units, paragraphs ("paragraph:1", …) and its "sign-off". */
  units: Unit[];
  company: string;
  location: string;
  tone: Tone;
  /** The date for a letter without its own date unit, at UTC midnight. */
  date: Date;
  dateFormat: DateFormat;
  /** Overrides the "sign-off" unit; without either, the tone's sign-off. */
  signOff?: string;
  /** The language it's written in (SET-7); without one, it's read from the text. */
  language?: string | null;
};

export type PreparedLetter = {
  theme: ThemeData;
  resolved: ResolvedTheme;
  paper: PaperSize;
  labels: DocumentLabels;
  lang?: string;
  dir: "ltr" | "rtl";
  senderName: string;
  /** One entry per contact line; each is the items between " · ". */
  contactLines: string[][];
  date: string;
  recipient: string[];
  greeting: string;
  paragraphs: { id: string; text: string }[];
  closing: string;
  signature: string;
};

/**
 * A cover letter (APP-10) in a single column, whatever the theme's resume
 * layout: sender, date, recipient, greeting, paragraphs, sign-off and the
 * name under it. Every part is a unit (E8-11); a letter written before that
 * shows the parts it lacks as letters always had them.
 */
export function prepareLetter(input: LetterInput): PreparedLetter {
  const { theme, paper, basics, units, company, location, tone, date, dateFormat, signOff, language } = input;
  const resolved = resolveTheme(theme, ThemeLayout.SINGLE);
  const labels = documentLabels(language);
  const parts = withFrame(units, defaultFrame({ basics, company, location, date, dateFormat, labels }));
  const text = (id: FrameUnitId) => parts.find((unit) => unit.id === id)?.text.trim() ?? "";
  const paragraphs = parts.filter((unit) => !isFrameUnit(unit.id) && unit.id !== SIGN_OFF_UNIT);
  const closing = signOff ?? parts.find((unit) => unit.id === SIGN_OFF_UNIT)?.text.trim() ?? TONES[tone].signOff;
  const [name = "", ...contactLines] = text("sender").split("\n");
  const { lang, dir } = documentLanguage(language, paragraphs.map((unit) => unit.text).join(" "));

  return {
    theme,
    resolved,
    paper,
    labels,
    lang,
    dir,
    senderName: name.trim(),
    contactLines: contactLines.map(contactsIn).filter((items) => items.length > 0),
    date: text("date"),
    recipient: text("recipient")
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean),
    greeting: text("greeting"),
    paragraphs: paragraphs.filter((unit) => unit.text.trim()).map((unit) => ({ id: unit.id, text: unit.text.trim() })),
    closing,
    signature: text("signature"),
  };
}
