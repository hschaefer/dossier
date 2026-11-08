import { LineRuleType, TextRun } from "docx";

import type { PaperSize } from "@/generated/prisma/enums";
import { hasRightToLeft } from "@/lib/document-language";
import { CONTACT_SEPARATOR } from "@/lib/resume-content";
import { SIDEBAR, THEME_FONTS, type ResolvedTheme } from "@/lib/themes";

import { contentWidthTwips, lineBoxTwips, unitsToHalfPoints, unitsToPoints, unitsToTwips } from "./geometry";
import {
  facesFor,
  SCRIPT_FONT_KEYS,
  WORD_FONT_KEYS,
  wordFont,
  type FontQuery,
  type WordFace,
} from "./fonts";

// Runs are where a theme's typography becomes real Word formatting: a family
// per weight, an explicit script font per grapheme cluster, and a language.
// `wordFont` already resolves the nearest prepared face, so a 600-weight run
// names the semibold family rather than relying on Word's synthetic bold.

/** The document's ink, text and muted colours (document.module.css). */
export const INK = "#1d1c1a";
export const TEXT = "#2b2a27";
export const MUTED = "#69655d";

/** A "#rrggbb" colour without its hash, as Word stores it. */
export const hex = (color: string) => color.replace("#", "").toLowerCase();

/** `accent` mixed `amount` (0–1) into `base`, as CSS color-mix would. */
export function mixHex(accent: string, amount: number, base = "#ffffff"): string {
  const channel = (color: string, start: number) => Number.parseInt(color.slice(start, start + 2), 16);
  const mixed = [1, 3, 5].map((start) =>
    Math.round(channel(accent, start) * amount + channel(base, start) * (1 - amount)),
  );
  return `#${mixed.map((value) => value.toString(16).padStart(2, "0")).join("")}`;
}

export type FontRole = "head" | "body";

/** Word's language properties for the document. */
export type DocLanguage = { value: string; eastAsia?: string; bidirectional?: string };

export type DocContext = {
  paper: PaperSize;
  resolved: ResolvedTheme;
  headKey: string;
  bodyKey: string;
  lang?: string;
  dir: "ltr" | "rtl";
  language: DocLanguage;
  fonts: FontCollector;
};

/** The BCP 47 tag to put on runs, with the slots CJK and bidi text use. */
function wordLanguage(lang?: string): DocLanguage {
  const value = lang ?? "en-US";
  const base = value.split("-")[0];
  return {
    value: base === "en" && !value.includes("-") ? "en-US" : value,
    eastAsia: ["ja", "ko", "zh"].includes(base) ? value : undefined,
    bidirectional: ["ar", "fa", "he", "ur"].includes(base) ? value : undefined,
  };
}

/** Remembers every face a run names, for embedding, and resolves it. */
export class FontCollector {
  private readonly queries: FontQuery[] = [];

  constructor(private readonly headKey: string, private readonly bodyKey: string) {}

  /** The family for a role at a weight, and whether Word must slant it. */
  face(role: FontRole, weight: number, italic = false, optical: "text" | "display" = "text") {
    return this.query(role === "head" ? this.headKey : this.bodyKey, weight, italic, optical);
  }

  /** The family for a manifest key (a theme font or a script fallback). */
  query(key: string, weight: number, italic = false, optical: "text" | "display" = "text") {
    const query: FontQuery = { key, weight, style: italic ? "italic" : "normal", optical };
    this.queries.push(query);
    return wordFont(query);
  }

  /** Every distinct face used, for `embedFonts`. */
  faces(): WordFace[] {
    return facesFor(this.queries);
  }
}

export function createContext(input: {
  paper: PaperSize;
  resolved: ResolvedTheme;
  headFont: string;
  bodyFont: string;
  lang?: string;
  dir: "ltr" | "rtl";
}): DocContext {
  return {
    paper: input.paper,
    resolved: input.resolved,
    headKey: WORD_FONT_KEYS[input.headFont as keyof typeof WORD_FONT_KEYS],
    bodyKey: WORD_FONT_KEYS[input.bodyFont as keyof typeof WORD_FONT_KEYS],
    lang: input.lang,
    dir: input.dir,
    language: wordLanguage(input.lang),
    fonts: new FontCollector(
      WORD_FONT_KEYS[input.headFont as keyof typeof WORD_FONT_KEYS],
      WORD_FONT_KEYS[input.bodyFont as keyof typeof WORD_FONT_KEYS],
    ),
  };
}

/** `true` for a right-to-left document; left out for a left-to-right one. */
export const bidi = (ctx: DocContext) => (ctx.dir === "rtl" ? true : undefined);

/** The theme's body size in Word's half-points. */
export const bodyHalfPoints = (ctx: DocContext) => unitsToHalfPoints(ctx.resolved.density.fontSize, ctx.paper);

/** A size in page units as Word's half-points. */
export const halfPointsIn = (paper: PaperSize, units: number) => unitsToHalfPoints(units, paper);

/** A length in page units as twips. */
export const twipsIn = (paper: PaperSize, units: number) => unitsToTwips(units, paper);

/** The text width on the page with the theme's side padding. */
export const textWidth = (ctx: DocContext) => contentWidthTwips(ctx.paper, ctx.resolved.density.paddingX);

/**
 * A paragraph's line box for text `units` tall: its CSS line-height as Word's
 * "at least" spacing, in twips. Word's own multiple spacing measures the
 * font's height instead of the size, so the CSS number alone made every line
 * about a third taller than the PDF. Word still grows the line when the font
 * needs more room than the CSS gives it (the 1.1 line height of a name over a
 * font whose own height is 1.3 em), where the browser would overlap instead of
 * clip.
 */
export function lineBox(ctx: DocContext, units: number, lineHeight = ctx.resolved.density.lineHeight) {
  return { line: lineBoxTwips(lineHeight, unitsToPoints(units, ctx.paper)), lineRule: LineRuleType.AT_LEAST };
}

/**
 * The character scale, in percent, that keeps a rounded Word size at its CSS
 * width. Word takes half-points only, so a 10.5-unit body (9.77 pt) is drawn
 * at 10 pt; without this, every line is 2.4% wider and wraps sooner than the
 * PDF. Word's scale squeezes each character's advance and glyph width by the
 * same percentage, so lines break where the PDF's do.
 */
function sizeScale(ctx: DocContext, units: number): number | undefined {
  const drawn = unitsToHalfPoints(units, ctx.paper) / 2;
  const percent = Math.round((unitsToPoints(units, ctx.paper) / drawn) * 100);
  return percent === 100 ? undefined : percent;
}

export type RunStyle = {
  role?: FontRole;
  weight?: number;
  italic?: boolean;
  optical?: "text" | "display";
  /** The font size in page units; the body size when left out. */
  units?: number;
  color?: string;
  allCaps?: boolean;
  /** Word's `w:spacing`, in twentieths of a point (tracking on top of the size scale). */
  characterSpacing?: number;
  /** Mark a piece of text left-to-right inside a right-to-left document. */
  ltr?: boolean;
};

type ScriptChunk = { text: string; key: string; rtl: boolean };

const GRAPHEMES = new Intl.Segmenter(undefined, { granularity: "grapheme" });

/** The language a script font's own glyphs are written in. */
const SCRIPT_LANGUAGES: Record<string, string> = {
  [SCRIPT_FONT_KEYS.arabic]: "ar",
  [SCRIPT_FONT_KEYS.devanagari]: "hi",
  [SCRIPT_FONT_KEYS.japanese]: "ja",
  [SCRIPT_FONT_KEYS.korean]: "ko",
  [SCRIPT_FONT_KEYS.chinese]: "zh-Hans",
};

function runLanguage(ctx: DocContext, key: string, rtl: boolean): DocLanguage {
  const script = SCRIPT_LANGUAGES[key];
  if (script) {
    return {
      value: script,
      eastAsia: /^(ja|ko|zh)/.test(script) ? script : undefined,
      bidirectional: script === "ar" ? script : undefined,
    };
  }
  if (rtl) return { value: ctx.language.value, bidirectional: ctx.language.bidirectional ?? "ar" };
  return ctx.language;
}

const ARABIC = /[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]/u;
const OTHER_RTL = /[\u0590-\u05FF\u0700-\u074F\u0780-\u07BF]/u;
const DEVANAGARI = /[\u0900-\u097F\uA8E0-\uA8FF]/u;
const GREEK = /[\u0370-\u03FF\u1F00-\u1FFF]/u;
const KANA = /[\u3040-\u30FF\u31F0-\u31FF\uFF66-\uFF9D]/u;
const HANGUL = /[\u1100-\u11FF\u3130-\u318F\uA960-\uA97F\uAC00-\uD7AF]/u;
const HAN = /[\u2E80-\u2FFF\u3000-\u303F\u3400-\u4DBF\u4E00-\u9FFF\uF900-\uFAFF]/u;
const LETTER = /\p{L}/u;

/**
 * The font key and direction for one grapheme. `null` is a neutral (space,
 * punctuation, digit) that keeps the script around it. Han characters use the
 * document language's CJK face, as document.module.css does.
 */
function scriptOf(grapheme: string, ctx: DocContext, role: FontRole): { key: string; rtl: boolean } | null {
  const base = ctx.lang?.split("-")[0];
  if (ARABIC.test(grapheme)) return { key: SCRIPT_FONT_KEYS.arabic, rtl: true };
  if (OTHER_RTL.test(grapheme)) return { key: role === "head" ? ctx.headKey : ctx.bodyKey, rtl: true };
  if (DEVANAGARI.test(grapheme)) return { key: SCRIPT_FONT_KEYS.devanagari, rtl: false };
  if (GREEK.test(grapheme)) {
    const serif = THEME_FONTS[(role === "head" ? ctx.headKey : ctx.bodyKey) as keyof typeof THEME_FONTS].serif;
    return { key: serif ? WORD_FONT_KEYS.GARAMOND : WORD_FONT_KEYS.PLEX_SANS, rtl: false };
  }
  if (KANA.test(grapheme)) return { key: SCRIPT_FONT_KEYS.japanese, rtl: false };
  if (HANGUL.test(grapheme)) return { key: SCRIPT_FONT_KEYS.korean, rtl: false };
  if (HAN.test(grapheme)) {
    return { key: base === "ja" ? SCRIPT_FONT_KEYS.japanese : base === "ko" ? SCRIPT_FONT_KEYS.korean : SCRIPT_FONT_KEYS.chinese, rtl: false };
  }
  if (LETTER.test(grapheme)) return { key: role === "head" ? ctx.headKey : ctx.bodyKey, rtl: false };
  return null;
}

/** Text split at script changes, one run per font. */
function scriptChunks(text: string, ctx: DocContext, role: FontRole): ScriptChunk[] {
  const chunks: ScriptChunk[] = [];
  let key = role === "head" ? ctx.headKey : ctx.bodyKey;
  let rtl = ctx.dir === "rtl";
  for (const { segment } of GRAPHEMES.segment(text)) {
    const script = scriptOf(segment, ctx, role);
    if (script) {
      key = script.key;
      rtl = script.rtl;
    }
    const last = chunks.at(-1);
    if (last && last.key === key && last.rtl === rtl) last.text += segment;
    else chunks.push({ text: segment, key, rtl });
  }
  return chunks;
}

/**
 * A string as Word runs. Newlines become line breaks inside one paragraph, and
 * each grapheme cluster takes its script's font. A character with no prepared
 * font for its script keeps the theme face rather than disappearing.
 */
export function runsFor(ctx: DocContext, text: string, style: RunStyle = {}): TextRun[] {
  const role = style.role ?? "body";
  const units = style.units ?? ctx.resolved.density.fontSize;
  const size = unitsToHalfPoints(units, ctx.paper);
  const scale = sizeScale(ctx, units);
  const color = style.color ?? TEXT;
  const runs: TextRun[] = [];
  text.split("\n").forEach((line, lineIndex) => {
    let first = true;
    for (const chunk of scriptChunks(line, ctx, role)) {
      const resolved = ctx.fonts.query(chunk.key, style.weight ?? 400, style.italic ?? false, style.optical ?? "text");
      const rtl = style.ltr ? false : chunk.rtl;
      const joined = chunk.rtl || chunk.key === SCRIPT_FONT_KEYS.devanagari;
      runs.push(
        new TextRun({
          text: chunk.text,
          break: first && lineIndex > 0 ? 1 : undefined,
          font: resolved.family,
          size,
          // Joined scripts keep their own proportions and spacing.
          scale: joined ? undefined : scale,
          // Pair kerning at every size, as the browser shapes the PDF (Word's
          // own default leaves it off, which sets letters further apart).
          kern: 2,
          color: hex(color),
          // Tracking pulls joined Arabic and Devanagari apart, as the CSS does.
          characterSpacing: joined ? 0 : style.characterSpacing,
          italics: resolved.italic || undefined,
          bold: resolved.bold || undefined,
          allCaps: style.allCaps || undefined,
          rightToLeft: style.ltr ? false : rtl || undefined,
          language: runLanguage(ctx, chunk.key, rtl),
        }),
      );
      first = false;
    }
    // An empty line is still a line break.
    if (line === "" && lineIndex > 0) {
      runs.push(new TextRun({ text: "", break: 1, size, font: ctx.fonts.face(role, 400).family }));
    }
  });
  return runs;
}

/** `items` as one line separated by "  ·  ", each kept left to right when it has no RTL letters. */
export function contactRuns(ctx: DocContext, items: readonly string[], style: RunStyle): TextRun[] {
  return items.flatMap((item, index) => [
    ...(index > 0 ? runsFor(ctx, CONTACT_SEPARATOR, style) : []),
    ...runsFor(ctx, item, { ...style, ltr: !hasRightToLeft(item) }),
  ]);
}

/** The sidebar's fixed width in twips, panel padding included for a tinted sidebar. */
export const sidebarWidth = (ctx: DocContext) => twipsIn(ctx.paper, SIDEBAR.width + (ctx.resolved.tintedSidebar ? 24 : 0));

/** The gap between a sidebar and the main column in twips. */
export const sidebarGap = (ctx: DocContext) => twipsIn(ctx.paper, SIDEBAR.gap);
