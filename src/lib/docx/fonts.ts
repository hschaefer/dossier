import type { ThemeFont } from "@/generated/prisma/enums";

import manifest from "@/fonts/word/manifest.json";

// The static TrueType faces Word embeds, generated from the bundled WOFF2
// fonts by scripts/word-fonts.py. Word has no CSS font stack or variable
// weights: every weight is a family of its own, and the face choices here
// mirror resolveTheme's head/body fonts and the script fallbacks.

export type WordFace = {
  key: string;
  family: string;
  weight: number;
  style: "normal" | "italic";
  optical: "text";
  file: string;
};

export type FontRole = "head" | "body";

export type FontQuery = {
  key: string;
  weight: number;
  style?: "normal" | "italic";
  /** Renderer-role compatibility: both use the static source's default axes. */
  optical?: "text" | "display";
};

const FACES = (manifest as { faces: WordFace[] }).faces;

/** The face files by query, so a missing asset is a test failure rather than a surprise at download time. */
const BY_KEY = new Map<string, WordFace[]>();
for (const face of FACES) {
  const list = BY_KEY.get(face.key) ?? [];
  list.push(face);
  BY_KEY.set(face.key, list);
}

/** The manifest key for each theme font, as scripts/word-fonts.py writes it. */
export const WORD_FONT_KEYS: Record<ThemeFont, string> = {
  NEWSREADER: "NEWSREADER",
  GARAMOND: "GARAMOND",
  SOURCE_SERIF: "SOURCE_SERIF",
  LORA: "LORA",
  LIBRE_CASLON: "LIBRE_CASLON",
  PLAYFAIR: "PLAYFAIR",
  FRAUNCES: "FRAUNCES",
  GEIST: "GEIST",
  PLEX_SANS: "PLEX_SANS",
  INTER: "INTER",
  MANROPE: "MANROPE",
  WORK_SANS: "WORK_SANS",
  SPACE_GROTESK: "SPACE_GROTESK",
  PLEX_MONO: "PLEX_MONO",
  JETBRAINS_MONO: "JETBRAINS_MONO",
};

/** The script fallbacks, as in src/fonts/index.ts. */
export const SCRIPT_FONT_KEYS = {
  arabic: "ARABIC",
  devanagari: "DEVANAGARI",
  japanese: "JAPANESE",
  korean: "KOREAN",
  chinese: "CHINESE",
} as const;

/** The bundles in the order a document's font stack lists them. */
export const ALL_FONT_KEYS = [...Object.values(WORD_FONT_KEYS), ...Object.values(SCRIPT_FONT_KEYS)];

/** The Newsreader roman/italic pair: an italic face under its own family, so Word draws true italics. */
const ITALIC_KEYS: Record<string, string> = { NEWSREADER: "NEWSREADER_ITALIC" };

const nearest = (values: number[], wanted: number) =>
  values.reduce((best, value) => (Math.abs(value - wanted) < Math.abs(best - wanted) ? value : best), values[0]);

/**
 * The nearest face of `key` for a weight and style. Fonts without an italic
 * asset return their regular face: the caller adds Word's italic flag, as a
 * browser synthesises it.
 */
export function resolveFace(query: FontQuery): { face: WordFace; syntheticItalic: boolean; syntheticBold: boolean } {
  const wantedStyle = query.style ?? "normal";
  const italic = wantedStyle === "italic";
  const key = (italic && ITALIC_KEYS[query.key]) || query.key;
  const faces = BY_KEY.get(key);
  if (!faces?.length) throw new Error(`No Word font is prepared for ${query.key}.`);

  const styled = faces.filter((face) => face.style === (italic && ITALIC_KEYS[query.key] ? "italic" : "normal"));
  const pool = styled.length > 0 ? styled : faces.filter((face) => face.style === "normal");
  const weights = [...new Set(pool.map((face) => face.weight))];
  const weight = nearest(weights, query.weight);
  // The browser's static faces already pin all non-weight axes at defaults.
  // Keep the renderer's optical role input, but never invent a display variant.
  const face = pool.find((face) => face.weight === weight)!;
  return {
    face,
    syntheticItalic: italic && face.style !== "italic",
    // A family with fewer faces (CJK, Plex Mono) says the nearest
    // weight; bold it rather than silently rendering 600 as regular.
    syntheticBold: query.weight - face.weight >= 200,
  };
}

/** The family name a run names in DOCX, and whether Word should slant or embolden it. */
export function wordFont(query: FontQuery): { family: string; italic: boolean; bold: boolean; face: WordFace } {
  const { face, syntheticItalic, syntheticBold } = resolveFace(query);
  return { family: face.family, italic: syntheticItalic, bold: syntheticBold, face };
}

/** The faces a document needs, for embedding: every family a run can name. */
export function facesFor(queries: readonly FontQuery[]): WordFace[] {
  const faces = new Map<string, WordFace>();
  for (const query of queries) {
    const { face } = resolveFace(query);
    faces.set(`${face.family}:${face.file}`, face);
  }
  return [...faces.values()];
}

/** Every face in the manifest, for checks and documentation. */
export const WORD_FACES = FACES;
