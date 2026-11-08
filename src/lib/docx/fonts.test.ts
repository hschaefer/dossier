import { createHash } from "node:crypto";
import { existsSync, readFileSync } from "node:fs";
import path from "node:path";
import { describe, expect, it } from "vitest";

import { ThemeFont } from "@/generated/prisma/enums";
import manifest from "@/fonts/word/manifest.json";
import nextConfig from "../../../next.config";

import {
  ALL_FONT_KEYS,
  SCRIPT_FONT_KEYS,
  WORD_FACES,
  WORD_FONT_KEYS,
  facesFor,
  resolveFace,
  wordFont,
} from "./fonts";

const fontsRoot = path.join(process.cwd(), "src", "fonts");

describe("Word font manifest", () => {
  it("has a prepared family for every theme font and script fallback", () => {
    for (const font of Object.values(ThemeFont)) expect(WORD_FONT_KEYS[font]).toBeTruthy();
    expect(new Set(ALL_FONT_KEYS).size).toBe(ALL_FONT_KEYS.length);
    for (const key of ALL_FONT_KEYS) {
      expect(WORD_FACES.some((face) => face.key === key)).toBe(true);
    }
  });

  it("ships every face the manifest names", () => {
    for (const face of WORD_FACES) {
      expect(existsSync(path.join(fontsRoot, face.file)), face.file).toBe(true);
    }
  });

  it("records the exact current static sources and generated assets", () => {
    const browserFonts = readFileSync(path.join(fontsRoot, "index.ts"), "utf8");
    for (const face of manifest.faces) {
      expect(face.source).not.toContain("Variable");
      expect(browserFonts).toContain(`./${face.source}`);
      expect(face.optical).toBe("text");
      expect(createHash("sha256").update(readFileSync(path.join(fontsRoot, face.source))).digest("hex")).toBe(face.sourceSha256);
      const generated = readFileSync(path.join(fontsRoot, face.file));
      expect(createHash("sha256").update(generated).digest("hex")).toBe(face.sha256);
      // sfnt TrueType, not a WOFF2 file renamed with a .ttf extension.
      expect(generated.readUInt32BE(0)).toBe(0x00010000);
      const tables = Array.from({ length: generated.readUInt16BE(4) }, (_, i) => generated.toString("ascii", 12 + i * 16, 16 + i * 16));
      expect(tables).toContain("glyf");
      expect(tables).not.toContain("fvar");
      expect(face.embedding).toBe("editable");
    }
  });

  it("traces Word assets alongside upstream's playwright runtime files", () => {
    expect(nextConfig.outputFileTracingIncludes?.["/**"]).toEqual(expect.arrayContaining([
      "./src/fonts/word/**/*", "./node_modules/playwright-core/**/*",
    ]));
  });

  it("keeps weights separate instead of collapsing 500 and 600 into 400 or 700", () => {
    const regular = wordFont({ key: "GEIST", weight: 400 }).family;
    const medium = wordFont({ key: "GEIST", weight: 500 }).family;
    const semibold = wordFont({ key: "GEIST", weight: 600 }).family;
    const bold = wordFont({ key: "GEIST", weight: 700 }).family;
    expect(new Set([regular, medium, semibold, bold]).size).toBe(4);
    expect(medium).toContain("Medium");
    expect(semibold).toContain("SemiBold");
    expect(bold).toContain("Bold");
  });

  it("uses the nearest available weight when a family has fewer faces", () => {
    expect(resolveFace({ key: "PLEX_MONO", weight: 600 }).face.weight).toBe(600);
    expect(resolveFace({ key: "CHINESE", weight: 700 }).face.weight).toBe(600);
    // A simpler face is emboldened by Word rather than left looking regular.
    expect(wordFont({ key: "CHINESE", weight: 600 }).bold).toBe(false);
    expect(wordFont({ key: "PLEX_MONO", weight: 700 }).bold).toBe(false);
    expect(wordFont({ key: "GEIST", weight: 600 }).bold).toBe(false);
  });

  it("uses the browser's default axes for text and display roles", () => {
    const text = resolveFace({ key: "NEWSREADER", weight: 500, optical: "text" }).face;
    const display = resolveFace({ key: "NEWSREADER", weight: 500, optical: "display" }).face;
    expect(text.optical).toBe("text");
    expect(display).toBe(text);
  });

  it("uses Newsreader's real italic and slants other families synthetically", () => {
    const newsreader = wordFont({ key: "NEWSREADER", weight: 400, style: "italic" });
    expect(newsreader.family).toContain("Italic");
    expect(newsreader.italic).toBe(false);
    const geist = wordFont({ key: "GEIST", weight: 400, style: "italic" });
    expect(geist.family).toBe("Geist");
    expect(geist.italic).toBe(true);
  });

  it("does not use an OFL Reserved Font Name in a derived family", () => {
    const reserved = /\b(Plex|Lora|Playfair Display|Source)\b/i;
    for (const face of WORD_FACES) {
      expect(face.family, face.family).not.toMatch(reserved);
    }
    expect(WORD_FONT_KEYS.PLEX_SANS).toBe("PLEX_SANS");
    expect(wordFont({ key: WORD_FONT_KEYS.PLEX_SANS, weight: 400 }).family).toBe("Dossier Grotesk");
    expect(wordFont({ key: WORD_FONT_KEYS.LORA, weight: 400 }).family).toBe("Dossier Literary");
    expect(wordFont({ key: WORD_FONT_KEYS.PLAYFAIR, weight: 400 }).family).toBe("Dossier Display");
  });

  it("returns one face per family for embedding", () => {
    const faces = facesFor([
      { key: "GEIST", weight: 400 },
      { key: "GEIST", weight: 400, optical: "text" },
      { key: SCRIPT_FONT_KEYS.arabic, weight: 400 },
    ]);
    expect(faces).toHaveLength(2);
    expect(new Set(faces.map((face) => face.family)).size).toBe(2);
  });
});
