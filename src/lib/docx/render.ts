import "server-only";

import { Document, Packer, type FileChild, type INumberingOptions } from "docx";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";

import type { PreparedLetter, PreparedResume } from "@/lib/document-prep";

import { embedFonts } from "./embed";
import { pageProperties } from "./layout";
import { letterChildren } from "./letter";
import { bulletNumbering } from "./numbering";
import { resumeChildren } from "./resume";
import { createContext, type DocContext } from "./runs";
import { documentStyles } from "./styles";

// The DOCX entry points. Everything is built from the prepared document and
// the resolved theme; no HTML or PDF is parsed, and every user string reaches
// Word as a TextRun, which escapes it as XML text rather than markup.

/**
 * docx 9.15.1 writes the font table, the obfuscated `.odttf` parts and
 * `w:embedTrueTypeFonts`, which Word uses without the fonts installed. Its
 * public Document options cannot set `w:saveSubsetFonts`, though: add that to
 * `word/settings.xml` so a save keeps every glyph, not only the ones this text
 * used, and editing new text still has the embedded coverage.
 */
function keepEmbeddedFonts(zip: Uint8Array): Uint8Array {
  const files = unzipSync(zip);
  const settings = files["word/settings.xml"];
  if (!settings) return zip;
  const xml = strFromU8(settings);
  const flags = [
    ...(xml.includes("w:embedTrueTypeFonts") ? [] : ["<w:embedTrueTypeFonts/>"]),
    ...(xml.includes("w:saveSubsetFonts") ? [] : ['<w:saveSubsetFonts w:val="false"/>']),
  ].join("");
  if (!flags) return zip;
  files["word/settings.xml"] = strToU8(xml.replace(/(<w:settings\b[^>]*>)/, `$1${flags}`));
  return zipSync(files);
}

/** A packed document with its fonts embedded and the save flags set. */
async function pack(ctx: DocContext, children: FileChild[], title: string, numbering?: INumberingOptions): Promise<Uint8Array> {
  const document = new Document({
    title,
    creator: "Dossier",
    description: "Exported from Dossier",
    styles: documentStyles(ctx),
    numbering,
    fonts: await embedFonts(ctx.fonts.faces()),
    sections: [{ properties: pageProperties(ctx), children }],
  });
  return keepEmbeddedFonts(new Uint8Array(await Packer.toBuffer(document)));
}

/** A resume as an editable Word document, with its fonts embedded. */
export async function renderResumeDocx(prepared: PreparedResume): Promise<Uint8Array> {
  const ctx = createContext({
    paper: prepared.paper,
    resolved: prepared.resolved,
    headFont: prepared.theme.headFont,
    bodyFont: prepared.theme.bodyFont,
    lang: prepared.lang,
    dir: prepared.dir,
  });
  return pack(ctx, await resumeChildren(ctx, prepared), "Resume", bulletNumbering(ctx));
}

/** A cover letter as an editable Word document, with its fonts embedded. */
export async function renderLetterDocx(prepared: PreparedLetter): Promise<Uint8Array> {
  const ctx = createContext({
    paper: prepared.paper,
    resolved: prepared.resolved,
    headFont: prepared.theme.headFont,
    bodyFont: prepared.theme.bodyFont,
    lang: prepared.lang,
    dir: prepared.dir,
  });
  return pack(ctx, letterChildren(ctx, prepared), "Cover letter");
}
