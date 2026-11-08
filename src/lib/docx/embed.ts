import "server-only";

import { readFile } from "node:fs/promises";
import path from "node:path";

import type { WordFace } from "./fonts";

// Font bytes stay in memory once read: the files never change between builds
// and a download should not touch the disk for each family. The promise map
// also keeps concurrent exports to one read.

const cache = new Map<string, Promise<Buffer>>();

function readFace(file: string): Promise<Buffer> {
  const cached = cache.get(file);
  if (cached) return cached;
  const reading = readFile(path.join(process.cwd(), "src", "fonts", file)).then((bytes) => Buffer.from(bytes));
  cache.set(file, reading);
  // A missing or unreadable file is worth retrying; don't cache the rejection.
  reading.catch(() => cache.delete(file));
  return reading;
}

export type EmbeddedFont = { name: string; data: Buffer };

/** The font-table entries for the faces a document uses. */
export async function embedFonts(faces: readonly WordFace[]): Promise<EmbeddedFont[]> {
  const unique = new Map<string, WordFace>();
  for (const face of faces) unique.set(face.family, face);
  return Promise.all([...unique.values()].map(async (face) => ({ name: face.family, data: await readFace(face.file) })));
}

/** For tests and diagnostics: how many bytes the embedded faces take. */
export async function embeddedFontBytes(faces: readonly WordFace[]): Promise<number> {
  const embedded = await embedFonts(faces);
  return embedded.reduce((total, font) => total + font.data.length, 0);
}
