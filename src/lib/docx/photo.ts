import { AlignmentType, ImageRun, Paragraph } from "docx";

import { PhotoPosition, PhotoShape } from "@/generated/prisma/enums";
import { isPhotoDataUrl } from "@/lib/photo";
import { PHOTO_SHAPES } from "@/lib/themes";

import { unitsToPixels, unitsToPoints, unitsToTwips } from "./geometry";
import { bidi, twipsIn, type DocContext } from "./runs";

// The profile photo. The stored image is always 3:4; square shapes show the
// middle of it through ImageRun's crop percentages instead of stretching the
// face. Word can only clip a picture to a rectangle, so a circle or rounded
// square is baked into the PNG's alpha channel here, which Word draws as
// transparent corners over the banner or the paper.

export type PhotoPlace = "header" | "sidebar";

const TYPES: Record<string, "jpg" | "png"> = { jpeg: "jpg", png: "png" };

/** The corner radius of the rounded shape, in page units (document.module.css). */
const ROUNDED_RADIUS = 12;

type Sharp = typeof import("sharp").default;

let sharpModule: Promise<Sharp | null> | undefined;

/**
 * Load sharp for masks and WebP conversion. If unavailable, JPEG/PNG photos
 * keep Word's rectangular crop and unconvertible WebP photos are omitted.
 */
function loadSharp(): Promise<Sharp | null> {
  sharpModule ??= import("sharp").then(
    (module) => module.default,
    () => null,
  );
  return sharpModule;
}

/**
 * The photo as Word draws it. The circle and rounded shapes are clipped here,
 * because Word cannot; those come out as a square PNG with transparent corners.
 */
export async function photoImage(ctx: DocContext, src: string | null, name: string, place: PhotoPlace): Promise<ImageRun | null> {
  if (!src || !isPhotoDataUrl(src)) return null;
  const match = /^data:image\/(jpeg|png|webp);base64,(.+)$/.exec(src);
  if (!match) return null;

  const { shape } = ctx.resolved.photo;
  const { width, height } = PHOTO_SHAPES[shape][place];
  // The source is 3:4. Trimming 12.5% off the top and bottom of a 3:4 image
  // leaves a square, the middle of the crop the preview shows.
  const square = shape !== PhotoShape.PORTRAIT;
  const clipped = shape === PhotoShape.CIRCLE || shape === PhotoShape.ROUNDED;

  let data: Uint8Array = Buffer.from(match[2], "base64");
  let type = TYPES[match[1]];
  let masked = false;
  const sharp = clipped || match[1] === "webp" ? await loadSharp() : null;
  if (sharp) {
    try {
      if (match[1] === "webp") {
        data = await sharp(data).png().toBuffer();
        type = "png";
      }
      if (clipped) {
        // Enough pixels for 300 dpi on paper, so a masked PNG stays small.
        const printWidth = Math.round((unitsToPoints(width, ctx.paper) / 72) * 300);
        data = await maskedPhoto(sharp, data, shape === PhotoShape.CIRCLE ? "circle" : "rounded", ROUNDED_RADIUS / width, printWidth);
        type = "png";
        masked = true;
      }
    } catch {
      // Keep a Word-compatible image if masking fails. Failed WebP conversion
      // leaves no compatible type and is omitted below.
    }
  }
  // JPEG and PNG are the formats Word's drawing ML accepts directly; WebP is
  // not, and Word prompts to repair a package that claims one. Leave it out
  // rather than crash or write a misleading image part.
  if (!type) return null;

  return new ImageRun({
    type,
    data,
    transformation: { width: unitsToPixels(width, ctx.paper), height: unitsToPixels(height, ctx.paper) },
    // A masked image is already the square crop; the others still crop in Word.
    crop: masked || !square ? undefined : { top: 12.5, bottom: 12.5 },
    altText: { name: "Profile photo", description: name, title: name },
  });
}

/**
 * The middle square of the photo, clipped to a circle or a rounded square and
 * returned as a PNG with transparency where the shape cuts it away. It is
 * scaled to `target` pixels across (print resolution) to keep the package
 * small: a photo as PNG is much larger than its JPEG source.
 */
async function maskedPhoto(
  sharp: Sharp,
  image: Uint8Array,
  shape: "circle" | "rounded",
  radius: number,
  target: number,
): Promise<Uint8Array> {
  const { width = 0, height = 0 } = await sharp(image).metadata();
  if (!width || !height) throw new Error("The photo has no readable size.");
  const side = Math.min(width, height);
  const size = Math.max(1, Math.min(side, target));
  const square = await sharp(image)
    .extract({
      left: Math.round((width - side) / 2),
      top: Math.round((height - side) / 2),
      width: side,
      height: side,
    })
    .resize(size, size, { fit: "cover" })
    .png()
    .toBuffer();
  const corner = shape === "circle" ? size / 2 : Math.round(radius * size);
  const mask = Buffer.from(
    `<svg xmlns="http://www.w3.org/2000/svg" width="${size}" height="${size}">` +
      `<rect width="${size}" height="${size}" rx="${corner}" ry="${corner}" fill="#fff"/></svg>`,
  );
  return await sharp(square).composite([{ input: mask, blend: "dest-in" }]).png({ compressionLevel: 9 }).toBuffer();
}

/** A photo above the name follows the header's alignment; a sidebar photo stays at its start. */
export function photoParagraph(ctx: DocContext, image: ImageRun, after = 8): Paragraph {
  return new Paragraph({
    alignment: ctx.resolved.photo.position === PhotoPosition.ABOVE_NAME && ctx.resolved.centered ? AlignmentType.CENTER : undefined,
    bidirectional: bidi(ctx),
    spacing: { before: 0, after: twipsIn(ctx.paper, after) },
    children: [image],
  });
}

/** The width a photo needs in a header's two-column table, in twips. */
export function photoWidth(ctx: DocContext, place: PhotoPlace): number {
  const { width } = PHOTO_SHAPES[ctx.resolved.photo.shape][place];
  return unitsToTwips(width, ctx.paper);
}
