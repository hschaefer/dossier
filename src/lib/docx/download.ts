import "server-only";

import { contentDisposition } from "@/lib/file-name";

/** Word's media type for a .docx package. */
export const DOCX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

/** A generated DOCX as a download named `fileName`. */
export function docxDownload(docx: Uint8Array, fileName: string): Response {
  return new Response(Buffer.from(docx), {
    headers: {
      "content-type": DOCX_CONTENT_TYPE,
      "content-disposition": contentDisposition(fileName),
    },
  });
}
