"use client";

import { ChevronDown } from "lucide-react";
import { useRef, type ReactNode } from "react";

import { buttonClassName, type ButtonSize, type ButtonVariant } from "@/components/ui/button";
import {
  DropdownMenu,
  DropdownMenuContent,
  DropdownMenuItem,
  DropdownMenuLabel,
  DropdownMenuTrigger,
} from "@/components/ui/dropdown-menu";

import { useFileDownload } from "./download-pdf-button";

/** Hover help and the menu's explanation of what a Word download changes. */
export const WORD_EXPORT_NOTE =
  "Editable Word document with embedded fonts. Line and page breaks may differ from the PDF. PDF fit-to-one-page does not apply to Word.";

/**
 * Download a document as a split button: the main part exports the PDF, the
 * arrow opens every format the document offers. The PDF stays a plain link,
 * so it still downloads without JavaScript; the menu needs JavaScript like
 * the other export controls.
 */
export function DownloadSplitButton({
  pdfHref,
  wordHref,
  pdfProgress = "Rendering your PDF on the server. It'll download in a moment.",
  wordProgress = "Preparing your Word document…",
  variant = "primary",
  size = "md",
  children = "Download PDF",
}: {
  pdfHref: string;
  wordHref: string;
  /** The toast while the PDF renders, e.g. "Rendering your cover letter as a PDF…". */
  pdfProgress?: string;
  /** The toast while the Word document prepares, e.g. "Preparing your cover letter as a Word document…". */
  wordProgress?: string;
  variant?: ButtonVariant;
  size?: ButtonSize;
  /** The main button's label; it exports the PDF. */
  children?: ReactNode;
}) {
  const pdf = useFileDownload({ href: pdfHref, held: true, fallbackName: "document.pdf", progress: pdfProgress });
  const word = useFileDownload({ href: wordHref, held: false, fallbackName: "document.docx", progress: wordProgress });
  const busy = pdf.busy || word.busy;
  // A synchronous lock also covers clicks before React renders the busy state.
  const running = useRef(false);
  const download = async (format: "pdf" | "word") => {
    if (running.current) return;
    running.current = true;
    try {
      await (format === "pdf" ? pdf : word).download();
    } finally {
      running.current = false;
    }
  };
  const split = variant === "secondary";
  const mainClassName = split
    ? "rounded-none border-r-0 shadow-none! focus-visible:-outline-offset-2"
    : "rounded-none shadow-none! hover:text-on-primary focus-visible:-outline-offset-2";
  const triggerClassName = split
    ? "rounded-none px-2 shadow-none! focus-visible:-outline-offset-2"
    : "rounded-none border-l border-on-primary/25 px-2 shadow-none! focus-visible:-outline-offset-2";

  return (
    <div className={`inline-flex overflow-hidden rounded-control ${split ? "shadow-control" : "shadow-cta"}`}>
      <a
        href={pdfHref}
        aria-disabled={busy || undefined}
        onClick={(event) => {
          event.preventDefault();
          void download("pdf");
        }}
        className={buttonClassName(variant, size, mainClassName)}
      >
        {pdf.busy ? "Rendering…" : word.busy ? "Preparing…" : children}
      </a>
      <DropdownMenu>
        <DropdownMenuTrigger asChild>
          <button type="button" aria-label="More export options" className={buttonClassName(variant, size, triggerClassName)}>
            <ChevronDown size={16} aria-hidden="true" />
          </button>
        </DropdownMenuTrigger>
        <DropdownMenuContent align="end" className="min-w-[17rem]">
          <DropdownMenuLabel>Export</DropdownMenuLabel>
          <DropdownMenuItem disabled={busy} onSelect={() => void download("pdf")}>
            <span className="flex flex-col gap-0.5">
              <span>PDF (.pdf)</span>
              <span className="text-xs text-muted">The document as previewed</span>
            </span>
          </DropdownMenuItem>
          <DropdownMenuItem disabled={busy} onSelect={() => void download("word")} title={WORD_EXPORT_NOTE}>
            <span className="flex flex-col gap-0.5">
              <span>Word (.docx)</span>
              <span className="text-xs text-muted">{WORD_EXPORT_NOTE}</span>
            </span>
          </DropdownMenuItem>
        </DropdownMenuContent>
      </DropdownMenu>
    </div>
  );
}
