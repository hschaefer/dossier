"use client";

import { useRef, useState, type ReactNode } from "react";

import { toast } from "@/components/toast";
import { buttonClassName, type ButtonSize, type ButtonVariant } from "@/components/ui/button";
import { messages } from "@/lib/messages";
import { downloadFrom, saveFile } from "@/lib/save-file";

/** The file name a download response asks for, preferring the UTF-8 one (RFC 6266). */
export function fileNameFrom(disposition: string | null): string | null {
  if (!disposition) return null;
  const encoded = /filename\*=UTF-8''([^;]+)/i.exec(disposition)?.[1];
  if (encoded) {
    try {
      return decodeURIComponent(encoded);
    } catch {
      // Fall back to the plain name.
    }
  }
  return /filename="([^"]*)"/i.exec(disposition)?.[1] ?? null;
}

/**
 * The download behind every export control: fetch `href`, show a progress
 * toast and a failure toast, then save the file. A held file (as /api/pdf
 * renders one) comes back as a link; any other file is fetched and saved from
 * the page. `busy` keeps a second click from starting another export.
 */
export function useFileDownload({
  href,
  held,
  fallbackName,
  progress,
}: {
  href: string;
  /** Whether `href` answers a POST by holding the file on the server first, as /api/pdf does. */
  held: boolean;
  /** The name to save under when the server gives none. */
  fallbackName: string;
  /** The toast while it prepares, e.g. "Rendering your cover letter as a PDF…". */
  progress: string;
}) {
  const [busy, setBusy] = useState(false);
  const running = useRef(false);

  const download = async () => {
    if (running.current) return;
    running.current = true;
    setBusy(true);
    toast(progress);
    try {
      const response = await fetch(href, held ? { method: "POST" } : undefined);
      if (!response.ok) {
        const body = (await response.json().catch(() => null)) as { error?: string } | null;
        toast(body?.error ?? messages.unexpected);
        return;
      }
      if (held) {
        const { url } = (await response.json()) as { url: string };
        downloadFrom(url);
      } else {
        saveFile(await response.blob(), fileNameFrom(response.headers.get("content-disposition")) ?? fallbackName);
      }
    } catch {
      toast(messages.offline);
    } finally {
      running.current = false;
      setBusy(false);
    }
  };

  return { busy, download };
}

/**
 * Downloads a file from `href`, holding a moment's work behind a progress
 * toast and showing a failure in a toast instead of leaving the page. A held
 * file (as /api/pdf renders one) comes back as a link; any other file is
 * fetched and saved from the page. Without JavaScript it's a plain link.
 */
export function DownloadFileButton({
  href,
  held,
  fallbackName,
  busyLabel,
  progress,
  variant = "primary",
  size = "md",
  children,
}: {
  href: string;
  /**
   * Whether `href` answers a POST by holding the file on the server and
   * sending the link to download it from, as /api/pdf does. Browsers save
   * that reliably; otherwise the file is fetched and saved from the page.
   */
  held: boolean;
  /** The name to save under when the server gives none. */
  fallbackName: string;
  /** What the link says while the file is on its way. */
  busyLabel: string;
  /** The toast while it prepares, e.g. "Rendering your cover letter as a PDF…". */
  progress: string;
  variant?: ButtonVariant;
  size?: ButtonSize;
  children: ReactNode;
}) {
  const { busy, download } = useFileDownload({ href, held, fallbackName, progress });

  return (
    <a
      href={href}
      aria-disabled={busy || undefined}
      onClick={(event) => {
        event.preventDefault();
        void download();
      }}
      className={buttonClassName(variant, size, variant === "primary" ? "hover:text-on-primary" : undefined)}
    >
      {busy ? busyLabel : children}
    </a>
  );
}

/**
 * Downloads a PDF from /api/pdf. Rendering takes a moment, so it says so,
 * and a failure shows up as a toast instead of an error page.
 */
export function DownloadPdfButton({
  href,
  held = true,
  variant = "primary",
  size = "md",
  progress = "Rendering your PDF on the server. It'll download in a moment.",
  children,
}: {
  href: string;
  /** Whether `href` holds the file on the server first; see DownloadFileButton. */
  held?: boolean;
  variant?: ButtonVariant;
  size?: ButtonSize;
  /** The toast while it renders, e.g. "Rendering your cover letter as a PDF…". */
  progress?: string;
  children: ReactNode;
}) {
  return (
    <DownloadFileButton
      href={href}
      held={held}
      fallbackName="document.pdf"
      busyLabel="Rendering…"
      progress={progress}
      variant={variant}
      size={size}
    >
      {children}
    </DownloadFileButton>
  );
}
