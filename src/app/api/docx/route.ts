import { NextResponse } from "next/server";

import { DocumentKind, type ThemeLayout } from "@/generated/prisma/enums";
import { loadDocuments } from "@/lib/application-documents";
import { prepareLetter, prepareResume } from "@/lib/document-prep";
import { docxDownload } from "@/lib/docx/download";
import { renderLetterDocx, renderResumeDocx } from "@/lib/docx/render";
import { resolveFileName } from "@/lib/file-name";
import { loadProfile, loadProfileName } from "@/lib/profile-data";
import { loadSettings } from "@/lib/settings-data";
import { loadTheme } from "@/lib/theme-data";
import { BUILT_IN_THEMES, withoutSidebar } from "@/lib/themes";
import { route } from "@/lib/route";

// Editable Word downloads (docs/word-export.md): the same prepared
// documents the PDF prints, drawn by the native DOCX renderer instead of
// Chromium. The file goes straight to the browser without a download cache.

/** The documents an application can download, by their `doc=` value. */
const TAILORED = { resume: DocumentKind.RESUME, letter: DocumentKind.COVER_LETTER } as const;

/** Application ids are cuids; anything else never reaches a database query. */
const APPLICATION_ID = /^[a-z0-9]{1,64}$/i;

const notFound = (error: string) => NextResponse.json({ error }, { status: 404 });

/**
 * Downloads a document as an editable .docx: `?doc=base` for the base resume,
 * `?doc=resume&application={id}` or `?doc=letter&application={id}` for an
 * application's tailored resume or cover letter. The file is named with
 * Settings → Document file name.
 */
export const GET = route(async (request) => {
  const doc = request.nextUrl.searchParams.get("doc");
  const [name, settings] = await Promise.all([loadProfileName(), loadSettings()]);
  const date = new Date();

  if (doc === "base") {
    const [profile, theme] = await Promise.all([loadProfile(), loadTheme(settings.baseThemeId)]);
    const prepared = prepareResume({
      theme: theme ?? BUILT_IN_THEMES[0],
      profile,
      paper: settings.paper,
      language: settings.resumeLanguage,
      includeReferences: settings.includeRefs,
      sections: settings.resumeSections,
      // PDF fitting measures Chromium pages; editable Word uses normal spacing.
      fit: 0,
    });
    return docxDownload(
      await renderResumeDocx(prepared),
      resolveFileName(settings.fileNamePattern, { name, doc: DocumentKind.RESUME, date }, "docx"),
    );
  }

  if (doc !== "resume" && doc !== "letter") return notFound("There's no such document to download.");
  const id = request.nextUrl.searchParams.get("application") ?? "";
  if (!APPLICATION_ID.test(id)) return notFound("That application no longer exists.");
  let loaded: Awaited<ReturnType<typeof loadDocuments>>;
  try {
    loaded = await loadDocuments(id);
  } catch (error) {
    // The shared page loader calls Next's notFound(). Keep this file API's
    // JSON 404 response, without swallowing other loader failures.
    if (error instanceof Error && "digest" in error && error.digest === "NEXT_HTTP_ERROR_FALLBACK;404") {
      return notFound("That application no longer exists.");
    }
    throw error;
  }
  const { application, profile, settings: documentSettings, resume, letter, includeReferences, sections, letterDate } = loaded;
  const document = doc === "resume" ? resume : letter;
  if (!document) {
    return notFound(doc === "resume" ? "This application has no tailored resume." : "This application has no cover letter.");
  }

  const { theme, units } = document;
  const fileName = resolveFileName(
    settings.fileNamePattern,
    { name, company: application.company, role: application.role, doc: TAILORED[doc], date },
    "docx",
  );

  if (doc === "letter") {
    const prepared = prepareLetter({
      theme,
      paper: documentSettings.paper,
      basics: profile.basics,
      units,
      company: application.company,
      location: application.location,
      tone: application.tone,
      date: letterDate,
      dateFormat: documentSettings.dateFormat,
      language: application.language,
    });
    return docxDownload(await renderLetterDocx(prepared), fileName);
  }

  const prepared = prepareResume({
    theme,
    profile,
    paper: documentSettings.paper,
    units,
    layout: (application.followThemeLayout ? undefined : withoutSidebar(theme.layout)) as ThemeLayout | undefined,
    showPhoto: application.showPhoto,
    language: application.language,
    includeReferences,
    sections,
    fit: 0,
  });
  return docxDownload(await renderResumeDocx(prepared), fileName);
});
