import type { Metadata } from "next";
import { notFound } from "next/navigation";

import { CoverLetterDocument } from "@/components/documents/cover-letter-document";
import { Tone } from "@/generated/prisma/enums";
import { demoLetterDate, demoLetterUnits, demoProfile } from "@/lib/demo-profile";
import { verifyPrintAccess } from "@/lib/pdf/print-access";
import { loadSettings } from "@/lib/settings-data";
import { BUILT_IN_THEMES, printCss } from "@/lib/themes";

export const metadata: Metadata = { title: "Theme letter sample", robots: { index: false } };

/**
 * One built-in theme's demo cover letter on its own, so the sample matrix can
 * print the Chromium reference for the Word letter (`/api/dev/word-samples`).
 * Development only.
 */
export default async function PrintThemeLetterPage({ params }: PageProps<"/print/theme/[theme]/letter">) {
  if (process.env.NODE_ENV === "production") notFound();
  const { theme: id } = await params;
  await verifyPrintAccess(`/print/theme/${id}/letter`);
  const theme = BUILT_IN_THEMES.find((entry) => entry.id === id);
  if (!theme) notFound();
  const settings = await loadSettings();
  const profile = demoProfile();

  return (
    <>
      <style>{printCss(theme, settings.paper)}</style>
      <CoverLetterDocument
        theme={theme}
        paper={settings.paper}
        basics={profile.basics}
        units={demoLetterUnits()}
        company="Ledgerly"
        location="Amsterdam"
        tone={Tone.WARM}
        date={demoLetterDate}
        dateFormat={settings.dateFormat}
        language={settings.resumeLanguage}
      />
    </>
  );
}
