import type { Metadata } from "next";
import { notFound } from "next/navigation";

import { ResumeDocument } from "@/components/documents/resume-document";
import { demoProfile } from "@/lib/demo-profile";
import { verifyPrintAccess } from "@/lib/pdf/print-access";
import { loadSettings } from "@/lib/settings-data";
import { BUILT_IN_THEMES, printCss } from "@/lib/themes";

export const metadata: Metadata = { title: "Theme references sample", robots: { index: false } };

/**
 * One built-in theme's demo base resume with references forced on, so the
 * sample matrix can print the Chromium reference for a Word resume whose main
 * column ends in a nested references table. Development only.
 */
export default async function PrintThemeReferencesPage({ params }: PageProps<"/print/theme/[theme]/references">) {
  if (process.env.NODE_ENV === "production") notFound();
  const { theme: id } = await params;
  await verifyPrintAccess(`/print/theme/${id}/references`);
  const theme = BUILT_IN_THEMES.find((entry) => entry.id === id);
  if (!theme) notFound();
  const settings = await loadSettings();

  return (
    <>
      <style>{printCss(theme, settings.paper)}</style>
      <ResumeDocument
        theme={theme}
        profile={demoProfile()}
        paper={settings.paper}
        includeReferences
        language={settings.resumeLanguage}
      />
    </>
  );
}
