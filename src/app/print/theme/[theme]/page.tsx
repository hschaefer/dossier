import type { Metadata } from "next";
import { notFound } from "next/navigation";

import { ResumeDocument } from "@/components/documents/resume-document";
import { verifyPrintAccess } from "@/lib/pdf/print-access";
import { loadProfile } from "@/lib/profile-data";
import { loadSettings } from "@/lib/settings-data";
import { BUILT_IN_THEMES, printCss } from "@/lib/themes";

export const metadata: Metadata = { title: "Theme sample", robots: { index: false } };

/**
 * One built-in theme's base resume on its own, so the sample matrix can print
 * the Chromium reference for every theme (`/api/dev/word-samples`). Development
 * only: the app's own documents print from /print/base and /print/{doc}/{id}.
 */
export default async function PrintThemeSamplePage({ params }: PageProps<"/print/theme/[theme]">) {
  if (process.env.NODE_ENV === "production") notFound();
  const { theme: id } = await params;
  await verifyPrintAccess(`/print/theme/${id}`);
  const theme = BUILT_IN_THEMES.find((entry) => entry.id === id);
  if (!theme) notFound();
  const [profile, settings] = await Promise.all([loadProfile(), loadSettings()]);

  return (
    <>
      <style>{printCss(theme, settings.paper)}</style>
      <ResumeDocument
        theme={theme}
        profile={profile}
        paper={settings.paper}
        includeReferences={settings.includeRefs}
        sections={settings.resumeSections}
        language={settings.resumeLanguage}
      />
    </>
  );
}
