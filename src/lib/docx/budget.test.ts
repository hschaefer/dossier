import { Unzip } from "fflate";
import { describe, expect, it } from "vitest";

import { demoProfile } from "@/lib/demo-profile";
import { prepareLetter, prepareResume } from "@/lib/document-prep";
import { BUILT_IN_THEMES } from "@/lib/themes";

import { renderLetterDocx, renderResumeDocx } from "./render";

// The size budget is incremental embedded-font overhead, not the whole
// package. ZIP entries keep their compressed size, so this reads it without
// unpacking the fonts.

const theme = (id: string) => BUILT_IN_THEMES.find((entry) => entry.id === id)!;
const today = new Date(2026, 8, 28);

function fontOverhead(zip: Uint8Array): { bytes: number; files: string[] } {
  let bytes = 0;
  const files: string[] = [];
  const unzip = new Unzip((file) => {
    if (file.name.startsWith("word/fonts/") && file.name.endsWith(".odttf")) {
      bytes += file.size ?? 0;
      files.push(file.name);
    }
  });
  unzip.push(zip, true);
  return { bytes, files };
}

async function measure(zip: Uint8Array) {
  const { bytes, files } = fontOverhead(zip);
  return { total: zip.length, fonts: bytes, files: files.length };
}

describe("Word export size", () => {
  it("keeps a single-column Editorial resume and letter within the 1 MB font budget", async () => {
    const profile = demoProfile();
    const resume = await measure(
      await renderResumeDocx(prepareResume({ theme: theme("editorial"), profile, paper: "A4", today, includeReferences: true })),
    );
    const letter = await measure(
      await renderLetterDocx(
        prepareLetter({
          theme: theme("editorial"),
          paper: "A4",
          basics: profile.basics,
          units: [],
          company: "Ledgerly",
          location: "Utrecht",
          tone: "WARM",
          date: today,
          dateFormat: "DAY_MONTH_YEAR",
        }),
      ),
    );
    console.log("Word sizes", { resume, letter });
    expect(resume.fonts).toBeLessThan(1_000_000);
    expect(letter.fonts).toBeLessThan(1_000_000);
  });

  it("keeps a tinted-sidebar Studio resume within the 1 MB font budget", async () => {
    const profile = demoProfile();
    const studio = await measure(
      await renderResumeDocx(prepareResume({ theme: theme("studio"), profile, paper: "A4", today, includeReferences: true })),
    );
    console.log("Word sizes", { studio });
    expect(studio.fonts).toBeLessThan(1_000_000);
  });

  it("reports, but does not fail on, a CJK fixture that needs a large face", async () => {
    const profile = { ...demoProfile(), basics: { ...demoProfile().basics, name: "測試者", summary: "テスト 测试 테스트." } };
    const cjk = await measure(
      await renderResumeDocx(prepareResume({ theme: theme("editorial"), profile, paper: "A4", today })),
    );
    console.log("Word sizes", { cjk });
    expect(cjk.files).toBeGreaterThan(0);
  });
});
