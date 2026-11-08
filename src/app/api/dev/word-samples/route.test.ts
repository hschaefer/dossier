import { strFromU8, unzipSync } from "fflate";
import { NextRequest } from "next/server";
import { beforeEach, describe, expect, it, vi } from "vitest";

import { getSession } from "@/lib/auth/session";
import { renderLetterDocx, renderResumeDocx } from "@/lib/docx/render";
import { renderPdf } from "@/lib/pdf/render";
import { loadProfile } from "@/lib/profile-data";
import { loadSettings } from "@/lib/settings-data";
import { DEFAULT_SECTION_LAYOUT } from "@/lib/resume-sections";
import { BUILT_IN_THEMES } from "@/lib/themes";

import { GET } from "./route";

vi.mock("@/lib/auth/session", () => ({ getSession: vi.fn() }));
vi.mock("@/lib/docx/render", () => ({
  renderResumeDocx: vi.fn(async () => new Uint8Array([80, 75, 3, 4])),
  renderLetterDocx: vi.fn(async () => new Uint8Array([80, 75, 3, 4, 76])),
}));
vi.mock("@/lib/pdf/render", () => ({ renderPdf: vi.fn(async () => new Uint8Array([37, 80, 68, 70])) }));
vi.mock("@/lib/profile-data", () => ({ loadProfile: vi.fn() }));
vi.mock("@/lib/settings-data", () => ({ loadSettings: vi.fn() }));

const profile = {
  basics: {
    name: "Jane Doe",
    headline: "",
    email: "",
    phone: "",
    dob: "",
    website: "",
    address: "",
    city: "",
    country: "",
    linkedin: "",
    summary: "",
    showAge: false,
  },
  photo: null,
  skills: [],
  languages: [],
  experience: [],
  projects: [],
  education: [],
  certifications: [],
  customSections: [],
  publications: [],
  references: [],
  facts: [],
};

const get = () => GET(new NextRequest("http://localhost:3000/api/dev/word-samples"), undefined);

describe("GET /api/dev/word-samples", () => {
  beforeEach(() => {
    vi.clearAllMocks();
    vi.mocked(getSession).mockResolvedValue({ persistent: false, issuedAt: new Date(), expiresAt: new Date() });
    vi.mocked(loadProfile).mockResolvedValue(profile);
    vi.mocked(loadSettings).mockResolvedValue({
      paper: "A4",
      resumeLanguage: "English",
      includeRefs: false,
      dateFormat: "DAY_MONTH_YEAR",
      resumeSections: { ...DEFAULT_SECTION_LAYOUT, off: ["skills"], titles: { summary: "About me" } },
    } as never);
  });

  it("zips a Word file and a Chromium PDF for every built-in theme", async () => {
    const response = await get();
    expect(response.status).toBe(200);
    expect(response.headers.get("content-type")).toBe("application/zip");
    expect(response.headers.get("content-disposition")).toContain(".zip");

    const files = unzipSync(new Uint8Array(await response.arrayBuffer()));
    for (const theme of BUILT_IN_THEMES) {
      expect(files[`docx/${theme.id}.docx`], theme.id).toEqual(new Uint8Array([80, 75, 3, 4]));
      expect(files[`chromium/${theme.id}.pdf`], theme.id).toEqual(new Uint8Array([37, 80, 68, 70]));
    }
    // Demo stress fixtures: two letters and a references-enabled resume.
    for (const name of ["letter-editorial", "letter-signal"]) {
      expect(files[`docx/${name}.docx`], name).toEqual(new Uint8Array([80, 75, 3, 4, 76]));
      expect(files[`chromium/${name}.pdf`], name).toEqual(new Uint8Array([37, 80, 68, 70]));
    }
    expect(files["docx/references-studio.docx"]).toEqual(new Uint8Array([80, 75, 3, 4]));
    expect(files["chromium/references-studio.pdf"]).toEqual(new Uint8Array([37, 80, 68, 70]));
    // The Windows converter and a manifest that explains the folders.
    expect(strFromU8(files["docx-to-pdf.ps1"])).toContain("ExportAsFixedFormat");
    const manifest = JSON.parse(strFromU8(files["manifest.json"]));
    expect(manifest.paper).toBe("A4");
    expect(manifest.documents.map((entry: { theme: string }) => entry.theme)).toEqual(BUILT_IN_THEMES.map((theme) => theme.id));
    expect(manifest.extras.map((entry: { docx: string }) => entry.docx)).toEqual([
      "docx/letter-editorial.docx",
      "docx/letter-signal.docx",
      "docx/references-studio.docx",
    ]);

    expect(renderLetterDocx).toHaveBeenCalledTimes(2);
    expect(renderResumeDocx).toHaveBeenNthCalledWith(1, expect.objectContaining({
      sections: expect.objectContaining({ off: ["skills"], titles: { summary: "About me" } }),
    }));
    expect(renderPdf).toHaveBeenCalledTimes(BUILT_IN_THEMES.length + 3);
    expect(vi.mocked(renderPdf).mock.calls[0][0]).toMatch(/^\/print\/theme\//);
    expect(vi.mocked(renderPdf).mock.calls.map(([url]) => url)).toEqual(
      expect.arrayContaining(["/print/theme/editorial/letter", "/print/theme/signal/letter", "/print/theme/studio/references"]),
    );
  });

  it("needs a session", async () => {
    vi.mocked(getSession).mockResolvedValue(null);
    expect((await get()).status).toBe(401);
  });

  it("is unavailable in production", async () => {
    vi.stubEnv("NODE_ENV", "production");
    try {
      expect((await get()).status).toBe(404);
      expect(renderResumeDocx).not.toHaveBeenCalled();
      expect(loadProfile).not.toHaveBeenCalled();
    } finally {
      vi.unstubAllEnvs();
    }
  });
});
