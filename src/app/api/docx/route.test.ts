import { NextRequest } from "next/server";
import { beforeEach, describe, expect, it, vi } from "vitest";

import { getSession } from "@/lib/auth/session";
import { DOCX_CONTENT_TYPE } from "@/lib/docx/download";
import { renderLetterDocx, renderResumeDocx } from "@/lib/docx/render";
import { renderPdf } from "@/lib/pdf/render";
import { prisma } from "@/lib/prisma";
import { loadProfile, loadProfileName } from "@/lib/profile-data";
import { settingsFromRows } from "@/lib/settings";
import { loadSettings } from "@/lib/settings-data";
import { loadTheme } from "@/lib/theme-data";
import { BUILT_IN_THEMES, resolveTheme, withoutSidebar } from "@/lib/themes";

import { GET } from "./route";

vi.mock("@/lib/auth/session", () => ({ getSession: vi.fn() }));
vi.mock("@/lib/pdf/render", () => ({ renderPdf: vi.fn() }));
vi.mock("@/lib/docx/render", () => ({
  renderResumeDocx: vi.fn(async () => new Uint8Array([80, 75, 3, 4])),
  renderLetterDocx: vi.fn(async () => new Uint8Array([80, 75, 3, 4])),
}));
vi.mock("@/lib/profile-data", () => ({ loadProfile: vi.fn(), loadProfileName: vi.fn() }));
vi.mock("@/lib/settings-data", () => ({ loadSettings: vi.fn() }));
vi.mock("@/lib/theme-data", () => ({ loadTheme: vi.fn(), THEME_FIELDS: {} }));
vi.mock("@/lib/prisma", () => ({ prisma: { application: { findUnique: vi.fn() } } }));

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
  references: [{ id: "ref1", name: "Referee", relation: "Manager", company: "Northwind", email: "ref@example.test", phone: "" }],
  facts: [],
};
const units = [{ id: "summary", label: "Summary", text: "Hello" }];
const extracted = {
  title: null, company: null, role: null, location: null, mode: null,
  salary: null, deadline: null, requirements: [], language: null,
};

const get = (query: string, headers?: Record<string, string>) =>
  GET(new NextRequest(`http://localhost:3000/api/docx${query}`, { headers }), undefined);

const application = {
  company: "Ledgerly",
  role: "Senior Designer",
  location: "Utrecht",
  language: "English",
  tone: "WARM",
  sentAt: null,
  followThemeLayout: false,
  showPhoto: null,
  extracted: null,
  frozen: null,
  sections: null,
  documents: [
    { kind: "RESUME", units, theme: BUILT_IN_THEMES[0] },
    { kind: "COVER_LETTER", units, theme: BUILT_IN_THEMES[0] },
  ],
};

beforeEach(() => {
  vi.mocked(renderPdf).mockReset().mockResolvedValue(new TextEncoder().encode("%PDF /Type /Page "));
  vi.mocked(getSession).mockResolvedValue({ persistent: false, issuedAt: new Date(), expiresAt: new Date() });
  vi.mocked(loadProfileName).mockResolvedValue("Jane Doe");
  vi.mocked(loadProfile).mockClear().mockResolvedValue(profile as never);
  vi.mocked(loadSettings).mockResolvedValue(settingsFromRows([]));
  vi.mocked(loadTheme).mockResolvedValue(BUILT_IN_THEMES[0]);
  vi.mocked(prisma.application.findUnique).mockReset().mockResolvedValue(application as never);
  vi.mocked(renderResumeDocx).mockClear();
  vi.mocked(renderLetterDocx).mockClear();
});

describe("GET /api/docx", () => {
  it("downloads the base resume as a Word document named by the pattern", async () => {
    const response = await get("?doc=base");
    expect(response.status).toBe(200);
    expect(renderResumeDocx).toHaveBeenCalledTimes(1);
    expect(response.headers.get("content-type")).toBe(DOCX_CONTENT_TYPE);
    expect(response.headers.get("content-disposition")).toBe(
      "attachment; filename=\"Jane Doe - Resume.docx\"; filename*=UTF-8''Jane%20Doe%20%E2%80%93%20Resume.docx",
    );
    expect(response.headers.get("cache-control")).toBe("private, no-store");
    expect(new Uint8Array(await response.arrayBuffer())).toEqual(new Uint8Array([80, 75, 3, 4]));
    expect(renderPdf).not.toHaveBeenCalled();
  });

  it("downloads a tailored resume and cover letter, named for the application", async () => {
    const resume = await get("?doc=resume&application=app1");
    expect(resume.status).toBe(200);
    expect(renderResumeDocx).toHaveBeenCalledTimes(1);
    expect(resume.headers.get("content-disposition")).toContain(
      "filename*=UTF-8''Jane%20Doe%20%E2%80%93%20Ledgerly%20%E2%80%93%20Resume.docx",
    );

    const letter = await get("?doc=letter&application=app1");
    expect(letter.status).toBe(200);
    expect(renderLetterDocx).toHaveBeenCalledTimes(1);
    expect(letter.headers.get("content-disposition")).toContain(
      "filename*=UTF-8''Jane%20Doe%20%E2%80%93%20Ledgerly%20%E2%80%93%20Cover%20letter.docx",
    );
  });

  it("fills the role token in a tailored Word file name, as the PDF name does", async () => {
    vi.mocked(loadSettings).mockResolvedValue(
      settingsFromRows([{ key: "fileNamePattern", value: "{name} – {role} – {doc}" }]),
    );
    const response = await get("?doc=resume&application=app1");
    expect(response.status).toBe(200);
    expect(response.headers.get("content-disposition")).toContain(
      "filename*=UTF-8''Jane%20Doe%20%E2%80%93%20Senior%20Designer%20%E2%80%93%20Resume.docx",
    );
  });

  it("uses the document's theme, layout override and photo switch", async () => {
    vi.mocked(loadProfile).mockResolvedValue({ ...profile, photo: { src: "data:image/png;base64,iVBORw==", onResumes: true } } as never);
    vi.mocked(prisma.application.findUnique).mockResolvedValue({
      ...application,
      followThemeLayout: true,
      showPhoto: false,
      documents: [{ kind: "RESUME", units, theme: { ...BUILT_IN_THEMES[0], layout: "SPLIT" } }],
    } as never);
    await get("?doc=resume&application=app1");
    expect(renderResumeDocx).toHaveBeenCalledWith(
      expect.objectContaining({
        resolved: expect.objectContaining({ split: true, photo: expect.objectContaining({ position: expect.any(String) }) }),
        content: expect.objectContaining({ photo: null }),
      }),
    );
  });

  it("removes the sidebar for a tailored resume unless it follows the theme layout", async () => {
    const theme = { ...BUILT_IN_THEMES[0], layout: "SPLIT_END" as const };
    vi.mocked(prisma.application.findUnique).mockResolvedValue({
      ...application,
      documents: [{ kind: "RESUME", units, theme }],
    } as never);
    await get("?doc=resume&application=app1");
    expect(renderResumeDocx).toHaveBeenCalledWith(expect.objectContaining({
      resolved: resolveTheme(theme, withoutSidebar(theme.layout) as never, 0),
    }));
  });

  it("passes base sections through and shares PDF's first fitting level", async () => {
    const sections = { order: ["skills", "summary"], off: ["education"], titles: { skills: "Expertise" } };
    vi.mocked(loadSettings).mockResolvedValue(settingsFromRows([
      { key: "resumeSections", value: JSON.stringify(sections) },
      { key: "fitBaseResume", value: "true" },
    ]));
    vi.mocked(renderPdf).mockImplementation(async (path) => new TextEncoder().encode(
      `%PDF ${"/Type /Page ".repeat(path === "/print/base?fit=2" ? 1 : 2)}`,
    ));
    await get("?doc=base");
    expect(renderResumeDocx).toHaveBeenCalledWith(expect.objectContaining({
      sections,
      resolved: resolveTheme(BUILT_IN_THEMES[0], undefined, 2),
    }));
    expect(vi.mocked(renderPdf).mock.calls.map(([path]) => path)).toEqual(["/print/base", "/print/base?fit=1", "/print/base?fit=2"]);
  });

  it("keeps normal formatting when PDF already fits, and uses the tightest level when it cannot fit", async () => {
    vi.mocked(loadSettings).mockResolvedValue(settingsFromRows([{ key: "fitBaseResume", value: "true" }]));
    await get("?doc=base");
    expect(renderResumeDocx).toHaveBeenLastCalledWith(expect.objectContaining({
      resolved: resolveTheme(BUILT_IN_THEMES[0], undefined, 0),
    }));
    expect(renderPdf).toHaveBeenCalledTimes(1);

    vi.mocked(renderPdf).mockClear().mockResolvedValue(new TextEncoder().encode("%PDF /Type /Page /Type /Page "));
    await get("?doc=base");
    expect(renderResumeDocx).toHaveBeenLastCalledWith(expect.objectContaining({
      resolved: resolveTheme(BUILT_IN_THEMES[0], undefined, 4),
    }));
    expect(renderPdf).toHaveBeenCalledTimes(5);
  });

  it("does not apply base fitting to tailored resumes or letters", async () => {
    vi.mocked(loadSettings).mockResolvedValue(settingsFromRows([{ key: "fitBaseResume", value: "true" }]));
    await get("?doc=resume&application=app1");
    await get("?doc=letter&application=app1");
    expect(renderPdf).not.toHaveBeenCalled();
  });

  it("does not silently export unfitted Word when PDF measurement fails", async () => {
    vi.mocked(loadSettings).mockResolvedValue(settingsFromRows([{ key: "fitBaseResume", value: "true" }]));
    vi.mocked(renderPdf).mockRejectedValueOnce(new Error("Chromium unavailable"));
    const log = vi.spyOn(console, "error").mockImplementation(() => {});
    try {
      expect((await get("?doc=base")).status).toBe(500);
      expect(renderResumeDocx).not.toHaveBeenCalled();
    } finally {
      log.mockRestore();
    }
  });

  it("uses application sections rather than base sections, including renamed and disabled sections", async () => {
    const sections = { order: ["projects", "summary"], off: ["skills"], titles: { projects: "Selected work" } };
    vi.mocked(prisma.application.findUnique).mockResolvedValue({ ...application, sections } as never);
    await get("?doc=resume&application=app1");
    expect(renderResumeDocx).toHaveBeenCalledWith(expect.objectContaining({ sections }));
  });

  it("falls back to base sections for an application without a valid layout", async () => {
    const sections = { order: ["languages", "skills"], off: ["projects"] };
    vi.mocked(loadSettings).mockResolvedValue(settingsFromRows([{ key: "resumeSections", value: JSON.stringify(sections) }]));
    vi.mocked(prisma.application.findUnique).mockResolvedValue({ ...application, sections: { invalid: true } } as never);
    await get("?doc=resume&application=app1");
    expect(renderResumeDocx).toHaveBeenCalledWith(expect.objectContaining({ sections }));
  });

  it.each([
    [false, true, false, false],
    [true, false, false, false],
    [true, true, false, true],
    [false, false, true, true],
  ])("applies reference policy (setting=%s, posting=%s, explicit=%s)", async (setting, asks, explicit, included) => {
    vi.mocked(loadSettings).mockResolvedValue(settingsFromRows([{ key: "includeRefs", value: String(setting) }]));
    vi.mocked(prisma.application.findUnique).mockResolvedValue({
      ...application,
      extracted: { ...extracted, asksForReferences: asks },
      sections: { order: [], off: [], withReferences: explicit },
    } as never);
    await get("?doc=resume&application=app1");
    const prepared = vi.mocked(renderResumeDocx).mock.calls[0][0];
    expect(prepared.content.references.map((reference) => reference.name)).toEqual(included ? ["Referee"] : []);
  });

  it("uses the sent copy's profile, themes, paper, date format and reference policy without mutating inputs", async () => {
    const frozenTheme = { ...BUILT_IN_THEMES[0], accent: "#123456", layout: "SPLIT" as const };
    const frozen = {
      profile: { ...profile, basics: { ...profile.basics, name: "Sent Name", email: "sent@example.test" } },
      themes: { RESUME: frozenTheme, COVER_LETTER: frozenTheme },
      paper: "LETTER",
      dateFormat: "NUMERIC",
      includeReferences: true,
      followedBase: true,
    };
    const sent = {
      ...application,
      frozen,
      sentAt: new Date(2025, 0, 2, 12),
      sections: { order: ["references", "summary"], off: ["skills"] },
    };
    const before = structuredClone({ sent, profile });
    const livePhoto = { src: "data:image/png;base64,iVBORw==", onResumes: true };
    vi.mocked(loadProfile).mockResolvedValue({ ...profile, photo: livePhoto } as never);
    vi.mocked(prisma.application.findUnique).mockResolvedValue(sent as never);
    vi.mocked(loadSettings).mockResolvedValue(settingsFromRows([{ key: "includeRefs", value: "false" }]));
    expect((await get("?doc=resume&application=app1")).status).toBe(200);
    expect(renderResumeDocx).toHaveBeenCalledWith(expect.objectContaining({
      paper: "LETTER",
      theme: frozenTheme,
      resolved: resolveTheme(frozenTheme, withoutSidebar("SPLIT") as never, 0),
      sections: sent.sections,
      content: expect.objectContaining({ name: "Sent Name", photo: livePhoto.src, references: [expect.objectContaining({ name: "Referee" })] }),
    }));
    expect((await get("?doc=letter&application=app1")).status).toBe(200);
    expect(renderLetterDocx).toHaveBeenCalledWith(expect.objectContaining({
      paper: "LETTER",
      theme: frozenTheme,
      senderName: "Sent Name",
      contactLines: [["sent@example.test"]],
      date: "02-01-2025",
    }));
    expect({ sent, profile }).toEqual(before);
  });

  it("keeps a sent reference exclusion unless sections explicitly add references", async () => {
    const frozen = {
      profile,
      themes: { RESUME: BUILT_IN_THEMES[0], COVER_LETTER: null },
      paper: "A4",
      dateFormat: "NUMERIC",
      includeReferences: false,
      followedBase: true,
    };
    vi.mocked(loadSettings).mockResolvedValue(settingsFromRows([{ key: "includeRefs", value: "true" }]));
    vi.mocked(prisma.application.findUnique).mockResolvedValue({ ...application, frozen, extracted: { ...extracted, asksForReferences: true } } as never);
    await get("?doc=resume&application=app1");
    expect(vi.mocked(renderResumeDocx).mock.calls[0][0].content.references).toEqual([]);
    vi.mocked(prisma.application.findUnique).mockResolvedValue({
      ...application, frozen, sections: { order: [], off: [], withReferences: true },
    } as never);
    await get("?doc=resume&application=app1");
    expect(vi.mocked(renderResumeDocx).mock.calls[1][0].content.references).toHaveLength(1);
  });

  it("answers 404 for documents it doesn't know, missing applications and missing documents", async () => {
    expect((await get("?doc=resume:abc")).status).toBe(404);
    expect((await get("")).status).toBe(404);
    vi.mocked(prisma.application.findUnique).mockResolvedValue(null);
    expect((await get("?doc=resume&application=gone")).status).toBe(404);
    expect((await get("?doc=resume&application=..%2Fbase")).status).toBe(404);
    vi.mocked(prisma.application.findUnique).mockResolvedValue({ ...application, documents: [] } as never);
    expect((await get("?doc=letter&application=app1")).status).toBe(404);
    expect(renderResumeDocx).not.toHaveBeenCalled();
    expect(renderLetterDocx).not.toHaveBeenCalled();
  });

  it("treats units from an older version as a missing document", async () => {
    vi.mocked(prisma.application.findUnique).mockResolvedValue({
      ...application,
      documents: [{ kind: "RESUME", units: [{ nope: true }], theme: null }],
    } as never);
    const response = await get("?doc=resume&application=app1");
    expect(response.status).toBe(404);
    expect(await response.json()).toEqual({ error: "This application has no tailored resume." });
  });

  it("rejects malformed application IDs before querying applications", async () => {
    expect((await get("?doc=resume&application=..%2Fbase")).status).toBe(404);
    expect((await get("?doc=letter")).status).toBe(404);
    expect(prisma.application.findUnique).not.toHaveBeenCalled();
  });

  it("does not disguise unexpected loader failures as missing applications", async () => {
    vi.mocked(prisma.application.findUnique).mockRejectedValueOnce(new Error("database offline"));
    const log = vi.spyOn(console, "error").mockImplementation(() => {});
    try {
      expect((await get("?doc=resume&application=app1")).status).toBe(500);
      expect(renderResumeDocx).not.toHaveBeenCalled();
    } finally {
      log.mockRestore();
    }
  });

  it("needs a session and the app's own origin", async () => {
    const other = await get("?doc=base", { origin: "https://evil.example", host: "localhost:3000" });
    expect(other.status).toBe(403);
    vi.mocked(getSession).mockResolvedValue(null);
    expect((await get("?doc=base")).status).toBe(401);
    expect(renderResumeDocx).not.toHaveBeenCalled();
    expect(prisma.application.findUnique).not.toHaveBeenCalled();
    expect(loadProfile).not.toHaveBeenCalled();
  });

  it("turns a failed render into a friendly 500 without the document in the log", async () => {
    vi.mocked(renderResumeDocx).mockRejectedValueOnce(new Error("font missing"));
    const log = vi.spyOn(console, "error").mockImplementation(() => {});
    const response = await get("?doc=base");
    expect(response.status).toBe(500);
    expect(await response.json()).toEqual({ error: expect.stringMatching(/^Something went wrong/) });
    expect(log).toHaveBeenCalled();
    expect(JSON.stringify(log.mock.calls)).not.toContain("Jane Doe");
    log.mockRestore();
  });
});
