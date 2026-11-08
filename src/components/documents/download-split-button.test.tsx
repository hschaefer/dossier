import { render, screen, waitFor } from "@testing-library/react";
import userEvent from "@testing-library/user-event";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";

import { toast } from "@/components/toast";

import { DownloadSplitButton, WORD_EXPORT_NOTE } from "./download-split-button";

vi.mock("@/components/toast", () => ({ toast: vi.fn() }));

describe("DownloadSplitButton", () => {
  const click = vi.fn();
  // Counts across tests: the last file offered stays until the next download.
  let made = 0;

  beforeEach(() => {
    vi.mocked(toast).mockReset();
    click.mockReset();
    URL.createObjectURL = vi.fn(() => `blob:file-${++made}`);
    URL.revokeObjectURL = vi.fn();
    vi.spyOn(HTMLAnchorElement.prototype, "click").mockImplementation(function (this: HTMLAnchorElement) {
      click(this.download, this.getAttribute("href"));
    });
  });
  afterEach(() => vi.unstubAllGlobals());

  const renderSplit = (props: Partial<Parameters<typeof DownloadSplitButton>[0]> = {}) =>
    render(
      <DownloadSplitButton
        pdfHref="/api/pdf?doc=base"
        wordHref="/api/docx?doc=base"
        {...props}
      />,
    );

  it("exports the PDF from the main button", async () => {
    vi.stubGlobal("fetch", vi.fn(async () => Response.json({ url: "/api/pdf/held-1" })));
    renderSplit();
    const link = screen.getByRole("link", { name: "Download PDF" });
    expect(link).toHaveAttribute("href", "/api/pdf?doc=base");
    await userEvent.click(link);
    expect(fetch).toHaveBeenCalledWith("/api/pdf?doc=base", { method: "POST" });
    expect(click).toHaveBeenCalledWith("", "/api/pdf/held-1");
  });

  it("lists PDF and Word from the arrow", async () => {
    const user = userEvent.setup();
    renderSplit();
    await user.click(screen.getByRole("button", { name: "More export options" }));
    expect(screen.getByRole("menuitem", { name: /PDF \(\.pdf\)/ })).toBeInTheDocument();
    const word = screen.getByRole("menuitem", { name: /Word \(\.docx\)/ });
    expect(word).toHaveAttribute("title", WORD_EXPORT_NOTE);
    expect(word).toHaveTextContent("Line and page breaks may differ from the PDF.");
    expect(word).toHaveTextContent("PDF fit-to-one-page does not apply to Word.");
  });

  it("exports Word from the menu with its own progress and file name", async () => {
    const user = userEvent.setup();
    vi.stubGlobal(
      "fetch",
      vi.fn(
        async () =>
          new Response("PK", {
            headers: { "content-disposition": "attachment; filename*=UTF-8''Jane%20%E2%80%93%20Resume.docx" },
          }),
      ),
    );
    renderSplit();
    await user.click(screen.getByRole("button", { name: "More export options" }));
    await user.click(screen.getByRole("menuitem", { name: /Word \(\.docx\)/ }));
    expect(toast).toHaveBeenCalledWith("Preparing your Word document…");
    expect(fetch).toHaveBeenCalledWith("/api/docx?doc=base", undefined);
    await waitFor(() => expect(click).toHaveBeenCalledWith("Jane – Resume.docx", expect.stringMatching(/^blob:/)));
  });

  it("exports the PDF from the menu too", async () => {
    const user = userEvent.setup();
    vi.stubGlobal("fetch", vi.fn(async () => Response.json({ url: "/api/pdf/held-2" })));
    renderSplit();
    await user.click(screen.getByRole("button", { name: "More export options" }));
    await user.click(screen.getByRole("menuitem", { name: /PDF \(\.pdf\)/ }));
    expect(toast).toHaveBeenCalledWith("Rendering your PDF on the server. It'll download in a moment.");
    expect(fetch).toHaveBeenCalledWith("/api/pdf?doc=base", { method: "POST" });
    await waitFor(() => expect(click).toHaveBeenCalledWith("", "/api/pdf/held-2"));
  });

  it.each(["PDF", "Word"] as const)("blocks both formats during a pending %s export and permits retry after failure", async (format) => {
    const user = userEvent.setup();
    let finish!: (response: Response) => void;
    const pending = new Promise<Response>((resolve) => { finish = resolve; });
    const fetchFile = vi.fn()
      .mockReturnValueOnce(pending)
      .mockResolvedValueOnce(new Response("PK"));
    vi.stubGlobal("fetch", fetchFile);
    renderSplit();
    if (format === "PDF") {
      await user.click(screen.getByRole("link", { name: "Download PDF" }));
    } else {
      await user.click(screen.getByRole("button", { name: "More export options" }));
      await user.click(screen.getByRole("menuitem", { name: /Word \(\.docx\)/ }));
    }
    const main = screen.getByRole("link");
    expect(main).toHaveAttribute("aria-disabled", "true");
    await user.click(main);
    await user.click(screen.getByRole("button", { name: "More export options" }));
    const pdf = screen.getByRole("menuitem", { name: /PDF \(\.pdf\)/ });
    const word = screen.getByRole("menuitem", { name: /Word \(\.docx\)/ });
    expect(pdf).toHaveAttribute("aria-disabled", "true");
    expect(word).toHaveAttribute("aria-disabled", "true");
    await user.click(pdf);
    await user.click(word);
    expect(fetchFile).toHaveBeenCalledTimes(1);

    finish(Response.json({ error: "Please retry." }, { status: 500 }));
    await waitFor(() => expect(main).not.toHaveAttribute("aria-disabled"));
    expect(toast).toHaveBeenCalledWith("Please retry.");
    expect(word).not.toHaveAttribute("aria-disabled");
    await user.click(word);
    await waitFor(() => expect(click).toHaveBeenCalledWith("document.docx", expect.stringMatching(/^blob:/)));
    expect(fetchFile).toHaveBeenCalledTimes(2);
    expect(fetchFile).toHaveBeenLastCalledWith("/api/docx?doc=base", undefined);
  });

  it("opens and closes the accessible format menu from the keyboard", async () => {
    const user = userEvent.setup();
    renderSplit();
    const trigger = screen.getByRole("button", { name: "More export options" });
    trigger.focus();
    await user.keyboard("{ArrowDown}");
    expect(trigger).toHaveAttribute("aria-expanded", "true");
    expect(screen.getByRole("menuitem", { name: /PDF \(\.pdf\)/ })).toHaveFocus();
    await user.keyboard("{Escape}");
    expect(screen.queryByRole("menu")).not.toBeInTheDocument();
    expect(trigger).toHaveFocus();
  });
});
