# Word export

The **Download PDF** button's arrow offers **Word (.docx)** for the base
resume, tailored resumes and cover letters. PDF remains the default. Word
documents contain editable text, links and lists, with embedded fonts, so
they can be adjusted manually before sending. Settings → **Document file
name** supplies the same naming pattern for PDF and Word.

The PDF and browser preview remain the layout reference. Word and Chromium
wrap lines and paginate differently; identical page counts are not promised.
Microsoft Word for Windows is the primary target. Do not assume equivalent
rendering in Word for Mac, web/mobile Word, LibreOffice or Google Docs.

## Content and layout

- Base resumes use the current profile, section settings and theme. Tailored
  documents use their saved units, including the content retained after PDF
  trimming, not regenerated text.
- Export respects section order, hidden sections, custom headings,
  certifications, volunteering, custom sections, references and photo rules.
  Sent applications use the same frozen profile, theme, paper and date
  settings as their PDF.
- **Fit to one page shares formatting between PDF and Word.** For the base
  resume, PDF measurement selects the first fitting compression level (or
  the tightest available), and Word uses the same margins, gaps, line height
  and body size. Nothing is removed. With fitting off, Word uses normal theme
  spacing and does not need PDF rendering. Word pages are not measured, so a
  one-page PDF does not guarantee a one-page Word document. The preview is
  labelled PDF preview; the Word export menu explains the pagination caveat.
- Layout tables preserve sidebars, side headings and timeline layouts.
  Long tables may paginate differently from the PDF's independent columns.
- Circle and rounded-square photos are masked locally with `sharp`. If the
  image cannot be processed, export falls back to a rectangular crop. Word
  cannot display unconverted WebP photos in the unmasked layouts.
- Downloading Word does not change the document or its history. Email
  attachments, bulk PDF downloads and backups retain their existing formats.

## Implementation and fonts

`src/lib/document-prep.ts` shares plain content preparation between the React
and native TypeScript DOCX renderers. `src/lib/docx/` uses pinned `docx@9.15.1`
(MIT); it does not convert HTML or PDF, upload data, or call an AI service.
`GET /api/docx?doc=base`, `?doc=resume&application={id}` and
`?doc=letter&application={id}` require the existing session and same-origin
checks, with private/no-store download responses. No schema migration is
needed.

Generated TrueType faces in `src/fonts/word/` are production assets, not test
artifacts. They come from the same static WOFF2 faces as the browser/PDF,
retain their licence metadata, and use renamed derived families where OFL
Reserved Font Names require it. See [font provenance and rebuild
instructions](../src/fonts/README.md#word-export-fonts). Production tracing
and desktop packaging include them for offline use. No fonts are downloaded
at runtime.

Only the faces a document uses are embedded. Multilingual documents can be
several megabytes, especially when they need multiple CJK faces. Embedded
fonts keep their bundled glyph coverage: characters outside that coverage
use Word's machine-dependent fallback. In particular, the bundled Chinese
face is a GB 2312 subset, not full Chinese coverage.

## Reproducing visual checks

In development, while signed in, `/dev/documents` links to
`/api/dev/word-samples`. It downloads a ZIP containing a Word export and
Chromium PDF for each built-in theme using the current profile and section
settings. Two synthetic cover letters and a references-enabled Studio resume
exercise letter framing and nested sidebar tables. Samples use normal theme
spacing, without PDF's automatic fit-to-one-page adjustment.

**The ZIP contains profile text. Keep it local; do not commit or attach it to
a public issue.** Generated comparisons and machine-specific verification
tooling do not belong in source control.

On a Windows machine with desktop Word, extract the ZIP to a local disk,
close Word, and run:

```powershell
powershell -ExecutionPolicy Bypass -File .\docx-to-pdf.ps1 -In .\docx -Out .\word
```

Compare `word/*.pdf` with `chromium/*.pdf`. The script requires an interactive
desktop session; use `-Visible` to investigate Word dialogs. Network shares
can stall Word's PDF conversion.

Automated tests inspect the DOCX package's content, styles, geometry,
numbering, relationships and embedded fonts; they are not visual acceptance:

```bash
docker/bin/yarn test
docker/bin/yarn typecheck
docker/bin/yarn lint
```

After changing fonts or layout, repeat comparisons in real Word on a machine
without the bundled fonts installed. Include A4 and Letter, all layout
types, new/custom and hidden sections, long multi-page content, references,
photos, mixed scripts, and edit/save/reopen. Validate both standalone and
desktop packaging offline. Comparisons of the earlier pre-1.12 prototype do
not establish visual parity for the updated font assets or new sections.
