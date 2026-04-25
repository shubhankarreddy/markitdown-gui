# Design Brief

Updated: 2026-04-25

## Product Positioning

MarkItDown GUI should position itself as a local Windows workspace for turning mixed source material into AI-ready Markdown.

The strongest promise is not simply "GUI for MarkItDown". The stronger promise is:

- local and privacy-friendly
- Windows-first and practical
- queue-based and reviewable
- useful for batch cleanup, Git workflows, and RAG preparation

## Primary Users

- AI builders preparing documents for retrieval and prompt pipelines
- docs and operations teams cleaning exported files into reusable Markdown
- researchers and students converting PDFs, slide decks, and notes
- Obsidian and knowledge-base users who want inspectable Markdown instead of opaque proprietary files

## Competitive Landscape

- [Microsoft MarkItDown](https://github.com/microsoft/markitdown)
  Baseline engine and format coverage benchmark. Strong conversion capability, weaker desktop workflow.
- [Docling](https://github.com/docling-project/docling)
  Quality benchmark for structure-aware document conversion and AI-facing messaging.
- [TO MD](https://tomd.io/)
  Fast UX benchmark for low-friction conversion and preview.
- [AnythingMD](https://anythingmd.com/about)
  Messaging benchmark for "AI-ready Markdown" positioning.
- [TurnMD](https://turnmd.com/)
  Simplicity benchmark for focused, one-job conversion experiences.
- [Obsidian Web Clipper](https://obsidian.md/clipper)
  Expectation benchmark for web capture, local ownership, and durable Markdown output.
- [Pandoc](https://pandoc.org/index.html)
  Conversion breadth benchmark and long-standing trust anchor for technical users.

## Information Architecture

The app should read as one clear flow:

1. Collect sources
2. Review queue
3. Choose output
4. Convert
5. Inspect and reuse

LLM enrichment should remain visible but clearly secondary to the core workflow.

## UI Direction

- Keep the left rail focused on workflow controls, not generic settings.
- Make the queue feel like the center of the app, not a side list.
- Keep one dominant primary action at all times.
- Use a smaller set of stronger surface levels so the hierarchy reads quickly.
- Use Windows-friendly typography with a more intentional heading treatment.
- Make progress, scan state, and retry state visible inline, not just in the status bar.

## Wireframe

```text
[ Brand / Promise ]
[ Collect sources ]
  Drop files or folders
  Add Files | Add Folder
  URL input | Clip URL

[ Output folder ]
[ Queue ]
  Ready / Saved / Issues
  Selected item actions
  Retry / Remove / Clear

[ Sticky progress + primary convert action ]
------------------------------------------------------------
[ Selected item header ]
  Name | Status | Path

[ Markdown ] [ Preview ] [ Details ]
  Main inspection workspace

[ Open source ] [ Open saved ] [ Open folder ] [ Copy ] [ Save As ]
```

## Current Product Call

The PyQt app is the most feature-complete client today. It already supports the richer queue, folder scanning, rendered preview, and saved preferences flow. Until the native WPF client reaches feature parity, the repo should make that distinction obvious.

## Testing And Troubleshooting

Minimum manual QA set for every release:

- single file conversion
- mixed queue conversion
- folder scan
- URL capture
- retry failed item
- auto-save failure handling
- open source / open saved / save as
- light theme and dark theme
- build from clean environment

Golden fixture set should include:

- PDF
- DOCX
- PPTX
- XLSX
- HTML or URL
- notebook
- one intentionally broken file

## Ongoing Updates

- Review competitor messaging monthly.
- Keep screenshots and README aligned with the current app.
- Track the top three UX metrics manually:
  - time to first successful conversion
  - retry success rate
  - number of clicks from add source to saved Markdown
- Maintain a short changelog section for visible UX improvements.
