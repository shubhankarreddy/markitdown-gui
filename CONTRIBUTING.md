# Contributing

Thanks for helping make MarkItDown GUI more useful. The project goal is simple: provide a practical local Windows workflow for converting mixed source material into reusable Markdown.

## Good First Contributions

- Improve conversion workflow clarity without adding visual clutter.
- Add small reliability fixes around queue, retry, folder scan, and auto-save behavior.
- Test more file formats and document what works or fails.
- Improve packaging and installer scripts.
- Add screenshots, demo GIFs, and release notes.

## Local Setup

1. Install Python 3.11 or newer.
2. Run `setup.bat`.
3. Run `venv\Scripts\python.exe markitdown_app.py`.
4. Run `build.bat` before release testing.

## Design Principles

- Keep the app local-first and reviewable.
- Prefer practical workflow improvements over decoration.
- Make failures recoverable with retry and clear status.
- Keep the queue, output folder, and preview easy to understand at a glance.

## Before Opening A Pull Request

- Run `venv\Scripts\python.exe -m py_compile markitdown_app.py`.
- Test at least one file conversion manually.
- If changing UI layout, test at a normal laptop width and a wide desktop width.
- Update `CHANGELOG.md` when the change is user-visible.
