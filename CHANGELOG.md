# Changelog

All notable user-visible changes should be documented here.

## Unreleased

- Reworked the left rail to avoid clipping and keep conversion controls practical.
- Made the queue compact so about six filenames are visible at once.
- Kept the main convert action sticky and easier to reach.
- Added viewport width syncing so scroll content does not overflow horizontally.
- Fixed PyInstaller packaging so Magika model files are included in frozen builds.
- Installed the common MarkItDown document extras for PDF, DOCX, PPTX, XLS/XLSX, Outlook, audio, and YouTube inputs.
- Added a hidden frozen-app conversion self-test switch for release verification.
- Made the sticky primary action recover failed queue items when no ready items remain.
- Normalized PDF bullet extraction artifacts into proper Markdown `- ` list items.
- Switched the default build to PyInstaller one-folder output for faster startup.
- Switched the desktop UI back to a dark-first default theme.
- Added contributor and roadmap docs for GitHub publishing.
