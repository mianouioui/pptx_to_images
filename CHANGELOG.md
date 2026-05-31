# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.0] - 2026-05-31

### Added

- Export progress feedback: each external step (LibreOffice→PDF, PDF→PNG, macOS PDFKit, PowerPoint, Keynote) now shows a live spinner with elapsed seconds in a terminal, finishing with `✓`/`✗` and the step's duration. So even a slow step (notably LibreOffice's first cold start) no longer looks like a hang.
- Batch runs now prefix each file with an `[index/total]` counter.

### Changed

- `run_command` wraps every external command with the new progress indicator. In a non-interactive context (pipe/redirect) the indicator degrades to static start/finish lines, so logs stay clean.
- Timeout errors now state the limit in seconds.
- The “LibreOffice not found” error now includes install instructions (Homebrew / download link), and the README gained an “Installing LibreOffice” section (Chinese / English / French).
- New `PPTX2IMG_ENGINE` environment variable sets the default engine (e.g. `PPTX2IMG_ENGINE=keynote`), so the double-click launcher / interactive mode can use Keynote without passing `--engine`. An explicit `--engine` still overrides it.

## [1.0.1] - 2026-05-31

### Fixed

- `--force` reuses an existing output folder, but now clears leftover `pageNNN.png` from the previous run before exporting, so re-rendering a shortened deck no longer leaves stale trailing pages behind.
- The macOS PDFKit fallback now renders into a temporary subfolder and goes through the shared numbering/cleanup step like every other engine, so it honors `--force` and no longer leaves half-written images on failure.
- Batch conversion no longer aborts on the first failing file; it skips the file, continues with the rest, lists every failure at the end, and exits non-zero if any failed.
- Each engine's `.raw_*` working folder is now removed via `try/finally`, so a failed render no longer leaves it behind in the output folder.
- Opening the output after a multi-file run with `-o` now uses the expanded/resolved path.

### Changed

- macOS launcher `pptx_to_images.command` no longer prints its own banner in interactive mode; the Python entry point now shows the only banner.
- Windows launcher `pptx_to_images.cmd` accepts multiple dropped files at once (previously only the first file was processed).
- `--engine keynote` now states, in `--help` and at runtime, that `--width` has no effect because Keynote exports at the slide's native resolution.
- `iter_presentations` filters by file extension before touching the filesystem, avoiding unnecessary `stat` calls.

## [1.0.0] - 2026-05-31

### Added

- `pptx_to_images.py`: Python source entry point for exporting each slide in a `.pptx` / `.ppt` file to PNG images.
- `pptx_to_images.command`: macOS launcher with drag-and-drop terminal guidance.
- `pptx_to_images.cmd`: Windows launcher that uses PowerShell to automate Microsoft PowerPoint export.
- Automatic rendering engine selection in the Python source:
  - Windows: PowerPoint first, LibreOffice fallback when using the Python source.
  - macOS/Linux: LibreOffice to PDF, then `pdftoppm` or macOS PDFKit to PNG.
  - macOS: Keynote fallback when LibreOffice is unavailable.
- Batch processing for folders, including recursive scanning with `--recursive`.
- Output image naming as `page001.png`, `page002.png`, and so on.
- Options for `--width`, `--prefix`, `-o/--output`, `--force`, `--dry-run`, `--open`, and `--engine`.
- README documentation aligned with the companion PPTX tool family, including Chinese, English, and French sections.
- Per-version release notes under `docs/`.

### Notes

- The tool requires a real presentation renderer because slide-to-image export cannot be done by simply reading PPTX XML. Supported renderers are Microsoft PowerPoint, LibreOffice, and Keynote on macOS.
- The Python source has been tested on macOS with LibreOffice + `pdftoppm`, and with the macOS PDFKit fallback. The Windows PowerPoint launcher should be verified on a real Windows machine before release-critical use.

[1.1.0]: https://github.com/mianouioui/pptx_to_images/releases/tag/V1.1.0
[1.0.1]: https://github.com/mianouioui/pptx_to_images/releases/tag/V1.0.1
[1.0.0]: https://github.com/mianouioui/pptx_to_images/releases/tag/V1.0.0
