# Lao document generation — instructions for AI coding agents

Add this file to your project root as `AGENTS.md` (or merge into an existing one)
so AI assistants (Codex, Cursor, Copilot, Claude, etc.) generate Lao documents
that wrap correctly. Full recipe with ready helpers:
https://phetsarath.mts.la/ai/SKILL.md

## The one rule you must not skip

Lao is written without spaces. Microsoft Word wraps Lao at word boundaries ONLY
for runs marked as complex script. Generated .docx files (python-docx, docx4j,
templates) must put these in every Lao run's `w:rPr`:

- `<w:rFonts w:cs="Phetsarath OT" .../>` — complex-script font
- `<w:sz .../>` **and** `<w:szCs .../>` — both sizes
- `<w:b/>+<w:bCs/>`, `<w:i/>+<w:iCs/>` — both style flags when bold/italic
- `<w:cs/>` — the complex-script flag itself
- `<w:lang w:bidi="lo-LA"/>`

…and `<w:themeFontLang w:bidi="lo-LA"/>` in `word/settings.xml`.

Without `<w:cs/>`, Word breaks Lao words at arbitrary letter boundaries
(ອັກສອນ → ອັ|ກສອນ).

## Other formats

- **PowerPoint (.pptx)**: also set the DrawingML complex-script typeface
  (`<a:cs typeface="Phetsarath OT"/>`), not just `a:latin`.
- **HTML**: mark Lao content `lang="lo"`; never use `word-break: break-all`;
  use `line-height ≥ 1.8` so stacked vowels/tones don't clip.
- **PDF**: use HarfBuzz-based renderers (WeasyPrint, headless browsers);
  ReportLab cannot shape Lao.

## Font

Use **Phetsarath OT v4.103** (official Lao font, free, SIL OFL):
https://phetsarath.mts.la — families `Phetsarath OT` (Regular/Italic/Bold/Bold
Italic) plus `Phetsarath OT Light` / `… Medium` / `… SemiBold` / `… Black`.
Webfont: https://phetsarath.mts.la/fonts/phetsarath.css
