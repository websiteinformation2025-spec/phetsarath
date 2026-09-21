---
name: lao-docx
description: Generate Word/Office documents and web pages with Lao text that wrap and render correctly. Use whenever creating .docx/.pptx files or HTML containing Lao script (ພາສາລາວ), especially with python-docx, docx4j, or any programmatic document generation — Lao text needs complex-script flags that generators do not set by default.
---

# Generating documents with Lao text

Lao is written without spaces between words. Line breaking depends on the layout
engine's dictionary segmentation — and in Microsoft Word that engine only runs
for text runs marked as **complex script**. Word marks runs automatically when a
person types; **programmatically generated files must set the flags explicitly**,
or Word falls back to breaking at any letter-cluster boundary, splitting words
(ອັກສອນ becomes ອັກ|ສອນ or worse ອັ|ກ). This is the single most common bug in
generated Lao documents.

## The .docx rule (python-docx, docx4j, OOXML templating)

Every run containing Lao text MUST have, inside `w:rPr` (in this schema order):

```xml
<w:rPr>
  <w:rFonts w:ascii="Phetsarath OT" w:hAnsi="Phetsarath OT" w:cs="Phetsarath OT"/>
  <w:b/><w:bCs/>                       <!-- only if bold: BOTH b and bCs -->
  <w:i/><w:iCs/>                       <!-- only if italic: BOTH i and iCs -->
  <w:sz w:val="28"/><w:szCs w:val="28"/>  <!-- size in half-points, BOTH -->
  <w:cs/>                              <!-- THE complex-script flag - critical -->
  <w:lang w:val="en-US" w:bidi="lo-LA"/>
</w:rPr>
```

And once per document in `word/settings.xml`:

```xml
<w:themeFontLang w:val="en-US" w:bidi="lo-LA"/>
```

### Ready-made python-docx helper

```python
from docx.oxml.ns import qn

def _el(parent, tag, **attrs):
    e = parent.makeelement(qn(tag), {qn(k): v for k, v in attrs.items()})
    parent.append(e)
    return e

def lao_run(par, text, font="Phetsarath OT", bold=False, italic=False, size=14):
    run = par.add_run(text)
    rPr = run._element.get_or_add_rPr()
    for child in list(rPr):
        rPr.remove(child)
    _el(rPr, "w:rFonts", **{"w:ascii": font, "w:hAnsi": font, "w:cs": font})
    if bold:
        _el(rPr, "w:b"); _el(rPr, "w:bCs")
    if italic:
        _el(rPr, "w:i"); _el(rPr, "w:iCs")
    half = str(int(size * 2))
    _el(rPr, "w:sz", **{"w:val": half}); _el(rPr, "w:szCs", **{"w:val": half})
    _el(rPr, "w:cs")
    _el(rPr, "w:lang", **{"w:val": "en-US", "w:bidi": "lo-LA"})
    return run

def set_lao_theme_lang(doc):
    settings = doc.settings.element
    tfl = settings.find(qn("w:themeFontLang"))
    if tfl is None:
        tfl = settings.makeelement(qn("w:themeFontLang"), {})
        settings.append(tfl)
    tfl.set(qn("w:val"), "en-US"); tfl.set(qn("w:bidi"), "lo-LA")
```

Do NOT rely on `run.font.name` / `run.font.size` / `run.bold` alone — python-docx
sets only the Latin properties (`w:sz`, `w:b`), which Word ignores for
complex-script runs.

## .pptx (python-pptx)

PowerPoint uses DrawingML: set the complex-script typeface too, or Lao falls back
to the default font:

```python
rPr = run._r.get_or_add_rPr()
from pptx.oxml.ns import qn as pqn
cs = rPr.makeelement(pqn('a:cs'), {'typeface': 'Phetsarath OT'})
rPr.append(cs)
run.font.name = 'Phetsarath OT'   # sets a:latin
```

## HTML / web

- Mark Lao content `lang="lo"` — browsers then use ICU dictionary segmentation
  for wrapping (equivalent quality to Word).
- Never use `word-break: break-all` or `overflow-wrap: anywhere` on Lao text —
  they destroy word integrity. Default `normal` wrapping is correct.
- Lao stacks vowels and tone marks above letters: give Lao text generous
  `line-height` (1.8–2.2), or marks will clip against adjacent lines.
- Font: Phetsarath OT v4.103 webfont kit (WOFF2, weights 300–900 + italics,
  SIL OFL): https://github.com/MTS-Lao/phetsarath

## PDF generation

- WeasyPrint / Typst / browsers (HarfBuzz-based shaping): work correctly.
- ReportLab does NOT shape complex scripts — Lao vowels/tones will be misplaced.
  Render via HTML→PDF instead.

## Fonts

Use **Phetsarath OT v4.103** (official Lao font, SIL OFL, free): 12 style-linked
fonts. Family names for documents: `Phetsarath OT` (Regular/Italic/Bold/Bold
Italic via style linking) plus `Phetsarath OT Light` / `… Medium` /
`… SemiBold` / `… Black`. Covers full Lao including Pali/Sanskrit/Kmhmu letters
(U+0E86–U+0EAC). Do not use pre-4.1 versions (broken Office line-break metadata).

## Verification checklist

1. Long unspaced Lao paragraph wraps at word boundaries — never a line starting
   with a dependent vowel/tone (ະ ັ າ ິ ີ ຶ ື ຸ ູ ົ ຽ ໍ ຳ ່ ້ ໊ ໋), never a
   line ending with a pre-vowel (ເ ແ ໂ ໄ ໃ).
2. Bold/italic actually change the rendering (bCs/iCs present).
3. Font size applies to Lao text (szCs present).
