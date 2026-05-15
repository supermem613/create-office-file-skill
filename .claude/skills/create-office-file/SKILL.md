---
name: create-office-file
description: Convert markdown to PowerPoint (.pptx) or Word (.docx) documents.
---

## Usage

```bash
node $SKILL_DIR/scripts/create-office-file.mjs -i <input.md> -o <output.pptx|docx>
```

Format is auto-detected from the output file extension. Use `-f pptx` or `-f docx` to override.

Reads from stdin if no `-i` is provided:

```bash
cat notes.md | node $SKILL_DIR/scripts/create-office-file.mjs -o notes.docx
```

### Custom Theme via Template

Apply colors and fonts from an existing Office file:

```bash
node $SKILL_DIR/scripts/create-office-file.mjs -i input.md -o output.pptx --template corporate.pptx
```

The `--template` (`-t`) option extracts theme colors (12 OOXML scheme colors) and fonts (major/minor) from the provided `.pptx` or `.docx` file and applies them to the output. Cross-format works: a `.pptx` template can style a `.docx` output and vice versa.

When a `.docx` template is provided, the script also extracts:
- **Full styles.xml** — heading styles, title style, font defaults, and spacing are faithfully reproduced
- **Headers and footers** — `header1.xml`, `footer1.xml`, etc. are carried over with correct `sectPr` references, so page numbers and classification labels appear in the output
- **Numbering.xml** — preserved when the template binds heading styles to a multilevel list; markdown lists are merged in at offset numIds so they never collide with template-defined numbering

#### Heading auto-numbering (default)

When the template's `Heading1`..`HeadingN` styles bind to a multilevel `numId` (the typical "1.", "1.1", "1.1.1" pattern in corporate Word templates), the script preserves that wiring and lets Word render the numbers:

- The `numPr` references on heading styles are kept (instead of being stripped).
- Markdown heading levels are promoted by one so `##` maps to `Heading1`, `###` to `Heading2`, etc. The first `#` still becomes `Title`. This matches the convention where the document title is `# Title` and the first numbered section is `## 1. Foo`.
- Any leading numeric prefix in the markdown heading text (e.g. `## 1. Foo`, `### 1.2 Bar`, `#### 1.2.3. Baz`) is stripped automatically so the template's auto-numbering is the only thing the reader sees.

If the template has no heading-bound numbering, the previous behavior applies: `numPr` is stripped from styles and markdown headings keep their literal text.

#### Sensitivity-labeled templates

Templates protected with a Microsoft Information Protection (MIP) sensitivity label are stored as OLE compound files and cannot be read as ZIPs. The script detects this and surfaces a clear, actionable error rather than a generic "EOCD not found". Workarounds: open the template in Word, use `File > Info > Sensitivity` to remove the label and save a copy, then point `--template` at the unlabeled copy. Alternatively, `--save-template` against an unlabeled copy produces a portable bundle you can reuse.

### Saving a Portable Template Bundle

Some templates are encrypted (Microsoft Information Protection labels), live on a remote drive, or are content-heavy `.dotx`/`.docx` sources you would rather not redistribute. The `--save-template` flag extracts only the parts the script actually consumes (theme, styles, numbering, headers, footers, sectPr) and packages them into a small, portable `.docx` bundle:

```bash
node $SKILL_DIR/scripts/create-office-file.mjs --template corporate.dotx --save-template corporate-bundle.docx
```

The bundle is a valid `.docx` openable in Word (the body is a single empty paragraph that picks up the template's section setup). Pass it back as `--template` on subsequent runs to get the same styling without needing the original source.

## PPTX Slide Splitting

| Markdown | Result |
|----------|--------|
| `# Heading` | Title slide (large centered text) |
| `## Heading` | Content slide with title bar |
| `---` | Explicit slide break |
| Content between headings | Body text / bullets on current slide |

## DOCX Heading Mapping

The first `#` heading in the document is styled as **Title** (large, themed). Subsequent `#` headings use `Heading1`. `##` → Heading2, ... `######` → Heading6. All other content maps to styled paragraphs, lists, code blocks, or tables.

## DOCX List Numbering

Each separate list in the markdown (bullet or ordered) gets its own numbering instance with restart. Lists separated by headings, paragraphs, or other content will restart their numbering from 1, not continue from the previous list.

## Supported Markdown

- Headings (`#` through `######`)
- Paragraphs
- **Bold** (`**text**`), *Italic* (`*text*`), ***Bold+Italic*** (`***text***`)
- Bullet lists (`- item`), Numbered lists (`1. item`) — nested up to 3 levels via indentation
- Fenced code blocks (with language hint)
- Tables (`| col | col |`)
- Links (`[text](url)`)
- Images (`![alt text](path/to/image.png)`) — PNG and JPEG, local files
- Horizontal rules (`---`)
- Inline code (`` `code` ``)

## Not Supported

Nested emphasis, reference-style links, HTML-in-markdown, remote images (URLs), nested lists beyond 3 levels, template slide/page layout passthrough (only theme colors and fonts are extracted).

## Reference Files

For OOXML internals (only needed when debugging or extending the script):

| File | Covers |
|------|--------|
| `references/pptx.md` | PresentationML, DrawingML, slide shapes, tables |
| `references/docx.md` | WordprocessingML, paragraphs, runs, lists, tables |
