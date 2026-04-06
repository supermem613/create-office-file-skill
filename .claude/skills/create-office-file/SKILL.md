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

## PPTX Slide Splitting

| Markdown | Result |
|----------|--------|
| `# Heading` | Title slide (large centered text) |
| `## Heading` | Content slide with title bar |
| `---` | Explicit slide break |
| Content between headings | Body text / bullets on current slide |

## DOCX Heading Mapping

`#` → Heading1, `##` → Heading2, ... `######` → Heading6. All other content maps to styled paragraphs, lists, code blocks, or tables.

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
