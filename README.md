# Create Office File Skill

**Zero-dependency markdown → PowerPoint / Word converter for Claude Code, GitHub Copilot CLI, and other AI coding agents**

[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

---

## What It Does

This skill teaches AI coding agents to convert markdown into Office documents — `.pptx` (PowerPoint) and `.docx` (Word) — using only Node.js built-in modules. No npm install, no Python, no system tools. Built from first principles: ZIP archive creation, CRC-32 checksums, and OOXML XML generation are all implemented in a single script.

## Capabilities

### PowerPoint (.pptx)
- *"Create a presentation from this markdown with slides for each section"*
- *"Turn these meeting notes into a slide deck"*
- *"Generate a PPTX with title slide, bullet points, code blocks, and a data table"*

### Word (.docx)
- *"Convert this markdown report to a Word document"*
- *"Create a DOCX with headings, formatted text, numbered lists, and tables"*
- *"Turn this README into a Word document I can email"*

### Custom Themes

Have a corporate template you want to match? Just tell the agent:

- *"Create a presentation from these notes, using the colors and fonts from corporate.pptx"*
- *"Convert this markdown to a Word doc styled like our brand-template.docx"*
- *"Make a slide deck about Q3 results, themed to match marketing.pptx"*

The agent uses the `--template` option under the hood to extract theme colors and fonts from any existing `.pptx` or `.docx` file. Cross-format works — a `.pptx` template can style a `.docx` output.

When using a `.docx` template, the agent also extracts the full `styles.xml` (heading styles, title style, fonts, spacing) and any headers/footers (page numbers, classification labels, etc.) so the output matches the template's look and feel.

### DOCX Title and List Behavior

- The first `#` heading becomes the document **Title** (large, themed). Subsequent `#` headings use Heading 1.
- Each separate list restarts its numbering. Lists separated by headings or paragraphs won't continue counting from a previous list.

### Supported Markdown
- Headings (`#` through `######`)
- Bold (`**text**`), Italic (`*text*`), Bold+Italic (`***text***`)
- Bullet lists (`- item`), Numbered lists (`1. item`) — with nesting up to 3 levels
- Fenced code blocks (` ``` `)
- Tables (`| col | col |`)
- Links (`[text](url)`)
- Images (`![alt](path.png)`) — PNG and JPEG, local files
- Horizontal rules (`---`) — slide break in PPTX, section break in DOCX
- Inline code (`` `code` ``)

## Test Drive

Clone the repo and take it for a spin:

```bash
git clone https://github.com/supermem613/create-office-file-skill
cd create-office-file-skill
copilot   # or: claude
```

Try: *"Create a PowerPoint presentation about the benefits of AI coding assistants"*

## Install

### Claude Code

```claude
/install supermem613/create-office-file-skill
```

### Copilot CLI / Other AI Coding Agents

Copy the skill directory into your project:

**macOS / Linux:**

```bash
git clone https://github.com/supermem613/create-office-file-skill /tmp/create-office-file-skill
cp -r /tmp/create-office-file-skill/.claude/skills/create-office-file .claude/skills/
```

**Windows (PowerShell):**

```powershell
git clone https://github.com/supermem613/create-office-file-skill $env:TEMP\create-office-file-skill
Copy-Item -Recurse $env:TEMP\create-office-file-skill\.claude\skills\create-office-file .claude\skills\
```

The skill is auto-discovered from `.claude/skills/`. No `npm install` required.

## How It Works

The script builds Office files from first principles:

1. **Markdown → AST** — Regex-based parser converts markdown to a structured tree
2. **AST → OOXML XML** — Hard-coded, verified XML skeletons (slide master, theme, styles) with generated content XML
3. **XML → ZIP** — Custom ZIP writer using Node.js `Buffer` + `zlib.deflateRawSync` + CRC-32

No external dependencies. The entire implementation is a single `.mjs` file.

## Evals

Evals verify generated files contain the expected text, formatting, and structure by parsing OOXML XML directly — no Office installation needed, cross-platform, instant.

```
Run evals/run-evals.md
```

Results are written to `evals/results/report.md`.

## Tests

```bash
npm test        # Static validation (no Office required)
```

Tests validate file structure, markdown parsing, ZIP generation, CRC-32 correctness, and OOXML structure — all without needing Office installed.

## Prerequisites

- **Node.js 18+** (the only requirement)

## Not Supported Yet

| Feature | Why | Workaround |
|---------|-----|-----------|
| Nested lists (3+ levels) | 3 levels covers virtually all real usage | Flatten deeper nesting to level 3 |
| PDF output | Different format entirely | Use Office's "Save as PDF" |
| Template slide layouts | Script generates PPTX structure from scratch | Edit generated file in Office |
| Complex tables (merged cells) | OOXML complexity explosion | Simple grid tables only |
| Remote images (URLs) | Keeps converter synchronous and offline | Pre-download images before conversion |
| Nested emphasis | Regex parser limitation | Use single emphasis level per span |
| Reference-style links | Rarely used in AI-generated content | Use inline links |
| HTML-in-markdown | Security and complexity concerns | Use pure markdown |

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, project structure, and how to modify the script, evals, and reference docs.

## License

[MIT](LICENSE)
