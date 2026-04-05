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

Evals verify generated files open correctly in Microsoft Office and contain the expected content. They use PowerShell COM automation to open files in the actual Word and PowerPoint applications.

```
Run evals/run-evals.md
```

Results are written to `evals/results/report.md`.

> **Note:** Evals require Windows with Microsoft Office installed (Word + PowerPoint). See [`evals/run-evals.md`](evals/run-evals.md) for details.

## Tests

```bash
npm test        # Static validation (no Office required)
```

Tests validate file structure, markdown parsing, ZIP generation, CRC-32 correctness, and OOXML structure — all without needing Office installed.

## Prerequisites

- **Node.js 18+** (the only requirement)
- **Microsoft Office** (Word + PowerPoint) — only for evals, not for usage

## Not Supported Yet

| Feature | Why | Workaround |
|---------|-----|-----------|
| Nested lists (3+ levels) | 3 levels covers virtually all real usage | Flatten deeper nesting to level 3 |
| Custom themes/colors | Hard-coded theme covers 95% of cases | Edit theme constant in script |
| PDF output | Different format entirely | Use Office's "Save as PDF" |
| Template-based editing | Script generates from scratch only | Edit generated file in Office |
| Complex tables (merged cells) | OOXML complexity explosion | Simple grid tables only |
| Remote images (URLs) | Keeps converter synchronous and offline | Pre-download images before conversion |

## Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development setup, project structure, and how to modify the script, evals, and reference docs.

## License

[MIT](LICENSE)
