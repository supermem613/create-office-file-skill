# Contributing

Guide for engineers working on the Create Office File Skill.

---

## Quick Start

```bash
git clone https://github.com/supermem613/create-office-file-skill
cd create-office-file-skill
npm test
```

**Prerequisites:** Node.js 18+. No `npm install` needed — the skill has zero npm dependencies.

## Project Structure

```
.claude/skills/create-office-file/
  SKILL.md                    # Agent-facing skill definition (~2K tokens)
  scripts/
    create-office-file.mjs    # Single script — ZIP + CRC-32 + MD parser + OOXML
  references/                 # OOXML domain docs (lazy-loaded by agent)
    pptx.md                   #   PresentationML reference
    docx.md                   #   WordprocessingML reference
docs/                         # Human-facing documentation
  architecture.md             #   Design decisions and diagrams
evals/
  run-evals.md                # 18 evals — agent-executable spec (Windows-only)
  verify-output.mjs           # Office COM verification (Windows-only)
tests/
  test-scripts.js             # Static validation (no Office required)
```

## Architecture

This is a **skill**, not an MCP server. The agent reads `SKILL.md`, learns the CLI interface, and calls `scripts/create-office-file.mjs` directly. No runtime server, no tool registration.

**Zero dependencies.** The script uses only Node.js built-in modules (`fs`, `path`, `zlib`). ZIP archive creation, CRC-32, markdown parsing, and OOXML XML generation are all implemented from first principles.

**Token budget:** `SKILL.md` is kept small (~2K tokens). The 2 reference files in `references/` are loaded on demand by the agent only when needed for debugging or extending functionality.

See [`docs/architecture.md`](docs/architecture.md) for diagrams and design decisions.

## Development Workflow

### Running Tests

```bash
npm test        # Static validation — no Office, no network
```

Static tests (`test-scripts.js`) validate:
- File existence and shebang line
- CRC-32 correctness against known values
- ZIP structure (valid archive with correct entries)
- Markdown parser (headings, bold, italic, lists, code blocks, tables)
- PPTX structure (slide master, layout, theme, correct slide count)
- DOCX structure (document.xml, styles.xml, numbering.xml)
- No external npm dependencies in the main script
- CLI error handling (missing args, bad format)

### Running Evals

Evals test generated files against real Microsoft Office applications using PowerShell COM automation. **Windows-only** — requires Word and PowerPoint installed.

```
Run evals/run-evals.md
```

Results are written to `evals/results/report.md`.

### Verifying Output Manually

Generate test files and open them:

```bash
node .claude/skills/create-office-file/scripts/create-office-file.mjs -i test-input.md -o test.pptx
node .claude/skills/create-office-file/scripts/create-office-file.mjs -i test-input.md -o test.docx
```

Check that they open in PowerPoint/Word without a repair dialog and that formatting renders correctly.

## Modifying the Script

The entire implementation lives in `.claude/skills/create-office-file/scripts/create-office-file.mjs`. Rules:

1. **No npm dependencies.** Only Node.js built-ins (`fs`, `path`, `zlib`). The test suite enforces this.
2. **Cross-platform.** No shell dependencies, no OS-specific paths. Node.js only.
3. **Hard-coded XML skeletons.** The OOXML templates (slide master, theme, styles, numbering) are string constants in the script. They are verified-correct and should only change if Office compatibility requires it.
4. **Content generation is separate.** The functions that convert AST nodes to XML are isolated from the skeleton templates. Add new markdown features by adding AST node types and their corresponding XML generators.

### Internal Architecture

| Section | Purpose | ~Lines |
|---------|---------|--------|
| CRC-32 | Lookup table checksum | 25 |
| ZipWriter | ZIP archive from Buffer + zlib | 100 |
| XML utilities | Escaping | 5 |
| Markdown parser | Regex-based MD → AST | 250 |
| PPTX generator | PresentationML + hard-coded skeletons | 400 |
| DOCX generator | WordprocessingML + hard-coded skeletons | 300 |
| CLI | Arg parsing, stdin/file, format detection | 50 |

## Modifying SKILL.md

`SKILL.md` is what the agent reads. It's the most sensitive file in the repo — small changes affect every agent interaction.

- Keep it under ~2K tokens. Move detailed OOXML docs to `references/`.
- Test changes by invoking the skill from Claude Code and observing agent behavior.

## Modifying Reference Files

The 2 files in `references/` are loaded on demand when the agent needs OOXML-specific knowledge (typically for debugging or extending the script).

| File | Covers |
|------|--------|
| `pptx.md` | PresentationML structure, DrawingML text formatting, slide shapes, tables |
| `docx.md` | WordprocessingML structure, paragraph/run formatting, lists, tables |

## Adding Evals

Evals live in `evals/run-evals.md`. Each eval has:

- A number and name
- The markdown input (inline or file reference)
- The command to generate the output
- The verification command (COM-based text/formatting extraction)
- A pass condition

All generated test files use the `EVAL_TEST_` prefix and are cleaned up after the run.

## Code Style

- `'use strict'` at the top of the script
- Shebang line (`#!/usr/bin/env node`) on every script
- Errors go to stderr, file output goes to the file system
- Silence means success — no verbose output by default
- No comments explaining obvious code. Comment intent, not mechanics.
