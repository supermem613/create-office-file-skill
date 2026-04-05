# Architecture & Design Decisions

## What This Is

A skill that teaches AI agents (Claude Code, GitHub Copilot, Codex) to convert markdown into Office documents (.pptx, .docx). The entire implementation is a single Node.js script with zero external dependencies — ZIP creation, CRC-32 checksums, markdown parsing, and OOXML XML generation are all built from first principles.

## Why From First Principles

| Alternative | Why Not |
|------------|---------|
| pptxgenjs, docx (npm) | External dependencies, supply chain risk, user explicitly requested first principles |
| python-pptx, python-docx | Python dependency, not cross-platform without Python installed |
| LibreOffice CLI | Heavy system dependency, not available on all machines |
| Open XML SDK (.NET) | .NET dependency |

The zero-dependency approach means: `git clone` + `node script.mjs` — nothing else required.

## How Office Files Work

Both `.pptx` and `.docx` are **ZIP archives containing XML files** following the Office Open XML (OOXML/ECMA-376) standard.

### PPTX Structure
```
[Content_Types].xml
_rels/.rels
ppt/presentation.xml
ppt/_rels/presentation.xml.rels
ppt/slides/slide1.xml              ← content goes here
ppt/slides/_rels/slide1.xml.rels
ppt/slideMasters/slideMaster1.xml  ← required skeleton
ppt/slideLayouts/slideLayout1.xml  ← required skeleton
ppt/theme/theme1.xml               ← required skeleton
```

**Critical:** PowerPoint requires slide master + layout + theme or it shows a repair dialog.

### DOCX Structure
```
[Content_Types].xml
_rels/.rels
word/document.xml                   ← content goes here
word/_rels/document.xml.rels
word/styles.xml                     ← heading/code styles
word/numbering.xml                  ← bullet/numbered list defs
```

DOCX is more forgiving — minimum viable is just 3 files.

## Script Architecture

```
┌─────────────────────────────────────────────────┐
│         scripts/create-office-file.mjs              │
│                                                   │
│  ┌──────────┐  ┌────────────┐  ┌──────────────┐ │
│  │ CRC-32   │  │ ZipWriter  │  │ Markdown     │ │
│  │ (25 ln)  │  │ (100 ln)   │  │ Parser       │ │
│  │          │  │ Buffer +   │  │ (250 ln)     │ │
│  │ Lookup   │  │ zlib +     │  │ Regex-based  │ │
│  │ table    │  │ deflateRaw │  │ MD → AST     │ │
│  └──────────┘  └────────────┘  └──────┬───────┘ │
│                      ▲                │          │
│               ZIP data               AST         │
│                      │                │          │
│  ┌──────────────┐    │                │          │
│  │ Image Utils  │    │                │          │
│  │ (65 ln)      │    │                │          │
│  │ PNG/JPEG hdr │    │                │          │
│  │ EMU sizing   │    │                │          │
│  │ Path resolve │    │                │          │
│  └──────┬───────┘    │                │          │
│         │            │                │          │
│         ▼            │                │          │
│         ┌────────────┴───────┐        │          │
│         │                    │        │          │
│  ┌──────┴───────┐  ┌────────┴────┐   │          │
│  │ PPTX Gen     │  │ DOCX Gen    │   │          │
│  │ (400 ln)     │  │ (300 ln)    │◀──┘          │
│  │              │  │             │               │
│  │ Hard-coded:  │  │ Hard-coded: │               │
│  │ • SlideMstr  │  │ • styles    │               │
│  │ • SlideLayout│  │ • numbering │               │
│  │ • Theme      │  │             │               │
│  │              │  │ Generated:  │               │
│  │ Generated:   │  │ • document  │               │
│  │ • slides     │  │   .xml body │               │
│  │ • media/*    │  │ • media/*   │               │
│  └──────────────┘  └─────────────┘               │
│                                                   │
│  ┌──────────────────────────────────────────────┐│
│  │ CLI: -i input.md -o output.pptx|docx         ││
│  └──────────────────────────────────────────────┘│
└─────────────────────────────────────────────────┘
```

## Template-Constrained Design

The key architectural decision (driven by 5-lens evaluation — see POR):

**Don't build a general-purpose OOXML generator.** Instead:

1. **Hard-code known-good XML skeletons** — Slide master, layout, theme, styles, numbering as string constants. These are verified to open without repair dialogs and never change.

2. **Generate only content XML** — Slide content shapes and document body paragraphs. This is the small, testable surface area.

3. **Define a supported markdown subset** — Refuse/ignore unsupported features rather than producing broken output.

This collapses OOXML correctness risk from "infinite spec surface" to "finite, tested templates."

## Verification Architecture (Windows-only)

```
┌──────────────────────────────────────────┐
│         evals/verify-output.mjs           │
│                                           │
│  1. Detect file type (.pptx / .docx)     │
│  2. Spawn PowerShell with COM script     │
│  3. PowerShell opens file in Office app  │
│  4. COM API extracts:                    │
│     • Text per slide/paragraph           │
│     • Formatting (bold, italic, font)    │
│     • Slide/paragraph count              │
│  5. PowerShell outputs JSON to stdout    │
│  6. Node parses and returns structured   │
│     verification result                  │
└──────────────────────────────────────────┘
```

Uses `Word.Application` and `PowerPoint.Application` COM objects via PowerShell. Gracefully errors with a clear message on non-Windows systems or when Office is not installed.

## Design Decisions Log

| Decision | Chosen | Why |
|----------|--------|-----|
| Zero deps vs npm libs | Zero deps | User requirement; ZIP + CRC-32 are simple enough |
| Single file vs modules | Single file | Skill portability — one `.mjs` to copy |
| Template-first vs general OOXML | Template-first | 4/5 evaluation lenses converged on this |
| PPTX slide splitting | `# heading` = title slide, `## heading` = content slide | Mirrors how humans write slides in markdown |
| DOCX heading mapping | `# = Heading1` through `######` = Heading6 | Standard convention |
| Verification approach | PowerShell COM | Uses actual Office apps — the ground truth |
| Eval platform | Windows-only | COM automation requires Office; documented clearly |

