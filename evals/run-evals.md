# Create Office File Skill — Evals

Evals verify that generated `.pptx` and `.docx` files open correctly in Microsoft Office and contain the expected text and formatting.

## Platform Requirement

**Windows-only.** Evals use PowerShell COM automation (`Word.Application`, `PowerPoint.Application`) to open files in actual Microsoft Office and inspect content. Requires:
- Windows 10/11
- Microsoft Office (Word + PowerPoint) installed
- Node.js 18+

On non-Windows systems or without Office, evals will report a skip with a clear message.

## How to Run

```
Run evals/run-evals.md
```

## Execution Model

The agent generates Office files with `create-office-file.mjs` (in the skill), then verifies them with `verify-output.mjs` (in evals).

**Script paths** (relative to repo root):
- Generator: `node .claude/skills/create-office-file/scripts/create-office-file.mjs`
- Verifier: `node evals/verify-output.mjs`

For brevity in each eval below:
- `$GEN` = `.claude/skills/create-office-file/scripts/create-office-file.mjs`
- `$VER` = `evals/verify-output.mjs`

**How to run each eval:**
1. Write a markdown file to a temp location (or use inline stdin)
2. Generate the Office file: `node $GEN -i <input.md> -o <output>`
3. Verify with COM: `node $VER <output>`
4. Parse the JSON output and check the pass condition
5. Clean up temp files

**Temporary files:** Write all eval files to `os.tmpdir()`, not inside the repo. Prefix with `EVAL_TEST_`.

**Pass/Fail:** `verify-output.mjs` outputs JSON with extracted text and formatting. Check the JSON fields against the pass conditions below. Exit code 0 = verification succeeded, non-zero = error.

## Setup

No auth or external services needed. All evals are local and offline.

## Scoring

For each eval: generate, verify, check pass condition.
- ✅ **PASS** — file generated, COM verification succeeded, pass condition met
- ❌ **FAIL** — generation failed, COM verification failed, or pass condition not met
- ⏭️ **SKIP** — not on Windows or Office not installed

After all evals, write the report to `evals/results/report.md`.

---

## PPTX — Structure (3)

### 01 — Title slide
**Input:** `# My Presentation Title`
**Generate:** `node $GEN -o EVAL_TEST_01.pptx` (stdin)
**Verify:** `node $VER EVAL_TEST_01.pptx`
**Pass if:** JSON `.slideCount` is 1 AND any slide shape `.text` contains `My Presentation Title`

### 02 — Multiple slides from headings
**Input:**
```markdown
# Title
## Slide One
Content A
## Slide Two
Content B
```
**Generate:** Write input to `EVAL_TEST_02.md`, then `node $GEN -i EVAL_TEST_02.md -o EVAL_TEST_02.pptx`
**Verify:** `node $VER EVAL_TEST_02.pptx`
**Pass if:** JSON `.slideCount` is 3

### 03 — Slide break with ---
**Input:**
```markdown
## Part 1
A

---

## Part 2
B
```
**Generate & verify as above**
**Pass if:** JSON `.slideCount` is 2

---

## PPTX — Formatting (5)

### 04 — Bold text
**Input:** `## Test\n**Bold words here**`
**Pass if:** Any run in slide 1 has `.bold` = true AND `.text` contains `Bold words here`

### 05 — Italic text
**Input:** `## Test\n*Italic words here*`
**Pass if:** Any run has `.italic` = true AND `.text` contains `Italic words here`

### 06 — Bullet list
**Input:** `## Test\n- Alpha\n- Beta\n- Gamma`
**Pass if:** Slide text contains `Alpha`, `Beta`, `Gamma`

### 07 — Numbered list
**Input:** `## Test\n1. First\n2. Second\n3. Third`
**Pass if:** Slide text contains `First`, `Second`, `Third`

### 08 — Code block with monospace font
**Input:**
````markdown
## Test
```
const x = 42;
```
````
**Pass if:** Any run has `.fontName` containing `Consolas` AND `.text` contains `const x = 42`

---

## PPTX — Content (2)

### 09 — Table
**Input:**
```markdown
## Data
| Name | Age |
|------|-----|
| Alice | 30 |
| Bob | 25 |
```
**Pass if:** Slide text contains `Alice` AND `Bob` AND `Name` AND `Age`

### 10 — Mixed content slide
**Input:**
```markdown
## Overview
Here is **bold** and *italic* text.

- Point one
- Point two

And a [link](https://example.com).
```
**Pass if:** Slide text contains `bold`, `italic`, `Point one`, `Point two`, `link`

---

## DOCX — Structure (2)

### 11 — Headings map to styles
**Input:**
```markdown
# Heading One
## Heading Two
### Heading Three
Paragraph text.
```
**Generate:** `node $GEN -i EVAL_TEST_11.md -o EVAL_TEST_11.docx`
**Verify:** `node $VER EVAL_TEST_11.docx`
**Pass if:** Paragraphs include one with `.style` containing `Heading 1` and text `Heading One`, one with style `Heading 2` and text `Heading Two`, one with style `Heading 3` and text `Heading Three`

### 12 — Paragraph count
**Input:** `# Title\nPara one.\n\nPara two.\n\nPara three.`
**Pass if:** JSON `.paragraphCount` >= 4 (title + 3 paragraphs)

---

## DOCX — Formatting (4)

### 13 — Bold text
**Input:** `**This is bold**`
**Pass if:** Any run has `.bold` = true AND `.text` contains `This is bold` (or word fragments thereof)

### 14 — Italic text
**Input:** `*This is italic*`
**Pass if:** Any run has `.italic` = true

### 15 — Bullet list
**Input:** `- Apple\n- Banana\n- Cherry`
**Pass if:** Paragraph text contains `Apple`, `Banana`, `Cherry`
**Note:** Word COM reports list paragraphs with `List Paragraph` or `List Bullet` style — check style name contains `List` or the paragraph text is correct.

### 16 — Code block
**Input:**
````markdown
```
function hello() { return 42; }
```
````
**Pass if:** Any run has `.fontName` containing `Consolas`

---

## DOCX — Content (2)

### 17 — Table
**Input:**
```markdown
| City | Pop |
|------|-----|
| NYC | 8M |
| LA | 4M |
```
**Pass if:** Document text contains `City`, `Pop`, `NYC`, `LA`

### 18 — Full document
**Input:**
```markdown
# Report Title

## Introduction
This report covers **important** findings.

## Key Points
1. First finding
2. Second finding
3. Third finding

## Data

| Metric | Value |
|--------|-------|
| Score | 95 |
| Grade | A |

## Conclusion
*Thank you* for reading.
```
**Pass if:** All of: `Report Title`, `Introduction`, `important`, `First finding`, `Score`, `Grade`, `Conclusion`, `Thank you` appear in document text

---

## PPTX — Images (2)

### 19 — PPTX with embedded PNG
**Input:**
```markdown
## Photo Slide

![Test Image](test-image.png)
```
**Setup:** Copy `tests/test-image.png` to the temp directory next to the markdown file.
**Generate:** Write input to `EVAL_TEST_19.md` in temp dir (with test-image.png alongside), then `node $GEN -i EVAL_TEST_19.md -o EVAL_TEST_19.pptx`
**Verify:** `node $VER EVAL_TEST_19.pptx`
**Pass if:** JSON has no `.error`, AND any shape in any slide has `.isPicture` = true OR `.shapeType` = 13

### 20 — PPTX with image and text
**Input:**
```markdown
## Mixed Slide
Here is some text content.

![Chart](test-image.png)
```
**Setup:** Same as 19 — ensure test-image.png is alongside the markdown file.
**Generate & verify as above**
**Pass if:** Any slide has a shape with `.isPicture` = true AND another shape with `.hasText` = true containing `some text content`

---

## DOCX — Images (2)

### 21 — DOCX with embedded PNG
**Input:**
```markdown
# Document with Image

![Test Image](test-image.png)
```
**Setup:** Copy `tests/test-image.png` to temp dir next to the markdown file.
**Generate:** `node $GEN -i EVAL_TEST_21.md -o EVAL_TEST_21.docx`
**Verify:** `node $VER EVAL_TEST_21.docx`
**Pass if:** JSON `.inlineShapeCount` >= 1

### 22 — DOCX with image and surrounding text
**Input:**
```markdown
# Report

First paragraph of text.

![Data Chart](test-image.png)

Second paragraph with conclusions.
```
**Setup:** Same as 21.
**Generate & verify as above**
**Pass if:** JSON `.inlineShapeCount` >= 1 AND paragraph text contains `First paragraph` AND `Second paragraph`

---

## Report

After completing all evals, write `evals/results/report.md`:

```
# Eval Report — [date]

**Platform:** Windows [version]
**Office:** [Word/PowerPoint version]
**Overall:** [passed]/22 ([percentage]%) — [failed] failed, [skipped] skipped

## Summary

| Category           | Pass | Fail | Skip | Total |
|--------------------|------|------|------|-------|
| PPTX Structure     |      |      |      | 3     |
| PPTX Formatting    |      |      |      | 5     |
| PPTX Content       |      |      |      | 2     |
| PPTX Images        |      |      |      | 2     |
| DOCX Structure     |      |      |      | 2     |
| DOCX Formatting    |      |      |      | 4     |
| DOCX Content       |      |      |      | 2     |
| DOCX Images        |      |      |      | 2     |

## Results

| #  | Eval                  | Score | Notes |
|----|-----------------------|-------|-------|
| 01 | Title slide           |       |       |
| 02 | Multiple slides       |       |       |
| 03 | Slide break           |       |       |
| 04 | Bold text (PPTX)      |       |       |
| 05 | Italic text (PPTX)    |       |       |
| 06 | Bullet list (PPTX)    |       |       |
| 07 | Numbered list (PPTX)  |       |       |
| 08 | Code block (PPTX)     |       |       |
| 09 | Table (PPTX)          |       |       |
| 10 | Mixed content (PPTX)  |       |       |
| 11 | Heading styles (DOCX) |       |       |
| 12 | Paragraph count       |       |       |
| 13 | Bold text (DOCX)      |       |       |
| 14 | Italic text (DOCX)    |       |       |
| 15 | Bullet list (DOCX)    |       |       |
| 16 | Code block (DOCX)     |       |       |
| 17 | Table (DOCX)          |       |       |
| 18 | Full document (DOCX)  |       |       |
| 19 | Image (PPTX)          |       |       |
| 20 | Image + text (PPTX)   |       |       |
| 21 | Image (DOCX)          |       |       |
| 22 | Image + text (DOCX)   |       |       |

## Failures
[Details for any ❌]
```
