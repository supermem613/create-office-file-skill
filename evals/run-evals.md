# Create Office File Skill — Evals

Evals verify that the `create-office-file` skill produces correct `.pptx` and `.docx` files.

## Collateral

All eval inputs are pre-built in `evals/inputs/`. Image evals reference `test-image.png` which is also in `evals/inputs/` alongside the markdown files. Template evals use `evals/create-template.mjs` and `evals/create-docx-template.mjs` to generate templates.

## Numbering Scheme

| Range   | Purpose    |
|---------|------------|
| 100–199 | DOCX evals |
| 200–299 | PPTX evals |

## Verifier

`evals/verify-output.mjs` parses OOXML XML directly — no COM, no Office, cross-platform, ~80ms per file.

```
node evals/verify-output.mjs <file.pptx|docx>
```

Outputs JSON with text, formatting, structure, theme, and header/footer info.

`evals/verify-pass-conditions.mjs` checks all result files against pass conditions and prints a summary table.

```
node evals/verify-pass-conditions.mjs
```

## How to Run

```
Run evals/run-evals.md
```

## Execution Model

For each eval below:
1. Send the **Prompt** to a sub-agent using the exact sub-agent prompt template below
2. Run `node evals/verify-output.mjs` on the generated file
3. Evaluate the JSON output against the **Pass condition**
4. Record ✅ PASS or ❌ FAIL

Output files go to `evals/results/`.

### Sub-agent prompt template

Every eval MUST be executed by a sub-agent (general-purpose task agent) that receives
ONLY the following prompt — nothing else. The sub-agent must discover how to use the
skill from the skill interface alone. **Do NOT** pass repo knowledge, script paths,
or implementation details to the sub-agent.

```
You have access to the "create-office-file" skill.
Invoke the skill, then follow its instructions to complete this task:

{eval prompt here}

RULES:
- You MUST invoke the create-office-file skill to learn how to complete this task.
- Do NOT look inside .claude/skills/*/scripts/ or read any .mjs files.
- Do NOT run any script directly unless the skill's own instructions tell you to.
```

If an eval has a **Setup** step, the runner (not the sub-agent) executes setup
before launching the sub-agent.

---

## DOCX — Structure (2)

### 100 — Headings map to styles
**Prompt:** Create a Word document from `evals/inputs/100.md` and save it to `evals/results/100.docx`
**Pass if:** Paragraphs include style `Title` or `Heading 1` with text `Heading One`, style `Heading 2` with `Heading Two`, style `Heading 3` with `Heading Three`

### 101 — Paragraph count
**Prompt:** Create a Word document from `evals/inputs/101.md` and save it to `evals/results/101.docx`
**Pass if:** `.paragraphCount` >= 4

---

## DOCX — Formatting (4)

### 102 — Bold text
**Prompt:** Create a Word document from `evals/inputs/102.md` and save it to `evals/results/102.docx`
**Pass if:** Any run has `.bold` == true

### 103 — Italic text
**Prompt:** Create a Word document from `evals/inputs/103.md` and save it to `evals/results/103.docx`
**Pass if:** Any run has `.italic` == true

### 104 — Bullet list
**Prompt:** Create a Word document from `evals/inputs/104.md` and save it to `evals/results/104.docx`
**Pass if:** Paragraph text contains `Apple`, `Banana`, `Cherry`

### 105 — Code block
**Prompt:** Create a Word document from `evals/inputs/105.md` and save it to `evals/results/105.docx`
**Pass if:** Any run has `.fontName` containing `Consolas`

---

## DOCX — Content (2)

### 106 — Table
**Prompt:** Create a Word document from `evals/inputs/106.md` and save it to `evals/results/106.docx`
**Pass if:** Document text contains `City`, `Pop`, `NYC`, `LA`

### 107 — Full document
**Prompt:** Create a Word document from `evals/inputs/107.md` and save it to `evals/results/107.docx`
**Pass if:** Document text contains all of: `Report Title`, `Introduction`, `important`, `First finding`, `Score`, `Grade`, `Conclusion`, `Thank you`

---

## DOCX — Images (2)

### 108 — DOCX with embedded PNG
**Prompt:** Create a Word document from `evals/inputs/108.md` and save it to `evals/results/108.docx`
**Pass if:** `.inlineShapeCount` >= 1

### 109 — DOCX with image and surrounding text
**Prompt:** Create a Word document from `evals/inputs/109.md` and save it to `evals/results/109.docx`
**Pass if:** `.inlineShapeCount` >= 1 AND paragraph text contains `First paragraph` AND `Second paragraph`

---

## DOCX — Nested Lists (1)

### 110 — Nested mixed list
**Prompt:** Create a Word document from `evals/inputs/110.md` and save it to `evals/results/110.docx`
**Pass if:** Paragraph text contains all of: `Bullet level 0`, `Bullet level 1`, `Bullet level 2`, `Back to level 0`, `Number level 0`, `Number level 1`, `Number level 2`

---

## DOCX — Custom Theme (2)

### 111 — DOCX with theme
**Prompt:** Create a Word document from `evals/inputs/111.md` and save it to `evals/results/111.docx`
**Pass if:** No `.error`; paragraph text contains `Document with Theme` and `Paragraph with`

### 112 — DOCX from template has custom styles
**Setup:** Run `node evals/create-template.mjs evals/results/eval27_template.pptx` first.
**Prompt:** Create a Word document from `evals/inputs/112.md` using `evals/results/eval27_template.pptx` as template, and save it to `evals/results/112.docx`
**Pass if:** No `.error`; `.theme.fonts.minor` == `Verdana` AND `.theme.colors.accent1` == `E94560`

---

## DOCX — Title Style (1)

### 113 — First H1 uses Title style
**Prompt:** Create a Word document from `evals/inputs/113.md` and save it to `evals/results/113.docx`
**Pass if:** First paragraph has `.style` == `Title` AND text contains `My Document Title`; a later paragraph has `.style` containing `Heading 1` AND text contains `Another Top-Level Heading`

---

## DOCX — List Restart (2)

### 114 — Ordered lists restart numbering
**Prompt:** Create a Word document from `evals/inputs/114.md` and save it to `evals/results/114.docx`
**Pass if:** Document text contains `Ship built-in skills`, `Skill marketplace`, `OOB skills visible`

### 115 — Bullet lists are independent
**Prompt:** Create a Word document from `evals/inputs/115.md` and save it to `evals/results/115.docx`
**Pass if:** Paragraph text contains `Alpha`, `Beta`, `Gamma`, `Delta`

---

## DOCX — Template Styles (1)

### 116 — DOCX from template preserves heading styles
**Setup:** Ensure `evals/results/eval27_template.pptx` exists (same template as eval 27).
**Prompt:** Create a Word document from `evals/inputs/116.md` using `evals/results/eval27_template.pptx` as template, and save it to `evals/results/116.docx`
**Pass if:** No `.error`; first H1 paragraph has `.style` == `Title`

---

## DOCX — Headers & Footers (1)

### 117 — DOCX from template preserves footer
**Setup:** Run `node evals/create-docx-template.mjs evals/results/eval100_template.docx` to create a DOCX template with a footer.
**Prompt:** Create a Word document from `evals/inputs/117.md` using `evals/results/eval100_template.docx` as template, and save it to `evals/results/117.docx`
**Pass if:** No `.error`; `.headersFooters` has an entry with `.name` containing `footer` AND `.text` containing `CONFIDENTIAL`

---

## PPTX — Structure (3)

### 200 — Title slide
**Prompt:** Create a PowerPoint from `evals/inputs/200.md` and save it to `evals/results/200.pptx`
**Pass if:** `.slideCount` == 1 AND any shape `.text` contains `My Presentation Title`

### 201 — Multiple slides from headings
**Prompt:** Create a PowerPoint from `evals/inputs/201.md` and save it to `evals/results/201.pptx`
**Pass if:** `.slideCount` == 3

### 202 — Slide break with ---
**Prompt:** Create a PowerPoint from `evals/inputs/202.md` and save it to `evals/results/202.pptx`
**Pass if:** `.slideCount` == 2

---

## PPTX — Formatting (5)

### 203 — Bold text
**Prompt:** Create a PowerPoint from `evals/inputs/203.md` and save it to `evals/results/203.pptx`
**Pass if:** Any run has `.bold` == true AND `.text` contains `Bold words here`

### 204 — Italic text
**Prompt:** Create a PowerPoint from `evals/inputs/204.md` and save it to `evals/results/204.pptx`
**Pass if:** Any run has `.italic` == true AND `.text` contains `Italic words here`

### 205 — Bullet list
**Prompt:** Create a PowerPoint from `evals/inputs/205.md` and save it to `evals/results/205.pptx`
**Pass if:** Slide text contains `Alpha`, `Beta`, `Gamma`

### 206 — Numbered list
**Prompt:** Create a PowerPoint from `evals/inputs/206.md` and save it to `evals/results/206.pptx`
**Pass if:** Slide text contains `First`, `Second`, `Third`

### 207 — Code block with monospace font
**Prompt:** Create a PowerPoint from `evals/inputs/207.md` and save it to `evals/results/207.pptx`
**Pass if:** Any run has `.fontName` containing `Consolas` AND `.text` contains `const x = 42`

---

## PPTX — Content (2)

### 208 — Table
**Prompt:** Create a PowerPoint from `evals/inputs/208.md` and save it to `evals/results/208.pptx`
**Pass if:** Slide text contains `Alice`, `Bob`, `Name`, `Age`

### 209 — Mixed content slide
**Prompt:** Create a PowerPoint from `evals/inputs/209.md` and save it to `evals/results/209.pptx`
**Pass if:** Slide text contains `bold`, `italic`, `Point one`, `Point two`, `link`

---

## PPTX — Images (2)

### 210 — PPTX with embedded PNG
**Prompt:** Create a PowerPoint from `evals/inputs/210.md` and save it to `evals/results/210.pptx`
**Pass if:** Any shape has `.isPicture` == true OR `.shapeType` == 13

### 211 — PPTX with image and text
**Prompt:** Create a PowerPoint from `evals/inputs/211.md` and save it to `evals/results/211.pptx`
**Pass if:** Any shape has `.isPicture` == true AND another shape `.text` contains `some text content`

---

## PPTX — Nested Lists (1)

### 212 — Nested bullet list
**Prompt:** Create a PowerPoint from `evals/inputs/212.md` and save it to `evals/results/212.pptx`
**Pass if:** Shape text contains `Top level item`, `Second level item`, `Third level item`, `Back to top`

---

## PPTX — Custom Theme (2)

### 213 — PPTX with default theme
**Prompt:** Create a PowerPoint from `evals/inputs/213.md` and save it to `evals/results/213.pptx`
**Pass if:** No `.error`; shape text contains `Default Theme Test` and `Some content`

### 214 — PPTX from template has custom colors
**Setup:** Run `node evals/create-template.mjs evals/results/eval27_template.pptx` first.
**Prompt:** Create a PowerPoint from `evals/inputs/214.md` using `evals/results/eval27_template.pptx` as template, and save it to `evals/results/214.pptx`
**Pass if:** No `.error`; `.theme.colors.accent1` == `E94560` AND `.theme.fonts.major` == `Georgia`; `.theme.colors.accent1` != `4472C4`

---

## Report

After all evals, run `node evals/verify-pass-conditions.mjs` and write `evals/results/report.md` with the summary table and per-eval results.
