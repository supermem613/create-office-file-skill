#!/usr/bin/env node
// verify-pass-conditions.mjs — Check all eval results against pass conditions
// Run after generating files: node evals/verify-pass-conditions.mjs
// Assumes output files exist in evals/results/
'use strict';

import { execSync } from 'child_process';
import { existsSync, readFileSync } from 'fs';
import { resolve } from 'path';

function verify(file) {
  const abs = resolve(file);
  if (!existsSync(abs)) return { error: `File not found: ${file}` };
  try {
    const out = execSync(`node evals/verify-output.mjs "${abs}"`, { encoding: 'utf8', timeout: 10000 });
    return JSON.parse(out);
  } catch (e) {
    return { error: e.message };
  }
}

function allText(j) {
  return JSON.stringify(j);
}

function hasRun(runs, pred) {
  if (!runs) return false;
  return runs.some(pred);
}

function flatRuns(j) {
  if (j.type === 'pptx') {
    const runs = [];
    for (const s of j.slides || [])
      for (const sh of s.shapes || [])
        for (const p of sh.paragraphs || [])
          for (const r of p.runs || [])
            runs.push(r);
    return runs;
  }
  const runs = [];
  for (const p of j.paragraphs || [])
    for (const r of p.runs || [])
      runs.push(r);
  return runs;
}

// Each eval: { file, check(j) => { pass, notes } }
const evals = [
  // DOCX — Structure
  { id: '100', name: 'Heading styles (DOCX)', file: 'evals/results/100.docx',
    check: j => {
      const ps = j.paragraphs || [];
      const h1 = ps.find(p => p.text.includes('Heading One'));
      const h2 = ps.find(p => p.text.includes('Heading Two'));
      const h3 = ps.find(p => p.text.includes('Heading Three'));
      const hasH1 = h1 && (h1.style === 'Title' || h1.style.toLowerCase().includes('heading 1'));
      const hasH2 = h2 && h2.style.toLowerCase().includes('heading 2');
      const hasH3 = h3 && h3.style.toLowerCase().includes('heading 3');
      return { pass: hasH1 && hasH2 && hasH3, notes: `H1=${h1?.style}, H2=${h2?.style}, H3=${h3?.style}` };
    }},
  { id: '101', name: 'Paragraph count (DOCX)', file: 'evals/results/101.docx',
    check: j => ({ pass: j.paragraphCount >= 4, notes: `paragraphCount=${j.paragraphCount}` })},
  // DOCX — Formatting
  { id: '102', name: 'Bold text (DOCX)', file: 'evals/results/102.docx',
    check: j => {
      const pass = hasRun(flatRuns(j), r => r.bold);
      return { pass, notes: 'bold found' };
    }},
  { id: '103', name: 'Italic text (DOCX)', file: 'evals/results/103.docx',
    check: j => {
      const pass = hasRun(flatRuns(j), r => r.italic);
      return { pass, notes: 'italic found' };
    }},
  { id: '104', name: 'Bullet list (DOCX)', file: 'evals/results/104.docx',
    check: j => {
      const t = allText(j);
      const pass = t.includes('Apple') && t.includes('Banana') && t.includes('Cherry');
      return { pass, notes: 'bullet items' };
    }},
  { id: '105', name: 'Code block (DOCX)', file: 'evals/results/105.docx',
    check: j => {
      const pass = hasRun(flatRuns(j), r => r.fontName?.includes('Consolas'));
      return { pass, notes: 'code font' };
    }},
  // DOCX — Content
  { id: '106', name: 'Table (DOCX)', file: 'evals/results/106.docx',
    check: j => {
      const t = allText(j);
      const pass = t.includes('City') && t.includes('Pop') && t.includes('NYC') && t.includes('LA');
      return { pass, notes: 'table content' };
    }},
  { id: '107', name: 'Full document (DOCX)', file: 'evals/results/107.docx',
    check: j => {
      const t = allText(j);
      const words = ['Report Title','Introduction','important','First finding','Score','Grade','Conclusion','Thank you'];
      const pass = words.every(w => t.includes(w));
      return { pass, notes: `${words.filter(w => t.includes(w)).length}/${words.length} words` };
    }},
  // DOCX — Images
  { id: '108', name: 'Image (DOCX)', file: 'evals/results/108.docx',
    check: j => ({ pass: j.inlineShapeCount >= 1, notes: `inlineShapeCount=${j.inlineShapeCount}` })},
  { id: '109', name: 'Image + text (DOCX)', file: 'evals/results/109.docx',
    check: j => {
      const t = allText(j);
      const pass = j.inlineShapeCount >= 1 && t.includes('First paragraph') && t.includes('Second paragraph');
      return { pass, notes: 'image+text' };
    }},
  // DOCX — Nested Lists
  { id: '110', name: 'Nested list (DOCX)', file: 'evals/results/110.docx',
    check: j => {
      const t = allText(j);
      const words = ['Bullet level 0','Bullet level 1','Bullet level 2','Back to level 0','Number level 0','Number level 1','Number level 2'];
      const pass = words.every(w => t.includes(w));
      return { pass, notes: 'nested mixed' };
    }},
  // DOCX — Custom Theme
  { id: '111', name: 'Theme (DOCX)', file: 'evals/results/111.docx',
    check: j => {
      const t = allText(j);
      const pass = !j.error && t.includes('Document with Theme') && t.includes('Paragraph with');
      return { pass, notes: 'docx theme' };
    }},
  { id: '112', name: 'Template theme (DOCX)', file: 'evals/results/112.docx',
    check: j => {
      const pass = !j.error
        && j.theme?.fonts?.minor === 'Verdana'
        && j.theme?.colors?.accent1 === 'E94560';
      return { pass, notes: `minor=${j.theme?.fonts?.minor}, accent1=${j.theme?.colors?.accent1}` };
    }},
  // DOCX — Title Style
  { id: '113', name: 'Title style (DOCX)', file: 'evals/results/113.docx',
    check: j => {
      const ps = j.paragraphs || [];
      const first = ps.find(p => p.text.includes('My Document Title'));
      const second = ps.find(p => p.text.includes('Another Top-Level Heading'));
      const pass = first?.style === 'Title' && second?.style?.toLowerCase().includes('heading 1');
      return { pass, notes: `first=${first?.style}, second=${second?.style}` };
    }},
  // DOCX — List Restart
  { id: '114', name: 'List restart (DOCX)', file: 'evals/results/114.docx',
    check: j => {
      const t = allText(j);
      const pass = t.includes('Ship built-in skills') && t.includes('Skill marketplace') && t.includes('OOB skills visible');
      return { pass, notes: 'list restart' };
    }},
  { id: '115', name: 'Bullet independent (DOCX)', file: 'evals/results/115.docx',
    check: j => {
      const t = allText(j);
      const pass = t.includes('Alpha') && t.includes('Beta') && t.includes('Gamma') && t.includes('Delta');
      return { pass, notes: 'bullet indep' };
    }},
  // DOCX — Template Styles
  { id: '116', name: 'Template styles (DOCX)', file: 'evals/results/116.docx',
    check: j => {
      const ps = j.paragraphs || [];
      const titlePara = ps.find(p => p.style === 'Title');
      const pass = !j.error && titlePara != null;
      return { pass, notes: `title style=${titlePara?.style || 'none'}` };
    }},
  // DOCX — Headers & Footers
  { id: '117', name: 'Footer from template (DOCX)', file: 'evals/results/117.docx',
    check: j => {
      const hf = j.headersFooters || [];
      const footer = hf.find(h => h.name.includes('footer'));
      const pass = !j.error && footer != null && footer.text.includes('CONFIDENTIAL');
      return { pass, notes: footer ? `footer="${footer.text}"` : 'no footer' };
    }},
  // PPTX — Structure
  { id: '200', name: 'Title slide', file: 'evals/results/200.pptx',
    check: j => {
      const pass = j.slideCount === 1 && allText(j).includes('My Presentation Title');
      return { pass, notes: `slideCount=${j.slideCount}` };
    }},
  { id: '201', name: 'Multiple slides', file: 'evals/results/201.pptx',
    check: j => ({ pass: j.slideCount === 3, notes: `slideCount=${j.slideCount}` })},
  { id: '202', name: 'Slide break', file: 'evals/results/202.pptx',
    check: j => ({ pass: j.slideCount === 2, notes: `slideCount=${j.slideCount}` })},
  // PPTX — Formatting
  { id: '203', name: 'Bold text (PPTX)', file: 'evals/results/203.pptx',
    check: j => {
      const runs = flatRuns(j);
      const pass = hasRun(runs, r => r.bold && r.text.includes('Bold words here'));
      return { pass, notes: 'bold runs' };
    }},
  { id: '204', name: 'Italic text (PPTX)', file: 'evals/results/204.pptx',
    check: j => {
      const runs = flatRuns(j);
      const pass = hasRun(runs, r => r.italic && r.text.includes('Italic words here'));
      return { pass, notes: 'italic runs' };
    }},
  { id: '205', name: 'Bullet list (PPTX)', file: 'evals/results/205.pptx',
    check: j => {
      const t = allText(j);
      const pass = t.includes('Alpha') && t.includes('Beta') && t.includes('Gamma');
      return { pass, notes: 'bullet items' };
    }},
  { id: '206', name: 'Numbered list (PPTX)', file: 'evals/results/206.pptx',
    check: j => {
      const t = allText(j);
      const pass = t.includes('First') && t.includes('Second') && t.includes('Third');
      return { pass, notes: 'numbered items' };
    }},
  { id: '207', name: 'Code block (PPTX)', file: 'evals/results/207.pptx',
    check: j => {
      const runs = flatRuns(j);
      const pass = hasRun(runs, r => r.fontName?.includes('Consolas') && r.text.includes('const x = 42'));
      return { pass, notes: 'code font' };
    }},
  // PPTX — Content
  { id: '208', name: 'Table (PPTX)', file: 'evals/results/208.pptx',
    check: j => {
      const t = allText(j);
      const pass = t.includes('Alice') && t.includes('Bob') && t.includes('Name') && t.includes('Age');
      return { pass, notes: 'table content' };
    }},
  { id: '209', name: 'Mixed content (PPTX)', file: 'evals/results/209.pptx',
    check: j => {
      const t = allText(j);
      const pass = t.includes('bold') && t.includes('italic') && t.includes('Point one') && t.includes('Point two') && t.includes('link');
      return { pass, notes: 'mixed content' };
    }},
  // PPTX — Images
  { id: '210', name: 'Image (PPTX)', file: 'evals/results/210.pptx',
    check: j => {
      const shapes = j.slides?.flatMap(s => s.shapes || []) || [];
      const pass = shapes.some(s => s.isPicture || s.shapeType === 13);
      return { pass, notes: 'picture shape' };
    }},
  { id: '211', name: 'Image + text (PPTX)', file: 'evals/results/211.pptx',
    check: j => {
      const shapes = j.slides?.flatMap(s => s.shapes || []) || [];
      const hasPic = shapes.some(s => s.isPicture);
      const hasText = allText(j).includes('some text content');
      return { pass: hasPic && hasText, notes: 'pic+text' };
    }},
  // PPTX — Nested Lists
  { id: '212', name: 'Nested list (PPTX)', file: 'evals/results/212.pptx',
    check: j => {
      const t = allText(j);
      const pass = t.includes('Top level item') && t.includes('Second level item') && t.includes('Third level item') && t.includes('Back to top');
      return { pass, notes: 'nested list' };
    }},
  // PPTX — Custom Theme
  { id: '213', name: 'Default theme (PPTX)', file: 'evals/results/213.pptx',
    check: j => {
      const t = allText(j);
      const pass = !j.error && t.includes('Default Theme Test') && t.includes('Some content');
      return { pass, notes: 'default theme' };
    }},
  { id: '214', name: 'Template theme (PPTX)', file: 'evals/results/214.pptx',
    check: j => {
      const pass = !j.error
        && j.theme?.colors?.accent1 === 'E94560'
        && j.theme?.fonts?.major === 'Georgia'
        && j.theme?.colors?.accent1 !== '4472C4';
      return { pass, notes: `accent1=${j.theme?.colors?.accent1}, major=${j.theme?.fonts?.major}` };
    }},
];

// Run all checks
console.log('');
console.log('# Eval Pass-Condition Results');
console.log('');

const categories = {
  'DOCX Structure': ['100','101'],
  'DOCX Formatting': ['102','103','104','105'],
  'DOCX Content': ['106','107'],
  'DOCX Images': ['108','109'],
  'DOCX Nested Lists': ['110'],
  'DOCX Custom Theme': ['111','112'],
  'DOCX Title Style': ['113'],
  'DOCX List Restart': ['114','115'],
  'DOCX Template Styles': ['116'],
  'DOCX Headers & Footers': ['117'],
  'PPTX Structure': ['200','201','202'],
  'PPTX Formatting': ['203','204','205','206','207'],
  'PPTX Content': ['208','209'],
  'PPTX Images': ['210','211'],
  'PPTX Nested Lists': ['212'],
  'PPTX Custom Theme': ['213','214'],
};

const results = [];
for (const ev of evals) {
  const j = verify(ev.file);
  if (j.error && !ev.check) {
    results.push({ ...ev, pass: false, notes: j.error });
    continue;
  }
  const { pass, notes } = ev.check(j);
  results.push({ ...ev, pass, notes });
}

// Summary table
console.log('| Category | Pass | Fail | Total |');
console.log('|----------|------|------|-------|');
let totalPass = 0, totalFail = 0;
for (const [cat, ids] of Object.entries(categories)) {
  const catResults = results.filter(r => ids.includes(r.id));
  const p = catResults.filter(r => r.pass).length;
  const f = catResults.filter(r => !r.pass).length;
  totalPass += p;
  totalFail += f;
  console.log(`| ${cat.padEnd(22)} | ${String(p).padStart(4)} | ${String(f).padStart(4)} | ${String(p + f).padStart(5)} |`);
}
console.log(`| **Total**              | **${totalPass}** | **${totalFail}** | **${totalPass + totalFail}** |`);
console.log('');

// Per-eval results
console.log('| #   | Eval                        | Result | Notes |');
console.log('|-----|-----------------------------|--------|-------|');
for (const r of results) {
  const icon = r.pass ? '✅' : '❌';
  console.log(`| ${r.id.padStart(3)} | ${r.name.padEnd(27)} | ${icon}     | ${r.notes} |`);
}

console.log('');
console.log(`**${totalPass}/${totalPass + totalFail}** passed`);
if (totalFail > 0) {
  console.log('');
  console.log('Failures:');
  for (const r of results.filter(r => !r.pass)) {
    console.log(`  ❌ ${r.id} — ${r.name}: ${r.notes}`);
  }
}

process.exit(totalFail > 0 ? 1 : 0);
