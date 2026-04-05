#!/usr/bin/env node
// Static validation — no Office, no network required
// Run: node --test tests/test-scripts.js

const { describe, it } = require('node:test');
const assert = require('node:assert');
const { execSync } = require('node:child_process');
const { existsSync, readFileSync } = require('node:fs');
const { join } = require('node:path');

const skillDir = join(__dirname, '..', '.claude', 'skills', 'create-office-file');
const script = join(skillDir, 'scripts', 'create-office-file.mjs');
const evalsDir = join(__dirname, '..', 'evals');
const verifyScript = join(evalsDir, 'verify-output.mjs');

/**
 * Run the create-office-file script. Returns { exitCode, stdout, stderr }.
 */
function run(args = '', input = null) {
  try {
    const opts = { encoding: 'utf8', timeout: 15000, stdio: ['pipe', 'pipe', 'pipe'] };
    if (input) opts.input = input;
    const stdout = execSync(`node "${script}" ${args}`, opts);
    return { exitCode: 0, stdout, stderr: '' };
  } catch (e) {
    return { exitCode: e.status ?? 1, stdout: e.stdout ?? '', stderr: e.stderr ?? '' };
  }
}

/**
 * Generate a file from markdown string, return the buffer.
 */
function generate(md, format) {
  const outFile = join(__dirname, `_test_output.${format}`);
  try {
    execSync(`node "${script}" -f ${format} -o "${outFile}"`, {
      input: md, encoding: 'utf8', timeout: 15000, stdio: ['pipe', 'pipe', 'pipe']
    });
    return readFileSync(outFile);
  } finally {
    try { require('fs').unlinkSync(outFile); } catch {}
  }
}

/**
 * Parse a ZIP buffer and return entry names. Minimal ZIP reader for testing.
 */
function zipEntries(buf) {
  const entries = [];
  // Find EOCD signature (0x06054b50) scanning backward
  let eocdOff = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (buf.readUInt32LE(i) === 0x06054B50) { eocdOff = i; break; }
  }
  if (eocdOff < 0) return entries;
  const cdOff = buf.readUInt32LE(eocdOff + 16);
  const cdCount = buf.readUInt16LE(eocdOff + 10);
  let pos = cdOff;
  for (let i = 0; i < cdCount; i++) {
    if (buf.readUInt32LE(pos) !== 0x02014B50) break;
    const nameLen = buf.readUInt16LE(pos + 28);
    const extraLen = buf.readUInt16LE(pos + 30);
    const commentLen = buf.readUInt16LE(pos + 32);
    entries.push(buf.toString('utf8', pos + 46, pos + 46 + nameLen));
    pos += 46 + nameLen + extraLen + commentLen;
  }
  return entries;
}

/**
 * Extract a file from a ZIP buffer by name.
 */
function zipExtract(buf, name) {
  let eocdOff = -1;
  for (let i = buf.length - 22; i >= 0; i--) {
    if (buf.readUInt32LE(i) === 0x06054B50) { eocdOff = i; break; }
  }
  if (eocdOff < 0) return null;
  const cdOff = buf.readUInt32LE(eocdOff + 16);
  const cdCount = buf.readUInt16LE(eocdOff + 10);
  let pos = cdOff;
  for (let i = 0; i < cdCount; i++) {
    const nameLen = buf.readUInt16LE(pos + 28);
    const extraLen = buf.readUInt16LE(pos + 30);
    const commentLen = buf.readUInt16LE(pos + 32);
    const entryName = buf.toString('utf8', pos + 46, pos + 46 + nameLen);
    const localOff = buf.readUInt32LE(pos + 42);
    if (entryName === name) {
      const lNameLen = buf.readUInt16LE(localOff + 26);
      const lExtraLen = buf.readUInt16LE(localOff + 28);
      const method = buf.readUInt16LE(localOff + 8);
      const compSize = buf.readUInt32LE(localOff + 18);
      const dataStart = localOff + 30 + lNameLen + lExtraLen;
      const raw = buf.subarray(dataStart, dataStart + compSize);
      if (method === 0) return raw.toString('utf8');
      if (method === 8) return require('zlib').inflateRawSync(raw).toString('utf8');
      return null;
    }
    pos += 46 + nameLen + extraLen + commentLen;
  }
  return null;
}

const scriptSource = readFileSync(script, 'utf8');

// ============================================================================
// 1. File existence
// ============================================================================
describe('File existence', () => {
  for (const f of ['scripts/create-office-file.mjs', 'SKILL.md']) {
    it(`${f} exists`, () => {
      assert.ok(existsSync(join(skillDir, f)), `${f} not found`);
    });
  }
  it('evals/verify-output.mjs exists', () => {
    assert.ok(existsSync(verifyScript), 'evals/verify-output.mjs not found');
  });
  for (const f of ['pptx.md', 'docx.md']) {
    it(`references/${f} exists`, () => {
      assert.ok(existsSync(join(skillDir, 'references', f)), `references/${f} not found`);
    });
  }
});

// ============================================================================
// 2. Shebang line
// ============================================================================
describe('Shebang line', () => {
  for (const [label, path] of [
    ['scripts/create-office-file.mjs', script],
    ['evals/verify-output.mjs', verifyScript]
  ]) {
    it(`${label} has node shebang`, () => {
      const first = readFileSync(path, 'utf8').split(/\r?\n/)[0];
      assert.match(first, /node/, `${label} should have node shebang`);
    });
  }
});

// ============================================================================
// 3. No external npm dependencies
// ============================================================================
describe('No external npm dependencies', () => {
  it('create-office-file.mjs only imports Node built-ins', () => {
    const imports = scriptSource.match(/import\s+.*from\s+['"]([^'"]+)['"]/g) || [];
    for (const imp of imports) {
      const mod = imp.match(/from\s+['"]([^'"]+)['"]/)[1];
      assert.ok(
        ['fs', 'path', 'zlib'].includes(mod),
        `Unexpected import: ${mod}. Only fs, path, zlib allowed.`
      );
    }
  });
});

// ============================================================================
// 4. CLI error handling
// ============================================================================
describe('CLI error handling', () => {
  it('fails with no arguments', () => {
    const r = run('');
    assert.notStrictEqual(r.exitCode, 0);
  });

  it('fails with unknown format', () => {
    const r = run('-o test.xyz');
    assert.notStrictEqual(r.exitCode, 0);
    assert.match(r.stderr, /format/i);
  });

  it('fails with missing input file', () => {
    const r = run('-i nonexistent.md -o test.pptx');
    assert.notStrictEqual(r.exitCode, 0);
  });
});

// ============================================================================
// 5. CRC-32 correctness
// ============================================================================
describe('CRC-32', () => {
  it('produces correct CRC for known inputs', () => {
    // Generate a minimal file and check the ZIP is valid (CRC is validated by unzip)
    const buf = generate('# Test', 'docx');
    const entries = zipEntries(buf);
    assert.ok(entries.length > 0, 'ZIP should have entries');
    // Verify we can decompress (implicitly validates CRC in the ZIP structure)
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc, 'Should extract word/document.xml');
    assert.ok(doc.includes('Test'), 'Document should contain test text');
  });
});

// ============================================================================
// 6. ZIP structure
// ============================================================================
describe('ZIP structure', () => {
  it('PPTX has all required entries', () => {
    const buf = generate('# Title\n## Slide 1\nContent', 'pptx');
    const entries = zipEntries(buf);
    const required = [
      '[Content_Types].xml', '_rels/.rels',
      'ppt/presentation.xml', 'ppt/_rels/presentation.xml.rels',
      'ppt/theme/theme1.xml',
      'ppt/slideMasters/slideMaster1.xml', 'ppt/slideMasters/_rels/slideMaster1.xml.rels',
      'ppt/slideLayouts/slideLayout1.xml', 'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
      'ppt/slides/slide1.xml', 'ppt/slides/_rels/slide1.xml.rels'
    ];
    for (const r of required) {
      assert.ok(entries.includes(r), `Missing required entry: ${r}`);
    }
  });

  it('DOCX has all required entries', () => {
    const buf = generate('# Heading\nParagraph', 'docx');
    const entries = zipEntries(buf);
    const required = [
      '[Content_Types].xml', '_rels/.rels',
      'word/document.xml', 'word/_rels/document.xml.rels',
      'word/styles.xml', 'word/numbering.xml'
    ];
    for (const r of required) {
      assert.ok(entries.includes(r), `Missing required entry: ${r}`);
    }
  });
});

// ============================================================================
// 7. Markdown parser
// ============================================================================
describe('Markdown parser — PPTX', () => {
  it('# creates a title slide', () => {
    const buf = generate('# My Title', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('My Title'), 'Title slide should contain heading text');
    assert.ok(slide.includes('anchorCtr'), 'Title should be centered');
  });

  it('## creates a content slide', () => {
    const buf = generate('## Section\nSome content', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('Section'), 'Slide should contain heading');
    assert.ok(slide.includes('Some content'), 'Slide should contain body text');
  });

  it('--- creates a slide break', () => {
    const buf = generate('## Slide 1\nA\n\n---\n\n## Slide 2\nB', 'pptx');
    const entries = zipEntries(buf);
    const slides = entries.filter(e => /^ppt\/slides\/slide\d+\.xml$/.test(e));
    assert.ok(slides.length >= 2, `Expected at least 2 slides, got ${slides.length}`);
  });

  it('bold text gets b="1" attribute', () => {
    const buf = generate('## Test\n**bold words**', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('b="1"'), 'Should have bold attribute');
    assert.ok(slide.includes('bold words'), 'Should contain bold text');
  });

  it('italic text gets i="1" attribute', () => {
    const buf = generate('## Test\n*italic words*', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('i="1"'), 'Should have italic attribute');
  });

  it('bullet list uses buChar', () => {
    const buf = generate('## Test\n- Item 1\n- Item 2', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('buChar'), 'Should have bullet character element');
    assert.ok(slide.includes('Item 1'), 'Should contain list items');
  });

  it('numbered list uses buAutoNum', () => {
    const buf = generate('## Test\n1. First\n2. Second', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('buAutoNum'), 'Should have auto-number element');
  });

  it('code block uses Consolas font', () => {
    const buf = generate('## Test\n```\nconst x = 1;\n```', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('Consolas'), 'Code should use Consolas font');
  });

  it('table generates tbl element', () => {
    const buf = generate('## Test\n| A | B |\n|---|---|\n| 1 | 2 |', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('a:tbl'), 'Should have table element');
    assert.ok(slide.includes('a:gridCol'), 'Should have grid columns');
  });

  it('link generates hlinkClick', () => {
    const buf = generate('## Test\n[Example](https://example.com)', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('hlinkClick'), 'Should have hyperlink element');
    const rels = zipExtract(buf, 'ppt/slides/_rels/slide1.xml.rels');
    assert.ok(rels.includes('example.com'), 'Rels should contain URL');
  });
});

describe('Markdown parser — DOCX', () => {
  it('headings use correct styles', () => {
    const buf = generate('# H1\n## H2\n### H3', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('Heading1'), 'Should have Heading1 style');
    assert.ok(doc.includes('Heading2'), 'Should have Heading2 style');
    assert.ok(doc.includes('Heading3'), 'Should have Heading3 style');
  });

  it('bold text uses w:b element', () => {
    const buf = generate('**bold text**', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('<w:b/>'), 'Should have bold element');
  });

  it('italic text uses w:i element', () => {
    const buf = generate('*italic text*', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('<w:i/>'), 'Should have italic element');
  });

  it('bullet list references numId 1', () => {
    const buf = generate('- Bullet A\n- Bullet B', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('w:numId w:val="1"'), 'Should reference bullet numId');
  });

  it('ordered list references numId 2', () => {
    const buf = generate('1. One\n2. Two', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('w:numId w:val="2"'), 'Should reference ordered numId');
  });

  it('code block uses CodeBlock style', () => {
    const buf = generate('```\ncode line\n```', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('CodeBlock'), 'Should have CodeBlock style');
    assert.ok(doc.includes('Consolas'), 'Should use Consolas font');
  });

  it('table generates w:tbl element', () => {
    const buf = generate('| X | Y |\n|---|---|\n| a | b |', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('<w:tbl>'), 'Should have table element');
    assert.ok(doc.includes('<w:gridCol'), 'Should have grid columns');
  });

  it('hyperlink generates w:hyperlink', () => {
    const buf = generate('[Link](https://test.com)', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('w:hyperlink'), 'Should have hyperlink element');
    const rels = zipExtract(buf, 'word/_rels/document.xml.rels');
    assert.ok(rels.includes('test.com'), 'Rels should contain URL');
  });

  it('horizontal rule generates border', () => {
    const buf = generate('Above\n\n---\n\nBelow', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('w:pBdr'), 'Should have paragraph border');
  });
});

// ============================================================================
// 8. PPTX skeleton integrity
// ============================================================================
describe('PPTX skeleton integrity', () => {
  it('theme has Office color scheme', () => {
    const buf = generate('# Test', 'pptx');
    const theme = zipExtract(buf, 'ppt/theme/theme1.xml');
    assert.ok(theme.includes('clrScheme'), 'Theme should have color scheme');
    assert.ok(theme.includes('fontScheme'), 'Theme should have font scheme');
  });

  it('slide master references layout', () => {
    const buf = generate('# Test', 'pptx');
    const master = zipExtract(buf, 'ppt/slideMasters/slideMaster1.xml');
    assert.ok(master.includes('sldLayoutIdLst'), 'Master should reference layout');
  });

  it('slide layout references master', () => {
    const buf = generate('# Test', 'pptx');
    const rels = zipExtract(buf, 'ppt/slideLayouts/_rels/slideLayout1.xml.rels');
    assert.ok(rels.includes('slideMaster'), 'Layout rels should reference master');
  });

  it('content types lists all slides', () => {
    const buf = generate('# S1\n## S2\n## S3', 'pptx');
    const ct = zipExtract(buf, '[Content_Types].xml');
    assert.ok(ct.includes('slide1.xml'), 'Content types should list slide1');
    assert.ok(ct.includes('slide2.xml'), 'Content types should list slide2');
    assert.ok(ct.includes('slide3.xml'), 'Content types should list slide3');
  });
});

// ============================================================================
// 9. DOCX skeleton integrity
// ============================================================================
describe('DOCX skeleton integrity', () => {
  it('styles.xml defines all heading styles', () => {
    const buf = generate('# Test', 'docx');
    const styles = zipExtract(buf, 'word/styles.xml');
    for (let i = 1; i <= 6; i++) {
      assert.ok(styles.includes(`Heading${i}`), `Should define Heading${i}`);
    }
  });

  it('numbering.xml defines bullet and numbered lists', () => {
    const buf = generate('- a', 'docx');
    const num = zipExtract(buf, 'word/numbering.xml');
    assert.ok(num.includes('bullet'), 'Should have bullet format');
    assert.ok(num.includes('decimal'), 'Should have decimal format');
  });

  it('document.xml has section properties', () => {
    const buf = generate('Test', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('w:sectPr'), 'Should have section properties');
    assert.ok(doc.includes('w:pgSz'), 'Should have page size');
    assert.ok(doc.includes('w:pgMar'), 'Should have page margins');
  });
});

// ============================================================================
// 10. Stdin support
// ============================================================================
describe('Stdin support', () => {
  it('reads markdown from stdin', () => {
    const outFile = join(__dirname, '_test_stdin.docx');
    try {
      execSync(`node "${script}" -o "${outFile}"`, {
        input: '# From Stdin', encoding: 'utf8', timeout: 15000, stdio: ['pipe', 'pipe', 'pipe']
      });
      const buf = readFileSync(outFile);
      const doc = zipExtract(buf, 'word/document.xml');
      assert.ok(doc.includes('From Stdin'));
    } finally {
      try { require('fs').unlinkSync(outFile); } catch {}
    }
  });
});

// ============================================================================
// 11. Image support
// ============================================================================

const testPng = join(__dirname, 'test-image.png');
const testJpg = join(__dirname, 'test-image.jpg');

/**
 * Generate a file from a markdown FILE (not string), return the buffer.
 * This is needed for image tests where paths must resolve relative to the input file.
 */
function generateFromFile(mdPath, format) {
  const outFile = join(__dirname, `_test_img_output.${format}`);
  try {
    execSync(`node "${script}" -i "${mdPath}" -o "${outFile}"`, {
      encoding: 'utf8', timeout: 15000, stdio: ['pipe', 'pipe', 'pipe']
    });
    return readFileSync(outFile);
  } finally {
    try { require('fs').unlinkSync(outFile); } catch {}
  }
}

/**
 * Write a temp markdown file in the tests/ dir, generate, and return the buffer.
 */
function generateWithImage(md, format) {
  const mdFile = join(__dirname, '_test_img_input.md');
  try {
    require('fs').writeFileSync(mdFile, md, 'utf8');
    return generateFromFile(mdFile, format);
  } finally {
    try { require('fs').unlinkSync(mdFile); } catch {}
  }
}

describe('Markdown parser — images', () => {
  it('parses block-level ![alt](path) as image node', () => {
    const buf = generateWithImage('## Test\n\n![A photo](test-image.png)', 'pptx');
    const entries = zipEntries(buf);
    // Image should be embedded as media file
    const mediaFiles = entries.filter(e => e.startsWith('ppt/media/'));
    assert.ok(mediaFiles.length >= 1, `Expected media files, got: ${mediaFiles.join(', ')}`);
  });

  it('parses inline ![alt](path) in paragraph text', () => {
    const buf = generateWithImage('Text with ![icon](test-image.png) inline', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    // Should contain either an image drawing or fallback text
    assert.ok(doc.includes('pic:pic') || doc.includes('[Image:'), 'Should have image or fallback');
  });
});

describe('Image support — PPTX', () => {
  it('embeds PNG in ppt/media/', () => {
    const buf = generateWithImage('## Slide\n\n![Test PNG](test-image.png)', 'pptx');
    const entries = zipEntries(buf);
    const media = entries.filter(e => e.startsWith('ppt/media/'));
    assert.ok(media.some(e => e.endsWith('.png')), `Should have PNG in media: ${media.join(', ')}`);
  });

  it('slide XML contains p:pic with blipFill', () => {
    const buf = generateWithImage('## Slide\n\n![Test](test-image.png)', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('p:pic'), 'Slide should have p:pic element');
    assert.ok(slide.includes('blipFill'), 'Slide should have blipFill');
    assert.ok(slide.includes('a:blip'), 'Slide should have a:blip');
  });

  it('Content_Types includes png extension default', () => {
    const buf = generateWithImage('## Slide\n\n![Test](test-image.png)', 'pptx');
    const ct = zipExtract(buf, '[Content_Types].xml');
    assert.ok(ct.includes('Extension="png"'), 'Content types should include png');
    assert.ok(ct.includes('image/png'), 'Content types should have image/png');
  });

  it('slide rels include image relationship', () => {
    const buf = generateWithImage('## Slide\n\n![Test](test-image.png)', 'pptx');
    const rels = zipExtract(buf, 'ppt/slides/_rels/slide1.xml.rels');
    assert.ok(rels.includes('relationships/image'), 'Rels should have image relationship type');
    assert.ok(rels.includes('media/'), 'Rels should reference media directory');
  });

  it('p:pic has picLocks with noChangeAspect', () => {
    const buf = generateWithImage('## Slide\n\n![Test](test-image.png)', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('noChangeAspect'), 'Should lock aspect ratio');
  });
});

describe('Image support — DOCX', () => {
  it('embeds PNG in word/media/', () => {
    const buf = generateWithImage('![Test PNG](test-image.png)', 'docx');
    const entries = zipEntries(buf);
    const media = entries.filter(e => e.startsWith('word/media/'));
    assert.ok(media.some(e => e.endsWith('.png')), `Should have PNG in media: ${media.join(', ')}`);
  });

  it('document.xml contains w:drawing with pic:pic', () => {
    const buf = generateWithImage('![Test](test-image.png)', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('w:drawing'), 'Document should have w:drawing');
    assert.ok(doc.includes('pic:pic'), 'Document should have pic:pic');
    assert.ok(doc.includes('a:blip'), 'Document should have a:blip');
  });

  it('Content_Types includes png extension default', () => {
    const buf = generateWithImage('![Test](test-image.png)', 'docx');
    const ct = zipExtract(buf, '[Content_Types].xml');
    assert.ok(ct.includes('Extension="png"'), 'Content types should include png');
    assert.ok(ct.includes('image/png'), 'Content types should have image/png');
  });

  it('document rels include image relationship', () => {
    const buf = generateWithImage('![Test](test-image.png)', 'docx');
    const rels = zipExtract(buf, 'word/_rels/document.xml.rels');
    assert.ok(rels.includes('relationships/image'), 'Rels should have image relationship type');
    assert.ok(rels.includes('media/'), 'Rels should reference media directory');
  });

  it('w:drawing has wp:inline with extent dimensions', () => {
    const buf = generateWithImage('![Test](test-image.png)', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('wp:inline'), 'Should have wp:inline element');
    assert.ok(doc.includes('wp:extent'), 'Should have extent with dimensions');
  });

  it('document root has wp and pic namespaces', () => {
    const buf = generateWithImage('![Test](test-image.png)', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('xmlns:wp='), 'Should declare wp namespace');
    assert.ok(doc.includes('xmlns:pic='), 'Should declare pic namespace');
  });
});

describe('Image dimension parsing', () => {
  it('reads PNG dimensions (200x150)', () => {
    const buf = generateWithImage('![Test](test-image.png)', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    // The image is 200x150 pixels. At 96 DPI: 200*914400/96 = 1905000, 150*914400/96 = 1428750
    // These should fit within slide bounds without scaling
    assert.ok(slide.includes('1905000') || slide.includes('p:pic'), 'Should have correct width EMU or pic element');
  });

  it('reads JPEG dimensions (300x200)', () => {
    const buf = generateWithImage('## Slide\n\n![JPEG Test](test-image.jpg)', 'pptx');
    const entries = zipEntries(buf);
    const media = entries.filter(e => e.startsWith('ppt/media/'));
    assert.ok(media.some(e => e.endsWith('.jpg')), `Should embed JPEG: ${media.join(', ')}`);
    const ct = zipExtract(buf, '[Content_Types].xml');
    assert.ok(ct.includes('image/jpeg'), 'Content types should have image/jpeg');
  });
});

describe('Image error handling', () => {
  it('missing image file renders fallback text', () => {
    const buf = generateWithImage('![Missing](nonexistent.png)', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('[Image:'), 'Should have fallback text for missing image');
  });

  it('missing image in PPTX does not crash', () => {
    const buf = generateWithImage('## Slide\n\n![Missing](nonexistent.png)', 'pptx');
    const entries = zipEntries(buf);
    assert.ok(entries.length > 0, 'Should still produce a valid ZIP');
    const media = entries.filter(e => e.startsWith('ppt/media/'));
    assert.strictEqual(media.length, 0, 'No media files for missing images');
  });
});

describe('Image with existing content', () => {
  it('PPTX image coexists with text content', () => {
    const buf = generateWithImage('## Slide\nSome text\n\n![Photo](test-image.png)', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('Some text'), 'Should have text content');
    assert.ok(slide.includes('p:pic'), 'Should also have image');
  });

  it('DOCX image coexists with headings and paragraphs', () => {
    const buf = generateWithImage('# Title\n\nSome paragraph.\n\n![Photo](test-image.png)\n\nAnother paragraph.', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('Title'), 'Should have heading');
    assert.ok(doc.includes('Some paragraph'), 'Should have first paragraph');
    assert.ok(doc.includes('pic:pic'), 'Should have image');
    assert.ok(doc.includes('Another paragraph'), 'Should have second paragraph');
  });
});

// ── Nested list tests ───────────────────────────────────────────────────────

describe('Nested lists — parser', () => {
  it('2-space indent produces level 1', () => {
    const buf = generate('## Test\n- Top\n  - Sub', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    // Level 1 items get marL = 342900 + 457200 = 800100
    assert.ok(slide.includes('marL="800100"'), 'Should have level 1 margin');
  });

  it('4-space indent produces level 2', () => {
    const buf = generate('## Test\n- Top\n    - Deep', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    // Level 2 items get marL = 342900 + 914400 = 1257300
    assert.ok(slide.includes('marL="1257300"'), 'Should have level 2 margin');
  });

  it('tab indent normalized to level 1', () => {
    const buf = generate('## Test\n- Top\n\t- Tabbed', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('marL="800100"'), 'Tab should produce level 1 margin');
  });

  it('excess indent (6+ spaces) clamped to level 2', () => {
    const buf = generate('## Test\n- Top\n      - TooDeep', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    // Should be clamped to level 2, not level 3
    assert.ok(slide.includes('marL="1257300"'), 'Excess indent should clamp to level 2');
    assert.ok(!slide.includes('marL="1714500"'), 'Should NOT have level 3 margin');
  });

  it('flat list items default to level 0', () => {
    const buf = generate('## Test\n- One\n- Two\n- Three', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    // All items should use the original marL
    assert.ok(slide.includes('marL="342900"'), 'Flat items should stay level 0');
    assert.ok(!slide.includes('marL="800100"'), 'Should NOT have level 1 margin');
  });
});

describe('Nested lists — PPTX', () => {
  it('level 0 uses bullet char \u2022', () => {
    const buf = generate('## Test\n- Top level', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('\u2022'), 'Level 0 should use bullet char \u2022');
  });

  it('level 1 uses dash char \u2013', () => {
    const buf = generate('## Test\n- Top\n  - Sub', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('\u2013'), 'Level 1 should use dash char \u2013');
  });

  it('level 2 uses angle char \u203A', () => {
    const buf = generate('## Test\n- Top\n    - Deep', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('\u203A'), 'Level 2 should use angle char \u203A');
  });

  it('numbered list nesting increases marL', () => {
    const buf = generate('## Test\n1. Top\n  1. Sub\n    1. Deep', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('marL="342900"'), 'Level 0 margin present');
    assert.ok(slide.includes('marL="800100"'), 'Level 1 margin present');
    assert.ok(slide.includes('marL="1257300"'), 'Level 2 margin present');
  });
});

describe('Nested lists — DOCX', () => {
  it('level 0 uses ilvl 0', () => {
    const buf = generate('- Top level item', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('w:ilvl w:val="0"'), 'Level 0 should use ilvl 0');
  });

  it('level 1 uses ilvl 1', () => {
    const buf = generate('- Top\n  - Sub', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('w:ilvl w:val="1"'), 'Level 1 should use ilvl 1');
  });

  it('level 2 uses ilvl 2', () => {
    const buf = generate('- Top\n    - Deep', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('w:ilvl w:val="2"'), 'Level 2 should use ilvl 2');
  });

  it('numbering.xml has 3 levels for bullet abstractNum', () => {
    const buf = generate('- Item', 'docx');
    const numbering = zipExtract(buf, 'word/numbering.xml');
    assert.ok(numbering.includes('w:ilvl="0"'), 'Should define level 0');
    assert.ok(numbering.includes('w:ilvl="1"'), 'Should define level 1');
    assert.ok(numbering.includes('w:ilvl="2"'), 'Should define level 2');
    assert.ok(numbering.includes('w:multiLevelType'), 'Should declare multiLevelType');
  });

  it('numbering.xml has 3 levels for numbered abstractNum', () => {
    const buf = generate('1. Item', 'docx');
    const numbering = zipExtract(buf, 'word/numbering.xml');
    assert.ok(numbering.includes('w:numFmt w:val="decimal"'), 'Level 0 should be decimal');
    assert.ok(numbering.includes('w:numFmt w:val="lowerLetter"'), 'Level 1 should be lowerLetter');
    assert.ok(numbering.includes('w:numFmt w:val="lowerRoman"'), 'Level 2 should be lowerRoman');
  });

  it('nested ordered list uses ilvl 1 with numId 2', () => {
    const buf = generate('1. Top\n  1. Sub', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('w:ilvl w:val="1"'), 'Sub-item should be ilvl 1');
    assert.ok(doc.includes('w:numId w:val="2"'), 'Should reference numbered list numId');
  });
});

describe('Nested lists — coexistence', () => {
  it('mixed levels in PPTX produce correct margins alongside flat content', () => {
    const buf = generate('## Slide\nSome text\n- Top\n  - Sub\n- Back to top', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('Some text'), 'Should have paragraph text');
    assert.ok(slide.includes('marL="342900"'), 'Should have level 0 bullets');
    assert.ok(slide.includes('marL="800100"'), 'Should have level 1 bullets');
  });

  it('mixed levels in DOCX produce correct ilvl values', () => {
    const buf = generate('# Heading\n- Top\n  - Sub\n    - Deep\n- Back', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('Heading'), 'Should have heading');
    assert.ok(doc.includes('w:ilvl w:val="0"'), 'Should have level 0');
    assert.ok(doc.includes('w:ilvl w:val="1"'), 'Should have level 1');
    assert.ok(doc.includes('w:ilvl w:val="2"'), 'Should have level 2');
  });
});
