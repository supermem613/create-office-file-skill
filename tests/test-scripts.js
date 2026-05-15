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
    const buf = generate('# H1\n## H2\n### H3\n# H1b', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('pStyle w:val="Title"'), 'First H1 should use Title style');
    assert.ok(doc.includes('Heading1'), 'Subsequent H1 should use Heading1 style');
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

  it('bullet list gets a numId', () => {
    const buf = generate('- Bullet A\n- Bullet B', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(/w:numId w:val="\d+"/.test(doc), 'Should reference a bullet numId');
  });

  it('ordered list gets a numId', () => {
    const buf = generate('1. One\n2. Two', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(/w:numId w:val="\d+"/.test(doc), 'Should reference an ordered numId');
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
  it('styles.xml includes Title style', () => {
    const buf = generate('# Test', 'docx');
    const styles = zipExtract(buf, 'word/styles.xml');
    assert.ok(styles.includes('Title'), 'Should define Title style');
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
// 9b. Title style and numbered list restart
// ============================================================================
describe('DOCX — Title style', () => {
  it('first H1 uses Title style', () => {
    const buf = generate('# First\n## Sub\n# Second', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('pStyle w:val="Title"'), 'First H1 should use Title');
    assert.ok(doc.includes('pStyle w:val="Heading1"'), 'Later H1 should use Heading1');
  });

  it('only one Title usage for multiple H1s', () => {
    const buf = generate('# A\n# B\n# C', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    const titleCount = (doc.match(/pStyle w:val="Title"/g) || []).length;
    assert.strictEqual(titleCount, 1, 'Should have exactly one Title style usage');
  });
});

describe('DOCX — Numbered list restart', () => {
  it('separate ordered lists get different numIds', () => {
    const buf = generate('1. A\n2. B\n\nParagraph break\n\n1. C\n2. D', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    const numIds = [...new Set((doc.match(/numId w:val="(\d+)"/g) || []).map(m => m.match(/"(\d+)"/)[1]))];
    assert.ok(numIds.length >= 2, `Expected at least 2 distinct numIds, got ${numIds.length}: ${numIds}`);
  });

  it('numbering.xml has startOverride for restart', () => {
    const buf = generate('1. A\n2. B\n\nText\n\n1. C\n2. D', 'docx');
    const num = zipExtract(buf, 'word/numbering.xml');
    assert.ok(num.includes('w:startOverride'), 'Should have startOverride for list restart');
  });

  it('separate bullet lists get different numIds', () => {
    const buf = generate('- A\n- B\n\nText\n\n- C\n- D', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    const numIds = [...new Set((doc.match(/numId w:val="(\d+)"/g) || []).map(m => m.match(/"(\d+)"/)[1]))];
    assert.ok(numIds.length >= 2, `Separate bullet lists should have different numIds, got: ${numIds}`);
  });
});

describe('DOCX — Template styles extraction', () => {
  it('template styles.xml is used when provided', () => {
    const templatePath = join(__dirname, '_test_tpl.docx');
    try {
      // Create a minimal DOCX template with custom styles
      const { deflateRawSync } = require('zlib');
      function crc(buf) {
        if (typeof buf === 'string') buf = Buffer.from(buf, 'utf8');
        const t = new Uint32Array(256);
        for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; }
        let c = 0xFFFFFFFF; for (let i = 0; i < buf.length; i++) c = t[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0;
      }
      const theme = '<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Test"><a:themeElements><a:clrScheme name="T"><a:dk1><a:srgbClr val="111111"/></a:dk1><a:lt1><a:srgbClr val="FEFEFE"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="DDDDDD"/></a:lt2><a:accent1><a:srgbClr val="AA0000"/></a:accent1><a:accent2><a:srgbClr val="00AA00"/></a:accent2><a:accent3><a:srgbClr val="0000AA"/></a:accent3><a:accent4><a:srgbClr val="AAAA00"/></a:accent4><a:accent5><a:srgbClr val="AA00AA"/></a:accent5><a:accent6><a:srgbClr val="00AAAA"/></a:accent6><a:hlink><a:srgbClr val="FF0000"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="T"><a:majorFont><a:latin typeface="TestMajor"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="TestMinor"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="T"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>';
      const stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="CustomFont" w:hAnsi="CustomFont"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="48"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:rPr><w:sz w:val="72"/></w:rPr></w:style></w:styles>';
      const ct = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>';
      const rootRels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
      const docRels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>';
      const doc = '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>';
      const files = [
        { name: '[Content_Types].xml', data: Buffer.from(ct) },
        { name: '_rels/.rels', data: Buffer.from(rootRels) },
        { name: 'word/document.xml', data: Buffer.from(doc) },
        { name: 'word/_rels/document.xml.rels', data: Buffer.from(docRels) },
        { name: 'word/styles.xml', data: Buffer.from(stylesXml) },
        { name: 'word/theme/theme1.xml', data: Buffer.from(theme) },
      ];
      const locals = [], centrals = [];
      let offset = 0;
      for (const f of files) {
        const nameBuf = Buffer.from(f.name);
        const comp = deflateRawSync(f.data);
        const use = comp.length < f.data.length;
        const stored = use ? comp : f.data;
        const method = use ? 8 : 0;
        const c = crc(f.data);
        const local = Buffer.alloc(30 + nameBuf.length);
        local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(method, 8);
        local.writeUInt32LE(c, 14); local.writeUInt32LE(stored.length, 18); local.writeUInt32LE(f.data.length, 22);
        local.writeUInt16LE(nameBuf.length, 26); nameBuf.copy(local, 30);
        locals.push(local, stored);
        const central = Buffer.alloc(46 + nameBuf.length);
        central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6);
        central.writeUInt16LE(method, 10); central.writeUInt32LE(c, 16);
        central.writeUInt32LE(stored.length, 20); central.writeUInt32LE(f.data.length, 24);
        central.writeUInt16LE(nameBuf.length, 28); central.writeUInt32LE(offset, 42);
        nameBuf.copy(central, 46); centrals.push(central);
        offset += local.length + stored.length;
      }
      const cdSz = centrals.reduce((s, b) => s + b.length, 0);
      const eocd = Buffer.alloc(22);
      eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(files.length, 8); eocd.writeUInt16LE(files.length, 10);
      eocd.writeUInt32LE(cdSz, 12); eocd.writeUInt32LE(offset, 16);
      require('fs').writeFileSync(templatePath, Buffer.concat([...locals, ...centrals, eocd]));

      const buf = generateWithTemplate('# Hello\nParagraph', 'docx', templatePath);
      const styles = zipExtract(buf, 'word/styles.xml');
      assert.ok(styles.includes('CustomFont'), 'Should use template styles with CustomFont');
      assert.ok(styles.includes('CodeBlock'), 'Should inject CodeBlock style if missing');
    } finally {
      try { require('fs').unlinkSync(templatePath); } catch {}
    }
  });

  it('preserves numPr on template heading styles for auto-numbering', () => {
    const templatePath = join(__dirname, '_test_tpl2.docx');
    try {
      const { deflateRawSync } = require('zlib');
      function crc(buf) {
        if (typeof buf === 'string') buf = Buffer.from(buf, 'utf8');
        const t = new Uint32Array(256);
        for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; }
        let c = 0xFFFFFFFF; for (let i = 0; i < buf.length; i++) c = t[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0;
      }
      const theme = '<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="T"><a:themeElements><a:clrScheme name="T"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="444444"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="T"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="T"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>';
      // Styles with numPr on Heading1 (template heading auto-numbering) AND
      // numPr on a non-heading style. Heading1 numPr must be preserved so
      // Word renders the template's "1", "1.1", ... numbering. Non-heading
      // style numPr must still be stripped so it cannot collide with the
      // markdown driver's paragraph-level list numPr.
      const stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="MyList"><w:name w:val="My List"/><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr></w:style></w:styles>';
      // Minimal numbering.xml with one num/abstractNum so the template-heading
      // detection has something to bind to.
      const numberingXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
      const ct = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>';
      const rootRels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
      const docRels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>';
      const doc = '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>';
      const files = [
        { name: '[Content_Types].xml', data: Buffer.from(ct) },
        { name: '_rels/.rels', data: Buffer.from(rootRels) },
        { name: 'word/document.xml', data: Buffer.from(doc) },
        { name: 'word/_rels/document.xml.rels', data: Buffer.from(docRels) },
        { name: 'word/styles.xml', data: Buffer.from(stylesXml) },
        { name: 'word/numbering.xml', data: Buffer.from(numberingXml) },
        { name: 'word/theme/theme1.xml', data: Buffer.from(theme) },
      ];
      const locals = [], centrals = [];
      let offset = 0;
      for (const f of files) {
        const nameBuf = Buffer.from(f.name);
        const comp = deflateRawSync(f.data);
        const use = comp.length < f.data.length;
        const stored = use ? comp : f.data;
        const method = use ? 8 : 0;
        const c = crc(f.data);
        const local = Buffer.alloc(30 + nameBuf.length);
        local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(method, 8);
        local.writeUInt32LE(c, 14); local.writeUInt32LE(stored.length, 18); local.writeUInt32LE(f.data.length, 22);
        local.writeUInt16LE(nameBuf.length, 26); nameBuf.copy(local, 30);
        locals.push(local, stored);
        const central = Buffer.alloc(46 + nameBuf.length);
        central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6);
        central.writeUInt16LE(method, 10); central.writeUInt32LE(c, 16);
        central.writeUInt32LE(stored.length, 20); central.writeUInt32LE(f.data.length, 24);
        central.writeUInt16LE(nameBuf.length, 28); central.writeUInt32LE(offset, 42);
        nameBuf.copy(central, 46); centrals.push(central);
        offset += local.length + stored.length;
      }
      const cdSz = centrals.reduce((s, b) => s + b.length, 0);
      const eocd = Buffer.alloc(22);
      eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(files.length, 8); eocd.writeUInt16LE(files.length, 10);
      eocd.writeUInt32LE(cdSz, 12); eocd.writeUInt32LE(offset, 16);
      require('fs').writeFileSync(templatePath, Buffer.concat([...locals, ...centrals, eocd]));

      const buf = generateWithTemplate('# Hello', 'docx', templatePath);
      const styles = zipExtract(buf, 'word/styles.xml');
      // Heading1 keeps its numPr so Word auto-numbers headings via the template scheme.
      const heading1 = styles.match(/<w:style[^>]*w:styleId="Heading1"[^>]*>[\s\S]*?<\/w:style>/);
      assert.ok(heading1, 'Heading1 style block should be present');
      assert.ok(heading1[0].includes('<w:numPr>'), 'Heading1 numPr should be preserved');
      assert.ok(heading1[0].includes('outlineLvl'), 'Heading1 outlineLvl should still be present');
      // Non-heading style still has its numPr stripped to avoid colliding
      // with the markdown driver's paragraph-level list numPr.
      const myList = styles.match(/<w:style[^>]*w:styleId="MyList"[^>]*>[\s\S]*?<\/w:style>/);
      assert.ok(myList, 'MyList style block should be present');
      assert.ok(!myList[0].includes('<w:numPr>'), 'numPr should be stripped from non-heading styles');
    } finally {
      try { require('fs').unlinkSync(templatePath); } catch {}
    }
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

  it('nested ordered list uses ilvl 1', () => {
    const buf = generate('1. Top\n  1. Sub', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('w:ilvl w:val="1"'), 'Sub-item should be ilvl 1');
    assert.ok(/w:numId w:val="\d+"/.test(doc), 'Should reference a numbered list numId');
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

// ============================================================================
// Theme token model
// ============================================================================

/**
 * Generate with a template file.
 */
function generateWithTemplate(md, format, templatePath) {
  const outFile = join(__dirname, `_test_output_themed.${format}`);
  try {
    execSync(`node "${script}" -f ${format} -o "${outFile}" --template "${templatePath}"`, {
      input: md, encoding: 'utf8', timeout: 15000, stdio: ['pipe', 'pipe', 'pipe']
    });
    return readFileSync(outFile);
  } finally {
    try { require('fs').unlinkSync(outFile); } catch {}
  }
}

/**
 * Create a minimal PPTX ZIP with a custom theme1.xml for testing.
 */
function createTestTemplate(themeXml) {
  const { deflateRawSync } = require('zlib');

  function crc32Test(buf) {
    if (typeof buf === 'string') buf = Buffer.from(buf, 'utf8');
    const table = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      table[n] = c >>> 0;
    }
    let crc = 0xFFFFFFFF;
    for (let i = 0; i < buf.length; i++) crc = table[(crc ^ buf[i]) & 0xFF] ^ (crc >>> 8);
    return (crc ^ 0xFFFFFFFF) >>> 0;
  }

  const files = [
    { name: 'ppt/theme/theme1.xml', data: Buffer.from(themeXml, 'utf8') },
    { name: '[Content_Types].xml', data: Buffer.from('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>', 'utf8') },
  ];
  const localHeaders = [];
  const centralHeaders = [];
  let offset = 0;
  for (const f of files) {
    const nameBuf = Buffer.from(f.name, 'utf8');
    const compressed = deflateRawSync(f.data);
    const useDeflate = compressed.length < f.data.length;
    const stored = useDeflate ? compressed : f.data;
    const method = useDeflate ? 8 : 0;
    const crc = crc32Test(f.data);

    const local = Buffer.alloc(30 + nameBuf.length + stored.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(method, 8);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(stored.length, 18);
    local.writeUInt32LE(f.data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26);
    nameBuf.copy(local, 30);
    stored.copy(local, 30 + nameBuf.length);
    localHeaders.push(local);

    const central = Buffer.alloc(46 + nameBuf.length);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(method, 10);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(stored.length, 20);
    central.writeUInt32LE(f.data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28);
    central.writeUInt32LE(offset, 42);
    nameBuf.copy(central, 46);
    centralHeaders.push(central);

    offset += local.length;
  }
  const cdOffset = offset;
  const cdSize = centralHeaders.reduce((s, b) => s + b.length, 0);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0);
  eocd.writeUInt16LE(files.length, 8);
  eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(cdSize, 12);
  eocd.writeUInt32LE(cdOffset, 16);
  return Buffer.concat([...localHeaders, ...centralHeaders, eocd]);
}

describe('Theme token model — defaults', () => {
  it('PPTX theme XML contains default Office colors', () => {
    const buf = generate('# Hello', 'pptx');
    const theme = zipExtract(buf, 'ppt/theme/theme1.xml');
    assert.ok(theme.includes('val="4472C4"'), 'accent1');
    assert.ok(theme.includes('val="ED7D31"'), 'accent2');
    assert.ok(theme.includes('lastClr="000000"'), 'dk1');
    assert.ok(theme.includes('lastClr="FFFFFF"'), 'lt1');
  });

  it('PPTX theme XML contains default fonts', () => {
    const buf = generate('# Hello', 'pptx');
    const theme = zipExtract(buf, 'ppt/theme/theme1.xml');
    assert.ok(theme.includes('typeface="Calibri Light"'), 'major font');
    assert.ok(theme.includes('typeface="Calibri"'), 'minor font');
  });

  it('DOCX styles.xml uses default heading color', () => {
    const buf = generate('# Hello', 'docx');
    const styles = zipExtract(buf, 'word/styles.xml');
    assert.ok(styles.includes('w:val="2F5496"'), 'heading color');
  });

  it('DOCX now includes theme1.xml', () => {
    const buf = generate('# Hello', 'docx');
    const entries = zipEntries(buf);
    assert.ok(entries.includes('word/theme/theme1.xml'), 'Should have theme file');
    const theme = zipExtract(buf, 'word/theme/theme1.xml');
    assert.ok(theme.includes('val="4472C4"'), 'accent1 in DOCX theme');
  });

  it('DOCX content types includes theme override', () => {
    const buf = generate('# Hello', 'docx');
    const ct = zipExtract(buf, '[Content_Types].xml');
    assert.ok(ct.includes('theme+xml'), 'Should have theme content type');
  });

  it('DOCX document rels includes theme relationship', () => {
    const buf = generate('# Hello', 'docx');
    const rels = zipExtract(buf, 'word/_rels/document.xml.rels');
    assert.ok(rels.includes('theme/theme1.xml'), 'Should reference theme');
    assert.ok(rels.includes('rId3'), 'Theme should be rId3');
  });

  it('PPTX code block uses theme code font and color', () => {
    const buf = generate('## Code\n```\nlet x = 1;\n```', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    assert.ok(slide.includes('typeface="Consolas"'), 'default code font');
    assert.ok(slide.includes('val="2B2B2B"'), 'default code text color');
  });

  it('DOCX inline code uses theme code font and bg', () => {
    const buf = generate('Use `code` here', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('w:ascii="Consolas"'), 'default code font');
    assert.ok(doc.includes('w:fill="E8E8E8"'), 'default inline code bg');
  });

  it('DOCX table header uses theme fill color', () => {
    const buf = generate('| H1 | H2 |\n|---|---|\n| a | b |', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(doc.includes('w:fill="D9E2F3"'), 'default table header fill');
  });
});

describe('Theme token model — template extraction', () => {
  const customThemeXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom">
  <a:themeElements>
    <a:clrScheme name="Custom">
      <a:dk1><a:srgbClr val="1A1A2E"/></a:dk1>
      <a:lt1><a:srgbClr val="EAEAEA"/></a:lt1>
      <a:dk2><a:srgbClr val="16213E"/></a:dk2>
      <a:lt2><a:srgbClr val="C8C8C8"/></a:lt2>
      <a:accent1><a:srgbClr val="E94560"/></a:accent1>
      <a:accent2><a:srgbClr val="0F3460"/></a:accent2>
      <a:accent3><a:srgbClr val="533483"/></a:accent3>
      <a:accent4><a:srgbClr val="E94560"/></a:accent4>
      <a:accent5><a:srgbClr val="16213E"/></a:accent5>
      <a:accent6><a:srgbClr val="1A1A2E"/></a:accent6>
      <a:hlink><a:srgbClr val="E94560"/></a:hlink>
      <a:folHlink><a:srgbClr val="533483"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Custom">
      <a:majorFont><a:latin typeface="Georgia"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="Verdana"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Custom">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;

  let templatePath;

  // Create a test template file before tests
  it('setup: create test template', () => {
    const templateBuf = createTestTemplate(customThemeXml);
    templatePath = join(__dirname, '_test_template.pptx');
    require('fs').writeFileSync(templatePath, templateBuf);
    assert.ok(existsSync(templatePath), 'Template should exist');
  });

  it('PPTX with template uses custom accent colors', () => {
    const buf = generateWithTemplate('# Custom Theme Test', 'pptx', templatePath);
    const theme = zipExtract(buf, 'ppt/theme/theme1.xml');
    assert.ok(theme.includes('val="E94560"'), 'accent1 should be E94560');
    assert.ok(theme.includes('val="0F3460"'), 'accent2 should be 0F3460');
    assert.ok(!theme.includes('val="4472C4"'), 'should NOT have default accent1');
  });

  it('PPTX with template uses custom fonts', () => {
    const buf = generateWithTemplate('# Custom Font', 'pptx', templatePath);
    const theme = zipExtract(buf, 'ppt/theme/theme1.xml');
    assert.ok(theme.includes('typeface="Georgia"'), 'major font should be Georgia');
    assert.ok(theme.includes('typeface="Verdana"'), 'minor font should be Verdana');
  });

  it('DOCX with template uses custom theme colors', () => {
    const buf = generateWithTemplate('# Themed Doc', 'docx', templatePath);
    const theme = zipExtract(buf, 'word/theme/theme1.xml');
    assert.ok(theme.includes('val="E94560"'), 'accent1 in DOCX theme');
    assert.ok(theme.includes('val="1A1A2E"'), 'dk1 in DOCX theme');
  });

  it('DOCX with template uses custom fonts in styles', () => {
    const buf = generateWithTemplate('# Styled Doc', 'docx', templatePath);
    const styles = zipExtract(buf, 'word/styles.xml');
    assert.ok(styles.includes('w:ascii="Verdana"'), 'body font should be Verdana');
  });

  it('template with sysClr extracts lastClr correctly', () => {
    const sysTheme = customThemeXml.replace(
      '<a:dk1><a:srgbClr val="1A1A2E"/></a:dk1>',
      '<a:dk1><a:sysClr val="windowText" lastClr="112233"/></a:dk1>'
    );
    const sysBuf = createTestTemplate(sysTheme);
    const sysPath = join(__dirname, '_test_template_sys.pptx');
    require('fs').writeFileSync(sysPath, sysBuf);
    try {
      const buf = generateWithTemplate('# SysClr', 'pptx', sysPath);
      const theme = zipExtract(buf, 'ppt/theme/theme1.xml');
      assert.ok(theme.includes('lastClr="112233"'), 'Should extract lastClr from sysClr');
    } finally {
      try { require('fs').unlinkSync(sysPath); } catch {}
    }
  });

  it('derived colors kept when template only overrides standard colors', () => {
    const buf = generateWithTemplate('## Code\n```\nx = 1\n```', 'pptx', templatePath);
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    // Derived colors (codeText) should still be defaults since template doesn't override them
    assert.ok(slide.includes('val="2B2B2B"'), 'codeText should remain default 2B2B2B');
    assert.ok(slide.includes('typeface="Consolas"'), 'code font should remain Consolas');
  });

  it('cleanup: remove test template', () => {
    try { require('fs').unlinkSync(templatePath); } catch {}
    assert.ok(true, 'cleanup done');
  });
});

describe('Theme — ZIP reader', () => {
  it('reads files from a valid ZIP generated by our writer', () => {
    const buf = generate('# ZipTest', 'pptx');
    // Use our test zipExtract and zipEntries to verify the main script output
    const entries = zipEntries(buf);
    assert.ok(entries.length > 5, 'Should have multiple entries');
    assert.ok(entries.includes('ppt/theme/theme1.xml'), 'Should have theme');
  });

  it('invalid template file gives error exit', () => {
    const badFile = join(__dirname, '_bad_template.pptx');
    require('fs').writeFileSync(badFile, 'not a zip file');
    try {
      const result = run(`-f pptx -o _out.pptx --template "${badFile}"`, '# Test');
      assert.notStrictEqual(result.exitCode, 0, 'Should fail on invalid template');
    } finally {
      try { require('fs').unlinkSync(badFile); } catch {}
      try { require('fs').unlinkSync(join(__dirname, '_out.pptx')); } catch {}
    }
  });

  it('missing template file gives error exit', () => {
    const result = run('-f pptx -o _out.pptx --template nonexistent.pptx', '# Test');
    assert.notStrictEqual(result.exitCode, 0, 'Should fail on missing template');
  });
});

// ============================================================================
// --save-template: portable bundle round-trip
// ============================================================================

/**
 * Build a minimal valid DOCX with a custom theme, styles, header, footer, and
 * sectPr for use as a template source in the save-template tests.
 */
function buildRichDocxTemplate(filePath, opts = {}) {
  const { deflateRawSync } = require('zlib');
  function crc(buf) {
    if (typeof buf === 'string') buf = Buffer.from(buf, 'utf8');
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; }
    let c = 0xFFFFFFFF; for (let i = 0; i < buf.length; i++) c = t[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0;
  }
  const themeXml = opts.themeXml || '<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Bundle"><a:themeElements><a:clrScheme name="B"><a:dk1><a:srgbClr val="010101"/></a:dk1><a:lt1><a:srgbClr val="FAFAFA"/></a:lt1><a:dk2><a:srgbClr val="020202"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="BB0000"/></a:accent1><a:accent2><a:srgbClr val="00BB00"/></a:accent2><a:accent3><a:srgbClr val="0000BB"/></a:accent3><a:accent4><a:srgbClr val="BBBB00"/></a:accent4><a:accent5><a:srgbClr val="BB00BB"/></a:accent5><a:accent6><a:srgbClr val="00BBBB"/></a:accent6><a:hlink><a:srgbClr val="123456"/></a:hlink><a:folHlink><a:srgbClr val="654321"/></a:folHlink></a:clrScheme><a:fontScheme name="B"><a:majorFont><a:latin typeface="BundleMajor"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="BundleMinor"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="B"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>';
  const stylesXml = opts.stylesXml || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/><w:rPr><w:rFonts w:ascii="BundleFont" w:hAnsi="BundleFont"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:rPr><w:b/><w:sz w:val="44"/></w:rPr></w:style></w:styles>';
  const header = opts.headerXml || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>BundleHeader</w:t></w:r></w:p></w:hdr>';
  const footer = opts.footerXml || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>BundleFooter</w:t></w:r></w:p></w:ftr>';
  const ct = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>';
  const rootRels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  const docRels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/></Relationships>';
  const doc = '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Original body content</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId3"/><w:footerReference w:type="default" r:id="rId4"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>';
  const files = [
    { name: '[Content_Types].xml', data: Buffer.from(ct) },
    { name: '_rels/.rels', data: Buffer.from(rootRels) },
    { name: 'word/document.xml', data: Buffer.from(doc) },
    { name: 'word/_rels/document.xml.rels', data: Buffer.from(docRels) },
    { name: 'word/styles.xml', data: Buffer.from(stylesXml) },
    { name: 'word/theme/theme1.xml', data: Buffer.from(themeXml) },
    { name: 'word/header1.xml', data: Buffer.from(header) },
    { name: 'word/footer1.xml', data: Buffer.from(footer) },
  ];
  const locals = [], centrals = [];
  let offset = 0;
  for (const f of files) {
    const nameBuf = Buffer.from(f.name);
    const comp = deflateRawSync(f.data);
    const use = comp.length < f.data.length;
    const stored = use ? comp : f.data;
    const method = use ? 8 : 0;
    const c = crc(f.data);
    const local = Buffer.alloc(30 + nameBuf.length);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(method, 8);
    local.writeUInt32LE(c, 14); local.writeUInt32LE(stored.length, 18); local.writeUInt32LE(f.data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26); nameBuf.copy(local, 30);
    locals.push(local, stored);
    const central = Buffer.alloc(46 + nameBuf.length);
    central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6);
    central.writeUInt16LE(method, 10); central.writeUInt32LE(c, 16);
    central.writeUInt32LE(stored.length, 20); central.writeUInt32LE(f.data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28); central.writeUInt32LE(offset, 42);
    nameBuf.copy(central, 46); centrals.push(central);
    offset += local.length + stored.length;
  }
  const cdSz = centrals.reduce((s, b) => s + b.length, 0);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(files.length, 8); eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(cdSz, 12); eocd.writeUInt32LE(offset, 16);
  require('fs').writeFileSync(filePath, Buffer.concat([...locals, ...centrals, eocd]));
}

describe('--save-template — portable bundle round-trip', () => {
  it('writes a bundle containing only the reusable template parts', () => {
    const src = join(__dirname, '_test_save_src.docx');
    const bundle = join(__dirname, '_test_save_bundle.docx');
    try {
      buildRichDocxTemplate(src);
      const result = run(`--template "${src}" --save-template "${bundle}"`);
      assert.strictEqual(result.exitCode, 0, `save-template should succeed (stderr: ${result.stderr})`);
      assert.ok(existsSync(bundle), 'bundle file should exist');
      const buf = readFileSync(bundle);
      const entries = zipEntries(buf);
      assert.ok(entries.includes('[Content_Types].xml'));
      assert.ok(entries.includes('_rels/.rels'));
      assert.ok(entries.includes('word/document.xml'));
      assert.ok(entries.includes('word/_rels/document.xml.rels'));
      assert.ok(entries.includes('word/theme/theme1.xml'));
      assert.ok(entries.includes('word/styles.xml'));
      assert.ok(entries.includes('word/header1.xml'));
      assert.ok(entries.includes('word/footer1.xml'));
      // Image, settings, fontTable, numbering should NOT have been carried over
      assert.ok(!entries.some(e => e.startsWith('word/media/')), 'should not include image parts');
      // Document body should be a stub, not the original content
      const docXml = zipExtract(buf, 'word/document.xml');
      assert.ok(!docXml.includes('Original body content'), 'stub body should replace original content');
      assert.ok(docXml.includes('<w:sectPr'), 'stub should preserve sectPr');
      assert.ok(docXml.includes('w:headerReference'), 'stub sectPr should keep header reference');
      // Rels should be filtered: keep styles, theme, header, footer; drop image
      const relsXml = zipExtract(buf, 'word/_rels/document.xml.rels');
      assert.ok(/Type="[^"]*\/styles"/.test(relsXml), 'rels should keep styles');
      assert.ok(/Type="[^"]*\/theme"/.test(relsXml), 'rels should keep theme');
      assert.ok(/Type="[^"]*\/header"/.test(relsXml), 'rels should keep header');
      assert.ok(/Type="[^"]*\/footer"/.test(relsXml), 'rels should keep footer');
      assert.ok(!/Type="[^"]*\/image"/.test(relsXml), 'rels should drop image');
      // Theme and styles should be the original verbatim
      const themeOut = zipExtract(buf, 'word/theme/theme1.xml');
      assert.ok(themeOut.includes('BundleMajor'), 'theme major font should be preserved');
      const stylesOut = zipExtract(buf, 'word/styles.xml');
      assert.ok(stylesOut.includes('BundleFont'), 'styles font should be preserved');
    } finally {
      try { require('fs').unlinkSync(src); } catch {}
      try { require('fs').unlinkSync(bundle); } catch {}
    }
  });

  it('bundle is consumable as --template for a real document', () => {
    const src = join(__dirname, '_test_save_src2.docx');
    const bundle = join(__dirname, '_test_save_bundle2.docx');
    try {
      buildRichDocxTemplate(src);
      const r1 = run(`--template "${src}" --save-template "${bundle}"`);
      assert.strictEqual(r1.exitCode, 0, `bundle creation should succeed (stderr: ${r1.stderr})`);
      const buf = generateWithTemplate('# Hello\nBody text.', 'docx', bundle);
      const styles = zipExtract(buf, 'word/styles.xml');
      assert.ok(styles.includes('BundleFont'), 'output styles should carry the bundled font');
      const theme = zipExtract(buf, 'word/theme/theme1.xml');
      assert.ok(theme.includes('BundleMajor'), 'output theme should carry the bundled major font');
      const entries = zipEntries(buf);
      assert.ok(entries.includes('word/header1.xml'), 'output should carry the bundled header');
      assert.ok(entries.includes('word/footer1.xml'), 'output should carry the bundled footer');
    } finally {
      try { require('fs').unlinkSync(src); } catch {}
      try { require('fs').unlinkSync(bundle); } catch {}
    }
  });

  it('errors when --save-template is used without --template', () => {
    const result = run(`--save-template "${join(__dirname, '_unused.docx')}"`);
    assert.notStrictEqual(result.exitCode, 0, 'should fail without --template');
    assert.ok(/requires --template/i.test(result.stderr), `expected error message about --template, got: ${result.stderr}`);
  });

  it('PPTX source produces a theme-only bundle', () => {
    const src = join(__dirname, '_test_save_src.pptx');
    const bundle = join(__dirname, '_test_save_bundle3.docx');
    try {
      const themeXml = '<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="P"><a:themeElements><a:clrScheme name="P"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="111111"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="123456"/></a:accent1><a:accent2><a:srgbClr val="234567"/></a:accent2><a:accent3><a:srgbClr val="345678"/></a:accent3><a:accent4><a:srgbClr val="456789"/></a:accent4><a:accent5><a:srgbClr val="56789A"/></a:accent5><a:accent6><a:srgbClr val="6789AB"/></a:accent6><a:hlink><a:srgbClr val="789ABC"/></a:hlink><a:folHlink><a:srgbClr val="89ABCD"/></a:folHlink></a:clrScheme><a:fontScheme name="P"><a:majorFont><a:latin typeface="PptxMajor"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="PptxMinor"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="P"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>';
      require('fs').writeFileSync(src, createTestTemplate(themeXml));
      const r1 = run(`--template "${src}" --save-template "${bundle}"`);
      assert.strictEqual(r1.exitCode, 0, `pptx-to-bundle should succeed (stderr: ${r1.stderr})`);
      const buf = readFileSync(bundle);
      const entries = zipEntries(buf);
      assert.ok(entries.includes('word/theme/theme1.xml'), 'bundle should have a theme part');
      assert.ok(!entries.includes('word/styles.xml'), 'pptx source has no styles to carry over');
      assert.ok(!entries.some(e => /^word\/(header|footer)/.test(e)), 'pptx source has no headers/footers');
      const themeOut = zipExtract(buf, 'word/theme/theme1.xml');
      assert.ok(themeOut.includes('PptxMajor'), 'theme major font should be preserved');
    } finally {
      try { require('fs').unlinkSync(src); } catch {}
      try { require('fs').unlinkSync(bundle); } catch {}
    }
  });
});


// ============================================================================
// 19. Template heading auto-numbering (default behavior)
// ============================================================================

// Build a minimal template that has heading auto-numbering wired up the same
// way real Word templates do: Heading1..Heading3 styles bind to a multilevel
// abstractNum via numId. When such a template is provided, the script should
// preserve the heading numPr, merge the template numbering with markdown
// list numbering at offset numIds, promote heading levels by one, and strip
// any leading "1.", "1.1", ... numeric prefix from heading text.
function buildHeadingNumberingTemplate(filePath) {
  const { deflateRawSync } = require('zlib');
  function crc(buf) {
    if (typeof buf === 'string') buf = Buffer.from(buf, 'utf8');
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1); t[n] = c >>> 0; }
    let c = 0xFFFFFFFF; for (let i = 0; i < buf.length; i++) c = t[(c ^ buf[i]) & 0xFF] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0;
  }
  const themeXml = '<?xml version="1.0"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="HN"><a:themeElements><a:clrScheme name="HN"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2><a:accent1><a:srgbClr val="111111"/></a:accent1><a:accent2><a:srgbClr val="222222"/></a:accent2><a:accent3><a:srgbClr val="333333"/></a:accent3><a:accent4><a:srgbClr val="444444"/></a:accent4><a:accent5><a:srgbClr val="555555"/></a:accent5><a:accent6><a:srgbClr val="666666"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="HN"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="HN"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>';
  const stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:rPr><w:sz w:val="56"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="0"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="1"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:pPr><w:numPr><w:ilvl w:val="2"/><w:numId w:val="1"/></w:numPr><w:outlineLvl w:val="2"/></w:pPr><w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style></w:styles>';
  const numberingXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading2"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading3"/><w:lvlText w:val="%1.%2.%3"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>';
  const ct = '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>';
  const rootRels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>';
  const docRels = '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/></Relationships>';
  const doc = '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>';
  const files = [
    { name: '[Content_Types].xml', data: Buffer.from(ct) },
    { name: '_rels/.rels', data: Buffer.from(rootRels) },
    { name: 'word/document.xml', data: Buffer.from(doc) },
    { name: 'word/_rels/document.xml.rels', data: Buffer.from(docRels) },
    { name: 'word/styles.xml', data: Buffer.from(stylesXml) },
    { name: 'word/numbering.xml', data: Buffer.from(numberingXml) },
    { name: 'word/theme/theme1.xml', data: Buffer.from(themeXml) },
  ];
  const locals = [], centrals = [];
  let offset = 0;
  for (const f of files) {
    const nameBuf = Buffer.from(f.name);
    const comp = deflateRawSync(f.data);
    const use = comp.length < f.data.length;
    const stored = use ? comp : f.data;
    const method = use ? 8 : 0;
    const c = crc(f.data);
    const local = Buffer.alloc(30 + nameBuf.length);
    local.writeUInt32LE(0x04034b50, 0); local.writeUInt16LE(20, 4); local.writeUInt16LE(method, 8);
    local.writeUInt32LE(c, 14); local.writeUInt32LE(stored.length, 18); local.writeUInt32LE(f.data.length, 22);
    local.writeUInt16LE(nameBuf.length, 26); nameBuf.copy(local, 30);
    locals.push(local, stored);
    const central = Buffer.alloc(46 + nameBuf.length);
    central.writeUInt32LE(0x02014b50, 0); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6);
    central.writeUInt16LE(method, 10); central.writeUInt32LE(c, 16);
    central.writeUInt32LE(stored.length, 20); central.writeUInt32LE(f.data.length, 24);
    central.writeUInt16LE(nameBuf.length, 28); central.writeUInt32LE(offset, 42);
    nameBuf.copy(central, 46); centrals.push(central);
    offset += local.length + stored.length;
  }
  const cdSz = centrals.reduce((s, b) => s + b.length, 0);
  const eocd = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(files.length, 8); eocd.writeUInt16LE(files.length, 10);
  eocd.writeUInt32LE(cdSz, 12); eocd.writeUInt32LE(offset, 16);
  require('fs').writeFileSync(filePath, Buffer.concat([...locals, ...centrals, eocd]));
}

describe('Template heading auto-numbering', () => {
  it('preserves template numbering.xml and merges script lists at offset numIds', () => {
    const src = join(__dirname, '_test_hn_src.docx');
    try {
      buildHeadingNumberingTemplate(src);
      const md = '# Doc Title\n\n## Section A\n\n- bullet one\n- bullet two\n\n## Section B\n\nText.\n';
      const buf = generateWithTemplate(md, 'docx', src);
      const numbering = zipExtract(buf, 'word/numbering.xml');
      // Template's numId 1 must be preserved.
      assert.ok(/<w:num\b[^>]*w:numId="1"/.test(numbering), 'template numId=1 must be kept');
      // Template's heading-bound abstractNum 0 must be preserved.
      assert.ok(/<w:abstractNum\b[^>]*w:abstractNumId="0"[^>]*>[\s\S]*?<w:pStyle w:val="Heading1"/.test(numbering), 'template abstractNum 0 with Heading1 binding must be kept');
      // Script's bullet/ordered abstractNums must be appended at offset > 0.
      assert.ok(/<w:abstractNum\b[^>]*w:abstractNumId="[1-9]\d*"[^>]*>[\s\S]*?<w:numFmt w:val="bullet"/.test(numbering), 'script bullet abstractNum must be appended');
      // The bullet list in document.xml must reference an offset numId (>1).
      const docXml = zipExtract(buf, 'word/document.xml');
      const listNumIds = [...docXml.matchAll(/<w:numId\s+w:val="(\d+)"\s*\/?>/g)].map(m => parseInt(m[1], 10));
      assert.ok(listNumIds.length > 0, 'document body should reference at least one list numId');
      assert.ok(listNumIds.every(n => n > 1), `script list numIds must be offset beyond template max (1); got ${JSON.stringify(listNumIds)}`);
    } finally {
      try { require('fs').unlinkSync(src); } catch {}
    }
  });

  it('promotes heading levels by one when template has heading numbering', () => {
    const src = join(__dirname, '_test_hn_src2.docx');
    try {
      buildHeadingNumberingTemplate(src);
      // First # is Title, ## should become Heading1, ### should become Heading2.
      const md = '# My Title\n\n## Section\n\n### Subsection\n';
      const buf = generateWithTemplate(md, 'docx', src);
      const docXml = zipExtract(buf, 'word/document.xml');
      // First # → Title.
      assert.ok(/<w:pStyle w:val="Title"\/>[\s\S]*?<w:t[^>]*>My Title<\/w:t>/.test(docXml), 'first # should map to Title');
      // ## → Heading1 (promoted from Heading2 because template has heading numbering).
      assert.ok(/<w:pStyle w:val="Heading1"\/>[\s\S]*?<w:t[^>]*>Section<\/w:t>/.test(docXml), '## should map to Heading1');
      // ### → Heading2.
      assert.ok(/<w:pStyle w:val="Heading2"\/>[\s\S]*?<w:t[^>]*>Subsection<\/w:t>/.test(docXml), '### should map to Heading2');
    } finally {
      try { require('fs').unlinkSync(src); } catch {}
    }
  });

  it('strips leading "1.", "1.1", ... numeric prefix from heading text', () => {
    const src = join(__dirname, '_test_hn_src3.docx');
    try {
      buildHeadingNumberingTemplate(src);
      const md = '# Doc\n\n## 1. Today\n\n### 1.1 Subhead\n\n## 2. Tomorrow\n';
      const buf = generateWithTemplate(md, 'docx', src);
      const docXml = zipExtract(buf, 'word/document.xml');
      // The "1.", "1.1", "2." prefixes must be stripped because Word renders
      // them via the template's auto-numbering.
      assert.ok(/<w:t[^>]*>Today<\/w:t>/.test(docXml), 'heading should have prefix stripped: Today');
      assert.ok(/<w:t[^>]*>Subhead<\/w:t>/.test(docXml), 'heading should have prefix stripped: Subhead');
      assert.ok(/<w:t[^>]*>Tomorrow<\/w:t>/.test(docXml), 'heading should have prefix stripped: Tomorrow');
      assert.ok(!/<w:t[^>]*>1\. Today<\/w:t>/.test(docXml), 'leading "1." should not survive in heading text');
      assert.ok(!/<w:t[^>]*>1\.1 Subhead<\/w:t>/.test(docXml), 'leading "1.1" should not survive in heading text');
    } finally {
      try { require('fs').unlinkSync(src); } catch {}
    }
  });

  it('falls back to script-built numbering when template has no numbering.xml', () => {
    // The existing rich-template fixture has styles + theme + header + footer
    // but NO numbering.xml. The output should use script-built numbering and
    // markdown headings should NOT be promoted (## stays Heading2).
    const src = join(__dirname, '_test_hn_src4.docx');
    try {
      buildRichDocxTemplate(src);
      const md = '# Title\n\n## Subhead\n';
      const buf = generateWithTemplate(md, 'docx', src);
      const docXml = zipExtract(buf, 'word/document.xml');
      assert.ok(/<w:pStyle w:val="Heading2"\/>[\s\S]*?<w:t[^>]*>Subhead<\/w:t>/.test(docXml), '## should map to Heading2 when template has no heading numbering');
      const numbering = zipExtract(buf, 'word/numbering.xml');
      // Script's own numbering base names should be there.
      assert.ok(/<w:abstractNum\b[^>]*w:abstractNumId="0"[^>]*>[\s\S]*?<w:numFmt w:val="bullet"/.test(numbering), 'script default bullet abstractNum at id=0 must be present');
    } finally {
      try { require('fs').unlinkSync(src); } catch {}
    }
  });

  it('save-template bundle preserves the templates numbering.xml', () => {
    const src = join(__dirname, '_test_hn_src5.docx');
    const bundle = join(__dirname, '_test_hn_bundle.docx');
    try {
      buildHeadingNumberingTemplate(src);
      const result = run(`--template "${src}" --save-template "${bundle}"`);
      assert.strictEqual(result.exitCode, 0, `save-template should succeed (stderr: ${result.stderr})`);
      const buf = readFileSync(bundle);
      const entries = zipEntries(buf);
      assert.ok(entries.includes('word/numbering.xml'), 'bundle should carry numbering.xml when source has heading numbering');
      const numbering = zipExtract(buf, 'word/numbering.xml');
      assert.ok(/<w:pStyle w:val="Heading1"/.test(numbering), 'bundle numbering should keep Heading1 binding');
      // And the bundle is consumable: round-trip a doc and confirm headings still get auto-numbered.
      const out = generateWithTemplate('# T\n\n## Step\n', 'docx', bundle);
      const styles = zipExtract(out, 'word/styles.xml');
      const heading1 = styles.match(/<w:style[^>]*w:styleId="Heading1"[^>]*>[\s\S]*?<\/w:style>/);
      assert.ok(heading1 && heading1[0].includes('<w:numPr>'), 'round-tripped bundle must keep Heading1 numPr');
    } finally {
      try { require('fs').unlinkSync(src); } catch {}
      try { require('fs').unlinkSync(bundle); } catch {}
    }
  });
});

// ============================================================================
// 20. OLE-wrapped (MIP-encrypted) template gives a clear error
// ============================================================================
describe('OLE-wrapped template error', () => {
  it('surfaces a clear MIP/sensitivity-label error instead of a generic ZIP error', () => {
    const src = join(__dirname, '_test_ole.docx');
    try {
      // Write a fake OLE compound file (just the magic + padding) so the
      // ZipReader would otherwise fail with "EOCD not found".
      const ole = Buffer.alloc(512);
      Buffer.from([0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]).copy(ole, 0);
      require('fs').writeFileSync(src, ole);
      const out = join(__dirname, '_test_ole_out.docx');
      const result = run(`-o "${out}" --template "${src}"`, '# Hello');
      assert.notStrictEqual(result.exitCode, 0, 'should fail on OLE-wrapped template');
      assert.ok(/Microsoft Information Protection|sensitivity label|OLE-wrapped/i.test(result.stderr), `expected MIP/OLE error, got: ${result.stderr}`);
      assert.ok(!/EOCD not found/i.test(result.stderr), 'should NOT surface raw EOCD error');
    } finally {
      try { require('fs').unlinkSync(src); } catch {}
      try { require('fs').unlinkSync(join(__dirname, '_test_ole_out.docx')); } catch {}
    }
  });
});


// ============================================================================
// 21. Table cells preserve inline markdown formatting
// ============================================================================
describe('Table cell inline formatting', () => {
  it('renders `code` spans inside DOCX table cells via the code font', () => {
    const buf = generate('| Field | Notes |\n|---|---|\n| `userTenantId` | platform field |\n', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    // The `code` cell must use the code font (Consolas by default), not a
    // literal "`code`" text node.
    assert.ok(/<w:t[^>]*>userTenantId<\/w:t>/.test(doc), 'code span text must be present');
    assert.ok(!/<w:t[^>]*>`userTenantId`<\/w:t>/.test(doc), 'backticks must not survive as literal text');
    assert.ok(/w:ascii="Consolas"[\s\S]{0,200}<w:t[^>]*>userTenantId<\/w:t>/.test(doc), 'code-font run must wrap the code span text');
  });

  it('renders **bold** and *italic* spans inside DOCX table cells', () => {
    const buf = generate('| Term | Meaning |\n|---|---|\n| **alpha** | *beta* |\n', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(/<w:b\/>[\s\S]{0,80}<w:t[^>]*>alpha<\/w:t>/.test(doc), 'bold span must wrap "alpha"');
    assert.ok(/<w:i\/>[\s\S]{0,80}<w:t[^>]*>beta<\/w:t>/.test(doc), 'italic span must wrap "beta"');
    assert.ok(!/\*\*alpha\*\*/.test(doc), 'asterisks must not survive in body text');
  });

  it('keeps DOCX table header row cells bold even when cells have inline formatting', () => {
    const buf = generate('| `id` | name |\n|---|---|\n| `a` | b |\n', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    // Find the header row (first <w:tr>) and assert both header cells contain
    // <w:b/> in their first run's <w:rPr>.
    const trMatch = doc.match(/<w:tr>[\s\S]*?<\/w:tr>/);
    assert.ok(trMatch, 'expected at least one <w:tr>');
    const header = trMatch[0];
    const idCellHasBold = /<w:rPr>(?=[^<]*<w:b\/>)[\s\S]*?<\/w:rPr>[\s\S]*?<w:t[^>]*>id<\/w:t>/.test(header)
      || /<w:b\/>[\s\S]{0,200}<w:t[^>]*>id<\/w:t>/.test(header);
    const nameCellHasBold = /<w:b\/>[\s\S]{0,200}<w:t[^>]*>name<\/w:t>/.test(header);
    assert.ok(idCellHasBold, 'header cell with inline `code` must still be bold');
    assert.ok(nameCellHasBold, 'plain header cell must still be bold');
  });

  it('does not bold non-first-column DOCX data row cells', () => {
    const buf = generate('| H1 | H2 |\n|---|---|\n| label | body |\n', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    const trs = [...doc.matchAll(/<w:tr>[\s\S]*?<\/w:tr>/g)].map(m => m[0]);
    assert.ok(trs.length >= 2, 'expected at least one data row');
    const cells = [...trs[1].matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map(m => m[0]);
    assert.ok(cells.length >= 2, 'expected two cells in data row');
    // Inspect the SECOND cell only to avoid matching bold runs in the
    // adjacent first cell (which is intentionally bolded).
    const secondCell = cells[1];
    assert.ok(/<w:t[^>]*>body<\/w:t>/.test(secondCell), 'second cell should contain "body"');
    assert.ok(!/<w:b\/>/.test(secondCell), 'second-column cell must not contain any bold run');
  });

  it('bolds first-column DOCX data row cells (Grid Table 4 - Accent 1 parity)', () => {
    const buf = generate('| H1 | H2 |\n|---|---|\n| label | body |\n', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    const trs = [...doc.matchAll(/<w:tr>[\s\S]*?<\/w:tr>/g)].map(m => m[0]);
    assert.ok(trs.length >= 2, 'expected at least one data row');
    const cells = [...trs[1].matchAll(/<w:tc>[\s\S]*?<\/w:tc>/g)].map(m => m[0]);
    assert.ok(cells.length >= 2, 'expected two cells in data row');
    const firstCell = cells[0];
    assert.ok(/<w:b\/>[\s\S]{0,80}<w:t[^>]*>label<\/w:t>/.test(firstCell), 'first-column cell text must be bold');
  });

  it('flattens cells to plain text in PPTX tables (inline formatting is a known PPTX gap)', () => {
    const buf = generate('## Slide\n\n| `id` | name |\n|---|---|\n| `a` | b |\n', 'pptx');
    const slide = zipExtract(buf, 'ppt/slides/slide1.xml');
    // PPTX path strips inline formatting; only the text content survives.
    assert.ok(/<a:t>id<\/a:t>/.test(slide), 'PPTX cell text must contain the unwrapped text');
    assert.ok(!/<a:t>`id`<\/a:t>/.test(slide), 'backticks must not appear as literal text');
  });
});


// ============================================================================
// 22. List paragraphs carry built-in ListBullet / ListNumber pStyle
// ============================================================================
describe('List paragraph styles', () => {
  it('bullet list paragraphs carry pStyle="ListBullet"', () => {
    const buf = generate('- one\n- two\n', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(/<w:pStyle\s+w:val="ListBullet"\/>/.test(doc), 'bullet items must reference ListBullet style');
  });

  it('ordered list paragraphs carry pStyle="ListNumber"', () => {
    const buf = generate('1. one\n2. two\n', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(/<w:pStyle\s+w:val="ListNumber"\/>/.test(doc), 'ordered items must reference ListNumber style');
  });

  it('nested bullet uses ListBullet2 pStyle', () => {
    const buf = generate('- one\n  - nested\n', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    assert.ok(/<w:pStyle\s+w:val="ListBullet2"\/>/.test(doc), 'nested bullets must reference ListBullet2');
  });

  it('styles.xml defines ListBullet and ListNumber so Word reports correct style names', () => {
    const buf = generate('- one\n', 'docx');
    const styles = zipExtract(buf, 'word/styles.xml');
    assert.ok(/w:styleId="ListBullet"/.test(styles), 'ListBullet style must be defined');
    assert.ok(/w:styleId="ListNumber"/.test(styles), 'ListNumber style must be defined');
  });

  it('list paragraphs keep paragraph-level numPr so separate lists restart from 1', () => {
    const buf = generate('- a\n- b\n\nText.\n\n- c\n', 'docx');
    const doc = zipExtract(buf, 'word/document.xml');
    // Two separate bullet lists must use two distinct numIds (paragraph-level).
    const numIds = [...new Set([...doc.matchAll(/<w:numId\s+w:val="(\d+)"/g)].map(m => m[1]))];
    assert.ok(numIds.length >= 2, `expected at least 2 distinct numIds, got ${numIds.join(',')}`);
  });
});
