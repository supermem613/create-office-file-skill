#!/usr/bin/env node
// verify-output.mjs — Verify generated .pptx/.docx by parsing OOXML XML directly
// Zero dependencies, cross-platform — no COM, no Office, never hangs
// Outputs JSON to stdout with extracted text, formatting, and structure info
'use strict';

import { readFileSync, existsSync } from 'fs';
import { resolve, extname } from 'path';
import { inflateRawSync } from 'zlib';

const file = process.argv[2];
if (!file) { console.error('Usage: node verify-output.mjs <file.pptx|docx>'); process.exit(1); }

const absPath = resolve(file);
if (!existsSync(absPath)) { console.error(`ERROR: file not found: ${absPath}`); process.exit(1); }

const ext = extname(absPath).toLowerCase();
if (ext !== '.pptx' && ext !== '.docx') { console.error('ERROR: file must be .pptx or .docx'); process.exit(1); }

// ============================================================================
// ZIP Reader
// ============================================================================

class ZipReader {
  constructor(buffer) {
    this.buf = buffer;
    this.entries = this._parse();
  }
  _parse() {
    const buf = this.buf;
    let eocd = -1;
    for (let i = buf.length - 22; i >= Math.max(0, buf.length - 65557); i--) {
      if (buf.readUInt32LE(i) === 0x06054b50) { eocd = i; break; }
    }
    if (eocd < 0) throw new Error('Invalid ZIP: EOCD not found');
    const cdSize = buf.readUInt32LE(eocd + 12);
    const cdOffset = buf.readUInt32LE(eocd + 16);
    const entries = new Map();
    let pos = cdOffset;
    while (pos < cdOffset + cdSize) {
      if (buf.readUInt32LE(pos) !== 0x02014b50) break;
      const method = buf.readUInt16LE(pos + 10);
      const compSize = buf.readUInt32LE(pos + 20);
      const nameLen = buf.readUInt16LE(pos + 28);
      const extraLen = buf.readUInt16LE(pos + 30);
      const commentLen = buf.readUInt16LE(pos + 32);
      const localOffset = buf.readUInt32LE(pos + 42);
      const name = buf.toString('utf8', pos + 46, pos + 46 + nameLen);
      entries.set(name, { method, compSize, localOffset });
      pos += 46 + nameLen + extraLen + commentLen;
    }
    return entries;
  }
  getFile(name) {
    const entry = this.entries.get(name);
    if (!entry) return null;
    const pos = entry.localOffset;
    if (this.buf.readUInt32LE(pos) !== 0x04034b50) return null;
    const nameLen = this.buf.readUInt16LE(pos + 26);
    const extraLen = this.buf.readUInt16LE(pos + 28);
    const dataStart = pos + 30 + nameLen + extraLen;
    const raw = this.buf.subarray(dataStart, dataStart + entry.compSize);
    if (entry.method === 0) return raw;
    if (entry.method === 8) return inflateRawSync(raw);
    return null;
  }
  listFiles() { return [...this.entries.keys()]; }
}

// ============================================================================
// XML helpers
// ============================================================================

function unesc(s) {
  return s.replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'");
}

function extractText(xml, nsPrefix) {
  const texts = [];
  const re = new RegExp(`<${nsPrefix}:t(?:\\s[^>]*)?>([^<]*)</${nsPrefix}:t>`, 'g');
  let m;
  while ((m = re.exec(xml)) !== null) texts.push(unesc(m[1]));
  return texts.join('');
}

// ============================================================================
// PPTX Verification
// ============================================================================

function extractTheme(zip, themeFile) {
  const xml = zip.getFile(themeFile)?.toString('utf8');
  if (!xml) return null;
  const theme = { colors: {}, fonts: {} };
  const colorNames = ['dk1','lt1','dk2','lt2','accent1','accent2','accent3','accent4','accent5','accent6','hlink','folHlink'];
  for (const cn of colorNames) {
    const re = new RegExp(`<a:${cn}>.*?(?:val="([A-Fa-f0-9]{6})".*?|lastClr="([A-Fa-f0-9]{6})".*?)<\\/a:${cn}>`, 's');
    const m = xml.match(re);
    if (m) theme.colors[cn] = (m[1] || m[2]).toUpperCase();
  }
  const majorMatch = xml.match(/<a:majorFont>.*?<a:latin[^>]*typeface="([^"]*)".*?<\/a:majorFont>/s);
  const minorMatch = xml.match(/<a:minorFont>.*?<a:latin[^>]*typeface="([^"]*)".*?<\/a:minorFont>/s);
  if (majorMatch) theme.fonts.major = unesc(majorMatch[1]);
  if (minorMatch) theme.fonts.minor = unesc(minorMatch[1]);
  return theme;
}

function extractHeadersFooters(zip) {
  const parts = [];
  for (const name of zip.listFiles()) {
    if (/^word\/(header|footer)\d+\.xml$/.test(name)) {
      const xml = zip.getFile(name)?.toString('utf8');
      if (!xml) continue;
      const texts = [];
      const re = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g;
      let m;
      while ((m = re.exec(xml)) !== null) texts.push(unesc(m[1]));
      parts.push({ name, text: texts.join('') });
    }
  }
  return parts;
}

function verifyPptx(zip) {
  const result = { type: 'pptx', file: absPath, slides: [], slideCount: 0, error: null, repairNeeded: false };

  const themeFile = zip.listFiles().find(f => /^ppt\/theme\/theme\d+\.xml$/.test(f));
  if (themeFile) result.theme = extractTheme(zip, themeFile);

  const presXml = zip.getFile('ppt/presentation.xml')?.toString('utf8');
  if (!presXml) { result.error = 'Missing ppt/presentation.xml'; return result; }

  const sldSzMatch = presXml.match(/<p:sldSz[^>]*cx="(\d+)"[^>]*cy="(\d+)"/);
  if (sldSzMatch) {
    result.slideWidth = Math.round(parseInt(sldSzMatch[1]) / 12700 * 1.333);
    result.slideHeight = Math.round(parseInt(sldSzMatch[2]) / 12700 * 1.333);
  }

  const slideFiles = zip.listFiles()
    .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => parseInt(a.match(/slide(\d+)/)[1]) - parseInt(b.match(/slide(\d+)/)[1]));

  result.slideCount = slideFiles.length;

  for (const slideFile of slideFiles) {
    const xml = zip.getFile(slideFile)?.toString('utf8');
    if (!xml) continue;
    const slideIdx = parseInt(slideFile.match(/slide(\d+)/)[1]);
    const slideInfo = { index: slideIdx, shapes: [] };

    // Pictures (<p:pic>)
    const picRegex = /<p:pic\b[^>]*>[\s\S]*?<\/p:pic>/g;
    let picMatch;
    while ((picMatch = picRegex.exec(xml)) !== null) {
      slideInfo.shapes.push({ name: 'Picture', hasText: false, hasTable: false, isPicture: true, shapeType: 13 });
    }

    // Regular shapes (<p:sp>)
    const spRegex = /<p:sp\b[^>]*>([\s\S]*?)<\/p:sp>/g;
    let spMatch;
    while ((spMatch = spRegex.exec(xml)) !== null) {
      const spXml = spMatch[1];
      const shapeInfo = { name: '', hasText: false, hasTable: false, isPicture: false, shapeType: 17 };

      const nameMatch = spXml.match(/<p:cNvPr[^>]*name="([^"]*)"/);
      if (nameMatch) shapeInfo.name = unesc(nameMatch[1]);

      const txBodyMatch = spXml.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/);
      if (txBodyMatch) {
        const txXml = txBodyMatch[1];
        const fullText = extractText(txXml, 'a');
        if (fullText.length > 0) {
          shapeInfo.hasText = true;
          shapeInfo.text = fullText;
          shapeInfo.paragraphs = [];

          const paraRegex = /<a:p\b[^>]*>([\s\S]*?)<\/a:p>/g;
          let pm;
          while ((pm = paraRegex.exec(txXml)) !== null) {
            const paraXml = pm[1];
            const paraInfo = { text: extractText(paraXml, 'a'), runs: [] };

            const runRegex = /<a:r>([\s\S]*?)<\/a:r>/g;
            let rm;
            while ((rm = runRegex.exec(paraXml)) !== null) {
              const runXml = rm[1];
              const runText = extractText(runXml, 'a');
              if (!runText) continue;

              const rprMatch = runXml.match(/<a:rPr([^>]*?)\/?>(?:[\s\S]*?<\/a:rPr>)?/);
              const rprAttrs = rprMatch ? rprMatch[0] : '';
              const bold = /\bb="1"/.test(rprAttrs);
              const italic = /\bi="1"/.test(rprAttrs);

              let fontName = 'Calibri';
              const latinMatch = runXml.match(/<a:latin[^>]*typeface="([^"]*)"/);
              if (latinMatch) fontName = unesc(latinMatch[1]);

              const sizeMatch = rprAttrs.match(/\bsz="(\d+)"/);
              const fontSize = sizeMatch ? parseInt(sizeMatch[1]) / 100 : 18;

              paraInfo.runs.push({ text: runText, bold, italic, fontName, fontSize });
            }
            shapeInfo.paragraphs.push(paraInfo);
          }
        }
      }
      slideInfo.shapes.push(shapeInfo);
    }

    // Tables (<p:graphicFrame> containing <a:tbl>)
    const gfRegex = /<p:graphicFrame\b[^>]*>([\s\S]*?)<\/p:graphicFrame>/g;
    let gfMatch;
    while ((gfMatch = gfRegex.exec(xml)) !== null) {
      const gfXml = gfMatch[1];
      if (!gfXml.includes('<a:tbl')) continue;
      const shapeInfo = { name: 'Table', hasText: true, hasTable: true, isPicture: false, shapeType: 19 };
      const cellTexts = [];
      const cells = [];
      const trRegex = /<a:tr\b[^>]*>([\s\S]*?)<\/a:tr>/g;
      let rowIdx = 0, trMatch;
      while ((trMatch = trRegex.exec(gfXml)) !== null) {
        rowIdx++;
        const tcRegex = /<a:tc\b[^>]*>([\s\S]*?)<\/a:tc>/g;
        let colIdx = 0, tcMatch;
        while ((tcMatch = tcRegex.exec(trMatch[1])) !== null) {
          colIdx++;
          const cellText = extractText(tcMatch[1], 'a');
          cellTexts.push(cellText);
          cells.push({ row: rowIdx, col: colIdx, text: cellText });
        }
      }
      shapeInfo.text = cellTexts.join(' ');
      shapeInfo.cells = cells;
      shapeInfo.tableRows = rowIdx;
      shapeInfo.tableCols = cells.length > 0 ? Math.max(...cells.map(c => c.col)) : 0;
      slideInfo.shapes.push(shapeInfo);
    }

    result.slides.push(slideInfo);
  }
  return result;
}

// ============================================================================
// DOCX Verification
// ============================================================================

function verifyDocx(zip) {
  const result = { type: 'docx', file: absPath, paragraphs: [], paragraphCount: 0, error: null, inlineShapeCount: 0 };

  const themeFile = zip.listFiles().find(f => /^word\/theme\/theme\d+\.xml$/.test(f));
  if (themeFile) result.theme = extractTheme(zip, themeFile);

  result.headersFooters = extractHeadersFooters(zip);

  const docXml = zip.getFile('word/document.xml')?.toString('utf8');
  if (!docXml) { result.error = 'Missing word/document.xml'; return result; }

  const drawingCount = (docXml.match(/<w:drawing>/g) || []).length;
  const blipCount = (docXml.match(/<a:blip\b/g) || []).length;
  result.inlineShapeCount = Math.max(drawingCount, blipCount);

  // Build styleId → name map from styles.xml
  const stylesXml = zip.getFile('word/styles.xml')?.toString('utf8') || '';
  const styleMap = new Map();
  const styleRegex = /<w:style\b[^>]*w:styleId="([^"]*)"[^>]*>([\s\S]*?)<\/w:style>/g;
  let sm;
  while ((sm = styleRegex.exec(stylesXml)) !== null) {
    const nameMatch = sm[2].match(/<w:name\s+w:val="([^"]*)"/);
    if (nameMatch) styleMap.set(sm[1], unesc(nameMatch[1]));
  }
  function styleName(id) {
    if (!id) return 'Normal';
    const mapped = styleMap.get(id);
    if (mapped) return mapped;
    if (id === 'Heading1') return 'Heading 1';
    if (id === 'Heading2') return 'Heading 2';
    if (id === 'Heading3') return 'Heading 3';
    if (id === 'Title') return 'Title';
    if (id === 'ListParagraph') return 'List Paragraph';
    return id;
  }

  // Parse paragraphs
  const paraRegex = /<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g;
  let pm;
  while ((pm = paraRegex.exec(docXml)) !== null) {
    const paraXml = pm[1];

    const pStyleMatch = paraXml.match(/<w:pStyle\s+w:val="([^"]*)"/);
    const style = styleName(pStyleMatch ? pStyleMatch[1] : null);

    const jcMatch = paraXml.match(/<w:jc\s+w:val="([^"]*)"/);
    const alignment = jcMatch ? jcMatch[1] : 'left';

    const runs = [];
    const runRegex = /<w:r\b[^>]*>([\s\S]*?)<\/w:r>/g;
    let rm;
    while ((rm = runRegex.exec(paraXml)) !== null) {
      const runXml = rm[1];
      const textMatch = runXml.match(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/);
      if (!textMatch) continue;
      const text = unesc(textMatch[1]);

      const rprMatch = runXml.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
      const rprXml = rprMatch ? rprMatch[1] : '';

      const bold = /<w:b\b/.test(rprXml) && !/<w:b\s+w:val="(0|false)"/.test(rprXml);
      const italic = /<w:i\b/.test(rprXml) && !/<w:i\s+w:val="(0|false)"/.test(rprXml);

      let fontName = 'Calibri';
      const fontMatch = rprXml.match(/<w:rFonts[^>]*w:ascii="([^"]*)"/);
      if (fontMatch) fontName = unesc(fontMatch[1]);

      let fontSize = 11;
      const szMatch = rprXml.match(/<w:sz\s+w:val="(\d+)"/);
      if (szMatch) fontSize = parseInt(szMatch[1]) / 2;

      runs.push({ text, bold, italic, fontName, fontSize });
    }

    result.paragraphs.push({ text: runs.map(r => r.text).join(''), style, alignment, runs });
  }

  result.paragraphCount = result.paragraphs.length;
  return result;
}

// ============================================================================
// Main
// ============================================================================

try {
  const buf = readFileSync(absPath);
  const zip = new ZipReader(buf);
  const result = ext === '.pptx' ? verifyPptx(zip) : verifyDocx(zip);
  console.log(JSON.stringify(result, null, 2));
} catch (e) {
  console.error(`ERROR: ${e.message}`);
  process.exit(3);
}
