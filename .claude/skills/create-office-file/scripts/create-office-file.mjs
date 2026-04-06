#!/usr/bin/env node
// create-office-file.mjs — Zero-dependency markdown → PPTX/DOCX converter
// Built from first principles: ZIP + CRC-32 + OOXML XML generation
// Uses only Node.js built-in modules: fs, path, zlib
'use strict';

import { writeFileSync, readFileSync, existsSync } from 'fs';
import { resolve, extname, dirname, join, basename } from 'path';
import { deflateRawSync, inflateRawSync } from 'zlib';

// ============================================================================
// CRC-32 (lookup table, polynomial 0xEDB88320)
// ============================================================================

const crcTable = new Uint32Array(256);
for (let n = 0; n < 256; n++) {
  let c = n;
  for (let k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
  crcTable[n] = c >>> 0;
}

function crc32(buf) {
  if (typeof buf === 'string') buf = Buffer.from(buf, 'utf8');
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < buf.length; i++) crc = crcTable[(crc ^ buf[i]) & 0xFF] ^ (crc >>> 8);
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

// ============================================================================
// ZIP Writer (from first principles using Buffer + zlib)
// ============================================================================

class ZipWriter {
  constructor() { this.entries = []; this.offset = 0; this.buffers = []; }

  addFile(name, data) {
    const buf = typeof data === 'string' ? Buffer.from(data, 'utf8') : data;
    const nameBuf = Buffer.from(name, 'utf8');
    const crc = crc32(buf);
    const compressed = deflateRawSync(buf, { level: 6 });
    const useDeflate = compressed.length < buf.length;
    const fileData = useDeflate ? compressed : buf;
    const method = useDeflate ? 8 : 0;

    // Local file header (30 bytes + name)
    const local = Buffer.alloc(30 + nameBuf.length);
    local.writeUInt32LE(0x04034B50, 0);   // signature
    local.writeUInt16LE(20, 4);           // version needed
    local.writeUInt16LE(0, 6);            // flags
    local.writeUInt16LE(method, 8);       // compression
    local.writeUInt16LE(0, 10);           // mod time
    local.writeUInt16LE(0x0021, 12);      // mod date (1980-01-01)
    local.writeUInt32LE(crc, 14);         // crc-32
    local.writeUInt32LE(fileData.length, 18); // compressed size
    local.writeUInt32LE(buf.length, 22);  // uncompressed size
    local.writeUInt16LE(nameBuf.length, 26); // file name length
    local.writeUInt16LE(0, 28);           // extra field length
    nameBuf.copy(local, 30);

    this.entries.push({ name: nameBuf, crc, compressedSize: fileData.length,
      uncompressedSize: buf.length, method, offset: this.offset });
    this.buffers.push(local, fileData);
    this.offset += local.length + fileData.length;
  }

  toBuffer() {
    const cdBufs = [];
    let cdSize = 0;
    for (const e of this.entries) {
      const cd = Buffer.alloc(46 + e.name.length);
      cd.writeUInt32LE(0x02014B50, 0);    // central dir signature
      cd.writeUInt16LE(20, 4);            // version made by
      cd.writeUInt16LE(20, 6);            // version needed
      cd.writeUInt16LE(0, 8);             // flags
      cd.writeUInt16LE(e.method, 10);     // compression
      cd.writeUInt16LE(0, 12);            // mod time
      cd.writeUInt16LE(0x0021, 14);       // mod date
      cd.writeUInt32LE(e.crc, 16);        // crc-32
      cd.writeUInt32LE(e.compressedSize, 20);
      cd.writeUInt32LE(e.uncompressedSize, 24);
      cd.writeUInt16LE(e.name.length, 28); // file name length
      cd.writeUInt16LE(0, 30);            // extra field length
      cd.writeUInt16LE(0, 32);            // comment length
      cd.writeUInt16LE(0, 34);            // disk number
      cd.writeUInt16LE(0, 36);            // internal attrs
      cd.writeUInt32LE(0, 38);            // external attrs
      cd.writeUInt32LE(e.offset, 42);     // local header offset
      e.name.copy(cd, 46);
      cdBufs.push(cd);
      cdSize += cd.length;
    }

    // End of central directory record (22 bytes)
    const eocd = Buffer.alloc(22);
    eocd.writeUInt32LE(0x06054B50, 0);    // EOCD signature
    eocd.writeUInt16LE(0, 4);             // disk number
    eocd.writeUInt16LE(0, 6);             // CD start disk
    eocd.writeUInt16LE(this.entries.length, 8);  // entries on this disk
    eocd.writeUInt16LE(this.entries.length, 10);  // total entries
    eocd.writeUInt32LE(cdSize, 12);       // CD size
    eocd.writeUInt32LE(this.offset, 16);  // CD offset
    eocd.writeUInt16LE(0, 20);            // comment length

    return Buffer.concat([...this.buffers, ...cdBufs, eocd]);
  }
}

// ============================================================================
// ZIP Reader (for template extraction)
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
// XML Utilities
// ============================================================================

function esc(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

// ============================================================================
// Image Utilities (PNG/JPEG header parsing, sizing)
// ============================================================================

const IMAGE_CONTENT_TYPES = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.gif': 'image/gif',
};

function readImageDimensions(buf) {
  // PNG: signature (8 bytes) + IHDR chunk (length=4, type=4, data=13)
  // Width at offset 16, height at offset 20 (big-endian uint32)
  if (buf.length >= 24 && buf[0] === 0x89 && buf[1] === 0x50 && buf[2] === 0x4E && buf[3] === 0x47) {
    return { width: buf.readUInt32BE(16), height: buf.readUInt32BE(20), format: 'png' };
  }
  // JPEG: scan for SOF0 (0xFFC0) or SOF2 (0xFFC2) marker
  // Height at marker+5, width at marker+7 (big-endian uint16)
  if (buf.length >= 2 && buf[0] === 0xFF && buf[1] === 0xD8) {
    let pos = 2;
    while (pos < buf.length - 9) {
      if (buf[pos] !== 0xFF) { pos++; continue; }
      const marker = buf[pos + 1];
      if (marker === 0xC0 || marker === 0xC2) {
        return { width: buf.readUInt16BE(pos + 7), height: buf.readUInt16BE(pos + 5), format: 'jpeg' };
      }
      // Skip variable-length segment
      if (marker >= 0xC0 && marker !== 0xFF) {
        const segLen = buf.readUInt16BE(pos + 2);
        pos += 2 + segLen;
      } else {
        pos++;
      }
    }
    return { width: 0, height: 0, format: 'jpeg' };
  }
  return null;
}

// Convert pixel dimensions to EMU, preserving aspect ratio within max bounds
const PX_TO_EMU = 914400 / 96; // 96 DPI assumption

function fitImageEMU(widthPx, heightPx, maxWidthEMU, maxHeightEMU) {
  let w = widthPx * PX_TO_EMU;
  let h = heightPx * PX_TO_EMU;
  if (w > maxWidthEMU) { const scale = maxWidthEMU / w; w = maxWidthEMU; h *= scale; }
  if (h > maxHeightEMU) { const scale = maxHeightEMU / h; h = maxHeightEMU; w *= scale; }
  return { cx: Math.round(w), cy: Math.round(h) };
}

// Default image size when dimensions can't be read: 4" × 3"
const DEFAULT_IMG_W = 4 * 914400;
const DEFAULT_IMG_H = 3 * 914400;

function resolveImage(imgPath, basePath) {
  const resolved = basePath ? resolve(dirname(basePath), imgPath) : resolve(imgPath);
  if (!existsSync(resolved)) return null;
  try {
    const buf = readFileSync(resolved);
    const ext = extname(resolved).toLowerCase();
    const contentType = IMAGE_CONTENT_TYPES[ext];
    if (!contentType) return null;
    const dims = readImageDimensions(buf);
    return { buffer: buf, ext, contentType, width: dims?.width || 0, height: dims?.height || 0 };
  } catch { return null; }
}

// ============================================================================
// Theme Model (centralized colors, fonts — populated from defaults or template)
// ============================================================================

function defaultTheme() {
  return {
    colors: {
      dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6',
      accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5',
      accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47',
      hlink: '0563C1', folHlink: '954F72',
      heading: '2F5496', codeText: '2B2B2B', codeBg: 'F5F5F5',
      inlineCodeBg: 'E8E8E8', tableHeaderFill: 'D9E2F3',
    },
    fonts: { major: 'Calibri Light', minor: 'Calibri', code: 'Consolas' },
  };
}

function buildThemeXml(theme) {
  const c = theme.colors;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="${c.dk1}"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="${c.lt1}"/></a:lt1>
      <a:dk2><a:srgbClr val="${c.dk2}"/></a:dk2>
      <a:lt2><a:srgbClr val="${c.lt2}"/></a:lt2>
      <a:accent1><a:srgbClr val="${c.accent1}"/></a:accent1>
      <a:accent2><a:srgbClr val="${c.accent2}"/></a:accent2>
      <a:accent3><a:srgbClr val="${c.accent3}"/></a:accent3>
      <a:accent4><a:srgbClr val="${c.accent4}"/></a:accent4>
      <a:accent5><a:srgbClr val="${c.accent5}"/></a:accent5>
      <a:accent6><a:srgbClr val="${c.accent6}"/></a:accent6>
      <a:hlink><a:srgbClr val="${c.hlink}"/></a:hlink>
      <a:folHlink><a:srgbClr val="${c.folHlink}"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont><a:latin typeface="${theme.fonts.major}"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
      <a:minorFont><a:latin typeface="${theme.fonts.minor}"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
      <a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
      <a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
      <a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;
}

function buildDocxStylesXml(theme) {
  const c = theme.colors;
  const f = theme.fonts;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="${f.minor}" w:hAnsi="${f.minor}" w:cs="${f.minor}"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr><w:spacing w:after="160" w:line="259" w:lineRule="auto"/></w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:styleId="Normal" w:default="1">
    <w:name w:val="Normal"/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="48"/><w:szCs w:val="48"/><w:color w:val="${c.heading}"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="240" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="36"/><w:szCs w:val="36"/><w:color w:val="${c.heading}"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="200" w:after="80"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/><w:color w:val="${c.heading}"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading4">
    <w:name w:val="heading 4"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="40"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading5">
    <w:name w:val="heading 5"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="120" w:after="40"/></w:pPr>
    <w:rPr><w:b/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading6">
    <w:name w:val="heading 6"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:pPr><w:keepNext/><w:keepLines/><w:spacing w:before="120" w:after="40"/></w:pPr>
    <w:rPr><w:b/><w:i/><w:sz w:val="22"/><w:szCs w:val="22"/></w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="CodeBlock">
    <w:name w:val="Code Block"/>
    <w:basedOn w:val="Normal"/>
    <w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/><w:shd w:val="clear" w:color="auto" w:fill="${c.codeBg}"/></w:pPr>
    <w:rPr><w:rFonts w:ascii="${f.code}" w:hAnsi="${f.code}" w:cs="${f.code}"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>
  </w:style>
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:rPr><w:color w:val="${c.hlink}"/><w:u w:val="single"/></w:rPr>
  </w:style>
</w:styles>`;
}

function parseThemeColors(xml) {
  const colors = {};
  for (const name of ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink']) {
    let m = xml.match(new RegExp(`<a:${name}>\\s*<a:srgbClr\\s+val="([A-Fa-f0-9]{6})"`, 's'));
    if (m) { colors[name] = m[1].toUpperCase(); continue; }
    m = xml.match(new RegExp(`<a:${name}>\\s*<a:sysClr[^>]*lastClr="([A-Fa-f0-9]{6})"`, 's'));
    if (m) colors[name] = m[1].toUpperCase();
  }
  return colors;
}

function parseThemeFonts(xml) {
  const fonts = {};
  const major = xml.match(/<a:majorFont>\s*<a:latin\s+typeface="([^"]+)"/s);
  if (major) fonts.major = major[1];
  const minor = xml.match(/<a:minorFont>\s*<a:latin\s+typeface="([^"]+)"/s);
  if (minor) fonts.minor = minor[1];
  return fonts;
}

function extractThemeFromTemplate(templatePath) {
  const buf = readFileSync(templatePath);
  const zip = new ZipReader(buf);
  let themeXml = null;
  for (const p of ['ppt/theme/theme1.xml', 'word/theme/theme1.xml']) {
    const data = zip.getFile(p);
    if (data) { themeXml = data.toString('utf8'); break; }
  }
  if (!themeXml) return {};
  const overrides = {};
  const colors = parseThemeColors(themeXml);
  if (Object.keys(colors).length > 0) overrides.colors = colors;
  const fonts = parseThemeFonts(themeXml);
  if (Object.keys(fonts).length > 0) overrides.fonts = fonts;
  return overrides;
}

function mergeTheme(base, overrides) {
  return {
    colors: { ...base.colors, ...(overrides.colors || {}) },
    fonts: { ...base.fonts, ...(overrides.fonts || {}) },
  };
}

// ============================================================================
// Markdown Parser (constrained subset → AST)
// ============================================================================

/**
 * AST node types:
 *   { type: 'heading', level: 1-6, children: [inline...] }
 *   { type: 'paragraph', children: [inline...] }
 *   { type: 'bullet_list', items: [{ level: 0, children: [inline...] }] }
 *   { type: 'ordered_list', items: [{ level: 0, children: [inline...] }] }
 *   { type: 'code_block', lang: string, text: string }
 *   { type: 'hr' }
 *   { type: 'table', headers: [string...], rows: [[string...]...] }
 *
 * Inline types:
 *   { type: 'text', text: string }
 *   { type: 'bold', text: string }
 *   { type: 'italic', text: string }
 *   { type: 'bold_italic', text: string }
 *   { type: 'code', text: string }
 *   { type: 'link', text: string, url: string }
 *   { type: 'image', alt: string, src: string }
 */

function parseInline(text) {
  const tokens = [];
  const re = /(\*\*\*(.+?)\*\*\*|\*\*(.+?)\*\*|\*(.+?)\*|`(.+?)`|!\[([^\]]*)\]\(([^)]+)\)|\[([^\]]+)\]\(([^)]+)\))/g;
  let last = 0;
  let m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) tokens.push({ type: 'text', text: text.slice(last, m.index) });
    if (m[2]) tokens.push({ type: 'bold_italic', text: m[2] });
    else if (m[3]) tokens.push({ type: 'bold', text: m[3] });
    else if (m[4]) tokens.push({ type: 'italic', text: m[4] });
    else if (m[5]) tokens.push({ type: 'code', text: m[5] });
    else if (m[6] !== undefined) tokens.push({ type: 'image', alt: m[6], src: m[7] });
    else if (m[8]) tokens.push({ type: 'link', text: m[8], url: m[9] });
    last = m.index + m[0].length;
  }
  if (last < text.length) tokens.push({ type: 'text', text: text.slice(last) });
  return tokens;
}

function parseMarkdown(md) {
  const lines = md.replace(/\r\n/g, '\n').split('\n');
  const ast = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    // Blank line
    if (line.trim() === '') { i++; continue; }

    // Fenced code block
    const codeFence = line.match(/^```(\w*)/);
    if (codeFence) {
      const lang = codeFence[1] || '';
      const codeLines = [];
      i++;
      while (i < lines.length && !lines[i].startsWith('```')) {
        codeLines.push(lines[i]);
        i++;
      }
      if (i < lines.length) i++; // skip closing ```
      ast.push({ type: 'code_block', lang, text: codeLines.join('\n') });
      continue;
    }

    // Horizontal rule
    if (/^(-{3,}|\*{3,}|_{3,})$/.test(line.trim())) {
      ast.push({ type: 'hr' });
      i++;
      continue;
    }

    // Heading
    const headingMatch = line.match(/^(#{1,6})\s+(.+)$/);
    if (headingMatch) {
      ast.push({ type: 'heading', level: headingMatch[1].length, children: parseInline(headingMatch[2]) });
      i++;
      continue;
    }

    // Table (look for | in current line and separator line next)
    if (line.includes('|') && i + 1 < lines.length && /^\|?[\s:]*-+[\s:]*/.test(lines[i + 1])) {
      const parseRow = r => r.replace(/^\|/, '').replace(/\|$/, '').split('|').map(c => c.trim());
      const headers = parseRow(line);
      i += 2; // skip header + separator
      const rows = [];
      while (i < lines.length && lines[i].includes('|') && lines[i].trim() !== '') {
        rows.push(parseRow(lines[i]));
        i++;
      }
      ast.push({ type: 'table', headers, rows });
      continue;
    }

    // Bullet list
    if (/^[\s]*[-*+]\s/.test(line)) {
      const items = [];
      while (i < lines.length && /^[\s]*[-*+]\s/.test(lines[i])) {
        const raw = lines[i];
        const indent = raw.match(/^(\s*)/)[1].replace(/\t/g, '  ').length;
        const level = Math.min(Math.floor(indent / 2), 2);
        items.push({ level, children: parseInline(raw.replace(/^[\s]*[-*+]\s/, '')) });
        i++;
      }
      ast.push({ type: 'bullet_list', items });
      continue;
    }

    // Ordered list
    if (/^[\s]*\d+[.)]\s/.test(line)) {
      const items = [];
      while (i < lines.length && /^[\s]*\d+[.)]\s/.test(lines[i])) {
        const raw = lines[i];
        const indent = raw.match(/^(\s*)/)[1].replace(/\t/g, '  ').length;
        const level = Math.min(Math.floor(indent / 2), 2);
        items.push({ level, children: parseInline(raw.replace(/^[\s]*\d+[.)]\s/, '')) });
        i++;
      }
      ast.push({ type: 'ordered_list', items });
      continue;
    }

    // Block-level image (standalone line: ![alt](path))
    const imgMatch = line.match(/^!\[([^\]]*)\]\(([^)]+)\)\s*$/);
    if (imgMatch) {
      ast.push({ type: 'image', alt: imgMatch[1], src: imgMatch[2] });
      i++;
      continue;
    }

    // Paragraph (accumulate consecutive non-empty, non-special lines)
    const paraLines = [];
    while (i < lines.length && lines[i].trim() !== '' &&
           !lines[i].match(/^#{1,6}\s/) && !lines[i].match(/^```/) &&
           !lines[i].match(/^(-{3,}|\*{3,}|_{3,})$/) &&
           !(/^[\s]*[-*+]\s/.test(lines[i])) &&
           !(/^[\s]*\d+[.)]\s/.test(lines[i]))) {
      paraLines.push(lines[i]);
      i++;
    }
    if (paraLines.length > 0) {
      ast.push({ type: 'paragraph', children: parseInline(paraLines.join(' ')) });
    }
  }

  return ast;
}

// ============================================================================
// PPTX Generator (PresentationML with hard-coded templates)
// ============================================================================

// Slide dimensions: 16:9 widescreen
const SLIDE_W = 12192000; // 13.333" in EMU
const SLIDE_H = 6858000;  // 7.5" in EMU

function pptxContentTypes(slideCount, imageExts) {
  let overrides = '';
  for (let i = 1; i <= slideCount; i++) {
    overrides += `  <Override PartName="/ppt/slides/slide${i}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>\n`;
  }
  let imgDefaults = '';
  for (const ext of imageExts) {
    const ct = IMAGE_CONTENT_TYPES['.' + ext];
    if (ct) imgDefaults += `  <Default Extension="${ext}" ContentType="${ct}"/>\n`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
${imgDefaults}  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
${overrides}</Types>`;
}

const PPTX_ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

function pptxPresentationRels(slideCount) {
  let rels = `  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`;
  for (let i = 1; i <= slideCount; i++) {
    rels += `\n  <Relationship Id="rId${i + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i}.xml"/>`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels}
</Relationships>`;
}

function pptxPresentation(slideCount) {
  let sldIdLst = '';
  for (let i = 1; i <= slideCount; i++) {
    sldIdLst += `    <p:sldId id="${255 + i}" r:id="rId${i + 2}"/>\n`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId1"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
${sldIdLst}  </p:sldIdLst>
  <p:sldSz cx="${SLIDE_W}" cy="${SLIDE_H}"/>
  <p:notesSz cx="${SLIDE_H}" cy="${SLIDE_W}"/>
</p:presentation>`;
}

// Theme XML and slide master/layout are now generated via buildThemeXml(theme)

// Minimal slide master
const PPTX_SLIDE_MASTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst>
    <p:sldLayoutId id="2147483649" r:id="rId1"/>
  </p:sldLayoutIdLst>
</p:sldMaster>`;

const PPTX_SLIDE_MASTER_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;

// Minimal slide layout
const PPTX_SLIDE_LAYOUT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:sldLayout>`;

const PPTX_SLIDE_LAYOUT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

function pptxSlideRels(links, images) {
  let extra = '';
  links.forEach((url, idx) => {
    extra += `\n  <Relationship Id="rId${idx + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${esc(url)}" TargetMode="External"/>`;
  });
  const imgBase = links.length + 2;
  images.forEach((img, idx) => {
    extra += `\n  <Relationship Id="rId${imgBase + idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${img.name}"/>`;
  });
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>${extra}
</Relationships>`;
}

// Convert inline AST nodes to DrawingML runs
function inlineToDrawingML(children, baseSz, links, theme) {
  const codeFont = theme.fonts.code;
  return children.map(c => {
    const t = esc(c.text || '');
    switch (c.type) {
      case 'bold': return `<a:r><a:rPr lang="en-US" sz="${baseSz}" b="1" dirty="0"/><a:t>${t}</a:t></a:r>`;
      case 'italic': return `<a:r><a:rPr lang="en-US" sz="${baseSz}" i="1" dirty="0"/><a:t>${t}</a:t></a:r>`;
      case 'bold_italic': return `<a:r><a:rPr lang="en-US" sz="${baseSz}" b="1" i="1" dirty="0"/><a:t>${t}</a:t></a:r>`;
      case 'code': return `<a:r><a:rPr lang="en-US" sz="${baseSz}" dirty="0"><a:latin typeface="${codeFont}"/><a:cs typeface="${codeFont}"/></a:rPr><a:t>${t}</a:t></a:r>`;
      case 'link': {
        links.push(c.url);
        const rId = `rId${links.length + 1}`;
        return `<a:r><a:rPr lang="en-US" sz="${baseSz}" dirty="0"><a:hlinkClick r:id="${rId}"/></a:rPr><a:t>${esc(c.text)}</a:t></a:r>`;
      }
      default: return `<a:r><a:rPr lang="en-US" sz="${baseSz}" dirty="0"/><a:t>${t}</a:t></a:r>`;
    }
  }).join('');
}

function buildTitleSlide(heading, shapeId, theme) {
  const links = [];
  const runs = inlineToDrawingML(heading.children, 4400, links, theme);
  const xml = `<p:sp>
  <p:nvSpPr><p:cNvPr id="${shapeId}" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="457200" y="2286000"/><a:ext cx="11277600" cy="2286000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="ctr" anchorCtr="1"/>
    <a:lstStyle/>
    <a:p><a:pPr algn="ctr"/>${runs}</a:p>
  </p:txBody>
</p:sp>`;
  return { xml, links, images: [] };
}

function buildSlideTitleShape(title, shapeId, links, theme) {
  if (!title) return { xml: '', nextId: shapeId };
  const runs = inlineToDrawingML(title.children, 2800, links, theme);
  return { xml: `<p:sp>
  <p:nvSpPr><p:cNvPr id="${shapeId}" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="457200" y="274638"/><a:ext cx="11277600" cy="1143000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr anchor="b"/>
    <a:lstStyle/>
    <a:p>${runs}</a:p>
  </p:txBody>
</p:sp>`, nextId: shapeId + 1 };
}

function buildContentSlide(title, bodyNodes, startShapeId, inputPath, theme) {
  const links = [];
  const images = []; // { name, buffer, rId }
  const { xml: titleShape, nextId } = buildSlideTitleShape(title, startShapeId, links, theme);
  let shapeId = nextId;

  // Max image bounds within slide content area
  const IMG_MAX_W = 10 * 914400; // 10 inches
  const IMG_MAX_H = 4 * 914400;  // 4 inches (leave room for title + text)

  // Separate images from text content
  const textNodes = bodyNodes.filter(n => n.type !== 'image');
  const imageNodes = bodyNodes.filter(n => n.type === 'image');

  // Body text content
  let bodyParas = '';
  for (const node of textNodes) {
    switch (node.type) {
      case 'paragraph':
        bodyParas += `<a:p>${inlineToDrawingML(node.children, 1800, links, theme)}</a:p>`;
        break;
      case 'bullet_list': {
        const buChars = ['\u2022', '\u2013', '\u203A']; // •, –, ›
        for (const item of node.items) {
          const lvl = item.level || 0;
          const marL = 342900 + lvl * 457200;
          bodyParas += `<a:p><a:pPr marL="${marL}" indent="-342900"><a:buChar char="${buChars[lvl] || buChars[0]}"/></a:pPr>${inlineToDrawingML(item.children, 1800, links, theme)}</a:p>`;
        }
        break;
      }
      case 'ordered_list': {
        for (const item of node.items) {
          const lvl = item.level || 0;
          const marL = 342900 + lvl * 457200;
          bodyParas += `<a:p><a:pPr marL="${marL}" indent="-342900"><a:buAutoNum type="arabicPeriod"/></a:pPr>${inlineToDrawingML(item.children, 1800, links, theme)}</a:p>`;
        }
        break;
      }
      case 'code_block':
        for (const codeLine of node.text.split('\n')) {
          bodyParas += `<a:p><a:r><a:rPr lang="en-US" sz="1400" dirty="0"><a:latin typeface="${theme.fonts.code}"/><a:cs typeface="${theme.fonts.code}"/><a:solidFill><a:srgbClr val="${theme.colors.codeText}"/></a:solidFill></a:rPr><a:t>${esc(codeLine)}</a:t></a:r></a:p>`;
        }
        break;
      case 'heading':
        bodyParas += `<a:p>${inlineToDrawingML(node.children, 2400, links, theme)}</a:p>`;
        break;
      default:
        break;
    }
  }

  // If no body paragraphs, add empty paragraph to keep slide valid
  if (!bodyParas) bodyParas = '<a:p><a:endParaRPr lang="en-US"/></a:p>';

  // Adjust body shape height if images are present
  const hasImages = imageNodes.length > 0;
  const bodyH = hasImages ? 2200000 : 4525963;
  const bodyShape = `<p:sp>
  <p:nvSpPr><p:cNvPr id="${shapeId++}" name="Content"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="11277600" cy="${bodyH}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0"/>
    <a:lstStyle/>
    ${bodyParas}
  </p:txBody>
</p:sp>`;

  // Build p:pic shapes for images
  let imgShapes = '';
  for (const imgNode of imageNodes) {
    const imgData = resolveImage(imgNode.src, inputPath);
    if (!imgData) {
      // Fallback: render alt text as a text paragraph
      bodyParas += `<a:p><a:r><a:rPr lang="en-US" sz="1400" i="1" dirty="0"/><a:t>[Image: ${esc(imgNode.alt || imgNode.src)}]</a:t></a:r></a:p>`;
      continue;
    }
    const imgName = `image${images.length + 1}${imgData.ext}`;
    const rId = `rId${links.length + 2 + images.length}`;
    images.push({ name: imgName, buffer: imgData.buffer, rId });

    const w = imgData.width || Math.round(DEFAULT_IMG_W / PX_TO_EMU);
    const h = imgData.height || Math.round(DEFAULT_IMG_H / PX_TO_EMU);
    const { cx, cy } = fitImageEMU(w, h, IMG_MAX_W, IMG_MAX_H);
    const imgX = Math.round(457200 + (11277600 - cx) / 2); // center horizontally
    const imgY = hasImages ? 3900000 : 1600200;

    imgShapes += `\n<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="${shapeId++}" name="${esc(imgNode.alt || imgName)}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="${rId}"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm><a:off x="${imgX}" y="${imgY}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`;
  }

  return { xml: titleShape + '\n' + bodyShape + imgShapes, links, images };
}

function buildTableSlide(title, table, startShapeId, theme) {
  const links = [];
  const { xml: titleShape, nextId } = buildSlideTitleShape(title, startShapeId, links, theme);
  let shapeId = nextId;

  const numCols = table.headers.length;
  const colW = Math.floor(11277600 / numCols);
  const gridCols = table.headers.map(() => `<a:gridCol w="${colW}"/>`).join('');

  const mkCell = (text, bold) => {
    const rPr = bold ? '<a:rPr lang="en-US" sz="1400" b="1" dirty="0"/>' : '<a:rPr lang="en-US" sz="1400" dirty="0"/>';
    return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r>${rPr}<a:t>${esc(text)}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`;
  };

  const headerRow = `<a:tr h="370840">${table.headers.map(h => mkCell(h, true)).join('')}</a:tr>`;
  const dataRows = table.rows.map(row =>
    `<a:tr h="370840">${row.map(cell => mkCell(cell, false)).join('')}</a:tr>`
  ).join('');

  const tableFrame = `<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="${shapeId++}" name="Table"/>
    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm><a:off x="457200" y="1600200"/><a:ext cx="11277600" cy="3000000"/></p:xfrm>
  <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblPr firstRow="1" bandRow="1"/>
        <a:tblGrid>${gridCols}</a:tblGrid>
        ${headerRow}
        ${dataRows}
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>`;

  return { xml: titleShape + '\n' + tableFrame, links, images: [] };
}

function wrapSlide(innerXml) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr/>
      ${innerXml}
    </p:spTree>
  </p:cSld>
</p:sld>`;
}

function astToSlides(ast, inputPath, theme) {
  const slides = []; // [{ xml, links, images }]
  let currentTitle = null;
  let currentBody = [];

  function flushSlide() {
    if (!currentTitle && currentBody.length === 0) return;
    // Check if body has a table as the primary content
    const tableNode = currentBody.find(n => n.type === 'table');
    const nonTableBody = currentBody.filter(n => n.type !== 'table');
    let result;
    if (tableNode && nonTableBody.length === 0) {
      result = buildTableSlide(currentTitle, tableNode, 2, theme);
    } else if (currentTitle && currentTitle.level === 1 && currentBody.length === 0) {
      result = buildTitleSlide(currentTitle, 2, theme);
    } else {
      result = buildContentSlide(currentTitle, currentBody, 2, inputPath, theme);
    }
    slides.push({ xml: wrapSlide(result.xml), links: result.links, images: result.images || [] });
    currentTitle = null;
    currentBody = [];
  }

  for (const node of ast) {
    if (node.type === 'hr') {
      flushSlide();
    } else if (node.type === 'heading' && node.level <= 2) {
      flushSlide();
      currentTitle = node;
    } else {
      currentBody.push(node);
    }
  }
  flushSlide();

  // If no slides were created, add one empty slide
  if (slides.length === 0) {
    slides.push({
      xml: wrapSlide('<p:sp><p:nvSpPr><p:cNvPr id="2" name="Empty"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="11277600" cy="4525963"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'),
      links: [],
      images: []
    });
  }

  return slides;
}

function generatePptx(ast, inputPath, theme) {
  const slides = astToSlides(ast, inputPath, theme);
  const zip = new ZipWriter();

  // Collect all unique image extensions for content types
  const imageExts = new Set();
  const allImages = [];
  slides.forEach(slide => {
    for (const img of slide.images) {
      imageExts.add(img.name.split('.').pop());
      allImages.push(img);
    }
  });

  zip.addFile('[Content_Types].xml', pptxContentTypes(slides.length, imageExts));
  zip.addFile('_rels/.rels', PPTX_ROOT_RELS);
  zip.addFile('ppt/presentation.xml', pptxPresentation(slides.length));
  zip.addFile('ppt/_rels/presentation.xml.rels', pptxPresentationRels(slides.length));
  zip.addFile('ppt/theme/theme1.xml', buildThemeXml(theme));
  zip.addFile('ppt/slideMasters/slideMaster1.xml', PPTX_SLIDE_MASTER);
  zip.addFile('ppt/slideMasters/_rels/slideMaster1.xml.rels', PPTX_SLIDE_MASTER_RELS);
  zip.addFile('ppt/slideLayouts/slideLayout1.xml', PPTX_SLIDE_LAYOUT);
  zip.addFile('ppt/slideLayouts/_rels/slideLayout1.xml.rels', PPTX_SLIDE_LAYOUT_RELS);

  slides.forEach((slide, i) => {
    zip.addFile(`ppt/slides/slide${i + 1}.xml`, slide.xml);
    zip.addFile(`ppt/slides/_rels/slide${i + 1}.xml.rels`, pptxSlideRels(slide.links, slide.images));
    for (const img of slide.images) {
      zip.addFile(`ppt/media/${img.name}`, img.buffer);
    }
  });

  return zip.toBuffer();
}

// ============================================================================
// DOCX Generator (WordprocessingML with hard-coded templates)
// ============================================================================

function docxContentTypes(imageExts) {
  let imgDefaults = '';
  for (const ext of imageExts) {
    const ct = IMAGE_CONTENT_TYPES['.' + ext];
    if (ct) imgDefaults += `  <Default Extension="${ext}" ContentType="${ct}"/>\n`;
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
${imgDefaults}  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;
}

const DOCX_ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

function docxDocumentRels(hyperlinks, images) {
  let rels = `  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>`;
  hyperlinks.forEach((url, idx) => {
    rels += `\n  <Relationship Id="rId${idx + 4}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${esc(url)}" TargetMode="External"/>`;
  });
  const imgBase = hyperlinks.length + 4;
  images.forEach((img, idx) => {
    rels += `\n  <Relationship Id="rId${imgBase + idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${img.name}"/>`;
  });
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
${rels}
</Relationships>`;
}

// DOCX styles are now generated via buildDocxStylesXml(theme)

const DOCX_NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="multilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="\u2022"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="o"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New" w:hint="default"/></w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="\u25A0"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings" w:hint="default"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="multilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:start w:val="1"/>
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="%3."/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;

function inlineToWordML(children, links, images, inputPath, theme) {
  return children.map(c => {
    const text = c.text || '';
    const needsSpace = text.startsWith(' ') || text.endsWith(' ');
    const t = needsSpace ? `<w:t xml:space="preserve">${esc(text)}</w:t>` : `<w:t>${esc(text)}</w:t>`;

    switch (c.type) {
      case 'bold': return `<w:r><w:rPr><w:b/></w:rPr>${t}</w:r>`;
      case 'italic': return `<w:r><w:rPr><w:i/></w:rPr>${t}</w:r>`;
      case 'bold_italic': return `<w:r><w:rPr><w:b/><w:i/></w:rPr>${t}</w:r>`;
      case 'code': return `<w:r><w:rPr><w:rFonts w:ascii="${theme.fonts.code}" w:hAnsi="${theme.fonts.code}" w:cs="${theme.fonts.code}"/><w:shd w:val="clear" w:color="auto" w:fill="${theme.colors.inlineCodeBg}"/></w:rPr>${t}</w:r>`;
      case 'link': {
        links.push(c.url);
        const rId = `rId${links.length + 3}`;
        return `<w:hyperlink r:id="${rId}"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>${esc(c.text)}</w:t></w:r></w:hyperlink>`;
      }
      case 'image': {
        if (!images || !inputPath) {
          return `<w:r><w:rPr><w:i/></w:rPr><w:t>[Image: ${esc(c.alt || c.src)}]</w:t></w:r>`;
        }
        const imgData = resolveImage(c.src, inputPath);
        if (!imgData) {
          return `<w:r><w:rPr><w:i/></w:rPr><w:t>[Image: ${esc(c.alt || c.src)}]</w:t></w:r>`;
        }
        const imgName = `image${images.length + 1}${imgData.ext}`;
        const rId = `rId${links.length + 4 + images.length}`;
        images.push({ name: imgName, buffer: imgData.buffer });
        const w = imgData.width || Math.round(DEFAULT_IMG_W / PX_TO_EMU);
        const h = imgData.height || Math.round(DEFAULT_IMG_H / PX_TO_EMU);
        const { cx, cy } = fitImageEMU(w, h, 6.5 * 914400, 9 * 914400);
        return `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${images.length}" name="${esc(c.alt || imgName)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="${esc(imgName)}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
      }
      default: return `<w:r>${t}</w:r>`;
    }
  }).join('');
}

function astToDocxBody(ast, links, images, inputPath, theme) {
  let body = '';
  let docPrId = 1; // unique ID counter for wp:docPr

  // DOCX max image width: page width minus margins = 12240 - 2*1440 = 9360 twips = 6.5" = 5943600 EMU
  const DOCX_IMG_MAX_W = 6.5 * 914400;
  const DOCX_IMG_MAX_H = 9 * 914400;

  for (const node of ast) {
    switch (node.type) {
      case 'heading': {
        const level = Math.min(node.level, 6);
        body += `<w:p><w:pPr><w:pStyle w:val="Heading${level}"/></w:pPr>${inlineToWordML(node.children, links, images, inputPath, theme)}</w:p>`;
        break;
      }
      case 'paragraph':
        body += `<w:p>${inlineToWordML(node.children, links, images, inputPath, theme)}</w:p>`;
        break;
      case 'image': {
        const imgData = resolveImage(node.src, inputPath);
        if (!imgData) {
          body += `<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>[Image: ${esc(node.alt || node.src)}]</w:t></w:r></w:p>`;
          break;
        }
        const imgName = `image${images.length + 1}${imgData.ext}`;
        const rId = `rId${links.length + 4 + images.length}`;
        images.push({ name: imgName, buffer: imgData.buffer });

        const w = imgData.width || Math.round(DEFAULT_IMG_W / PX_TO_EMU);
        const h = imgData.height || Math.round(DEFAULT_IMG_H / PX_TO_EMU);
        const { cx, cy } = fitImageEMU(w, h, DOCX_IMG_MAX_W, DOCX_IMG_MAX_H);
        const dpId = docPrId++;

        body += `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${dpId}" name="${esc(node.alt || imgName)}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="0" name="${esc(imgName)}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
        break;
      }
      case 'bullet_list':
        for (const item of node.items) {
          const lvl = item.level || 0;
          body += `<w:p><w:pPr><w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="1"/></w:numPr></w:pPr>${inlineToWordML(item.children, links, images, inputPath, theme)}</w:p>`;
        }
        break;
      case 'ordered_list':
        for (const item of node.items) {
          const lvl = item.level || 0;
          body += `<w:p><w:pPr><w:numPr><w:ilvl w:val="${lvl}"/><w:numId w:val="2"/></w:numPr></w:pPr>${inlineToWordML(item.children, links, images, inputPath, theme)}</w:p>`;
        }
        break;
      case 'code_block':
        for (const line of node.text.split('\n')) {
          body += `<w:p><w:pPr><w:pStyle w:val="CodeBlock"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="${theme.fonts.code}" w:hAnsi="${theme.fonts.code}" w:cs="${theme.fonts.code}"/><w:sz w:val="20"/></w:rPr><w:t xml:space="preserve">${esc(line)}</w:t></w:r></w:p>`;
        }
        break;
      case 'hr':
        body += `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:pBdr></w:pPr></w:p>`;
        break;
      case 'table': {
        const numCols = node.headers.length;
        const colW = Math.floor(9360 / numCols);
        const gridCols = node.headers.map(() => `<w:gridCol w:w="${colW}"/>`).join('');
        const borders = `<w:tblBorders>
          <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
          <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        </w:tblBorders>`;

        const headerRow = `<w:tr>${node.headers.map(h =>
          `<w:tc><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="${theme.colors.tableHeaderFill}"/></w:tcPr><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>${esc(h)}</w:t></w:r></w:p></w:tc>`
        ).join('')}</w:tr>`;

        const dataRows = node.rows.map(row =>
          `<w:tr>${row.map(cell =>
            `<w:tc><w:p><w:r><w:t>${esc(cell)}</w:t></w:r></w:p></w:tc>`
          ).join('')}</w:tr>`
        ).join('');

        body += `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/>${borders}</w:tblPr><w:tblGrid>${gridCols}</w:tblGrid>${headerRow}${dataRows}</w:tbl>`;
        break;
      }
    }
  }

  return body;
}

function generateDocx(ast, inputPath, theme) {
  const links = [];
  const images = [];
  const bodyContent = astToDocxBody(ast, links, images, inputPath, theme);

  // Collect unique image extensions
  const imageExts = new Set();
  for (const img of images) imageExts.add(img.name.split('.').pop());

  const document = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
  <w:body>
    ${bodyContent}
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

  const zip = new ZipWriter();
  zip.addFile('[Content_Types].xml', docxContentTypes(imageExts));
  zip.addFile('_rels/.rels', DOCX_ROOT_RELS);
  zip.addFile('word/document.xml', document);
  zip.addFile('word/_rels/document.xml.rels', docxDocumentRels(links, images));
  zip.addFile('word/styles.xml', buildDocxStylesXml(theme));
  zip.addFile('word/numbering.xml', DOCX_NUMBERING);
  zip.addFile('word/theme/theme1.xml', buildThemeXml(theme));
  for (const img of images) {
    zip.addFile(`word/media/${img.name}`, img.buffer);
  }

  return zip.toBuffer();
}

// ============================================================================
// CLI Entry Point
// ============================================================================

function usage() {
  console.error(`Usage: node create-office-file.mjs [options]

Options:
  -i, --input <file>       Input markdown file (or reads stdin)
  -o, --output <file>      Output file path (required)
  -f, --format <fmt>       Output format: pptx or docx (auto-detected from -o extension)
  -t, --template <file>    Template .pptx or .docx file (extracts theme colors and fonts)
  -h, --help               Show this help

Examples:
  node create-office-file.mjs -i slides.md -o presentation.pptx
  node create-office-file.mjs -i report.md -o document.docx
  node create-office-file.mjs -i slides.md -o out.pptx --template corporate.pptx
  cat notes.md | node create-office-file.mjs -o notes.docx`);
  process.exit(1);
}

function main() {
  const args = process.argv.slice(2);
  let input = null, output = null, format = null, template = null;

  for (let i = 0; i < args.length; i++) {
    switch (args[i]) {
      case '-i': case '--input': input = args[++i]; break;
      case '-o': case '--output': output = args[++i]; break;
      case '-f': case '--format':
        format = args[++i];
        if (format !== 'pptx' && format !== 'docx') { console.error(`Error: unsupported format '${format}', use pptx or docx`); usage(); }
        break;
      case '-t': case '--template': template = args[++i]; break;
      case '-h': case '--help': usage(); break;
      default:
        if (!output && args[i].match(/\.(pptx|docx)$/i)) output = args[i];
        else if (!input && existsSync(args[i])) input = args[i];
    }
  }

  if (!output) { console.error('Error: output file required (-o)'); usage(); }

  // Auto-detect format from extension
  if (!format) {
    const ext = extname(output).toLowerCase();
    if (ext === '.pptx') format = 'pptx';
    else if (ext === '.docx') format = 'docx';
    else { console.error('Error: cannot detect format, use -f pptx or -f docx'); usage(); }
  }

  // Build theme (defaults + optional template overrides)
  let theme = defaultTheme();
  if (template) {
    const templatePath = resolve(template);
    if (!existsSync(templatePath)) { console.error(`Error: template file not found: ${template}`); process.exit(1); }
    try {
      const overrides = extractThemeFromTemplate(templatePath);
      theme = mergeTheme(theme, overrides);
    } catch (e) {
      console.error(`Error: failed to read template: ${e.message}`); process.exit(1);
    }
  }

  // Read input
  let md;
  let inputPath = null;
  if (input) {
    inputPath = resolve(input);
    md = readFileSync(inputPath, 'utf8');
  } else if (!process.stdin.isTTY) {
    md = readFileSync(0, 'utf8'); // stdin
  } else {
    console.error('Error: no input file or stdin data'); usage();
  }

  const ast = parseMarkdown(md);
  const buf = format === 'pptx' ? generatePptx(ast, inputPath, theme) : generateDocx(ast, inputPath, theme);
  writeFileSync(resolve(output), buf);
}

main();
