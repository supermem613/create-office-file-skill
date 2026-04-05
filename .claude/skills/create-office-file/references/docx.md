# DOCX OOXML Reference

> Progressive disclosure reference for Word document generation.
> Read this when working on DOCX-specific code in scripts/create-office-file.mjs.

## File Structure

A valid .docx is a ZIP archive containing:

```
[Content_Types].xml          # MIME types for all parts
_rels/.rels                  # Root relationships
word/document.xml            # Main body content
word/_rels/document.xml.rels # Links to styles, numbering, etc.
word/styles.xml              # Style definitions (Heading1, etc.)
word/numbering.xml           # List/numbering definitions
```

**DOCX is forgiving:** Minimum viable is just `[Content_Types].xml`, `_rels/.rels`,
and `word/document.xml`. But we include `styles.xml` and `numbering.xml` for proper
heading rendering and list support.

## Namespaces

| Prefix | URI | Used For |
|--------|-----|----------|
| `w:` | `http://schemas.openxmlformats.org/wordprocessingml/2006/main` | Everything |
| `r:` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationship IDs |

## Document Structure

```xml
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <!-- paragraphs go here -->
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>  <!-- US Letter in twips -->
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>
```

Page dimensions in **twips** (1/20 of a point, 1 inch = 1440 twips):
- US Letter: 12240 × 15840 twips (8.5 × 11 inches)
- A4: 11906 × 16838 twips

## Paragraph and Run Structure

### Plain paragraph
```xml
<w:p>
  <w:r>
    <w:t>Hello, World!</w:t>
  </w:r>
</w:p>
```

### Formatted text
```xml
<w:p>
  <w:r>
    <w:rPr>
      <w:b/>           <!-- bold -->
      <w:i/>           <!-- italic -->
      <w:u w:val="single"/>  <!-- underline -->
    </w:rPr>
    <w:t>Bold italic underlined</w:t>
  </w:r>
</w:p>
```

**Important:** `<w:t>` needs `xml:space="preserve"` if text has leading/trailing spaces:
```xml
<w:t xml:space="preserve"> text with spaces </w:t>
```

### Mixed formatting in one paragraph
```xml
<w:p>
  <w:r><w:t xml:space="preserve">Normal </w:t></w:r>
  <w:r>
    <w:rPr><w:b/></w:rPr>
    <w:t>bold</w:t>
  </w:r>
  <w:r><w:t xml:space="preserve"> normal</w:t></w:r>
</w:p>
```

## Headings

Headings use paragraph styles defined in `styles.xml`:
```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="Heading1"/>
  </w:pPr>
  <w:r>
    <w:t>Heading Level 1</w:t>
  </w:r>
</w:p>
```

| Markdown | Style ID | Typical Size |
|----------|----------|-------------|
| `#` | `Heading1` | 24pt bold |
| `##` | `Heading2` | 18pt bold |
| `###` | `Heading3` | 14pt bold |
| `####` | `Heading4` | 12pt bold |
| `#####` | `Heading5` | 11pt bold |
| `######` | `Heading6` | 11pt bold italic |

## Lists (Bullet and Numbered)

Lists use `<w:numPr>` in paragraph properties, referencing definitions in `numbering.xml`. Nesting is controlled by `w:ilvl` (0-based indent level).

### Bullet list item
```xml
<w:p>
  <w:pPr>
    <w:numPr>
      <w:ilvl w:val="0"/>    <!-- indent level: 0-based, max 8 -->
      <w:numId w:val="1"/>   <!-- references numbering.xml abstractNumId -->
    </w:numPr>
  </w:pPr>
  <w:r>
    <w:t>Bullet item</w:t>
  </w:r>
</w:p>
```

### Numbered list item
```xml
<w:p>
  <w:pPr>
    <w:numPr>
      <w:ilvl w:val="0"/>
      <w:numId w:val="2"/>   <!-- different numId for numbered -->
    </w:numPr>
  </w:pPr>
  <w:r>
    <w:t>Numbered item</w:t>
  </w:r>
</w:p>
```

### numbering.xml structure (multi-level)
```xml
<w:numbering xmlns:w="...">
  <!-- Bullet list definition — 3 levels -->
  <w:abstractNum w:abstractNumId="0">
    <w:multiLevelType w:val="multilevel"/>
    <w:lvl w:ilvl="0">
      <w:start w:val="1"/>
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="•"/>
      <w:lvlJc w:val="left"/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol"/></w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="o"/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:numFmt w:val="bullet"/>
      <w:lvlText w:val="■"/>
      <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>
      <w:rPr><w:rFonts w:ascii="Wingdings" w:hAnsi="Wingdings"/></w:rPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="1">
    <w:abstractNumId w:val="0"/>
  </w:num>

  <!-- Numbered list definition — 3 levels -->
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="multilevel"/>
    <w:lvl w:ilvl="0">
      <w:numFmt w:val="decimal"/>
      <w:lvlText w:val="%1."/>
      <w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="1">
      <w:numFmt w:val="lowerLetter"/>
      <w:lvlText w:val="%2."/>
      <w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr>
    </w:lvl>
    <w:lvl w:ilvl="2">
      <w:numFmt w:val="lowerRoman"/>
      <w:lvlText w:val="%3."/>
      <w:pPr><w:ind w:left="2160" w:hanging="360"/></w:pPr>
    </w:lvl>
  </w:abstractNum>
  <w:num w:numId="2">
    <w:abstractNumId w:val="1"/>
  </w:num>
</w:numbering>
```

## Code Blocks

Code blocks use a monospace font in run properties:
```xml
<w:p>
  <w:pPr>
    <w:pStyle w:val="CodeBlock"/>
    <w:shd w:val="clear" w:color="auto" w:fill="F5F5F5"/>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="Consolas" w:hAnsi="Consolas" w:cs="Consolas"/>
      <w:sz w:val="20"/>  <!-- 10pt (size in half-points) -->
    </w:rPr>
    <w:t xml:space="preserve">const x = 42;</w:t>
  </w:r>
</w:p>
```

**Note:** `w:sz` is in **half-points** (20 = 10pt, 24 = 12pt, 48 = 24pt).

## Hyperlinks

```xml
<w:hyperlink r:id="rId4">
  <w:r>
    <w:rPr>
      <w:rStyle w:val="Hyperlink"/>
      <w:color w:val="0563C1"/>
      <w:u w:val="single"/>
    </w:rPr>
    <w:t>Click here</w:t>
  </w:r>
</w:hyperlink>
```
Requires a relationship in `word/_rels/document.xml.rels`:
```xml
<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  Target="https://example.com" TargetMode="External"/>
```

## Tables

```xml
<w:tbl>
  <w:tblPr>
    <w:tblStyle w:val="TableGrid"/>
    <w:tblW w:w="0" w:type="auto"/>
    <w:tblBorders>
      <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="4680"/>  <!-- column width in twips -->
  </w:tblGrid>
  <w:tr>                      <!-- header row -->
    <w:tc>
      <w:tcPr><w:shd w:val="clear" w:fill="D9E2F3"/></w:tcPr>
      <w:p>
        <w:r><w:rPr><w:b/></w:rPr><w:t>Header</w:t></w:r>
      </w:p>
    </w:tc>
  </w:tr>
  <w:tr>                      <!-- data row -->
    <w:tc>
      <w:p><w:r><w:t>Cell</w:t></w:r></w:p>
    </w:tc>
  </w:tr>
</w:tbl>
```

## Horizontal Rules / Section Breaks

A horizontal rule in DOCX renders as a paragraph border:
```xml
<w:p>
  <w:pPr>
    <w:pBdr>
      <w:bottom w:val="single" w:sz="6" w:space="1" w:color="auto"/>
    </w:pBdr>
  </w:pPr>
</w:p>
```

## Heading → Markdown Mapping

| Markdown | DOCX Style | Font Size (half-pts) |
|----------|-----------|---------------------|
| `#` | Heading1 | 48 (24pt) |
| `##` | Heading2 | 36 (18pt) |
| `###` | Heading3 | 28 (14pt) |
| `####` | Heading4 | 24 (12pt) |
| `#####` | Heading5 | 22 (11pt) |
| `######` | Heading6 | 22 (11pt) italic |

## Inline Images

Images use `<w:drawing>` with `<wp:inline>` inside a run:
```xml
<w:r>
  <w:drawing>
    <wp:inline distT="0" distB="0" distL="0" distR="0">
      <wp:extent cx="1905000" cy="1428750"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:docPr id="1" name="Picture 1"/>
      <wp:cNvGraphicFramePr>
        <a:graphicFrameLocks noChangeAspect="1"/>
      </wp:cNvGraphicFramePr>
      <a:graphic>
        <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
          <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:nvPicPr>
              <pic:cNvPr id="0" name="image1.png"/>
              <pic:cNvPicPr/>
            </pic:nvPicPr>
            <pic:blipFill>
              <a:blip r:embed="rId5"/>
              <a:stretch><a:fillRect/></a:stretch>
            </pic:blipFill>
            <pic:spPr>
              <a:xfrm>
                <a:off x="0" y="0"/>
                <a:ext cx="1905000" cy="1428750"/>
              </a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </pic:spPr>
          </pic:pic>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>
</w:r>
```

Requires:
- Image binary in `word/media/image1.png`
- Relationship: `<Relationship Id="rId5" Type=".../relationships/image" Target="media/image1.png"/>`
- Content type default: `<Default Extension="png" ContentType="image/png"/>`
- Additional namespaces on `<w:document>`: `xmlns:wp`, `xmlns:a`, `xmlns:pic`

### Sizing

- Dimensions in EMU: `pixels × 914400 / 96` (96 DPI)
- Max width: page width minus margins = 6.5" = 5943600 EMU
- Use `noChangeAspect="1"` to preserve aspect ratio

## Content Types for DOCX

```xml
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/numbering.xml"
    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>
</Types>
```
