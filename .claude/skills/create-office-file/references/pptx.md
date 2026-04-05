# PPTX OOXML Reference

> Progressive disclosure reference for PowerPoint generation.
> Read this when working on PPTX-specific code in scripts/create-office-file.mjs.

## File Structure

A valid .pptx is a ZIP archive containing:

```
[Content_Types].xml          # MIME types for all parts
_rels/.rels                  # Root relationships
ppt/presentation.xml         # Slide order, slide size
ppt/_rels/presentation.xml.rels  # Links to slides, master, theme
ppt/slides/slide1.xml        # Slide content
ppt/slides/_rels/slide1.xml.rels # Slide → layout relationship
ppt/slideLayouts/slideLayout1.xml
ppt/slideLayouts/_rels/slideLayout1.xml.rels
ppt/slideMasters/slideMaster1.xml
ppt/slideMasters/_rels/slideMaster1.xml.rels
ppt/theme/theme1.xml         # Color/font/effect definitions
```

**Critical:** PowerPoint requires slide master + layout + theme or it shows a repair dialog.

## Namespaces

| Prefix | URI | Used For |
|--------|-----|----------|
| `p:` | `http://schemas.openxmlformats.org/presentationml/2006/main` | Slide structure |
| `a:` | `http://schemas.openxmlformats.org/drawingml/2006/main` | Text, shapes, formatting |
| `r:` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationship IDs |

## Slide Dimensions

Default 16:9 widescreen:
- Width: `12192000` EMU (13.333 inches × 914400)
- Height: `6858000` EMU (7.5 inches × 914400)

## Shape (Text Box) Structure

```xml
<p:sp>
  <p:nvSpPr>
    <p:cNvPr id="2" name="TextBox 1"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm>
      <a:off x="457200" y="1600200"/>   <!-- position in EMU -->
      <a:ext cx="11277600" cy="4525963"/> <!-- size in EMU -->
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" rtlCol="0"/>
    <a:lstStyle/>
    <!-- paragraphs go here -->
  </p:txBody>
</p:sp>
```

## Text Formatting

### Paragraph with runs
```xml
<a:p>
  <a:r>
    <a:rPr lang="en-US" sz="1800" b="1" i="0" dirty="0"/>
    <a:t>Bold text</a:t>
  </a:r>
</a:p>
```

- `sz` = font size in 1/100 pt (1800 = 18pt, 2800 = 28pt, 4400 = 44pt)
- `b="1"` = bold
- `i="1"` = italic
- `u="sng"` = underline

### Bullet list paragraph
```xml
<a:p>
  <a:pPr marL="342900" indent="-342900">
    <a:buChar char="•"/>
  </a:pPr>
  <a:r>
    <a:rPr lang="en-US" sz="1800" dirty="0"/>
    <a:t>Bullet item</a:t>
  </a:r>
</a:p>
```

- `marL` = left margin in EMU
- `indent` = negative for hanging indent (bullet hangs left of text)
- Nested bullets: increase `marL` by 457200 per level

### Numbered list paragraph
```xml
<a:p>
  <a:pPr marL="342900" indent="-342900">
    <a:buAutoNum type="arabicPeriod"/>
  </a:pPr>
  <a:r>
    <a:rPr lang="en-US" sz="1800" dirty="0"/>
    <a:t>Numbered item</a:t>
  </a:r>
</a:p>
```

### Code block (monospace)
```xml
<a:r>
  <a:rPr lang="en-US" sz="1400" dirty="0">
    <a:latin typeface="Consolas"/>
    <a:cs typeface="Consolas"/>
  </a:rPr>
  <a:t>const x = 42;</a:t>
</a:r>
```

### Hyperlink
```xml
<a:r>
  <a:rPr lang="en-US" sz="1800" dirty="0">
    <a:hlinkClick r:id="rId2"/>
  </a:rPr>
  <a:t>Click here</a:t>
</a:r>
```
Requires a relationship entry in the slide's .rels file:
```xml
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
  Target="https://example.com" TargetMode="External"/>
```

## Slide Splitting Heuristic

| Markdown | PPTX Result |
|----------|-------------|
| `# Heading` | Title slide — large centered text |
| `## Heading` | Content slide with title bar |
| `---` | Explicit slide break |
| Content between headings | Body text / bullet points |

## Relationship IDs

Each slide needs a unique `rId` in `presentation.xml.rels`:
```xml
<Relationship Id="rId2" Type=".../relationships/slide" Target="slides/slide1.xml"/>
```

Convention: rId1 = slideMaster, rId2 = theme, rId3+ = slides.

## Content Types

Each slide needs an Override in `[Content_Types].xml`:
```xml
<Override PartName="/ppt/slides/slide1.xml"
  ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
```

## Table Structure

Tables in PPTX use `<a:tbl>` inside a `<a:graphicFrame>`:
```xml
<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="4" name="Table 1"/>
    <p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>
    <p:nvPr/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="457200" y="1600200"/>
    <a:ext cx="11277600" cy="2000000"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblGrid>
          <a:gridCol w="3759200"/>  <!-- repeat per column -->
        </a:tblGrid>
        <a:tr h="370840">           <!-- header row -->
          <a:tc>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p><a:r><a:rPr lang="en-US" b="1"/><a:t>Header</a:t></a:r></a:p>
            </a:txBody>
          </a:tc>
        </a:tr>
        <a:tr h="370840">           <!-- data row -->
          <a:tc>
            <a:txBody>
              <a:bodyPr/>
              <a:lstStyle/>
              <a:p><a:r><a:rPr lang="en-US"/><a:t>Cell</a:t></a:r></a:p>
            </a:txBody>
          </a:tc>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

## Image (Picture) Structure

Images in PPTX use `<p:pic>` shape with a blipFill referencing embedded media:
```xml
<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="4" name="Picture 3"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip r:embed="rId3"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="457200" y="1600200"/>
      <a:ext cx="1905000" cy="1428750"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>
```

Requires:
- Image binary in `ppt/media/image1.png`
- Relationship in slide rels: `<Relationship Id="rId3" Type=".../relationships/image" Target="../media/image1.png"/>`
- Content type default: `<Default Extension="png" ContentType="image/png"/>`

### Sizing

- Dimensions in EMU: `pixels × 914400 / 96` (96 DPI)
- Use `noChangeAspect="1"` in picLocks to preserve aspect ratio
- Cap to slide content area: ~10" × 5" (leaving room for title)
