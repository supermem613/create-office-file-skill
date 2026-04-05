#!/usr/bin/env node
// verify-output.mjs — Open generated Office files in Word/PowerPoint via COM and extract content
// Windows-only: uses PowerShell COM automation with Word.Application / PowerPoint.Application
// Outputs JSON to stdout with extracted text, formatting, and structure info
'use strict';

import { execSync } from 'child_process';
import { resolve, extname, join } from 'path';
import { existsSync, writeFileSync, unlinkSync } from 'fs';
import { tmpdir } from 'os';

const file = process.argv[2];
if (!file) { console.error('Usage: node verify-output.mjs <file.pptx|docx>'); process.exit(1); }
if (process.platform !== 'win32') { console.error('ERROR: verify-output.mjs requires Windows with Microsoft Office'); process.exit(2); }

const absPath = resolve(file);
if (!existsSync(absPath)) { console.error(`ERROR: file not found: ${absPath}`); process.exit(1); }

const ext = extname(absPath).toLowerCase();
if (ext !== '.pptx' && ext !== '.docx') { console.error('ERROR: file must be .pptx or .docx'); process.exit(1); }

// PowerShell script for PPTX verification via COM
const pptxScript = `
$ErrorActionPreference = 'Stop'
$path = '${absPath.replace(/'/g, "''")}'
$result = @{ type = 'pptx'; file = $path; slides = @(); error = $null; repairNeeded = $false }
try {
  $ppt = New-Object -ComObject PowerPoint.Application
  $ppt.DisplayAlerts = 2  # ppAlertsNone
  $pres = $ppt.Presentations.Open($path, $true, $false, $false)
  $result.slideCount = $pres.Slides.Count
  $result.slideWidth = $pres.PageSetup.SlideWidth
  $result.slideHeight = $pres.PageSetup.SlideHeight
  foreach ($slide in $pres.Slides) {
    $slideInfo = @{ index = $slide.SlideIndex; shapes = @() }
    foreach ($shape in $slide.Shapes) {
      $shapeInfo = @{ name = $shape.Name; hasText = $false; hasTable = $false; isPicture = $false; shapeType = $shape.Type }
      if ($shape.Type -eq 13) {
        $shapeInfo.isPicture = $true
        $shapeInfo.hasText = $false
      }
      elseif ($shape.HasTable) {
        $shapeInfo.hasTable = $true
        $tbl = $shape.Table
        $shapeInfo.tableRows = $tbl.Rows.Count
        $shapeInfo.tableCols = $tbl.Columns.Count
        $shapeInfo.cells = @()
        $cellTexts = @()
        for ($row = 1; $row -le $tbl.Rows.Count; $row++) {
          for ($col = 1; $col -le $tbl.Columns.Count; $col++) {
            $cell = $tbl.Cell($row, $col)
            $cellText = $cell.Shape.TextFrame.TextRange.Text
            $shapeInfo.cells += @{ row = $row; col = $col; text = $cellText }
            $cellTexts += $cellText
          }
        }
        $shapeInfo.hasText = $true
        $shapeInfo.text = ($cellTexts -join ' ')
      }
      elseif ($shape.HasTextFrame -and $shape.TextFrame.HasText) {
        $shapeInfo.hasText = $true
        $tr = $shape.TextFrame.TextRange
        $shapeInfo.text = $tr.Text
        $shapeInfo.paragraphs = @()
        for ($p = 1; $p -le $tr.Paragraphs().Count; $p++) {
          $para = $tr.Paragraphs($p)
          $paraInfo = @{ text = $para.Text; runs = @() }
          for ($r = 1; $r -le $para.Runs().Count; $r++) {
            $run = $para.Runs($r)
            $runInfo = @{
              text = $run.Text
              bold = [bool]$run.Font.Bold
              italic = [bool]$run.Font.Italic
              fontName = $run.Font.Name
              fontSize = $run.Font.Size
            }
            $paraInfo.runs += $runInfo
          }
          $shapeInfo.paragraphs += $paraInfo
        }
      }
      $slideInfo.shapes += $shapeInfo
    }
    $result.slides += $slideInfo
  }
  $pres.Close()
  $ppt.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt) | Out-Null
} catch {
  $result.error = $_.Exception.Message
  try { $ppt.Quit() } catch {}
}
$result | ConvertTo-Json -Depth 10 -Compress
`;

// PowerShell script for DOCX verification via COM
const docxScript = `
$ErrorActionPreference = 'Stop'
$path = '${absPath.replace(/'/g, "''")}'
$result = @{ type = 'docx'; file = $path; paragraphs = @(); error = $null; inlineShapeCount = 0 }
try {
  $word = New-Object -ComObject Word.Application
  $word.Visible = $false
  $word.DisplayAlerts = 0
  $doc = $word.Documents.Open($path, $false, $true)
  $result.paragraphCount = $doc.Paragraphs.Count
  $result.inlineShapeCount = $doc.InlineShapes.Count
  foreach ($para in $doc.Paragraphs) {
    $paraInfo = @{
      text = $para.Range.Text.TrimEnd([char]13)
      style = $para.Style.NameLocal
      alignment = $para.Alignment
      runs = @()
    }
    $words = $para.Range.Words
    for ($w = 1; $w -le [Math]::Min($words.Count, 50); $w++) {
      $wd = $words.Item($w)
      $runInfo = @{
        text = $wd.Text
        bold = [bool]$wd.Font.Bold
        italic = [bool]$wd.Font.Italic
        fontName = $wd.Font.Name
        fontSize = $wd.Font.Size
      }
      $paraInfo.runs += $runInfo
    }
    $result.paragraphs += $paraInfo
  }
  $doc.Close($false)
  $word.Quit()
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
} catch {
  $result.error = $_.Exception.Message
  try { $word.Quit() } catch {}
}
$result | ConvertTo-Json -Depth 10 -Compress
`;

const script = ext === '.pptx' ? pptxScript : docxScript;

// Write script to temp file to avoid shell escaping issues
const tmpScript = join(tmpdir(), `verify_office_${Date.now()}.ps1`);
try {
  writeFileSync(tmpScript, script, 'utf8');
  const raw = execSync(
    `powershell -NoProfile -ExecutionPolicy Bypass -File "${tmpScript}"`,
    { encoding: 'utf8', timeout: 60000, stdio: ['pipe', 'pipe', 'pipe'], windowsHide: true }
  );
  const result = JSON.parse(raw.trim());
  if (result.error) {
    console.error(`ERROR: Office COM failed: ${result.error}`);
    process.exit(3);
  }
  console.log(JSON.stringify(result, null, 2));
} catch (e) {
  const stderr = e.stderr || '';
  if (stderr.includes('New-Object') && stderr.includes('ComObject')) {
    console.error('ERROR: Microsoft Office not installed (COM object creation failed)');
  } else {
    console.error(`ERROR: PowerShell COM execution failed: ${e.message}`);
    if (stderr) console.error(stderr.slice(0, 500));
  }
  process.exit(3);
} finally {
  try { unlinkSync(tmpScript); } catch {}
}
