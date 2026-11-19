<#
.SYNOPSIS
Reference script for cloning the Golden 5-Page Lesson Plan Template (template_lesson_plan_v3.docx
or a user-provided reference .docx), updating paragraphs/tables in-place while preserving exact
user-tuned spacing, and exporting a 5-page PDF without Word COM hangs.

.NOTES
Key Word COM Anti-Hang Rules learned from conversation 65857ca1-9362-4289-91b8-62c4900f003d:
1. Always terminate stale WINWORD processes and remove ~$*.docx lock files before opening Word.
2. Mutate paragraphs and table cells using Range manipulation (Find-Para / Set-ParaTextPreserveFmt / Range.InsertAfter)
   instead of $cell.Range.Select() inside complex tables.
3. After calling $doc.Save(), set $doc.Saved = $true, call $doc.Close(0) and $word.Quit(0), and stop WINWORD if needed.
4. For PDF export ($doc.ExportAsFixedFormat($outPdf, 17)), open the saved .docx in Read-Only mode
   ($word.Documents.Open($outDocx, $false, $true)) so Word never blocks on background save locks.
#>

param(
    [string]$TemplateDocx = "$PSScriptRoot\..\resources\template_lesson_plan_v3.docx",
    [Parameter(Mandatory = $true)]
    [string]$OutDocx,
    [Parameter(Mandatory = $true)]
    [string]$OutPdf
)

$ErrorActionPreference = "Stop"

Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
$outDir = Split-Path -Parent $OutDocx
New-Item -ItemType Directory -Force -Path $outDir | Out-Null
Remove-Item (Join-Path $outDir "~$*.docx") -Force -ErrorAction SilentlyContinue

$workDocx = Join-Path $env:TEMP ("lesson_plan_work_" + [guid]::NewGuid().ToString("N") + ".docx")
Copy-Item -LiteralPath $TemplateDocx -Destination $workDocx -Force

function Find-Para($doc, [string]$startsWith) {
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        if ($p.Range.Text.Trim().StartsWith($startsWith)) { return $p }
    }
    throw "Paragraph not found starting with: $startsWith"
}

function Set-ParaTextPreserveFmt($para, [string]$newText, [int]$boldPrefixLen = 0) {
    $rng = $para.Range.Duplicate
    $rng.MoveEnd(1, -1) | Out-Null
    $fontSize = $para.Range.Font.Size
    $rng.Text = $newText
    $rng.Font.Name = "TH SarabunPSK"
    $rng.Font.NameBi = "TH SarabunPSK"
    $rng.Font.Size = [single]$fontSize
    $rng.Font.SizeBi = [single]$fontSize
    $rng.Font.Color = 0
    if ($boldPrefixLen -gt 0) {
        $rng.Font.Bold = 0; $rng.Font.BoldBi = 0
        $rngBold = $para.Range.Duplicate
        $rngBold.SetRange($para.Range.Start, $para.Range.Start + $boldPrefixLen)
        $rngBold.Font.Bold = 1; $rngBold.Font.BoldBi = 1
    }
}

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open($workDocx)

    # --- APPLY PARAGRAPH & TABLE UPDATES HERE ---
    # Example:
    # $t1 = $doc.Tables.Item(1)
    # $t1.Cell(3, 1).Range.Text = "เรื่อง ..."

    $doc.Repaginate()
    $pageCount = $doc.ComputeStatistics(2)
    Write-Output "Computed Word Page Count: $pageCount (Target: 5)"

    $doc.Save()
    $doc.Saved = $true
    $doc.Close(0)
}
finally {
    $word.Quit(0)
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

Copy-Item -LiteralPath $workDocx -Destination $OutDocx -Force
Remove-Item -LiteralPath $workDocx -Force -ErrorAction SilentlyContinue

# Export PDF from Read-Only open document (fast, zero lock risk)
$word2 = New-Object -ComObject Word.Application
$word2.Visible = $false
$word2.DisplayAlerts = 0
try {
    $doc2 = $word2.Documents.Open($OutDocx, $false, $true)
    $doc2.ExportAsFixedFormat($OutPdf, 17)
    $doc2.Close(0)
}
finally {
    $word2.Quit(0)
    Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}

Write-Output "Saved DOCX: $OutDocx"
Write-Output "Saved PDF:  $OutPdf"
