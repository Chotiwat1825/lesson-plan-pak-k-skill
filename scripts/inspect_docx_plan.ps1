param(
    [Parameter(Mandatory = $true)]
    [string]$DocxPath
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $DocxPath)) {
    throw "File not found: $DocxPath"
}

$resolvedPath = (Resolve-Path -LiteralPath $DocxPath).Path
$tempCopy = Join-Path $env:TEMP ("inspect_plan_" + [guid]::NewGuid().ToString("N") + [System.IO.Path]::GetExtension($resolvedPath))
Copy-Item -LiteralPath $resolvedPath -Destination $tempCopy -Force

$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0

try {
    $doc = $word.Documents.Open($tempCopy, $false, $true)
    $doc.Repaginate()
    $pages = $doc.ComputeStatistics(2)
    Write-Output "=== DOCUMENT OVERVIEW ==="
    Write-Output "Source: $resolvedPath"
    Write-Output "Total Pages: $pages"
    Write-Output ("Margins (pt): Top={0:N1} Bottom={1:N1} Left={2:N1} Right={3:N1}" -f $doc.PageSetup.TopMargin, $doc.PageSetup.BottomMargin, $doc.PageSetup.LeftMargin, $doc.PageSetup.RightMargin)
    Write-Output "Paragraphs: $($doc.Paragraphs.Count) | Tables: $($doc.Tables.Count) | Shapes: $($doc.Shapes.Count) | InlineShapes: $($doc.InlineShapes.Count)"

    Write-Output "`n=== PARAGRAPHS ==="
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {
        $p = $doc.Paragraphs.Item($i)
        $txt = ($p.Range.Text -replace "[\r\n\a]", " ").Trim()
        if ($txt.Length -gt 0) {
            $preview = if ($txt.Length -gt 110) { $txt.Substring(0, 110) + "..." } else { $txt }
            Write-Output ("P{0:D3} | align={1} sb={2} sa={3} li={4} fi={5} sz={6} bold={7} | {8}" -f $i, $p.Format.Alignment, $p.Format.SpaceBefore, $p.Format.SpaceAfter, $p.Format.LeftIndent, $p.Format.FirstLineIndent, $p.Range.Font.Size, $p.Range.Font.Bold, $preview)
        }
    }

    Write-Output "`n=== TABLES ==="
    for ($t = 1; $t -le $doc.Tables.Count; $t++) {
        $tbl = $doc.Tables.Item($t)
        Write-Output ("--- Table {0} (Rows={1}, Cols={2}) ---" -f $t, $tbl.Rows.Count, $tbl.Columns.Count)
        foreach ($cell in $tbl.Range.Cells) {
            $cTxt = ($cell.Range.Text -replace "[\r\n\a]", " ").Trim()
            $cPrev = if ($cTxt.Length -gt 90) { $cTxt.Substring(0, 90) + "..." } else { $cTxt }
            Write-Output ("  Cell({0},{1}) | sz={2} align={3} | {4}" -f $cell.RowIndex, $cell.ColumnIndex, $cell.Range.Font.Size, $cell.Range.ParagraphFormat.Alignment, $cPrev)
        }
    }

    $doc.Close(0)
}
finally {
    $word.Quit(0)
    Remove-Item -LiteralPath $tempCopy -Force -ErrorAction SilentlyContinue
}
