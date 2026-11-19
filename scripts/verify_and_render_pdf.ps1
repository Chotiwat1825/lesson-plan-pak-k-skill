param(
    [Parameter(Mandatory = $true)]
    [string]$PdfPath,

    [Parameter(Mandatory = $false)]
    [string]$OutDir = ""
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $PdfPath)) {
    throw "PDF not found: $PdfPath"
}

$resolvedPdf = (Resolve-Path -LiteralPath $PdfPath).Path
if ([string]::IsNullOrWhiteSpace($OutDir)) {
    $OutDir = Join-Path (Split-Path -Parent $resolvedPdf) "pdf_pages_check"
}
New-Item -ItemType Directory -Force -Path $OutDir | Out-Null

Write-Output "=== 1. PYPDF PAGE & FONT AUDIT ==="
$pyCode = @"
import sys
from pypdf import PdfReader

pdf_path = sys.argv[1]
r = PdfReader(pdf_path)
print(f"Total PDF Pages: {len(r.pages)} (Target: 5 pages A4)")
fonts = set()
for i, p in enumerate(r.pages):
    res = p.get('/Resources') or {}
    fdict = res.get('/Font') or {}
    for k in fdict:
        fobj = fdict[k].get_object()
        bf = str(fobj.get('/BaseFont', ''))
        if bf:
            fonts.add(bf)
    txt = (p.extract_text() or '').strip()
    first_line = txt.splitlines()[0] if txt else ''
    last_line = txt.splitlines()[-1] if txt else ''
    print(f"  Page {i+1}: chars={len(txt)} | first={first_line[:65]!r} | last={last_line[-55:]!r}")

print("Embedded Fonts:", sorted(fonts))
has_sarabun = any('Sarabun' in f for f in fonts)
print("TH SarabunPSK Embedded Check:", "PASS" if has_sarabun else "WARNING - Check font embedding")
print("5-Page Ceiling Check:", "PASS (Exactly 5 pages)" if len(r.pages) == 5 else f"FAIL (Found {len(r.pages)} pages)")
"@

$tempPy = Join-Path $env:TEMP "audit_pdf_plan.py"
Set-Content -LiteralPath $tempPy -Value $pyCode -Encoding UTF8
python -u $tempPy $resolvedPdf
Remove-Item -LiteralPath $tempPy -Force -ErrorAction SilentlyContinue

Write-Output "`n=== 2. RENDERING PNG PAGES FOR VISUAL READ_FILE AUDIT ==="
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType=WindowsRuntime]
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime]

function Await($WinRtTask, $ResultType) {
    $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.IsGenericMethod })[0].MakeGenericMethod($ResultType)
    $netTask = $asTask.Invoke($null, @($WinRtTask))
    $netTask.Wait(-1) | Out-Null
    $netTask.Result
}
function AwaitAction($WinRtAction) {
    $asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and -not $_.IsGenericMethod })[0]
    $netTask = $asTask.Invoke($null, @($WinRtAction))
    $netTask.Wait(-1) | Out-Null
}

$file = Await ([Windows.Storage.StorageFile]::GetFileFromPathAsync($resolvedPdf)) ([Windows.Storage.StorageFile])
$pdfDoc = Await ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])

for ($i = 0; $i -lt $pdfDoc.PageCount; $i++) {
    $page = $pdfDoc.GetPage($i)
    $memStream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
    $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
    $opts.DestinationWidth = 1400
    AwaitAction ($page.RenderToStreamAsync($memStream, $opts))
    $reader = New-Object Windows.Storage.Streams.DataReader($memStream.GetInputStreamAt(0))
    $bytes = New-Object byte[] $memStream.Size
    Await ($reader.LoadAsync($memStream.Size)) ([uint32]) | Out-Null
    $reader.ReadBytes($bytes)
    $pngPath = Join-Path $OutDir ("page_" + ($i + 1) + ".png")
    [System.IO.File]::WriteAllBytes($pngPath, $bytes)
    Write-Output "Rendered Page $($i + 1): $pngPath ($($bytes.Length) bytes)"
    $page.Dispose()
}
