$ErrorActionPreference = "Stop"

$fontsSourceDir = Join-Path $PSScriptRoot "..\resources\fonts"
$userFontsDir   = Join-Path $env:LOCALAPPDATA "Microsoft\Windows\Fonts"
$regPath        = "HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts"

New-Item -ItemType Directory -Force -Path $userFontsDir | Out-Null
if (-not (Test-Path $regPath)) {
    New-Item -Path $regPath -Force | Out-Null
}

$fontMap = @(
    @{ File = "THSarabunPSK Regular.ttf";    RegName = "TH SarabunPSK (TrueType)" },
    @{ File = "THSarabunPSK Bold.ttf";       RegName = "TH SarabunPSK Bold (TrueType)" },
    @{ File = "THSarabunPSK Italic.ttf";     RegName = "TH SarabunPSK Italic (TrueType)" },
    @{ File = "THSarabunPSK BoldItalic.ttf"; RegName = "TH SarabunPSK Bold Italic (TrueType)" }
)

foreach ($f in $fontMap) {
    $src = Join-Path $fontsSourceDir $f.File
    $dst = Join-Path $userFontsDir $f.File
    if (Test-Path -LiteralPath $src) {
        if (-not (Test-Path -LiteralPath $dst)) {
            Copy-Item -LiteralPath $src -Destination $dst -Force
            Write-Output "Installed font file: $dst"
        }
        New-ItemProperty -Path $regPath -Name $f.RegName -Value $dst -PropertyType String -Force | Out-Null
    }
}

Write-Output "TH SarabunPSK font family is installed and registered for current user."
