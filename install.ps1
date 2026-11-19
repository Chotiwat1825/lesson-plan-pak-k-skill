$ErrorActionPreference = "Stop"

$targetDir = Join-Path $HOME ".gemini\config\skills\lesson-plan-pak-k"
New-Item -ItemType Directory -Force -Path $targetDir | Out-Null

$sourceDir = $PSScriptRoot
if ((Resolve-Path $sourceDir).Path -ne (Resolve-Path $targetDir).Path) {
    Copy-Item -Path "$sourceDir\*" -Destination $targetDir -Recurse -Force
    Write-Output "Copied skill files to: $targetDir"
} else {
    Write-Output "Skill already located in: $targetDir"
}

# Ensure TH SarabunPSK font is installed
powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $targetDir "scripts\ensure_thsarabun_font.ps1")

# Ensure pypdf is installed for 5-page PDF verification
$hasPypdf = python -c "import importlib.util; print(importlib.util.find_spec('pypdf') is not None)" 2>$null
if ($hasPypdf -ne "True") {
    Write-Output "Installing pypdf for PDF verification..."
    pip install pypdf --quiet
}

Write-Output "`n[SUCCESS] ติดตั้ง Skill 'lesson-plan-pak-k' และฟอนต์ TH SarabunPSK เรียบร้อยแล้ว!"
