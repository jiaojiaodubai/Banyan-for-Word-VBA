<#
.SYNOPSIS
    Normalize VBA UserForm .frm files in place.

.DESCRIPTION
    This is a standalone normalization utility for checked-in .frm files. It
    reads each file as UTF-8 first, then falls back to the current Windows ANSI
    code page if needed, and rewrites CRLF line endings.

.PARAMETER Encoding
    Ansi      Current Windows ANSI code page. Recommended for VBE import.
    UTF8BOM   UTF-8 with BOM. Kept as a fallback for non-Chinese locales.

.EXAMPLE
    .\Convert-FrmEncoding.ps1

.EXAMPLE
    .\Convert-FrmEncoding.ps1 -Encoding UTF8BOM
#>

param(
    [ValidateSet("Ansi", "UTF8BOM")]
    [string]$Encoding = "Ansi"
)

$helperPath = Join-Path $PSScriptRoot "VbeEncodingHelpers.ps1"
. $helperPath

$projectRoot = $PSScriptRoot
$frmFiles = Get-ChildItem -Path $projectRoot -Recurse -Filter *.frm

if ($frmFiles.Count -eq 0) {
    Write-Host "No .frm files found under $projectRoot"
    exit 0
}

$ansiCodePage = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage
$targetEncoding = if ($Encoding -eq "UTF8BOM") { [System.Text.UTF8Encoding]::new($true) } else { $script:VbeAnsiEncoding }

foreach ($file in $frmFiles) {
    $content = Get-VbeImportText $file.FullName
    Write-TextWithEncoding $file.FullName $content $targetEncoding
    Write-Host "[OK] $($file.FullName)"
}

Write-Host ""
if ($Encoding -eq "Ansi") {
    Write-Host "Done. $($frmFiles.Count) .frm file(s) normalized as Ansi code page $ansiCodePage with CRLF."
}
else {
    Write-Host "Done. $($frmFiles.Count) .frm file(s) normalized as UTF8BOM with CRLF."
}
