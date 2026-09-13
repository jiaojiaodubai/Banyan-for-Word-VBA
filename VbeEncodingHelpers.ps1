try {
    [System.Text.Encoding]::RegisterProvider([System.Text.CodePagesEncodingProvider]::Instance)
}
catch {
    # Windows PowerShell 5.1 already has legacy code pages available.
}

$script:VbeAnsiEncoding = [System.Text.Encoding]::GetEncoding([System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage)
$script:VbeUtf8Strict = [System.Text.UTF8Encoding]::new($false, $true)
$script:VbeGbkEncoding = [System.Text.Encoding]::GetEncoding(
    936,
    [System.Text.EncoderFallback]::ExceptionFallback,
    [System.Text.DecoderFallback]::ExceptionFallback)

function Read-TextWithFallbackEncoding([string]$Path) {
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    try {
        $text = $script:VbeUtf8Strict.GetString($bytes)
    }
    catch {
        $text = $script:VbeAnsiEncoding.GetString($bytes)
    }

    return (Assert-VbeTextIsNotCorrupted $text $Path)
}

function Standardize-LineEndings([string]$Text) {
    return ($Text -replace "`r?`n", "`r`n")
}

function Assert-VbeTextIsNotCorrupted([string]$Text, [string]$Path) {
    $cleanText = $Text.TrimStart([char]0xFEFF)

    if ($cleanText.Contains([char]0xFFFD)) {
        throw "Source text contains Unicode replacement characters (U+FFFD): $Path. Restore the file from a valid UTF-8/GBK copy before importing."
    }

    return $cleanText
}

function Write-TextWithEncoding([string]$Path, [string]$Text, [System.Text.Encoding]$Encoding) {
    $targetDir = Split-Path -Parent $Path
    if (-not [string]::IsNullOrEmpty($targetDir) -and -not (Test-Path -LiteralPath $targetDir)) {
        New-Item -ItemType Directory -Path $targetDir | Out-Null
    }
    $cleanText = Assert-VbeTextIsNotCorrupted $Text $Path
    [System.IO.File]::WriteAllText($Path, (Standardize-LineEndings $cleanText), $Encoding)
}

function Get-VbeImportText([string]$SourcePath) {
    $text = Read-TextWithFallbackEncoding $SourcePath
    Assert-DeclarationSectionOrder $text $SourcePath
    return $text
}

# Module-level declarations must sit above the first procedure: with the VBE's
# Compile-On-Demand a Const/Dim/Type below one is not registered and every use
# fails with "Variable not defined" (in the *using* module). Fail the build here.
function Assert-DeclarationSectionOrder([string]$Text, [string]$Path) {
    $lines = (Standardize-LineEndings $Text) -split "`r`n"
    $ln = 0
    $depth = 0
    $firstProcedureLine = 0
    foreach ($line in $lines) {
        $ln++
        if ($line -match '^\s*(Public\s+|Private\s+|Friend\s+)?(Function|Sub|Property\s+(Get|Let|Set))\s') {
            if ($depth -eq 0 -and $firstProcedureLine -eq 0) { $firstProcedureLine = $ln }
            $depth++
            continue
        }
        if ($line -match '^\s*End\s+(Function|Sub|Property)\s*$') {
            if ($depth -gt 0) { $depth-- }
            continue
        }
        if ($depth -eq 0 -and $firstProcedureLine -gt 0 -and
            $line -match '^(Public\s+|Private\s+|Friend\s+)?(Const|Dim|Static|Type|Enum|Declare)\b') {
            throw ("$Path line ${ln}: module-level declaration below the first procedure (line $firstProcedureLine). " +
                   "Move it into the declaration section at the top of the module, otherwise the VBE does not " +
                   "register it and every use fails with 'Variable not defined'. Offending line: " + $line.Trim())
        }
    }
}

function Write-VbeImportText([string]$TargetPath, [string]$Text) {
    Write-TextWithEncoding $TargetPath $Text $script:VbeGbkEncoding
}

function Assert-BasAttributeHeader([string]$Text, [string]$ModuleName) {
    $text = Standardize-LineEndings $Text
    if ($text -match '(?m)^Attribute VB_Name\s*=') {
        return $text
    }

    $lines = $text -split "`r`n", 2
    if ($lines.Count -gt 0 -and $lines[0] -eq "Option Explicit") {
        if ($lines.Count -eq 1) {
            return "Attribute VB_Name = `"$ModuleName`"`r`nOption Explicit`r`n"
        }
        return "Attribute VB_Name = `"$ModuleName`"`r`nOption Explicit`r`n" + $lines[1]
    }

    return "Attribute VB_Name = `"$ModuleName`"`r`n" + $text
}

function Set-VbeDevModeConstant([string]$Text, [bool]$Enabled) {
    # Rewrites the DEV_MODE toggle line in src/modTest.bas text so the build
    # scripts can produce Debug (True) or Release (False) imports without
    # touching the checked-in source, which always stays at the release default.
    $text = Standardize-LineEndings $Text
    if ($text -notmatch '(?m)^Private Const DEV_MODE As Boolean = (True|False)\s*$') {
        throw "Expected a 'Private Const DEV_MODE As Boolean = ...' line in modTest.bas text."
    }
    $value = if ($Enabled) { "True" } else { "False" }
    return ($text -replace '(?m)^Private Const DEV_MODE As Boolean = (True|False)', "Private Const DEV_MODE As Boolean = $value")
}

function Copy-VbeFrmBinary([string]$SourcePath, [string]$TargetPath) {
    $sourceFrx = [System.IO.Path]::ChangeExtension($SourcePath, ".frx")
    if (Test-Path -LiteralPath $sourceFrx) {
        $targetFrx = [System.IO.Path]::ChangeExtension($TargetPath, ".frx")
        $targetDir = Split-Path -Parent $targetFrx
        if (-not [string]::IsNullOrEmpty($targetDir) -and -not (Test-Path -LiteralPath $targetDir)) {
            New-Item -ItemType Directory -Path $targetDir | Out-Null
        }
        Copy-Item -LiteralPath $sourceFrx -Destination $targetFrx -Force
    }
}
