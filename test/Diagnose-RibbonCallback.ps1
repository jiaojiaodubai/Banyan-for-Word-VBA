<#
.SYNOPSIS
    Diagnose the Word error "the macro cannot be found, or macros have been
    disabled due to macro security settings" (Chinese: 'you yu hong an quan she
    zhi, wu fa zhao dao hong, huo hong yi bei jin yong').

.DESCRIPTION
    Reads the embedded Ribbon (customUI) part of a .dotm/.docm package and the
    VBA project it ships, then verifies that EVERY callback referenced by the
    Ribbon XML (onLoad / onAction / getLabel / getVisible / getEnabled / ...)
    actually exists as a Public procedure in the imported VBA modules.

    A Ribbon callback that names a non-existent (or non-public) procedure makes
    Word show exactly the message in the title on startup - it is not a macro
    security problem.

.PARAMETER TemplatePath
    Path to the .dotm/.docm to inspect.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\test\Diagnose-RibbonCallback.ps1 -TemplatePath .\Banyan.dotm
#>

[CmdletBinding()]
param(
    [string]$TemplatePath = (Join-Path (Split-Path -Parent $PSScriptRoot) "Banyan.dotm")
)

$ErrorActionPreference = "Stop"

if (-not (Test-Path -LiteralPath $TemplatePath)) {
    throw "Template not found: $TemplatePath"
}
$TemplatePath = (Resolve-Path -LiteralPath $TemplatePath).Path

# --------------------------------------------------------------------------
# 1) Read the customUI XML from the OOXML package (customUI/customUI*.xml)
# --------------------------------------------------------------------------
$workRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("banyan-ribbon-diag-" + [System.Guid]::NewGuid().ToString("N"))
$zipCopy = Join-Path $workRoot "package.zip"
New-Item -ItemType Directory -Path $workRoot | Out-Null
Copy-Item -LiteralPath $TemplatePath -Destination $zipCopy
Expand-Archive -LiteralPath $zipCopy -DestinationPath (Join-Path $workRoot "x") -Force

$customUiFiles = @(Get-ChildItem -Path (Join-Path $workRoot "x") -Recurse -Filter "customUI*.xml" -File |
    Where-Object { $_.DirectoryName -notlike "*\_rels*" })

if ($customUiFiles.Count -eq 0) {
    Write-Host "[WARN] No customUI part found in $TemplatePath"; exit 2
}

$attributeNames = @("onLoad", "onAction", "getLabel", "getVisible", "getEnabled",
    "getScreentip", "getSupertip", "getImage", "getSize", "onShow", "onHide",
    "getPressed", "getContent", "onChange", "getDescription", "onGetItemLabel",
    "getItemCount", "getSelectedItemIndex", "onItemSelectionChanged")

$required = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)

foreach ($file in $customUiFiles) {
    Write-Host "[INFO] customUI part: $($file.FullName.Substring((Join-Path $workRoot 'x').Length + 1))"
    [xml]$xml = Get-Content -LiteralPath $file.FullName -Raw
    $nodes = $xml.SelectNodes("//*")
    foreach ($node in $nodes) {
        foreach ($attr in $attributeNames) {
            if ($node.HasAttribute($attr)) {
                $value = $node.GetAttribute($attr)
                if ($value -and $value -notmatch '^\s*$') {
                    [void]$required.Add($value)
                }
            }
        }
    }
}

Write-Host "[INFO] Ribbon callbacks referenced: $($required.Count)"
foreach ($r in ($required | Sort-Object)) { Write-Host "       - $r" }

# --------------------------------------------------------------------------
# 2) Enumerate Public procedures in the shipped VBA project (via Word COM)
# --------------------------------------------------------------------------
$preExisting = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$word = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open($TemplatePath, $false, $true)
    $vbp = $doc.VBProject

    $procedures = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::Ordinal)
    foreach ($component in $vbp.VBComponents) {
        $cm = $component.CodeModule
        if ($cm.CountOfLines -eq 0) { continue }
        for ($i = 1; $i -le $cm.CountOfLines; $i++) {
            $line = $cm.Lines($i, 1)
            if ($line -match '^\s*(Public\s+)?(Sub|Function)\s+([A-Za-z_]\w*)') {
                $kind = $Matches[2]
                $name = $Matches[3]
                $isPublic = ($line -match '^\s*Public\s+') -or ($kind -eq 'Function' -and $line -match '^\s*Function')
                # VBA module procedures default to Public; only "Private" is restricted.
                if ($line -notmatch '^\s*Private\s+') {
                    [void]$procedures.Add($name)
                }
            }
        }
    }
    $doc.Close($false)
}
finally {
    if ($null -ne $word) {
        try { $word.Quit() } catch { }
        try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch { }
    }
    Start-Sleep -Milliseconds 500
    Get-Process WINWORD -ErrorAction SilentlyContinue |
        Where-Object { $preExisting -notcontains $_.Id } |
        Stop-Process -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $workRoot -Recurse -Force -ErrorAction SilentlyContinue
}

# --------------------------------------------------------------------------
# 3) Report
# --------------------------------------------------------------------------
$missing = @($required | Where-Object { -not $procedures.Contains($_) } | Sort-Object)

Write-Host ""
Write-Host "=============== RESULT ==============="
if ($missing.Count -eq 0) {
    Write-Host "[PASS] Every Ribbon callback resolves to a Public VBA procedure."
    exit 0
}

Write-Host "[FAIL] Ribbon callback(s) with NO matching Public procedure in the VBA project:"
foreach ($m in $missing) { Write-Host "       - $m" }
Write-Host ""
Write-Host "Word will show the error: macro cannot be found / macros have been disabled"
Write-Host "(Chinese: 'you yu hong an quan she zhi, wu fa zhao dao hong, huo hong yi bei jin yong')"
exit 1
