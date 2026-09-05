<#
.SYNOPSIS
    Build Banyan.dotm from VBA source (release or debug) and open the template.

.DESCRIPTION
    This is the BUILD script for the Banyan Word add-in. It imports all VBA
    source files into Banyan.dotm, saves the template, and either leaves it open
    or closes it with -CloseWord. It also serves as the everyday dev loop: after
    a source change, re-run it against the already-open template to refresh it.

    Word/VBE imports text modules using the current ANSI code page. This script
    prepares GBK (CP936) temporary copies before importing so Chinese strings are
    preserved on zh-CN Office installations.

    Standard .bas files in this repository are kept without an exported
    Attribute VB_Name header for readability. The temporary copies receive the
    correct VB_Name based on the source file name before import, so VBE creates
    modules with stable names instead of Module1/Module2.

    The script can update an already-open Banyan.dotm, or open a template path.
    It removes existing Banyan components with the same names before importing.

    DEV_MODE (src\modTest.bas) defaults to False (release build). Pass -DevMode
    to build a DEBUG template: the imported copy of modTest.bas is rewritten
    with DEV_MODE = True at build time while the checked-in source stays False.

.PARAMETER TemplatePath
    Path to Banyan.dotm. If omitted, the script first looks for an already-open
    Banyan.dotm in Word, then common local paths.

.PARAMETER CloseWord
    Save and close the template, then quit the Word instance created by this
    script. If the template was already open before running, Word is left open.

.PARAMETER DevMode
    Build a DEBUG template: import modTest.bas with DEV_MODE = True (shows the
    btnTest ribbon button and enables the DiagnosticsReraiseIfDev dev error
    policy). Default is a release build (DEV_MODE = False). This only affects
    the imported copy - src\modTest.bas is never modified.

.EXAMPLE
    # Release build (default)
    .\Import-BanyanDotm.ps1 -TemplatePath "$env:APPDATA\Microsoft\Templates\Banyan.dotm"

.EXAMPLE
    # Debug build for development/testing
    .\Import-BanyanDotm.ps1 -TemplatePath ".\Banyan.dotm" -DevMode

.EXAMPLE
    .\Import-BanyanDotm.ps1
#>

[CmdletBinding()]
param(
    [string]$TemplatePath,
    [switch]$CloseWord,
    [switch]$DevMode
)

$ErrorActionPreference = "Stop"
$helperPath = Join-Path $PSScriptRoot "VbeEncodingHelpers.ps1"
. $helperPath

if ($PSScriptRoot) {
    $projectRoot = $PSScriptRoot
}
else {
    $projectRoot = (Resolve-Path ".").Path
}
$importRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("banyan-dotm-import-" + [System.Guid]::NewGuid().ToString("N"))
$word = $null
$createdWord = $false
$doc = $null
$openedByScript = $false
$previousDisplayAlerts = $null

$moduleFiles = @(
    "src\Dictionary.cls",
    "src\JsonConverter.bas",
    "src\modJson.bas",
    "src\modDict.bas",
    "src\modI10n.bas",
    "src\modDiagnostics.bas",
    "src\modShell.bas",
    "src\modHttp.bas",
    "src\modProgress.bas",
    "src\modField.bas",
    "src\modChapterBreak.bas",
    "src\modPreference.bas",
    "src\modCitation.bas",
    "src\modBibliography.bas",
    "src\modRefresh.bas",
    "src\modConvert.bas",
    "src\modFinalize.bas",
    "src\frmSettings.frm",
    "src\modBusinessLogic.bas",
    "src\modRibbonCallbacks.bas",
    "src\modTest.bas",
    "test\testI10n.bas",
    "test\testJson.bas",
    "test\testDict.bas"
)

$removedModuleNames = @(
    "modRichText",
    "modConfig"
)

function Get-ModuleName([string]$RelativePath) {
    return [System.IO.Path]::GetFileNameWithoutExtension($RelativePath)
}

function Get-VbeImportPath([string]$RelativePath) {
    $sourcePath = Join-Path $projectRoot $RelativePath
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Missing source file: $RelativePath"
    }

    $targetPath = Join-Path $importRoot $RelativePath
    $extension = [System.IO.Path]::GetExtension($RelativePath).ToLowerInvariant()
    $text = Get-VbeImportText $sourcePath

    if ($extension -eq ".bas") {
        $text = Assert-BasAttributeHeader $text (Get-ModuleName $RelativePath)
        if ($DevMode -and (Get-ModuleName $RelativePath) -eq "modTest") {
            Write-Host "  (Debug build) modTest.bas imported with DEV_MODE = True"
            $text = Set-VbeDevModeConstant $text $true
        }
    }

    Write-VbeImportText $targetPath $text

    if ($extension -eq ".frm") {
        Copy-VbeFrmBinary $sourcePath $targetPath
    }

    return $targetPath
}

function Get-WordApplication {
    try {
        return [Runtime.InteropServices.Marshal]::GetActiveObject("Word.Application")
    }
    catch {
        $script:createdWord = $true
        return New-Object -ComObject Word.Application
    }
}

function Find-OpenTemplate($WordApplication, [string]$Path) {
    $resolvedPath = $null
    if ($Path) {
        $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction SilentlyContinue)
        if ($resolvedPath) {
            $resolvedPath = $resolvedPath.Path
        }
    }

    foreach ($item in $WordApplication.Documents) {
        try {
            if ($resolvedPath -and [string]::Equals($item.FullName, $resolvedPath, [System.StringComparison]::OrdinalIgnoreCase)) {
                return $item
            }
            if (-not $resolvedPath -and [string]::Equals($item.Name, "Banyan.dotm", [System.StringComparison]::OrdinalIgnoreCase)) {
                return $item
            }
        }
        catch {
        }
    }

    return $null
}

function Resolve-TemplatePath($WordApplication) {
    if ($TemplatePath) {
        return (Resolve-Path -LiteralPath $TemplatePath).Path
    }

    $candidates = @(
        (Join-Path $projectRoot "Banyan.dotm"),
        (Join-Path $projectRoot "src\Banyan.dotm"),
        (Join-Path $env:APPDATA "Microsoft\Templates\Banyan.dotm"),
        (Join-Path $env:APPDATA "Microsoft\Word\STARTUP\Banyan.dotm")
    )

    if ($null -ne $WordApplication) {
        try {
            if ($WordApplication.StartupPath) {
                $candidates += (Join-Path $WordApplication.StartupPath "Banyan.dotm")
            }
        }
        catch {
        }

        try {
            foreach ($addIn in $WordApplication.AddIns) {
                if ([string]::Equals($addIn.Name, "Banyan.dotm", [System.StringComparison]::OrdinalIgnoreCase)) {
                    $addInPath = Join-Path $addIn.Path $addIn.Name
                    if (Test-Path -LiteralPath $addInPath) {
                        return (Resolve-Path -LiteralPath $addInPath).Path
                    }
                }
            }
        }
        catch {
        }
    }

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return (Resolve-Path -LiteralPath $candidate).Path
        }
    }

    return $null
}

function Remove-ExistingComponent($VbProject, [string]$ComponentName) {
    for ($i = $VbProject.VBComponents.Count; $i -ge 1; $i--) {
        $component = $VbProject.VBComponents.Item($i)
        if ([string]::Equals($component.Name, $ComponentName, [System.StringComparison]::OrdinalIgnoreCase)) {
            $VbProject.VBComponents.Remove($component)
            return
        }
    }
}

function Remove-ScriptingRuntimeReference($VbProject) {
    $scriptingRuntimeGuid = "{420B2830-E718-11CF-893D-00A0C9054228}"

    for ($i = $VbProject.References.Count; $i -ge 1; $i--) {
        $reference = $VbProject.References.Item($i)
        try {
            if ([string]::Equals($reference.Guid, $scriptingRuntimeGuid, [System.StringComparison]::OrdinalIgnoreCase)) {
                $VbProject.References.Remove($reference)
                Write-Host "Removed Microsoft Scripting Runtime reference"
            }
        }
        catch {
            # Ignore references whose metadata cannot be read.
        }
    }
}

function Get-DocumentVbProject($Document) {
    try {
        return $Document.VBProject
    }
    catch {
        throw @"
Unable to access the VBA project in Banyan.dotm.

Enable Word's "Trust access to the VBA project object model" setting, then run this script again.
Original error: $($_.Exception.Message)
"@
    }
}

try {
    New-Item -ItemType Directory -Path $importRoot | Out-Null

    $word = Get-WordApplication
    $word.Visible = $true
    $previousDisplayAlerts = $word.DisplayAlerts
    $word.DisplayAlerts = 0

    $resolvedTemplatePath = Resolve-TemplatePath $word
    $doc = Find-OpenTemplate $word $resolvedTemplatePath

    if ($null -eq $doc) {
        if (-not $resolvedTemplatePath) {
            throw "Banyan.dotm was not found. Re-run with -TemplatePath `"C:\path\to\Banyan.dotm`"."
        }
        Write-Host "Opening $resolvedTemplatePath"
        $doc = $word.Documents.Open($resolvedTemplatePath)
        $openedByScript = $true
    }
    else {
        Write-Host "Using open template: $($doc.FullName)"
    }

    $vbProject = Get-DocumentVbProject $doc

    Remove-ScriptingRuntimeReference $vbProject

    foreach ($componentName in $removedModuleNames) {
        Write-Host "Removing stale component $componentName"
        Remove-ExistingComponent $vbProject $componentName
    }

    foreach ($relativePath in $moduleFiles) {
        $componentName = Get-ModuleName $relativePath
        Write-Host "Importing $relativePath as $componentName"
        Remove-ExistingComponent $vbProject $componentName
        $vbProject.VBComponents.Import((Get-VbeImportPath $relativePath)) | Out-Null
    }

    $doc.Save()
    $doc.Activate()
    $word.Visible = $true
    $word.DisplayAlerts = $previousDisplayAlerts
    $previousDisplayAlerts = $null

    if ($CloseWord) {
        $savedPath = $doc.FullName
        if ($openedByScript) {
            $doc.Close($true)
        }
        if ($createdWord) {
            $word.Quit() | Out-Null
        }
        if ($openedByScript) {
            Write-Host "Import OK. Template saved and closed: $savedPath"
        }
        else {
            Write-Host "Import OK. Template saved. It was already open, so Word was left open: $savedPath"
        }
    }
    else {
        Write-Host "Import OK. Template saved and left open: $($doc.FullName)"
    }
}
finally {
    if ($null -ne $word -and $null -ne $previousDisplayAlerts) {
        try {
            $word.DisplayAlerts = $previousDisplayAlerts
        }
        catch {
        }
    }
    Remove-Item -LiteralPath $importRoot -Recurse -Force -ErrorAction SilentlyContinue
}
