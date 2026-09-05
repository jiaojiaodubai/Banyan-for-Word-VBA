$ErrorActionPreference = "Stop"
$helperPath = Join-Path $PSScriptRoot "VbeEncodingHelpers.ps1"
. $helperPath

$root = (Resolve-Path ".").Path
$importRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("banyan-vbe-import-" + [System.Guid]::NewGuid().ToString("N"))
$word = $null

function Get-VbeImportPath([string]$RelativePath) {
    $sourcePath = Join-Path $root $RelativePath
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Missing $RelativePath"
    }

    $targetPath = Join-Path $importRoot $RelativePath
    $text = Get-VbeImportText $sourcePath
    Write-VbeImportText $targetPath $text

    if ([System.IO.Path]::GetExtension($RelativePath).Equals(".frm", [System.StringComparison]::OrdinalIgnoreCase)) {
        Copy-VbeFrmBinary $sourcePath $targetPath
    }

    return $targetPath
}

try {
    New-Item -ItemType Directory -Path $importRoot | Out-Null

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    $doc = $word.Documents.Add()
    $vbProject = $doc.VBProject

    $files = @(
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
        "test\testDict.bas",
        "test\testHttp.bas",
        "test\testPreference.bas",
        "test\testField.bas"
    )

    foreach ($relPath in $files) {
        Write-Host "Importing $relPath"
        $vbProject.VBComponents.Import((Get-VbeImportPath $relPath)) | Out-Null
    }

    Write-Host "Import OK"
    $doc.Close($false)
}
finally {
    if ($null -ne $word) {
        $word.Quit() | Out-Null
    }
    Remove-Item -LiteralPath $importRoot -Recurse -Force -ErrorAction SilentlyContinue
}
