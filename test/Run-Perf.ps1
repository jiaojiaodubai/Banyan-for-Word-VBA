# ============================================================================
# Run-Perf.ps1 - Diagnostic timing probe for the Banyan pipeline.
#
# NOT the dev gate. This imports all src modules plus the diagnostic probe
# test\testPerf.bas into a throwaway Word document (hidden by default). The
# default RunPerf entry measures local Word/VBA stages and excludes backend and
# dialog time. -BackendBibliographyOnly deliberately uses the live backend to
# generate production bibliography lines for a paired patch comparison.
#
# The probe is imported with DEV_MODE = True so a swallowed business error
# surfaces as a loud [FAIL] line instead of silently producing bogus "fast"
# timings.
#
# Usage: powershell -ExecutionPolicy Bypass -File .\test\Run-Perf.ps1
#        .\test\Run-Perf.ps1 -Sizes 10,50,100 -Repetitions 4 -Visible
#        .\test\Run-Perf.ps1 -BackendBibliographyOnly -Sizes 10,25,50 -Repetitions 2
# ============================================================================
[CmdletBinding()]
param(
    [ValidateNotNullOrEmpty()]
    [string]$Sizes = '10,50,100',

    [ValidateRange(1, 20)]
    [int]$Repetitions = 4,

    [switch]$Visible,

    [switch]$ComparisonOnly,

    [switch]$BackendBibliographyOnly
)

$ErrorActionPreference = 'Stop'
$helperPath = Join-Path $PSScriptRoot "..\VbeEncodingHelpers.ps1"
. $helperPath

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$importRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("banyan-perf-" + [System.Guid]::NewGuid().ToString("N"))
$parsedSizes = @()
foreach ($rawSize in ($Sizes -split ',')) {
    $parsedSize = 0
    if (-not [int]::TryParse($rawSize.Trim(), [ref]$parsedSize) -or $parsedSize -lt 1 -or $parsedSize -gt 1000) {
        throw "Sizes must be a comma-separated list of integers between 1 and 1000."
    }
    $parsedSizes += $parsedSize
}
$sizesCsv = ($parsedSizes -join ',')
New-Item -ItemType Directory -Path $importRoot | Out-Null

$srcFiles = @(
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
    "src\modBusinessLogic.bas",
    "src\modAutomation.bas",
    "src\modRibbonCallbacks.bas",
    "src\modTest.bas"
)

function Import-VbaModule([string]$relPath) {
    $sourcePath = Join-Path $root $relPath
    $targetPath = Join-Path $importRoot ([System.IO.Path]::GetFileName($relPath))
    $text = Get-VbeImportText $sourcePath
    if ($relPath.EndsWith(".bas") -and $text -notmatch "Attribute VB_Name") {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($relPath)
        $text = "Attribute VB_Name = `"$name`"`r`n" + $text
    }
    if ($relPath -eq "src\modTest.bas") {
        # Keep the dev-mode error policy active so business errors are loud.
        $text = Set-VbeDevModeConstant $text $true
    }
    Write-VbeImportText $targetPath $text
    Write-Host "Importing $relPath"
    $script:vbProject.VBComponents.Import($targetPath) | Out-Null
}

$script:preExistingWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$script:ownedWordPids = @()

function Test-LocalTcpPort([int]$Port, [int]$TimeoutMs = 250) {
    $client = [System.Net.Sockets.TcpClient]::new()
    try {
        $pending = $client.BeginConnect('127.0.0.1', $Port, $null, $null)
        if (-not $pending.AsyncWaitHandle.WaitOne($TimeoutMs)) { return $false }
        $client.EndConnect($pending)
        return $true
    }
    catch {
        return $false
    }
    finally {
        $client.Dispose()
    }
}

try {
    $totalSw = [System.Diagnostics.Stopwatch]::StartNew()
    $launchSw = [System.Diagnostics.Stopwatch]::StartNew()
    $word = New-Object -ComObject Word.Application
    $launchSw.Stop()
    $script:ownedWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue |
        Where-Object { $_.Id -notin $script:preExistingWordPids } |
        Select-Object -ExpandProperty Id)
    $word.Visible = [bool]$Visible
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $doc.Content.Text = "Banyan perf doc. "
    $script:vbProject = $doc.VBProject

    $importSw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Host "=== Injecting source modules ==="
    foreach ($rel in $srcFiles) { Import-VbaModule $rel }
    Write-Host ""

    Write-Host "=== Injecting perf probe ==="
    Import-VbaModule "test\testPerf.bas"
    $importSw.Stop()
    Write-Host ""

    Write-Output "==================== PERF REPORT ===================="
    Write-Output ("[HOST] Word launch: {0:N2}s; VBA import: {1:N2}s; visible: {2}" -f `
        $launchSw.Elapsed.TotalSeconds, $importSw.Elapsed.TotalSeconds, [bool]$Visible)
    foreach ($port in 23119, 23124) {
        $portState = if (Test-LocalTcpPort $port) { 'LISTENING' } else { 'CLOSED' }
        Write-Output ("[HOST] backend TCP {0}: {1}" -f $port, $portState)
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $macroSizes = $sizesCsv
        $macroRepetitions = $Repetitions
        if ($BackendBibliographyOnly) {
            $macroName = 'testPerf.RunBackendBibliographyPerf'
        } elseif ($ComparisonOnly) {
            $macroName = 'testPerf.RunComparisonPerf'
        } else {
            $macroName = 'testPerf.RunPerf'
        }
        $res = $word.Run($macroName, [ref]$macroSizes, [ref]$macroRepetitions)
    } catch {
        $res = "[FAIL] testPerf.RunPerf could not run: $($_.Exception.Message)`n"
    }
    $sw.Stop()
    Write-Output ("--- per-stage internal timings [VBA wall {0:N2}s] ---" -f $sw.Elapsed.TotalSeconds)
    Write-Output $res
    Write-Output ""
    Write-Output "====================================================="

    # Basic health checks (should all be 0 - the probe cleans up after itself).
    $leftover = "fields: " + $doc.Fields.Count + ", footnotes: " + $doc.Footnotes.Count
    Write-Output ("Leftover in doc after perf: " + $leftover)
    $totalSw.Stop()
    Write-Output ("Total runner wall time: {0:N2}s" -f $totalSw.Elapsed.TotalSeconds)
    if ($res -match '\[FAIL\]') { exit 1 }
    exit 0
} catch {
    Write-Output ("ERROR: " + $_.Exception.Message)
    Write-Output $_.ScriptStackTrace
    exit 1
} finally {
    try {
        if ($word) {
            while ($word.Documents.Count -gt 0) {
                $word.Documents.Item(1).Close(0) | Out-Null
            }
            $word.Quit()
        }
    } catch {}
    try {
        if ($script:vbProject) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:vbProject) }
    } catch {}
    try {
        if ($doc) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($doc) }
    } catch {}
    try {
        if ($word) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    $currentWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $leftoverPids = @($currentWordPids | Where-Object { $_ -in $script:ownedWordPids })
    foreach ($pidToKill in $leftoverPids) {
        Stop-Process -Id $pidToKill -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Milliseconds 500
    $remaining = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $remainingOwned = @($remaining | Where-Object { $_ -in $script:ownedWordPids })
    if ($remainingOwned.Count -gt 0) {
        Write-Warning "Perf-owned WINWORD processes still present after cleanup: $($remainingOwned -join ', ') - inspect manually."
    } else {
        Write-Host "Cleanup OK: no perf-owned WINWORD processes left."
    }

    for ($cleanupAttempt = 1; $cleanupAttempt -le 3; $cleanupAttempt++) {
        Remove-Item -LiteralPath $importRoot -Recurse -Force -ErrorAction SilentlyContinue
        if (-not (Test-Path -LiteralPath $importRoot)) { break }
        Start-Sleep -Milliseconds 250
    }
    if (Test-Path -LiteralPath $importRoot) {
        Write-Warning "Temporary VBA import directory could not be removed: $importRoot"
    }
}
