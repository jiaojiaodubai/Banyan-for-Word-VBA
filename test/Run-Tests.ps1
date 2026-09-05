# ============================================================================
# Run-Tests.ps1 - Full-suite test runner for Banyan (dev gate).
#
# Mirrors the "lint + mocha" habit from TS projects: after code changes, run
# this script. It injects ALL src modules plus ALL test modules (test\test*.bas)
# into a throwaway Word document through COM, runs every test module's
# RunTests(), and reports the aggregate. Exits non-zero if any test fails.
#
# The gate always runs the DEBUG configuration: modTest.bas is imported with
# DEV_MODE = True (see Set-VbeDevModeConstant) so DiagnosticsReraiseIfDev is
# active and swallowed business bugs surface as loud test failures.
#
# Usage:  powershell -ExecutionPolicy Bypass -File .\test\Run-Tests.ps1
# ============================================================================
$ErrorActionPreference = 'Stop'
$helperPath = Join-Path $PSScriptRoot "..\VbeEncodingHelpers.ps1"
. $helperPath

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$importRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("banyan-tests-" + [System.Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $importRoot | Out-Null

# Packaged modules (src) that the tests depend on.
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
    "src\modRibbonCallbacks.bas",
    "src\modTest.bas"
)

# Test modules, in run order. testField mutates the document, so run it last.
$testModules = @(
    "testI10n",
    "testJson",
    "testDict",
    "testHttp",
    "testPreference",
    "testField"
)

function Import-VbaModule([string]$relPath) {
    $sourcePath = Join-Path $root $relPath
    $targetPath = Join-Path $importRoot ([System.IO.Path]::GetFileName($relPath))
    $text = Get-VbeImportText $sourcePath
    if ($relPath.EndsWith(".bas") -and $text -notmatch "Attribute VB_Name") {
        $name = [System.IO.Path]::GetFileNameWithoutExtension($relPath)
        $text = "Attribute VB_Name = `"$name`"`r`n" + $text
    }
    # The dev gate always runs the DEBUG configuration: force DEV_MODE = True
    # in the imported modTest.bas so the dev-mode error policy
    # (DiagnosticsReraiseIfDev) is active and swallowed bugs fail loudly.
    if ($relPath -eq "src\modTest.bas") {
        $text = Set-VbeDevModeConstant $text $true
    }
    Write-VbeImportText $targetPath $text
    Write-Host "Importing $relPath"
    $script:vbProject.VBComponents.Import($targetPath) | Out-Null
}

# Word processes that exist before this script starts. Anything with the
# /Automation flag that appears AFTER this point is ours and MUST be killed
# in the finally block (a blocked Quit() would otherwise leave it behind).
$script:preExistingWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $doc.Content.Text = "Banyan test doc. "
    $script:vbProject = $doc.VBProject

    Write-Host "=== Injecting source modules ==="
    foreach ($rel in $srcFiles) { Import-VbaModule $rel }
    Write-Host ""

    Write-Host "=== Injecting test modules ==="
    foreach ($m in $testModules) { Import-VbaModule "test\$m.bas" }
    Write-Host ""

    # --- Run every test module ---
    $allReports = ""
    $totalPass = 0
    $totalFail = 0
    foreach ($m in $testModules) {
        Write-Host "Running $m.RunTests() ..."
        try {
            $report = $word.Run("$m.RunTests")
        } catch {
            $report = "[FAIL] $m - could not run: $($_.Exception.Message)`n"
        }
        $allReports += $report + "`n"

        $passCount = ([regex]::Matches($report, "\[PASS\]")).Count
        $failCount = ([regex]::Matches($report, "\[FAIL\]")).Count
        $totalPass += $passCount
        $totalFail += $failCount
        Write-Host ("  PASS: {0}  FAIL: {1}" -f $passCount, $failCount)
    }

    Write-Output ""
    Write-Output "==================== FULL TEST REPORT ===================="
    Write-Output $allReports
    Write-Output ("TOTAL  PASS: {0}  FAIL: {1}" -f $totalPass, $totalFail)

    $leftover = "fields: " + $doc.Fields.Count + ", footnotes: " + $doc.Footnotes.Count
    Write-Output ("Leftover in doc after tests: " + $leftover)

    if ($totalFail -gt 0) {
        Write-Output ""
        Write-Output "*** TEST FAILURES DETECTED - fix before committing ***"
        exit 1
    }
    Write-Output ""
    Write-Output "ALL TESTS PASSED"
    exit 0
} catch {
    Write-Output ("ERROR: " + $_.Exception.Message)
    Write-Output $_.ScriptStackTrace
    exit 1
} finally {
    # Robust cleanup: never leave Word (or a stray blank document) behind,
    # even if a step above threw before $doc was assigned. Each step is
    # guarded so a failure in one cannot skip the others.
    try {
        if ($word) {
            while ($word.Documents.Count -gt 0) {
                $word.Documents.Item(1).Close(0) | Out-Null
            }
            $word.Quit()
        }
    } catch {}
    try {
        if ($word) { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) }
    } catch {}
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    # Force-kill any Word automation instance this script spawned whose
    # Quit() did not complete (e.g. a modal dialog blocked it). Only touch
    # processes that did not exist before we started.
    $currentWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $leftoverPids = @($currentWordPids | Where-Object { $_ -notin $script:preExistingWordPids })
    foreach ($pidToKill in $leftoverPids) {
        Stop-Process -Id $pidToKill -Force -ErrorAction SilentlyContinue
    }
    Start-Sleep -Milliseconds 500
    $remaining = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    if ($remaining.Count -gt 0) {
        Write-Warning "WINWORD processes still present after cleanup: $($remaining -join ', ') - inspect manually."
    } else {
        Write-Host "Cleanup OK: no WINWORD processes left."
    }

    Remove-Item -Recurse -Force $importRoot -ErrorAction SilentlyContinue
}
