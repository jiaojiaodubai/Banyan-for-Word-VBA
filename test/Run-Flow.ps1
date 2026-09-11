# ============================================================================
# Run-Flow.ps1 - End-to-end agent-flow runner (live backend, no GUI dialogs).
#
# Injects all src modules into a throwaway Word document and drives
# modAutomation through the whole flow (style -> item snapshot -> pending
# citation/bibliography -> refresh) with rich-text assertions and a second
# refresh for idempotency. A live backend is required (modHttp uses 23119 and
# falls back to the dev port 23124). Modules are imported with DEV_MODE = True
# so swallowed business errors surface as loud failures.
#
# Usage: powershell -ExecutionPolicy Bypass -File .\test\Run-Flow.ps1
#        .\test\Run-Flow.ps1 -Visible -ItemKey 9W4UQYKB
#        .\test\Run-Flow.ps1 -StyleId "gb-t-7714-2025-numeric"
# ============================================================================
[CmdletBinding()]
param(
    [string]$StyleId = 'gb-t-7714-2025-numeric',

    # Empty = auto-pick a plain-title journalArticle from the snapshot.
    [string]$ItemKey = '',

    [ValidateRange(1, 500)]
    [int]$Limit = 100,

    [switch]$Visible,

    [switch]$SkipSecondRefresh
)

$ErrorActionPreference = 'Stop'
$helperPath = Join-Path $PSScriptRoot "..\VbeEncodingHelpers.ps1"
. $helperPath

$root = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$importRoot = Join-Path ([System.IO.Path]::GetTempPath()) ("banyan-flow-" + [System.Guid]::NewGuid().ToString("N"))
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
        $text = Set-VbeDevModeConstant $text $true
    }
    Write-VbeImportText $targetPath $text
    Write-Host "Importing $relPath"
    $script:vbProject.VBComponents.Import($targetPath) | Out-Null
}

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

function Get-AddinFieldsByPrefix($doc, [string]$prefix) {
    $matches = New-Object System.Collections.ArrayList
    $total = [int]$doc.Fields.Count
    for ($index = 1; $index -le $total; $index++) {
        $fld = $doc.Fields.Item($index)
        if ($fld.Type -eq 81) {
            # Word wraps the code text: " ADDIN BANYAN_CITATION <id> ".
            $code = ([string]$fld.Code.Text).Trim()
            if ($code -like "*$prefix*") { [void]$matches.Add($fld) }
        }
    }
    return , $matches
}

# Word.Application.Run marshals its optional arguments ByRef, so a COM call
# with arguments must pass PSReference values (plain values raise
# 'should be a System.Management.Automation.PSReference').
function Invoke-Macro1([string]$name, $a1) {
    $r1 = $a1
    return $script:word.Run($name, [ref]$r1)
}
function Invoke-Macro2([string]$name, $a1, $a2) {
    $r1 = $a1
    $r2 = $a2
    return $script:word.Run($name, [ref]$r1, [ref]$r2)
}

$script:preExistingWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$script:ownedWordPids = @()
$script:failures = New-Object System.Collections.ArrayList

function Add-Check([string]$name, [bool]$ok, [string]$detail = '') {
    $tag = if ($ok) { '[PASS]' } else { '[FAIL]' }
    Write-Output ("{0} {1}{2}" -f $tag, $name, $(if ($detail) { " - $detail" } else { '' }))
    if (-not $ok) { [void]$script:failures.Add($name) }
}

try {
    $word = New-Object -ComObject Word.Application
    $script:word = $word
    $script:ownedWordPids = @(Get-Process WINWORD -ErrorAction SilentlyContinue |
        Where-Object { $_.Id -notin $script:preExistingWordPids } |
        Select-Object -ExpandProperty Id)
    $word.Visible = [bool]$Visible
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Add()
    $doc.Content.Text = "Banyan automation flow doc. "
    $script:vbProject = $doc.VBProject

    Write-Host "=== Injecting source modules ==="
    foreach ($rel in $srcFiles) { Import-VbaModule $rel }
    Write-Host ""

    Write-Output "==================== FLOW REPORT ===================="
    foreach ($port in 23119, 23124) {
        $portState = if (Test-LocalTcpPort $port) { 'LISTENING' } else { 'CLOSED' }
        Write-Output ("[HOST] backend TCP {0}: {1}" -f $port, $portState)
    }

    # --- 1. Style: list route -> BANYAN_PREF --------------------------------
    $styleOk = [bool](Invoke-Macro2 'modAutomation.AutomationStyleApply' $StyleId '')
    Add-Check "style applied from /banyan/style/list ($StyleId)" $styleOk
    if (-not $styleOk) {
        Write-Output ("[INFO] HttpGetLastError: " + $word.Run('modHttp.HttpGetLastError'))
    }

    $prefStyleId = [string]$word.Run('modAutomation.AutomationPreferenceStyleId')
    Add-Check "BANYAN_PREF carries the style id" ($prefStyleId -eq $StyleId) $prefStyleId

    # --- 2. Item snapshot ----------------------------------------------------
    $snapshotCount = [int](Invoke-Macro1 'modAutomation.AutomationItemsSnapshot' $Limit)
    Add-Check "item snapshot cached ($snapshotCount items)" ($snapshotCount -gt 0)

    if (-not $ItemKey -and $snapshotCount -gt 0) {
        $candidates = @()
        for ($i = 1; $i -le $snapshotCount; $i++) {
            $key = [string](Invoke-Macro1 'modAutomation.AutomationItemKeyAt' $i)
            $type = [string](Invoke-Macro1 'modAutomation.AutomationItemTypeAt' $i)
            $title = [string](Invoke-Macro1 'modAutomation.AutomationItemTitleAt' $i)
            if ($key -and $title -and $title -notmatch '<') {
                $candidates += [pscustomobject]@{ Key = $key; Type = $type; Title = $title }
            }
        }
        $pick = $candidates | Where-Object { $_.Type -eq 'journalArticle' } | Select-Object -First 1
        if (-not $pick) { $pick = $candidates | Select-Object -First 1 }
        if ($pick) { $ItemKey = $pick.Key }
    }
    Add-Check "item selected" ([bool]$ItemKey) $ItemKey
    Write-Output ("[INFO] item key: {0}" -f $ItemKey)

    # --- 3. Pending citation + bibliography ---------------------------------
    $citeOk = [bool](Invoke-Macro1 'modAutomation.AutomationInsertPendingCitationForKey' $ItemKey)
    $citeFields = Get-AddinFieldsByPrefix $doc 'BANYAN_CITATION'
    Add-Check "pending citation inserted" ($citeOk -and $citeFields.Count -eq 1)
    if ($citeFields.Count -eq 1) {
        $pendingData = [string]$citeFields[0].Data
        $pendingText = [string]$citeFields[0].Result.Text
        Add-Check "citation starts as placeholder" `
            (($pendingData -match 'INTEXT_CITATION') -and ($pendingText -match 'INTEXT_CITATION'))
        Add-Check "pending source carries item key" ($pendingData -match [regex]::Escape($ItemKey))
    }

    $bibOk = [bool]$word.Run('modAutomation.AutomationInsertPendingBibliography')
    $bibFields = Get-AddinFieldsByPrefix $doc 'BANYAN_BIBLIOGRAPHY'
    Add-Check "pending bibliography inserted" ($bibOk -and $bibFields.Count -eq 1)
    if ($bibFields.Count -eq 1) {
        $bibText = [string]$bibFields[0].Result.Text
        Add-Check "bibliography starts as placeholder" ($bibText -match 'BIBLIOGRAPHY')
    }

    # --- 4. Refresh ----------------------------------------------------------
    $refresh1 = [bool]$word.Run('modAutomation.AutomationRefresh')
    Add-Check "refresh #1 succeeded" $refresh1
    if (-not $refresh1) {
        Write-Output ("[INFO] HttpGetLastError: " + $word.Run('modHttp.HttpGetLastError'))
    }

    # --- 5. Verify rendered rich text ---------------------------------------
    $citeFields = Get-AddinFieldsByPrefix $doc 'BANYAN_CITATION'
    Add-Check "citation field survived refresh" ($citeFields.Count -eq 1)
    if ($citeFields.Count -eq 1) {
        $fld = $citeFields[0]
        $resultText = [string]$fld.Result.Text
        $dataText = [string]$fld.Data
        Add-Check "citation rendered as [1]" ($resultText.Trim() -eq '[1]') $resultText
        Add-Check "citation data no longer placeholder" `
            (($dataText -notmatch 'INTEXT_CITATION') -and ($dataText -match '"superscript"'))
        Add-Check "citation content links to entry" ($dataText -match 'banyan://entry/')

        # Word reports wdUndefined (9999999) for some field-boundary characters
        # even when a mark covers the whole result, so require: none of the
        # characters is explicitly non-superscript, and at least one is on.
        $superOn = 0
        $superOff = 0
        $charCount = 0
        foreach ($ch in $fld.Result.Characters) {
            $charCount++
            $value = $ch.Font.Superscript
            if ($value -eq -1) { $superOn++ }
            elseif ($value -eq 0) { $superOff++ }
        }
        Add-Check "citation rendered superscript" ($charCount -gt 0 -and $superOn -ge 1 -and $superOff -eq 0) `
            ("on {0}/{1}, off {2}" -f $superOn, $charCount, $superOff)

        $hyperlinkOk = $false
        $linkDetail = ''
        if ($fld.Result.Hyperlinks.Count -ge 1) {
            $subAddress = [string]$fld.Result.Hyperlinks.Item(1).SubAddress
            $linkDetail = $subAddress
            $hyperlinkOk = ($subAddress -eq "Banyan_Entry_$ItemKey")
        }
        Add-Check "citation hyperlink targets bibliography bookmark" $hyperlinkOk $linkDetail
    }

    $bibFields = Get-AddinFieldsByPrefix $doc 'BANYAN_BIBLIOGRAPHY'
    Add-Check "bibliography entry field rendered" ($bibFields.Count -ge 1)
    if ($bibFields.Count -ge 1) {
        $bib = $bibFields[0]
        $bibResult = [string]$bib.Result.Text
        $bibData = [string]$bib.Data
        Add-Check "bibliography starts with [1]" ($bibResult -match '^\s*\[1\]') $bibResult
        Add-Check "bibliography data is an entry" ($bibData -match '"bibliography-entry"')
        Add-Check "no bibliography placeholder left" ($bibResult -notmatch 'BIBLIOGRAPHY')
    }
    $bookmarkNames = @()
    foreach ($bookmark in $doc.Bookmarks) { $bookmarkNames += [string]$bookmark.Name }
    Add-Check "bibliography bookmark exists" ($bookmarkNames -contains "Banyan_Entry_$ItemKey") `
        ($bookmarkNames -join ', ')
    # Word rules: start with a letter, letters/digits/underscore only, 40 chars max.
    $banyanBookmarks = @($bookmarkNames | Where-Object { $_ -like 'Banyan_*' })
    $invalidBookmarks = @($banyanBookmarks | Where-Object { $_ -notmatch '^[A-Za-z][A-Za-z0-9_]{0,39}$' })
    Add-Check "bookmark names follow Word rules" `
        (($banyanBookmarks.Count -ge 1) -and ($invalidBookmarks.Count -eq 0)) `
        ($banyanBookmarks -join ', ')

    # --- 6. Second refresh (idempotency) ------------------------------------
    if (-not $SkipSecondRefresh) {
        $refresh2 = [bool]$word.Run('modAutomation.AutomationRefresh')
        Add-Check "refresh #2 succeeded" $refresh2
        $citeFields2 = Get-AddinFieldsByPrefix $doc 'BANYAN_CITATION'
        if ($citeFields2.Count -eq 1) {
            Add-Check "citation stable after second refresh" `
                (([string]$citeFields2[0].Result.Text).Trim() -eq '[1]')
        }
        $bibFields2 = Get-AddinFieldsByPrefix $doc 'BANYAN_BIBLIOGRAPHY'
        Add-Check "bibliography stable after second refresh" ($bibFields2.Count -eq $bibFields.Count)
    }

    Write-Output ""
    $failCount = $script:failures.Count
    Write-Output ("RESULT: {0} failure(s){1}" -f $failCount, $(if ($failCount -gt 0) { ' -> ' + ($script:failures -join '; ') } else { '' }))
    Write-Output "====================================================="
    if ($failCount -gt 0) { exit 1 }
    exit 0
}
catch {
    Write-Output ("ERROR: " + $_.Exception.Message)
    Write-Output $_.ScriptStackTrace
    exit 1
}
finally {
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
        Write-Warning "Flow-owned WINWORD processes still present after cleanup: $($remainingOwned -join ', ') - inspect manually."
    } else {
        Write-Host "Cleanup OK: no flow-owned WINWORD processes left."
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
