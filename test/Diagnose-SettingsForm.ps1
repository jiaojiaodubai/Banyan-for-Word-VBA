<#
.SYNOPSIS
    Regression guard for the settings dialog (frmSettings) after a template build.

.DESCRIPTION
    Opens a temporary COPY of Banyan.dotm, injects an in-memory probe that drives
    frmSettings.cmdOK_Click (the OK button handler) against a synthetic preference
    object, and fails unless the probe compiles, runs and persists BANYAN_PREF
    with the style id.

    Why this exists: UserForm code is only compiled when its events run, so a
    broken form module (e.g. a call to an API that was deleted in a refactor)
    ships unnoticed - users hit "Compile error in hidden module: frmSettings" when
    they click OK. test\Run-Tests.ps1 cannot see it; this script can.

    The template file itself is never modified: the probe and bridge live in
    memory only and the copy is closed without saving.

    Requires Word's "Trust access to the VBA project object model".

    Notes on the harness quirks (verified on Word 16 zh-CN):
    - Application.Run cannot invoke UserForm module procedures, hence the
      standard-module bridge.
    - Macro lookup follows the ACTIVE document, so the copy is opened writable and
      activated before running.
    - PowerShell cannot enumerate CustomDocumentProperties reliably; the persisted
      preference is read back by a VBA function whose return value Run returns.
    - The whole COM phase runs in a child PowerShell with a timeout so a blocking
      VBE dialog can never hang the caller.

.PARAMETER TemplatePath
    Path to the template under test. Defaults to ..\Banyan.dotm next to test\.

.EXAMPLE
    .\test\Diagnose-SettingsForm.ps1
.EXAMPLE
    .\test\Diagnose-SettingsForm.ps1 -TemplatePath .\Banyan.dotm
#>
[CmdletBinding()]
param(
    [string]$TemplatePath,
    [switch]$Work
)

$ErrorActionPreference = 'Stop'

if (-not $TemplatePath) {
    $TemplatePath = Join-Path $PSScriptRoot '..\Banyan.dotm'
}

if (-not $Work) {
    # --- parent mode: run the COM phase in a child with a hard timeout ---
    $resolved = (Resolve-Path -LiteralPath $TemplatePath).Path
    $pre = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
    $out = Join-Path $env:TEMP ("banyan-settingsform-out-" + [Guid]::NewGuid().ToString("N") + ".txt")
    $err = Join-Path $env:TEMP ("banyan-settingsform-err-" + [Guid]::NewGuid().ToString("N") + ".txt")
    $child = Start-Process -FilePath 'powershell' `
        -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ('"' + $MyInvocation.MyCommand.Path + '"'), '-Work', '-TemplatePath', ('"' + $resolved + '"') `
        -RedirectStandardOutput $out -RedirectStandardError $err -PassThru -WindowStyle Hidden
    if ($child.WaitForExit(120000)) {
        $lines = @(Get-Content -LiteralPath $out -ErrorAction SilentlyContinue)
        $lines | Write-Host
        $errText = Get-Content -LiteralPath $err -ErrorAction SilentlyContinue
        if ($errText) { Write-Host '--- STDERR ---'; $errText }
        # Start-Process does not expose a readable process exit code here, so the
        # outcome is taken from the child's RESULT line.
        if (@($lines -match '^RESULT: PASS$').Count -gt 0) { $code = 0 }
        else {
            $code = 1
            if (@($lines -match '^RESULT:').Count -eq 0) { Write-Host 'FAIL: child produced no RESULT line' }
        }
    }
    else {
        Stop-Process -Id $child.Id -Force -ErrorAction SilentlyContinue
        Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $pre -notcontains $_.Id } | Stop-Process -Force -ErrorAction SilentlyContinue
        Write-Host 'FAIL: timed out - a modal VBE dialog is probably blocking Word'
        $code = 1
    }
    Remove-Item -LiteralPath $out, $err -Force -ErrorAction SilentlyContinue
    exit $code
}

# --- work mode: COM phase ---
if (-not (Test-Path -LiteralPath $TemplatePath)) {
    Write-Host ("FAIL: template not found: " + $TemplatePath)
    Write-Host 'RESULT: FAIL'
    exit 1
}

$pre = @(Get-Process WINWORD -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$workDir = Join-Path $env:TEMP ("banyan-settingsform-" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $workDir | Out-Null
$copy = Join-Path $workDir 'Banyan-diagnose.dotm'
Copy-Item -LiteralPath $TemplatePath $copy
$word = $null
$failed = $false
try {
    Write-Host ("Template: " + $TemplatePath)
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 1
    $doc = $word.Documents.Open($copy, $false, $false)
    $proj = $doc.VBProject
    $fs = $proj.VBComponents.Item('frmSettings')
    $cm = $fs.CodeModule
    $code = $cm.Lines(1, $cm.CountOfLines)

    # Static guard: APIs deleted on purpose must not be called again. The OK
    # handler must not reference the removed save-diagnostic plumbing.
    if ($code -match 'PreferenceGetLastSaveError') {
        Write-Host 'FAIL: frmSettings still calls PreferenceGetLastSaveError (deleted API) - the form cannot compile.'
        $failed = $true
    }
    if (-not $failed) {
        $probe = @(
            'Public Sub ProbeClickOK()',
            '    Dim style As Object',
            '    Dim pref As Object',
            '    Set style = New Dictionary',
            '    style("id") = "the-journal-of-international-studies"',
            '    style("title") = "TIS"',
            '    style("citationType") = "note-citation"',
            '    Set pref = New Dictionary',
            '    pref("syncItems") = False',
            '    pref("refreshAll") = True',
            '    Set pref("style") = style',
            '    Set m_pref = pref',
            '    Set m_style = CloneStyle(pref("style"))',
            '    cmdOK_Click',
            'End Sub'
        ) -join "`r`n"
        $cm.AddFromString($probe)

        $mt = $proj.VBComponents.Item('modTest').CodeModule
        $bridge = @(
            'Public Function ProbeSettingsOK() As String',
            '    On Error GoTo BEH',
            '    Dim f As frmSettings',
            '    Set f = New frmSettings',
            '    f.ProbeClickOK',
            '    ProbeSettingsOK = CStr(ActiveDocument.CustomDocumentProperties("BANYAN_PREF").Value)',
            '    Exit Function',
            'BEH:',
            '    ProbeSettingsOK = "ERROR " & CStr(Err.Number) & "|" & Err.Description',
            'End Function'
        ) -join "`r`n"
        $mt.AddFromString($bridge)

        $doc.Activate()
        $value = ''
        try { $value = [string]$word.Run('modTest.ProbeSettingsOK') }
        catch { $value = 'ERROR ' + $_.Exception.Message }

        Write-Host ('OK-button path returned: ' + $value)
        if ($value -match 'the-journal-of-international-studies') {
            Write-Host 'PASS: settings dialog OK path compiles, runs and saves the preference.'
        }
        else {
            Write-Host 'FAIL: settings dialog OK path did not persist the preference (compile or runtime error).'
            $failed = $true
        }
    }

    try { $doc.Close($false) } catch { }
}
catch {
    Write-Host ('FAIL: ' + $_.Exception.Message)
    $failed = $true
}
finally {
    if ($null -ne $word) {
        try { foreach ($d in @($word.Documents)) { try { $d.Close($false) } catch { } } } catch { }
        try { $word.Quit() } catch { }
    }
    Get-Process WINWORD -ErrorAction SilentlyContinue | Where-Object { $pre -notcontains $_.Id } | Stop-Process -Force -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $workDir -Recurse -Force -ErrorAction SilentlyContinue
}

if ($failed) {
    Write-Host 'RESULT: FAIL'
    exit 1
}
Write-Host 'RESULT: PASS'
exit 0
