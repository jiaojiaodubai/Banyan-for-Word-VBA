<#
.SYNOPSIS
    Generate frmSettings.frm/frx with the VBA UserForm designer.

.DESCRIPTION
    Hand-written MSForms control trees in .frm files are not reliably accepted
    by Word's VBE importer. This script asks Word/VBE to create the UserForm
    and its controls through the VBIDE Designer, exports the resulting .frm and
    .frx, then restores the code section from the existing src/frmSettings.frm.

    Word must allow "Trust access to the VBA project object model".
#>

$ErrorActionPreference = "Stop"

$projectRoot = $PSScriptRoot
$helperPath = Join-Path $projectRoot "VbeEncodingHelpers.ps1"
. $helperPath

$frmPath = Join-Path $projectRoot "src\frmSettings.frm"
$frxPath = Join-Path $projectRoot "src\frmSettings.frx"
$ansiCodePage = [System.Globalization.CultureInfo]::CurrentCulture.TextInfo.ANSICodePage

if (-not (Test-Path -LiteralPath $frmPath)) {
    throw "Missing $frmPath"
}

$existingText = Get-VbeImportText (Resolve-Path $frmPath)
$codeIndex = $existingText.IndexOf("Option Explicit")
if ($codeIndex -lt 0) {
    throw "Could not find Option Explicit in $frmPath"
}
$codeSection = $existingText.Substring($codeIndex) -replace "`r?`n", "`r`n"

$word = $null
$doc = $null

function Set-Common($ctrl, [double]$left, [double]$top, [double]$width, [double]$height) {
    $ctrl.Left = $left
    $ctrl.Top = $top
    $ctrl.Width = $width
    $ctrl.Height = $height
}

try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $doc = $word.Documents.Add()
    $vbProject = $doc.VBProject

    # vbext_ct_MSForm = 3
    $component = $vbProject.VBComponents.Add(3)
    $component.Name = "frmSettings"
    $form = $component.Designer
    $component.Properties.Item("Caption").Value = "Banyan Preferences"
    $component.Properties.Item("Width").Value = 430
    $component.Properties.Item("Height").Value = 330
    $component.Properties.Item("StartUpPosition").Value = 1

    $fraGlobal = $form.Controls.Add("Forms.Frame.1", "fraGlobal", $true)
    Set-Common $fraGlobal 12 12 390 66
    $fraGlobal.Caption = "Global"

    $chkSyncItems = $fraGlobal.Controls.Add("Forms.CheckBox.1", "chkSyncItems", $true)
    Set-Common $chkSyncItems 12 18 350 18
    $chkSyncItems.Caption = "Sync item metadata"

    $chkRefreshAll = $fraGlobal.Controls.Add("Forms.CheckBox.1", "chkRefreshAll", $true)
    Set-Common $chkRefreshAll 12 40 350 18
    $chkRefreshAll.Caption = "Refresh all chapters"

    $fraChapter = $form.Controls.Add("Forms.Frame.1", "fraChapter", $true)
    Set-Common $fraChapter 12 88 390 150
    $fraChapter.Caption = "Section"

    $lblStyle = $fraChapter.Controls.Add("Forms.Label.1", "lblStyle", $true)
    Set-Common $lblStyle 12 24 115 18
    $lblStyle.Caption = "Banyan style"

    $txtStyleTitle = $fraChapter.Controls.Add("Forms.TextBox.1", "txtStyleTitle", $true)
    Set-Common $txtStyleTitle 130 20 180 20
    $txtStyleTitle.Locked = $true

    $cmdFetchStyle = $fraChapter.Controls.Add("Forms.CommandButton.1", "cmdFetchStyle", $true)
    Set-Common $cmdFetchStyle 318 19 58 22
    $cmdFetchStyle.Caption = "Choose..."

    $lblBibTitleStyle = $fraChapter.Controls.Add("Forms.Label.1", "lblBibTitleStyle", $true)
    Set-Common $lblBibTitleStyle 12 64 115 18
    $lblBibTitleStyle.Caption = "Bibliography title"

    $txtBibTitleStyle = $fraChapter.Controls.Add("Forms.TextBox.1", "txtBibTitleStyle", $true)
    Set-Common $txtBibTitleStyle 130 60 180 20

    $cmdClearBibTitle = $fraChapter.Controls.Add("Forms.CommandButton.1", "cmdClearBibTitle", $true)
    Set-Common $cmdClearBibTitle 318 59 58 22
    $cmdClearBibTitle.Caption = "Clear"

    $lblBibEntryStyle = $fraChapter.Controls.Add("Forms.Label.1", "lblBibEntryStyle", $true)
    Set-Common $lblBibEntryStyle 12 104 115 18
    $lblBibEntryStyle.Caption = "Bibliography entry"

    $txtBibEntryStyle = $fraChapter.Controls.Add("Forms.TextBox.1", "txtBibEntryStyle", $true)
    Set-Common $txtBibEntryStyle 130 100 180 20

    $cmdClearBibEntry = $fraChapter.Controls.Add("Forms.CommandButton.1", "cmdClearBibEntry", $true)
    Set-Common $cmdClearBibEntry 318 99 58 22
    $cmdClearBibEntry.Caption = "Clear"

    $lblStatus = $form.Controls.Add("Forms.Label.1", "lblStatus", $true)
    Set-Common $lblStatus 12 254 245 18
    $lblStatus.Caption = ""

    $cmdOK = $form.Controls.Add("Forms.CommandButton.1", "cmdOK", $true)
    Set-Common $cmdOK 262 250 65 24
    $cmdOK.Caption = "OK"

    $cmdCancel = $form.Controls.Add("Forms.CommandButton.1", "cmdCancel", $true)
    Set-Common $cmdCancel 337 250 65 24
    $cmdCancel.Caption = "Cancel"

    Remove-Item -LiteralPath $frmPath -ErrorAction SilentlyContinue
    Remove-Item -LiteralPath $frxPath -ErrorAction SilentlyContinue
    $component.Export($frmPath)

    $doc.Close($false)
    $word.Quit()
    $doc = $null
    $word = $null

    $exportedText = Get-VbeImportText (Resolve-Path $frmPath)
    $exportedCodeIndex = $exportedText.IndexOf("Option Explicit")
    if ($exportedCodeIndex -lt 0) {
        $attributeEnd = $exportedText.IndexOf("Attribute VB_Exposed")
        if ($attributeEnd -lt 0) {
            throw "Could not find exported UserForm attributes"
        }
        $lineEnd = $exportedText.IndexOf("`n", $attributeEnd)
        $designerSection = $exportedText.Substring(0, $lineEnd + 1)
    }
    else {
        $designerSection = $exportedText.Substring(0, $exportedCodeIndex)
    }

    $finalText = ($designerSection.TrimEnd() + "`r`n" + $codeSection) -replace "`r?`n", "`r`n"
    Write-TextWithEncoding (Resolve-Path $frmPath) $finalText $script:VbeAnsiEncoding

    Write-Host "Generated $frmPath"
    Write-Host "Generated $frxPath"
    Write-Host "Encoding: ANSI code page $ansiCodePage, CRLF"
}
catch {
    if ($null -ne $doc) {
        try { $doc.Close($false) } catch {}
    }
    if ($null -ne $word) {
        try { $word.Quit() } catch {}
    }
    throw
}
