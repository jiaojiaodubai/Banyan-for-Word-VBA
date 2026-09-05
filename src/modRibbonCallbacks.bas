Option Explicit

' ============================================================================
' Module  : modRibbonCallbacks
' Purpose : Ribbon UI callback handlers for the Banyan custom tab.
'           Bridges Ribbon XML events to business-logic procedures.
'
' Dependencies : modI10n            (localization engine)
'                modBusinessLogic   (command handlers)
'
' Icons are hardcoded in Ribbon.xml via image="btnXxx" attributes.
' No VBA image callback is needed - manage icons in Office RibbonX Editor.
' ============================================================================

' --- Module-level variables ---
Private m_oRibbon As IRibbonUI          ' Ribbon reference for invalidation
Private m_ribbonI10nReady As Boolean


' --- Ribbon Load ---

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    ' Called by the Custom UI when the ribbon loads.
    ' 1) Store the reference so we can invalidate controls later.
    ' 2) Initialize localization only.
    '
    ' Note: PreferenceInit is NOT called here - the preference property is created
    ' lazily on first PreferenceGet/PreferenceSave to avoid silently
    ' modifying documents the user hasn't explicitly asked to process.
    Set m_oRibbon = ribbon
    EnsureRibbonI10n
End Sub

Public Sub OnRiibonLoad(ribbon As IRibbonUI)
    ' Backward-compatible wrapper for older Ribbon.xml builds that used this typo.
    OnRibbonLoad ribbon
End Sub


' --- Invalidation helpers ---

Public Sub InvalidateRibbon()
    ' Refreshes ALL controls (forces all get-callbacks to re-run).
    If Not m_oRibbon Is Nothing Then
        m_oRibbon.Invalidate
    End If
End Sub

Public Sub InvalidateControl(ByVal controlID As String)
    ' Refreshes a single control (forces its get-callbacks to re-run).
    If Not m_oRibbon Is Nothing Then
        m_oRibbon.InvalidateControl controlID
    End If
End Sub


' --- getLabel - Return localized display text for each button ---

Public Sub OnGetLabel(control As IRibbonControl, ByRef returnedVal)
    ' Resolves via modI10n.T("ribbon.<controlID>")
    ' Falls back to en-US, then to control.ID itself.
    EnsureRibbonI10n
    returnedVal = T("ribbon." & control.ID, control.ID)
End Sub


' --- getVisible - Show / hide buttons ---

Public Sub OnGetVisible(control As IRibbonControl, ByRef returnedVal)
    EnsureRibbonI10n

    ' Default: all buttons visible.  Add conditions as needed.
    Select Case control.ID
        Case "btnTest"
            ' Visible only in dev mode (modTest.IsDevMode)
            returnedVal = IsDevMode()
        Case "btnCitationPane"
            ' VBA is lacking a proper task pane API, so this button is hidden.
            returnedVal = false
        Case Else
            returnedVal = True
    End Select
End Sub


' --- onAction - Dispatch button clicks to business logic ---

Public Sub OnAction(control As IRibbonControl)
    EnsureRibbonI10n

    Select Case control.ID
        Case "btnCitation":       CitationAction
        Case "btnChapterBreak":   InsertChapterBreakFromPreference
        Case "btnBibliography":   BibliographyAction
        Case "btnCitationPane":   Exit Sub
        Case "btnRefresh":        RefreshAction
        Case "btnConvert":        ConvertAction
        Case "btnUnlink":         FinalizeAction
        Case "btnTest":           RunAllTests
        Case "btnSettings":       SettingsAction
        ' Other buttons still smoke-tested
        Case Else
            MsgBox T("ribbon." & control.ID, control.ID), vbInformation, "Banyan - " & control.ID
    End Select
End Sub


' --- EnsureRibbonI10n - Declare ribbon label translations (matches WPS) ---

Private Sub EnsureRibbonI10n()
    If m_ribbonI10nReady Then Exit Sub

    ' -- Chinese (Simplified) -----------------------------------------------
    I10nRegisterTable msoLanguageIDSimplifiedChinese, "ribbon", _
        "btnCitation",       "插入/编辑引注", _
        "btnChapterBreak",   "插入分隔符", _
        "btnBibliography",   "插入/编辑书目", _
        "btnCitationPane",   "打开引注窗格", _
        "btnRefresh",        "刷新", _
        "btnConvert",        "转换 Zotero 域", _
        "btnUnlink",         "定稿", _
        "btnTest", "运行测试", _
        "btnSettings",       "设置"

    ' -- English (United States) --------------------------------------------
    I10nRegisterTable msoLanguageIDEnglishUS, "ribbon", _
        "btnCitation",       "Insert/Edit Citation", _
        "btnChapterBreak",   "Insert Break", _
        "btnBibliography",   "Insert/Edit Bibliography", _
        "btnCitationPane",   "Open Citation Pane", _
        "btnRefresh",        "Refresh", _
        "btnConvert",        "Convert Zotero Fields", _
        "btnUnlink",         "Finalize", _
        "btnTest", "Run Tests", _
        "btnSettings",       "Preferences"

    m_ribbonI10nReady = True
End Sub
