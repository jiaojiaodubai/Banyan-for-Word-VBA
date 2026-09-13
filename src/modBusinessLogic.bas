Option Explicit

' ============================================================================
' Module  : modBusinessLogic
' Purpose : Settings command and shared ribbon-command helpers.
' ============================================================================

Private m_i18nReady As Boolean


Public Sub SettingsAction()
    EnsureBusinessLogicI10n

    On Error GoTo ErrHandler

    Dim dlg As frmSettings
    Dim pref As Object

    On Error GoTo PreferenceErr
    Set pref = PreferenceEnsure()
    On Error GoTo ErrHandler

    If pref Is Nothing Then
        ShowSettingsOfflineMessage
        Exit Sub
    End If

    Set dlg = New frmSettings
    Dim previousStyle As Object
    Set previousStyle = CloneStyleForRefresh(pref("style"))
    Dim previousTitleStyle As String
    Dim previousEntryStyle As String
    previousTitleStyle = DictKeyString(pref, "bibliographyTitleStyle")
    previousEntryStyle = DictKeyString(pref, "bibliographyEntryStyle")

    If dlg.EditPreference(pref) Then
        RefreshForStyleChange previousStyle, pref("style")
        ApplyBibliographyStyleChange previousTitleStyle, previousEntryStyle, pref
    End If

    Unload dlg
    Exit Sub
ErrHandler:
    DiagnosticsReraiseIfDev "modBusinessLogic.SettingsAction"
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim errLine As Long
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    errLine = Erl

    On Error Resume Next
    If Not dlg Is Nothing Then Unload dlg
    On Error GoTo 0

    DiagnosticShowError T("settingsAction.dialogTitle", "Banyan Preferences"), _
                        T("settingsAction.loadError", "Failed to open preferences."), _
                        errNumber, errSource, errDescription, errLine
    Exit Sub
PreferenceErr:
    Dim prefErrNumber As Long
    Dim prefErrSource As String
    Dim prefErrDescription As String
    Dim prefErrLine As Long
    prefErrNumber = Err.Number
    prefErrSource = Err.Source
    prefErrDescription = Err.Description
    prefErrLine = Erl

    On Error GoTo 0
    DiagnosticShowError T("settingsAction.dialogTitle", "Banyan Preferences"), _
                        T("settingsAction.loadOffline", _
                          "Unable to initialize preferences. Start Zotero and make sure the Banyan plugin is enabled, then try again."), _
                        prefErrNumber, prefErrSource, prefErrDescription, prefErrLine, _
                        SettingsActionOfflineText()
End Sub


' --- Local i18n ---

Private Sub EnsureBusinessLogicI10n()
    If m_i18nReady Then Exit Sub

    I10nRegisterTable msoLanguageIDSimplifiedChinese, "settingsAction", _
        "dialogTitle", "Banyan 设置", _
        "loadError", "打开设置失败。", _
        "loadOffline", "无法初始化设置。请启动 Zotero 客户端并确认 Banyan 插件已启用后重试。", _
        "requestUrl", "请求地址", _
        "port", "端口", _
        "httpError", "HTTP 错误", _
        "langId", "界面语言", _
        "applyingStyles", "正在应用书目样式…"

    I10nRegisterTable msoLanguageIDEnglishUS, "settingsAction", _
        "dialogTitle", "Banyan Preferences", _
        "loadError", "Failed to open preferences.", _
        "loadOffline", "Unable to initialize preferences. Start Zotero and make sure the Banyan plugin is enabled, then try again.", _
        "requestUrl", "Request URL", _
        "port", "Port", _
        "httpError", "HTTP error", _
        "langId", "UI language", _
        "applyingStyles", "Applying bibliography styles..."

    m_i18nReady = True
End Sub

Private Sub ShowSettingsOfflineMessage()
    DiagnosticShowMessage T("settingsAction.dialogTitle", "Banyan Preferences"), _
                          T("settingsAction.loadOffline", _
                            "Unable to initialize preferences. Start Zotero and make sure the Banyan plugin is enabled, then try again."), _
                          SettingsActionOfflineText(), _
                          vbExclamation
End Sub

' Word style names are a local presentation switch: apply them to the existing
' bibliography lines instead of running a refresh.
Private Sub ApplyBibliographyStyleChange(ByVal previousTitleStyle As String, _
                                         ByVal previousEntryStyle As String, _
                                         ByVal pref As Object)
    On Error GoTo ErrHandler

    Dim titleStyle As String
    Dim entryStyle As String
    titleStyle = DictKeyString(pref, "bibliographyTitleStyle")
    entryStyle = DictKeyString(pref, "bibliographyEntryStyle")
    If titleStyle = previousTitleStyle And entryStyle = previousEntryStyle Then Exit Sub

    Dim targetRange As Range
    Set targetRange = GetUpdateRange()
    If targetRange Is Nothing Then Exit Sub

    ProgressOpen T("settingsAction.applyingStyles", "Applying bibliography styles...")
    FieldBeginBatchUpdate
    RestyleBibliography targetRange, pref
    FieldEndBatchUpdate
    ProgressClose
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modBusinessLogic.ApplyBibliographyStyleChange"
    ProgressClose
    FieldEndBatchUpdate
End Sub

Private Function SettingsActionOfflineText() As String
    Dim httpError As String
    httpError = HttpGetLastError()

    SettingsActionOfflineText = T("settingsAction.requestUrl", "Request URL") & ": " & HttpBuildUrl("style") & vbCrLf & _
                                T("settingsAction.port", "Port") & ": " & HttpGetPort() & vbCrLf & _
                                T("settingsAction.langId", "UI language") & ": " & CStr(I10nGetCurrentLangID())
    If Len(httpError) > 0 Then
        SettingsActionOfflineText = SettingsActionOfflineText & vbCrLf & _
                                    T("settingsAction.httpError", "HTTP error") & ":" & vbCrLf & _
                                    httpError
    End If
End Function

Private Function CloneStyleForRefresh(ByVal source As Variant) As Object
    Dim result As Object
    Set result = New Dictionary
    ' Read via modDict (safe accessor, see README "Dictionary safety").
    result("id") = DictKeyString(source, "id")
    result("title") = DictKeyString(source, "title")
    result("citationType") = DictKeyString(source, "citationType")
    Set CloneStyleForRefresh = result
End Function
