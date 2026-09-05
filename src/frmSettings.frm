VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSettings 
   Caption         =   "Banyan Preferences"
   ClientHeight    =   6104
   ClientLeft      =   91
   ClientTop       =   406
   ClientWidth     =   8421.001
   OleObjectBlob   =   "frmSettings.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_pref As Object
Private m_style As Object
Private m_saved As Boolean
Private m_i18nReady As Boolean

Public Function EditPreference(ByVal pref As Object) As Boolean
    EnsureSettingsI10n

    Set m_pref = pref
    Set m_style = CloneStyle(pref("style"))

    ApplyLabels
    LoadPreference

    Me.Show vbModal
    EditPreference = m_saved
End Function

Private Sub UserForm_Initialize()
    EnsureSettingsI10n
    ApplyLabels
End Sub

Private Sub cmdFetchStyle_Click()
    Dim originalCaption As String
    originalCaption = cmdFetchStyle.Caption

    cmdFetchStyle.Enabled = False
    cmdFetchStyle.Caption = SettingsText("fetchingStyle", "Loading...")
    lblStatus.Caption = SettingsText("fetchingStyle", "Loading...")
    DoEvents

    On Error GoTo ErrHandler
    Dim style As Object
    Set style = PreferenceFetchStyle(m_style)
    If style Is Nothing Then
        MsgBox SettingsText("styleError", "Failed to fetch style."), vbExclamation, Me.Caption
    Else
        Set m_style = CloneStyle(style)
        RenderStyle
        lblStatus.Caption = SettingsText("styleUpdated", "Style updated.")
    End If

CleanUp:
    cmdFetchStyle.Caption = originalCaption
    cmdFetchStyle.Enabled = True
    Exit Sub
ErrHandler:
    MsgBox SettingsText("styleError", "Failed to fetch style.") & vbCrLf & Err.Description, vbExclamation, Me.Caption
    Resume CleanUp
End Sub

Private Sub cmdClearBibTitle_Click()
    txtBibTitleStyle.Text = ""
End Sub

Private Sub cmdClearBibEntry_Click()
    txtBibEntryStyle.Text = ""
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ErrHandler

    If m_style Is Nothing Then
        MsgBox SettingsText("styleRequired", "Please choose a Banyan style first."), vbExclamation, Me.Caption
        Exit Sub
    End If

    m_pref("syncItems") = CBool(chkSyncItems.Value)
    m_pref("refreshAll") = CBool(chkRefreshAll.Value)
    Set m_pref("style") = m_style
    m_pref("bibliographyTitleStyle") = Trim$(txtBibTitleStyle.Text)
    m_pref("bibliographyEntryStyle") = Trim$(txtBibEntryStyle.Text)

    If Not m_pref.Exists("extraSource") Then
        m_pref("extraSource") = Null
    End If

    If Not PreferenceSave(m_pref) Then
        DiagnosticShowMessage Me.Caption, _
                              SettingsText("saveError", "Failed to save preferences."), _
                              PreferenceGetLastSaveError(), _
                              vbExclamation
        Exit Sub
    End If

    m_saved = True
    Me.Hide
    Exit Sub
ErrHandler:
    MsgBox SettingsText("saveError", "Failed to save preferences.") & vbCrLf & Err.Description, vbExclamation, Me.Caption
End Sub

Private Sub cmdCancel_Click()
    m_saved = False
    Me.Hide
End Sub

Private Sub LoadPreference()
    chkSyncItems.Value = CBool(m_pref("syncItems"))
    chkRefreshAll.Value = CBool(m_pref("refreshAll"))
    txtBibTitleStyle.Text = CStr(m_pref("bibliographyTitleStyle"))
    txtBibEntryStyle.Text = CStr(m_pref("bibliographyEntryStyle"))
    lblStatus.Caption = ""
    RenderStyle
End Sub

Private Sub RenderStyle()
    If m_style Is Nothing Then
        txtStyleTitle.Text = ""
        txtStyleTitle.ControlTipText = SettingsText("styleUnset", "No style selected")
        Exit Sub
    End If

    Dim title As String
    Dim stId As String
    Dim citationType As String
    On Error Resume Next
    title = CStr(m_style("title"))
    stId = CStr(m_style("id"))
    citationType = CStr(m_style("citationType"))
    On Error GoTo 0

    txtStyleTitle.Text = title
    txtStyleTitle.ControlTipText = "id=" & stId & "; citationType=" & citationType
End Sub

Private Function CloneStyle(ByVal source As Variant) As Object
    Dim result As Object
    Set result = New Dictionary
    ' Read via modDict (safe accessor, see README "Dictionary safety").
    result("id") = DictKeyString(source, "id")
    result("title") = DictKeyString(source, "title")
    result("citationType") = DictKeyString(source, "citationType")
    Set CloneStyle = result
End Function

Private Sub ApplyLabels()
    Me.Caption = SettingsText("dialogTitle", "Banyan Preferences")
    fraGlobal.Caption = SettingsText("global", "Global")
    fraChapter.Caption = SettingsText("section", "Section")
    chkSyncItems.Caption = SettingsText("syncItems", "Sync item metadata when refreshing citations")
    chkRefreshAll.Caption = SettingsText("refreshAll", "Refresh all chapters")
    lblStyle.Caption = SettingsText("style", "Banyan style")
    cmdFetchStyle.Caption = SettingsText("fetchStyle", "Choose...")
    lblBibTitleStyle.Caption = SettingsText("bibTitleStyle", "Bibliography title")
    lblBibEntryStyle.Caption = SettingsText("bibEntryStyle", "Bibliography entry")
    cmdClearBibTitle.Caption = SettingsText("clear", "Clear")
    cmdClearBibEntry.Caption = SettingsText("clear", "Clear")
    cmdOK.Caption = SettingsText("ok", "OK")
    cmdCancel.Caption = SettingsText("cancel", "Cancel")
End Sub

Private Function SettingsText(ByVal key As String, ByVal fallback As String) As String
    EnsureSettingsI10n
    SettingsText = T("settingsForm." & key, fallback)
End Function

Private Sub EnsureSettingsI10n()
    If m_i18nReady Then Exit Sub

    I10nRegisterTable msoLanguageIDSimplifiedChinese, "settingsForm", _
        "dialogTitle",       "Banyan 设置", _
        "global",            "全局设置", _
        "section",           "章节设置", _
        "syncItems",         "刷新时同步条目元数据", _
        "refreshAll",        "刷新全部章节（否则仅当前章节）", _
        "style",             "引注样式（Banyan 样式）", _
        "styleUnset",        "未选择样式", _
        "fetchStyle",        "选择...", _
        "fetchingStyle",     "请求中...", _
        "styleUpdated",      "样式已更新。", _
        "styleRequired",     "请先选择 Banyan 样式。", _
        "styleError",        "获取样式失败。", _
        "bibTitleStyle",     "文献列表标题（Word 样式）", _
        "bibEntryStyle",     "文献列表题录（Word 样式）", _
        "clear",             "清除", _
        "ok",                "确定", _
        "cancel",            "取消", _
        "saveError",         "保存设置失败。"

    I10nRegisterTable msoLanguageIDEnglishUS, "settingsForm", _
        "dialogTitle",       "Banyan Preferences", _
        "global",            "Global", _
        "section",           "Section", _
        "syncItems",         "Sync item metadata when refreshing citations", _
        "refreshAll",        "Refresh all chapters (otherwise current chapter only)", _
        "style",             "Banyan style", _
        "styleUnset",        "No style selected", _
        "fetchStyle",        "Choose...", _
        "fetchingStyle",     "Loading...", _
        "styleUpdated",      "Style updated.", _
        "styleRequired",     "Please choose a Banyan style first.", _
        "styleError",        "Failed to fetch style.", _
        "bibTitleStyle",     "Bibliography title (Word style)", _
        "bibEntryStyle",     "Bibliography entry (Word style)", _
        "clear",             "Clear", _
        "ok",                "OK", _
        "cancel",            "Cancel", _
        "saveError",         "Failed to save preferences."

    m_i18nReady = True
End Sub
