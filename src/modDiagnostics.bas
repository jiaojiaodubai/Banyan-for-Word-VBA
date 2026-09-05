Option Explicit

' ============================================================================
' Module  : modDiagnostics
' Purpose : Shared diagnostic MessageBox helpers.
'
'           Callers provide the user-facing action message and optional context;
'           this module appends consistent, localized runtime error details.
' ============================================================================

Private m_i18nReady As Boolean


' --- Dev-mode error policy -------------------------------------------------
' When the dev switch (modTest.IsDevMode, the DEV_MODE constant) is ON, re-raise
' the current error so error handlers stop swallowing it - the VBE breaks /
' callers see the exact business failure point. No-op in production builds, so
' all existing best-effort fallbacks keep working.
'
' Call this as the FIRST statement of any ErrHandler that should swallow in
' production but throw in dev mode:
'     ErrHandler:
'         DiagnosticsReraiseIfDev "modX.FuncName"
'         <fallback assignment / exit>
'     End Sub
'
' Optional context: pass "modX.FuncName" of the enclosing procedure. The FIRST
' (deepest) re-raiser's identity is appended to the re-raised Err.Description as
' "[deepest: modX.FuncName]"; shallower re-raises pass it through unchanged, so
' the origin survives to the VBE break / COM caller (the re-raise chain otherwise
' hides it - especially on Mac, where the VBE has no "Break on All Errors"
' option).

Public Sub DiagnosticsReraiseIfDev(Optional ByVal context As String = "")
    If Not modTest.IsDevMode() Then Exit Sub
    If Err.Number = 0 Then Exit Sub
    ' Handlers run deepest-first, so the FIRST call appends the closest-to-
    ' origin identity; shallower re-raises see "[deepest:" already in the
    ' description and pass it through unchanged. Stateless by design.
    Dim desc As String
    desc = Err.Description
    If InStr(1, desc, "[deepest:", vbTextCompare) = 0 Then
        If Len(context) > 0 Then
            desc = desc & " [deepest: " & context & "]"
        End If
    End If
    Err.Raise Err.Number, Err.Source, desc
End Sub


' --- DiagnosticShowMessage - Generic message + optional diagnostic detail. ---

Public Sub DiagnosticShowMessage(ByVal title As String, _
                                 ByVal message As String, _
                                 Optional ByVal detail As String = "", _
                                 Optional ByVal style As VbMsgBoxStyle = vbInformation)
    MsgBox BuildPrompt(message, detail), style, title
End Sub


' --- DiagnosticShowError - MessageBox with standard Err details. ---

Public Sub DiagnosticShowError(ByVal title As String, _
                               ByVal message As String, _
                               ByVal errNumber As Long, _
                               ByVal errSource As String, _
                               ByVal errDescription As String, _
                               Optional ByVal errLine As Long = 0, _
                               Optional ByVal detail As String = "", _
                               Optional ByVal style As VbMsgBoxStyle = vbExclamation)
    Dim fullDetail As String
    fullDetail = detail
    If Len(fullDetail) > 0 Then fullDetail = fullDetail & vbCrLf
    fullDetail = fullDetail & DiagnosticErrorText(errNumber, errSource, errDescription, errLine)

    DiagnosticShowMessage title, message, fullDetail, style
End Sub


' --- DiagnosticErrorText - Localized text block for Err details. ---

Public Function DiagnosticErrorText(ByVal errNumber As Long, _
                                    ByVal errSource As String, _
                                    ByVal errDescription As String, _
                                    Optional ByVal errLine As Long = 0) As String
    EnsureDiagnosticsI10n

    DiagnosticErrorText = DText("errNumber", "Error number") & ": " & CStr(errNumber) & vbCrLf & _
                          DText("errSource", "Source") & ": " & NonEmptyText(errSource, DText("none", "(none)")) & vbCrLf & _
                          DText("errDescription", "Description") & ": " & NonEmptyText(errDescription, DText("empty", "(empty)")) & vbCrLf & _
                          DText("errLine", "Line") & ": " & CStr(errLine) & vbCrLf & _
                          DText("langId", "UI language") & ": " & CStr(I10nGetCurrentLangID())
End Function


' --- Local i18n ---

Private Sub EnsureDiagnosticsI10n()
    If m_i18nReady Then Exit Sub

    I10nRegisterTable msoLanguageIDSimplifiedChinese, "diagnostics", _
        "errNumber", "错误号", _
        "errSource", "来源", _
        "errDescription", "说明", _
        "errLine", "行号", _
        "langId", "界面语言", _
        "none", "（无）", _
        "empty", "（空）"

    I10nRegisterTable msoLanguageIDEnglishUS, "diagnostics", _
        "errNumber", "Error number", _
        "errSource", "Source", _
        "errDescription", "Description", _
        "errLine", "Line", _
        "langId", "UI language", _
        "none", "(none)", _
        "empty", "(empty)"

    m_i18nReady = True
End Sub

Private Function DText(ByVal key As String, ByVal fallback As String) As String
    EnsureDiagnosticsI10n
    DText = T("diagnostics." & key, fallback)
End Function

Private Function BuildPrompt(ByVal message As String, ByVal detail As String) As String
    BuildPrompt = message
    If Len(detail) > 0 Then
        BuildPrompt = BuildPrompt & vbCrLf & detail
    End If
End Function

Private Function NonEmptyText(ByVal value As String, ByVal fallback As String) As String
    If Len(value) > 0 Then
        NonEmptyText = value
    Else
        NonEmptyText = fallback
    End If
End Function
