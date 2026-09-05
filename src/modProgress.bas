Option Explicit

' ============================================================================
' Module  : modProgress
' Purpose : Best-effort wrapper for backend progress UI.
'
'           Mirrors WPS utils/progress.ts:
'             POST /banyan/progress { action:"open", reason }
'             POST /banyan/progress { action:"close" }
'
'           Calls are intentionally non-fatal. If Zotero/backend is unavailable
'           or the route is unsupported, the wrapped operation still proceeds.
' ============================================================================

Private m_progressDepth As Long


Public Sub ProgressOpen(ByVal reason As String)
    On Error GoTo ErrHandler

    m_progressDepth = m_progressDepth + 1
    If m_progressDepth <> 1 Then Exit Sub

    Dim body As Object
    Set body = New Dictionary
    body("action") = "open"
    body("reason") = reason

    HttpPost HttpBuildUrl("progress"), JsonStringify(body)
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modProgress.ProgressOpen"
End Sub


Public Sub ProgressClose()
    On Error GoTo ErrHandler

    If m_progressDepth <= 0 Then Exit Sub
    m_progressDepth = m_progressDepth - 1
    If m_progressDepth <> 0 Then Exit Sub

    Dim body As Object
    Set body = New Dictionary
    body("action") = "close"

    HttpPost HttpBuildUrl("progress"), JsonStringify(body)
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modProgress.ProgressClose"
    m_progressDepth = 0
End Sub
