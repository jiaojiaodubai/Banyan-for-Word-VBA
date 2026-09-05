Option Explicit

' ============================================================================
' Module  : testI10n
' Purpose : Tests for modI10n.
'
'           Part of the dev-only test suite under test/ (module name test*).
'           Run the FULL suite via test/Run-Tests.ps1; modTest.RunAllTests
'           aggregates a concise subset (see modTest).
'
' Public API:
'   RunTests() -> String report of [PASS]/[FAIL] lines (no dialogs)
' ============================================================================

Private m_testI10nReady As Boolean


Public Function RunTests() As String
    Dim report As String
    report = "testI10n" & vbCrLf & String(40, "-") & vbCrLf
    report = report & TestResult("T() lookup", TestI10nLookup()) & vbCrLf
    report = report & TestResult("T() en-US fallback", TestI10nFallback()) & vbCrLf
    report = report & TestResult("T() explicit fallback", TestI10nExplicitFallback()) & vbCrLf
    RunTests = report
End Function


Private Function TestI10nLookup() As Boolean
    ' Verify T() returns a registered value in the current language.
    RegisterTestI10n

    Dim label As String
    label = T("test.lookup", "FALLBACK")

    If I10nIsLangZH() Then
        TestI10nLookup = (label = "测试值")
    Else
        TestI10nLookup = (label = "Test Value")
    End If
End Function

Private Function TestI10nFallback() As Boolean
    ' Verify T() falls back to en-US before the explicit fallback.
    RegisterTestI10n

    Dim label As String
    label = T("test.enOnly", "FALLBACK")
    TestI10nFallback = (label = "English Only")
End Function

Private Function TestI10nExplicitFallback() As Boolean
    ' Verify T() falls back to the fallback arg for a non-existent key.
    Dim label As String
    label = T("test.nonexistent", "FALLBACK")
    TestI10nExplicitFallback = (label = "FALLBACK")
End Function


Private Sub RegisterTestI10n()
    If m_testI10nReady Then Exit Sub

    I10nRegisterTable msoLanguageIDSimplifiedChinese, "test", _
        "lookup", "测试值"

    I10nRegisterTable msoLanguageIDEnglishUS, "test", _
        "lookup", "Test Value", _
        "enOnly", "English Only"

    m_testI10nReady = True
End Sub

Private Function TestResult(ByVal name As String, ByVal passed As Boolean) As String
    If passed Then
        TestResult = "[PASS] " & name
    Else
        TestResult = "[FAIL] " & name
    End If
End Function
