Option Explicit

' ============================================================================
' Module  : modTest
' Purpose : Concise in-editor test runner for Banyan modules.
'
'           The full per-module test suites live under test/ (modules named
'           test*, each exposing RunTests() -> report string). During
'           development run the FULL suite via test\Run-Tests.ps1: it injects
'           all src + test modules through COM and requires every test to pass
'           (the equivalent of lint + mocha in a TS project).
'
'           This module only aggregates a CONCISE smoke set pulled from the
'           test modules (fast, pure, no document mutation, no backend) and
'           shows the result in a MsgBox from the dev test button.
'
'           The test button (btnTest) is visible only when
'           IsDevMode = True.
'
'           DEV_MODE defaults to False (release build: test button hidden and
'           the dev error policy disabled).  For development/testing do NOT
'           hand-edit this constant - build a DEBUG template instead:
'             Import-BanyanDotm.ps1 -DevMode   flips DEV_MODE to True in the
'                                              imported copy at build time
'                                              (this source file stays False)
'           test\Run-Tests.ps1 (the dev gate) also always imports modTest.bas
'           in Debug mode so the dev-mode error policy stays active in tests.
'
' Public API:
'   RunAllTests()       -> MsgBox summary of the concise test set
'   IsDevMode()         -> Boolean  (drives btnTest visibility + DiagnosticsReraiseIfDev)
' ============================================================================

' --- Development mode toggle (release default: False) ---
' Do not edit by hand: the build tooling rewrites this line (see the shared
' Set-VbeDevModeConstant helper in VbeEncodingHelpers.ps1).
Private Const DEV_MODE As Boolean = False


' --- IsDevMode   Controls button visibility via OnGetVisible ---

Public Function IsDevMode() As Boolean
    IsDevMode = DEV_MODE
End Function


' --- RunAllTests - Execute the concise smoke set and report results. ---

Public Sub RunAllTests()
    Dim report As String
    report = "Banyan Test Results" & vbCrLf & String(40, "-") & vbCrLf

    ' -- Concise smoke set pulled from the test modules ---------------------
    ' Fast, pure checks (no document mutation, no backend). The full suite
    ' (including testField / testHttp / testPreference) runs via
    ' test\Run-Tests.ps1.
    report = report & testI10n.RunTests() & vbCrLf
    report = report & testJson.RunTests() & vbCrLf
    report = report & testDict.RunTests() & vbCrLf

    ' -- Tally results ------------------------------------------------------
    Dim passed As Long
    Dim failed As Long
    Dim total As Long
    Dim line As Variant
    For Each line In Split(report, vbCrLf)
        If InStr(1, CStr(line), "[PASS]") > 0 Then
            passed = passed + 1
        ElseIf InStr(1, CStr(line), "[FAIL]") > 0 Then
            failed = failed + 1
        End If
    Next line
    total = passed + failed

    report = report & vbCrLf & String(40, "-") & vbCrLf
    report = report & "Total: " & total & "  |  "
    report = report & "PASS: " & passed & "  |  "
    report = report & "FAIL: " & failed

    If failed = 0 Then
        MsgBox report, vbInformation, "All Tests Passed"
    Else
        MsgBox report, vbExclamation, "Tests: " & failed & " failure(s)"
    End If
End Sub
