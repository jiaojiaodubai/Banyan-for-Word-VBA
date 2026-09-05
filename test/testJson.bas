Option Explicit

' ============================================================================
' Module  : testJson
' Purpose : Tests for modJson / JsonConverter.
'
'           Part of the dev-only test suite under test/ (module name test*).
'           Run the FULL suite via test/Run-Tests.ps1; modTest.RunAllTests
'           aggregates a concise subset (see modTest).
'
' Public API:
'   RunTests() -> String report of [PASS]/[FAIL] lines (no dialogs)
' ============================================================================


Public Function RunTests() As String
    Dim report As String
    report = "testJson" & vbCrLf & String(40, "-") & vbCrLf
    report = report & TestResult("ParseJson round-trip", TestJsonRoundTrip()) & vbCrLf
    report = report & TestResult("Stringify", TestJsonStringify()) & vbCrLf
    RunTests = report
End Function


Private Function TestJsonRoundTrip() As Boolean
    ' Parse a simple JSON string and verify the result
    On Error Resume Next
    Dim obj As Object
    Set obj = JsonParse("{""a"":1,""b"":""hello""}")
    If Err.Number <> 0 Then TestJsonRoundTrip = False: Exit Function
    On Error GoTo 0

    If obj Is Nothing Then TestJsonRoundTrip = False: Exit Function
    TestJsonRoundTrip = (obj("a") = 1 And obj("b") = "hello")
End Function

Private Function TestJsonStringify() As Boolean
    ' Convert a simple VBA object to JSON and verify
    On Error Resume Next
    Dim d As Object
    Set d = New Dictionary
    d("x") = 42
    d("y") = "test"
    Dim s As String
    s = JsonStringify(d)
    If Err.Number <> 0 Then TestJsonStringify = False: Exit Function
    On Error GoTo 0

    ' Both key orders are valid JSON
    TestJsonStringify = (s = "{""x"":42,""y"":""test""}" Or _
                         s = "{""y"":""test"",""x"":42}")
End Function


Private Function TestResult(ByVal name As String, ByVal passed As Boolean) As String
    If passed Then
        TestResult = "[PASS] " & name
    Else
        TestResult = "[FAIL] " & name
    End If
End Function
