Option Explicit

' ============================================================================
' Module  : testHttp
' Purpose : Tests for modHttp.
'
'           Part of the dev-only test suite under test/ (module name test*).
'           Run the FULL suite via test/Run-Tests.ps1.
'
' Public API:
'   RunTests() -> String report of [PASS]/[FAIL] lines (no dialogs)
' ============================================================================


Public Function RunTests() As String
    Dim report As String
    report = "testHttp" & vbCrLf & String(40, "-") & vbCrLf
    report = report & TestResult("BaseUrl format", TestHttpBaseUrl()) & vbCrLf
    report = report & TestResult("IsMac detection", TestHttpIsMac()) & vbCrLf
    report = report & TestResult("Decode base64+UTF8 Chinese", TestHttpDecodeChinese()) & vbCrLf
    report = report & TestResult("Decode base64+UTF8 emoji pair", TestHttpDecodeEmoji()) & vbCrLf
    report = report & TestResult("JsonStringify body ASCII-only", TestHttpJsonAsciiBody()) & vbCrLf
    RunTests = report
End Function


Private Function TestHttpBaseUrl() As Boolean
    Dim url As String
    url = HttpGetBaseUrl()
    TestHttpBaseUrl = (url Like "http://localhost:*/banyan")
End Function

Private Function TestHttpIsMac() As Boolean
    ' IsMac should return a Boolean (True or False) - always passes
    Dim b As Boolean
    b = IsMac()
    TestHttpIsMac = (b = True Or b = False)
End Function


Private Function TestHttpDecodeChinese() As Boolean
    ' base64("中文" as UTF-8) = 5Lit5paH ; must round-trip to the text.
    TestHttpDecodeChinese = (HttpDecodeUtf8B64("5Lit5paH") = "中文")
End Function

Private Function TestHttpDecodeEmoji() As Boolean
    ' base64("Emoji" U+1F600 as UTF-8) = 8J+YgA== ; decodes to a surrogate pair.
    Dim emoji As String
    emoji = ChrW(&HD83D) & ChrW(&HDE00)
    TestHttpDecodeEmoji = (HttpDecodeUtf8B64("8J+YgA==") = emoji)
End Function

Private Function TestHttpJsonAsciiBody() As Boolean
    ' JsonStringify must escape every non-ASCII char as \uXXXX so request
    ' bodies stay pure ASCII across the Mac shell channel.
    Dim d As Object
    Set d = New Dictionary
    d("text") = "中文测试ABC"
    Dim js As String
    js = LCase$(JsonStringify(d))
    Dim ok As Boolean
    ok = True
    Dim i As Long
    For i = 1 To Len(js)
        If AscW(Mid$(js, i, 1)) > 126 Then
            ok = False
            Exit For
        End If
    Next i
    TestHttpJsonAsciiBody = ok And (InStr(1, js, "\u4e2d\u6587\u6d4b\u8bd5", vbBinaryCompare) > 0)
End Function


Private Function TestResult(ByVal name As String, ByVal passed As Boolean) As String
    If passed Then
        TestResult = "[PASS] " & name
    Else
        TestResult = "[FAIL] " & name
    End If
End Function
