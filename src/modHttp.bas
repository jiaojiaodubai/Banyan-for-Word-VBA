Option Explicit

' ============================================================================
' Module  : modHttp
' Purpose : Cross-platform HTTP client for Banyan - minimal, synchronous.
'
'   Windows - COM late-bound (MSXML2.ServerXMLHTTP.6.0)
'   macOS   - curl via shared ShellExec - no external scripts required
'
' Public API:
'   HttpPost(url, body, headers)  -> response text ("" on error)
'   HttpGet(url)                  -> response text ("" on error)
'   HttpGetPort()                 -> fixed local server port
'   HttpGetBaseUrl()              -> "http://localhost:<port>/banyan"
'   IsMac()                       -> platform detection
'
' Cross-platform approach inspired by VBA-Web:
'   - WinHTTP error -> offline-detection pattern (ref: WebClient.cls)
'   - popen + curl for Mac (ref: WebHelpers.bas ExecuteInShell)
'     Copyright (c) Tim Hall, MIT license
'   - curl exit code -> WinHTTP error mapping table (ref: WebClient.cls)
'   - No external scripts needed on either platform.
' ============================================================================

' --- HTTP defaults ---
Private Const BASE_PATH     As String = "/banyan"
Private Const DEFAULT_ZOTERO_PORT As String = "23119"
Private Const FALLBACK_ZOTERO_PORT As String = "23124"
Private Const HTTP_RESOLVE_TIMEOUT_MS As Long = 10000
Private Const HTTP_CONNECT_TIMEOUT_MS As Long = 10000
Private Const HTTP_SEND_TIMEOUT_MS    As Long = 30000
Private Const HTTP_RECEIVE_TIMEOUT_MS As Long = 300000

Private m_lastError As String
Private m_lastTransportError As Boolean
Private m_activePort As String


' --- Platform detection ---

Public Function IsMac() As Boolean
    IsMac = (InStr(1, Application.System.OperatingSystem, "Mac", vbTextCompare) > 0)
End Function


' --- HttpPost - Send a POST request. ---
' url     - Full URL (e.g. "http://localhost:23119/banyan/citation")
' body    - Request body (JSON string, or "" for no body)
' headers - Optional late-bound Dictionary of extra header name->value pairs
' Default headers (Content-Type, Zotero-Allowed-Request,
' X-Banyan-Client) are always set.

Public Function HttpPost(ByVal url As String, _
                         Optional ByVal body As String = "", _
                         Optional ByVal headers As Variant) As String
    m_lastError = ""
    m_lastTransportError = False
    HttpPost = HttpPostOnce(url, body, headers)

    If m_lastTransportError Then
        Dim retryUrl As String
        retryUrl = AlternatePortUrl(url)
        If Len(retryUrl) > 0 Then
            m_lastError = ""
            m_lastTransportError = False
            HttpPost = HttpPostOnce(retryUrl, body, headers)
            If Not m_lastTransportError Then RememberPort retryUrl
        End If
    Else
        RememberPort url
    End If
End Function

Private Function HttpPostOnce(ByVal url As String, _
                              ByVal body As String, _
                              ByVal headers As Variant) As String
    If IsMac() Then
        HttpPostOnce = HttpPostMac(url, body, headers)
    Else
        HttpPostOnce = HttpPostWin(url, body, headers)
    End If
End Function


' --- HttpGet - Send a GET request. ---

Public Function HttpGet(ByVal url As String) As String
    m_lastError = ""
    m_lastTransportError = False
    HttpGet = HttpGetOnce(url)

    If m_lastTransportError Then
        Dim retryUrl As String
        retryUrl = AlternatePortUrl(url)
        If Len(retryUrl) > 0 Then
            m_lastError = ""
            m_lastTransportError = False
            HttpGet = HttpGetOnce(retryUrl)
            If Not m_lastTransportError Then RememberPort retryUrl
        End If
    Else
        RememberPort url
    End If
End Function

Private Function HttpGetOnce(ByVal url As String) As String
    If IsMac() Then
        HttpGetOnce = HttpGetMac(url)
    Else
        HttpGetOnce = HttpGetWin(url)
    End If
End Function


' --- Windows implementation  (MSXML2.ServerXMLHTTP.6.0, late-bound) ---

Private Function HttpPostWin(ByVal url As String, _
                              ByVal body As String, _
                              ByVal headers As Variant) As String
    On Error GoTo ErrHandler

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.setTimeouts HTTP_RESOLVE_TIMEOUT_MS, HTTP_CONNECT_TIMEOUT_MS, HTTP_SEND_TIMEOUT_MS, HTTP_RECEIVE_TIMEOUT_MS
    http.Open "POST", CleanUrl(url), False

    ' Default headers
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Zotero-Allowed-Request", "1"
    http.setRequestHeader "X-Banyan-Client", "Banyan for Word VBA"

    ApplyHeaders http, headers

    http.send body

    If http.status < 200 Or http.status >= 300 Then
        m_lastError = HttpStatusText("POST", CleanUrl(url), CLng(http.status), CStr(http.statusText), CStr(http.responseText))
        HttpPostWin = ""
        Set http = Nothing
        Exit Function
    End If

    HttpPostWin = http.responseText
    Set http = Nothing
    Exit Function
ErrHandler:
    m_lastTransportError = True
    m_lastError = HttpErrorText("POST", url, Err.Number, Err.Source, Err.Description)
    HttpPostWin = ""
End Function


Private Function HttpGetWin(ByVal url As String) As String
    On Error GoTo ErrHandler

    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")

    http.setTimeouts HTTP_RESOLVE_TIMEOUT_MS, HTTP_CONNECT_TIMEOUT_MS, HTTP_SEND_TIMEOUT_MS, HTTP_RECEIVE_TIMEOUT_MS
    http.Open "GET", CleanUrl(url), False

    http.send

    If http.status < 200 Or http.status >= 300 Then
        m_lastError = HttpStatusText("GET", CleanUrl(url), CLng(http.status), CStr(http.statusText), CStr(http.responseText))
        HttpGetWin = ""
        Set http = Nothing
        Exit Function
    End If

    HttpGetWin = http.responseText
    Set http = Nothing
    Exit Function
ErrHandler:
    m_lastTransportError = True
    m_lastError = HttpErrorText("GET", url, Err.Number, Err.Source, Err.Description)
    HttpGetWin = ""
End Function


' --- macOS implementation  (curl via popen - no external scripts) ---
'
' Inspired by VBA-Web WebHelpers.bas / WebClient.cls
' Copyright (c) Tim Hall, MIT license

Private Function HttpPostMac(ByVal url As String, _
                              ByVal body As String, _
                              ByVal headers As Variant) As String
    ' Build curl command and execute via popen.
    ' The JSON body is passed inline (single-quoted via QuoteArg) so no temp
    ' file is written - /tmp lies outside Word's macOS sandbox container, and
    ' writing there pops the system "grant folder access" dialog on first use.
    Dim cmd As String
    Dim result As String

    ' Build curl command
    cmd = "curl -s -X POST " & QuoteArg(url) & _
          " -H 'Content-Type: application/json'" & _
          " -H 'Zotero-Allowed-Request: 1'" & _
          " -H 'X-Banyan-Client: Banyan for Word VBA'" & _
          " -d " & QuoteArg(body) & _
          " --connect-timeout 10 --max-time 300"

    ' Append extra headers (read via modDict - safe accessor, see README)
    On Error Resume Next
    If Not IsMissing(headers) Then
        If Not IsEmpty(headers) Then
            If Not headers Is Nothing Then
                Dim hdrs As Object
                Set hdrs = headers
                If Not hdrs Is Nothing Then
                    Dim key As Variant
                    For Each key In hdrs.Keys
                        cmd = cmd & " -H '" & CStr(key) & ": " & DictKeyString(hdrs, CStr(key)) & "'"
                    Next key
                End If
            End If
        End If
    End If
    On Error GoTo 0

    ' Pipe the response through base64 so non-ASCII text is not mangled across
    ' the shell boundary; decode back to Unicode in VBA.
    result = HttpDecodeUtf8B64(ShellExec(cmd & " | base64"))
    If Len(result) = 0 Then
        m_lastTransportError = True
        m_lastError = "POST " & url & vbCrLf & "No response from local server."
    End If

    HttpPostMac = result
End Function


Private Function HttpGetMac(ByVal url As String) As String
    Dim cmd As String
    ' Pipe the response through base64 so non-ASCII text is not mangled across
    ' the shell boundary; decode back to Unicode in VBA.
    cmd = "curl -s " & QuoteArg(url) & " --connect-timeout 10 --max-time 300 | base64"
    HttpGetMac = HttpDecodeUtf8B64(ShellExec(cmd))
    If Len(HttpGetMac) = 0 Then
        m_lastTransportError = True
        m_lastError = "GET " & url & vbCrLf & "No response from local server."
    End If
End Function


' --- UTF-8-safe shell exchange on Mac ---------------------------------------
' On Mac, non-ASCII VBA strings are not reliably UTF-8 across the shell
' (Declare/popen) boundary, so text can be corrupted in transit.  Keep the
' shell channel pure ASCII and convert explicitly in VBA:
'   - responses are piped through `base64` (ASCII) and decoded below;
'   - request bodies are already ASCII-only because JsonStringify escapes
'     every non-ASCII char as \uXXXX (see JsonConverter.json_Encode), so
'     nothing non-ASCII ever crosses the shell.

Public Function HttpDecodeUtf8B64(ByVal b64 As String) As String
    ' Decode base64 (of a UTF-8 payload) into a proper Unicode VBA string.
    If Len(b64) = 0 Then Exit Function
    HttpDecodeUtf8B64 = Utf8ToUnicode(B64Decode(b64))
End Function

Private Function B64Decode(ByVal s As String) As String
    ' base64 -> raw bytes, returned as a VBA string holding one byte per char
    ' (each char code equals the byte value 0..255).
    Const ALPHABET As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
    Dim clean As String
    Dim i As Long
    Dim j As Long
    Dim c As String
    Dim idx As Long
    Dim quad As String
    Dim vals(0 To 3) As Long
    Dim out As String
    Dim b As Long

    clean = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If InStr(1, ALPHABET & "=", c, vbBinaryCompare) > 0 Then clean = clean & c
    Next i
    Do While Len(clean) Mod 4 <> 0
        clean = clean & "="
    Loop

    out = ""
    For i = 1 To Len(clean) Step 4
        quad = Mid$(clean, i, 4)
        For j = 0 To 3
            c = Mid$(quad, j + 1, 1)
            If c = "=" Then
                vals(j) = -1
            Else
                idx = InStr(1, ALPHABET, c, vbBinaryCompare)
                vals(j) = IIf(idx = 0, -1, idx - 1)
            End If
        Next j
        If vals(0) < 0 Or vals(1) < 0 Then GoTo NextQuad
        b = (vals(0) * 4) Or (vals(1) \ 16)                 ' byte 0
        out = out & ChrW(b)
        If vals(2) >= 0 Then
            b = ((vals(1) And 15) * 16) Or (vals(2) \ 4)    ' byte 1
            out = out & ChrW(b)
        End If
        If vals(3) >= 0 Then
            b = ((vals(2) And 3) * 64) Or vals(3)            ' byte 2
            out = out & ChrW(b)
        End If
NextQuad:
    Next i
    B64Decode = out
End Function

Private Function Utf8ToUnicode(ByVal bytes As String) As String
    ' Decode a raw UTF-8 byte string (one byte per char) into Unicode.
    Dim out As String
    Dim i As Long
    Dim n As Long
    Dim b0 As Long, b1 As Long, b2 As Long, b3 As Long
    Dim cp As Long

    out = ""
    n = Len(bytes)
    i = 1
    Do While i <= n
        b0 = AscW(Mid$(bytes, i, 1))
        i = i + 1
        If b0 < 128 Then
            out = out & ChrW(b0)
        ElseIf b0 < 224 Then
            b1 = ByteAt(bytes, i): i = i + 1
            cp = ((b0 And 31) * 64) Or (b1 And 63)
            out = out & ChrW(cp)
        ElseIf b0 < 240 Then
            b1 = ByteAt(bytes, i): i = i + 1
            b2 = ByteAt(bytes, i): i = i + 1
            cp = ((b0 And 15) * 4096) Or ((b1 And 63) * 64) Or (b2 And 63)
            out = out & ChrW(cp)
        Else
            b1 = ByteAt(bytes, i): i = i + 1
            b2 = ByteAt(bytes, i): i = i + 1
            b3 = ByteAt(bytes, i): i = i + 1
            cp = ((b0 And 7) * 262144) Or ((b1 And 63) * 4096) Or _
                 ((b2 And 63) * 64) Or (b3 And 63)
            cp = cp - &H10000
            out = out & ChrW(&HD800 + (cp \ 1024)) & ChrW(&HDC00 + (cp And 1023))
        End If
    Loop
    Utf8ToUnicode = out
End Function

Private Function ByteAt(ByVal s As String, ByVal i As Long) As Long
    If i <= Len(s) Then ByteAt = AscW(Mid$(s, i, 1))
End Function


' --- QuoteArg - Wrap a string in single-quotes for shell safety. ---

Private Function QuoteArg(ByVal s As String) As String
    ' Escape any internal single quotes: ' -> '\''
    s = Replace(s, "'", "'\''")
    QuoteArg = "'" & s & "'"
End Function


' --- ApplyHeaders - Iterate a Dictionary and set headers on the HTTP object. ---

Private Sub ApplyHeaders(ByRef http As Object, ByVal headers As Variant)
    If IsMissing(headers) Then Exit Sub
    If IsEmpty(headers) Then Exit Sub
    If headers Is Nothing Then Exit Sub

    ' Read via modDict (safe accessor, see README "Dictionary safety").
    Dim d As Object
    Set d = headers
    If d Is Nothing Then Exit Sub

    On Error Resume Next
    Dim key As Variant
    For Each key In d.Keys
        http.setRequestHeader CStr(key), DictKeyString(d, CStr(key))
    Next key
    On Error GoTo 0
End Sub


' --- HttpGetPort - Returns the active Zotero local server port. ---

Public Function HttpGetPort() As String
    If Len(m_activePort) = 0 Then
        HttpGetPort = DEFAULT_ZOTERO_PORT
    Else
        HttpGetPort = m_activePort
    End If
End Function

Private Function AlternatePortUrl(ByVal url As String) As String
    Dim primaryMarker As String
    Dim fallbackMarker As String
    primaryMarker = ":" & DEFAULT_ZOTERO_PORT & "/"
    fallbackMarker = ":" & FALLBACK_ZOTERO_PORT & "/"

    If InStr(1, url, primaryMarker, vbTextCompare) > 0 Then
        AlternatePortUrl = Replace(url, primaryMarker, fallbackMarker, 1, 1, vbTextCompare)
    ElseIf InStr(1, url, fallbackMarker, vbTextCompare) > 0 Then
        AlternatePortUrl = Replace(url, fallbackMarker, primaryMarker, 1, 1, vbTextCompare)
    End If
End Function

Private Sub RememberPort(ByVal url As String)
    If InStr(1, url, ":" & FALLBACK_ZOTERO_PORT & "/", vbTextCompare) > 0 Then
        m_activePort = FALLBACK_ZOTERO_PORT
    ElseIf InStr(1, url, ":" & DEFAULT_ZOTERO_PORT & "/", vbTextCompare) > 0 Then
        m_activePort = DEFAULT_ZOTERO_PORT
    End If
End Sub


' --- HttpGetBaseUrl - Returns  "http://localhost:<port>/banyan" ---

Public Function HttpGetBaseUrl() As String
    HttpGetBaseUrl = "http://localhost:" & HttpGetPort() & BASE_PATH
End Function


' --- HttpBuildUrl - Append a route path to the Banyan API base URL. ---

Public Function HttpBuildUrl(ByVal routePath As String) As String
    If Len(routePath) = 0 Then
        HttpBuildUrl = HttpGetBaseUrl()
    ElseIf Left$(routePath, 1) = "/" Then
        HttpBuildUrl = HttpGetBaseUrl() & routePath
    Else
        HttpBuildUrl = HttpGetBaseUrl() & "/" & routePath
    End If
End Function


' --- HttpGetLastError - Diagnostic detail for the last failed request. ---

Public Function HttpGetLastError() As String
    HttpGetLastError = m_lastError
End Function


Private Function HttpStatusText(ByVal method As String, _
                                ByVal url As String, _
                                ByVal statusCode As Long, _
                                ByVal statusText As String, _
                                ByVal responseText As String) As String
    HttpStatusText = method & " " & url & vbCrLf & _
                     "HTTP status: " & CStr(statusCode) & " " & NonEmpty(statusText, "(empty)")
    If Len(responseText) > 0 Then
        HttpStatusText = HttpStatusText & vbCrLf & _
                         "Response: " & Left$(responseText, 500)
    End If
End Function

Private Function CleanUrl(ByVal url As String) As String
    CleanUrl = Trim$(url)
    CleanUrl = Replace(CleanUrl, vbCr, "")
    CleanUrl = Replace(CleanUrl, vbLf, "")
    CleanUrl = Replace(CleanUrl, vbTab, "")
End Function

Private Function HttpErrorText(ByVal method As String, _
                               ByVal url As String, _
                               ByVal errNumber As Long, _
                               ByVal errSource As String, _
                               ByVal errDescription As String) As String
    HttpErrorText = method & " " & url & vbCrLf & _
                    "Error number: " & CStr(errNumber) & vbCrLf & _
                    "Low word: " & CStr(LowWord(errNumber)) & vbCrLf & _
                    "Source: " & NonEmpty(errSource, "(none)") & vbCrLf & _
                    "Description: " & NonEmpty(errDescription, "(empty)")
End Function

Private Function LowWord(ByVal value As Long) As Long
    LowWord = value Mod 65536
    If LowWord < 0 Then LowWord = LowWord + 65536
End Function

Private Function NonEmpty(ByVal value As String, ByVal fallback As String) As String
    If Len(value) > 0 Then
        NonEmpty = value
    Else
        NonEmpty = fallback
    End If
End Function
