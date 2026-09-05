Option Explicit

' ============================================================================
' Module  : testPreference
' Purpose : Tests for modPreference.
'
'           Part of the dev-only test suite under test/ (module name test*).
'           Run the FULL suite via test/Run-Tests.ps1.
'
' Public API:
'   RunTests() -> String report of [PASS]/[FAIL] lines (no dialogs)
' ============================================================================


Public Function RunTests() As String
    Dim report As String
    report = "testPreference" & vbCrLf & String(40, "-") & vbCrLf
    report = report & TestResult("PreferenceGet", TestPrefGet()) & vbCrLf
    report = report & TestResult("PreferenceSave round-trip", TestPrefSave()) & vbCrLf
    report = report & TestResult("Preference structure readable", TestPrefDisplay()) & vbCrLf
    ' Informational summary (not tallied) - replaces the old MessageBox so the
    ' COM automation flow in Run-Tests.ps1 never blocks on a dialog.
    report = report & PreferenceDisplayInfo() & vbCrLf
    RunTests = report
End Function


Private Function TestPrefGet() As Boolean
    ' Without a preference property, PreferenceGet returns Nothing (matching WPS).
    ' With a property (after PreferenceSave), it returns the full structure.
    On Error Resume Next
    Dim pref As Object
    Set pref = PreferenceGet()
    If Err.Number <> 0 Then TestPrefGet = False: Exit Function
    On Error GoTo 0

    ' If no part exists yet, Nothing is the correct return value
    If pref Is Nothing Then
        TestPrefGet = True
        Exit Function
    End If

    ' If a preference exists, verify the structure
    TestPrefGet = pref.Exists("syncItems") And _
                  pref.Exists("refreshAll") And _
                  pref.Exists("style") And _
                  pref.Exists("extraSource") And _
                  pref.Exists("bibliographyTitleStyle") And _
                  pref.Exists("bibliographyEntryStyle")
End Function

Private Function TestPrefSave() As Boolean
    ' Save + re-read should round-trip values.
    ' If no preference property exists yet, PreferenceSave creates one lazily.
    On Error Resume Next
    Dim pref As Object
    Set pref = PreferenceGet()
    If Err.Number <> 0 Then TestPrefSave = False: Exit Function
    On Error GoTo 0

    ' If no part exists, create a fresh one via save
    If pref Is Nothing Then
        Set pref = CreateFreshPreference()
        If Not PreferenceSave(pref) Then TestPrefSave = False: Exit Function
        ' Now re-read - should have a valid preference property
        Set pref = PreferenceGet()
        If pref Is Nothing Then TestPrefSave = False: Exit Function
    End If

    ' Toggle a value and save
    Dim original As Boolean
    original = CBool(pref("syncItems"))
    pref("syncItems") = Not original
    Dim ok As Boolean
    ok = PreferenceSave(pref)
    If Not ok Then TestPrefSave = False: Exit Function

    ' Re-read and verify
    Dim reloaded As Object
    Set reloaded = PreferenceGet()
    Dim restored As Boolean
    restored = (CBool(reloaded("syncItems")) = (Not original))

    ' Restore original
    pref("syncItems") = original
    PreferenceSave pref

    TestPrefSave = restored
End Function

Private Function TestPrefDisplay() As Boolean
    ' Read the current preference and assert its structure is readable.
    ' NOTE: no MessageBox here. Dialogs block the COM automation flow in
    ' Run-Tests.ps1; the formatted summary is appended to the report via
    ' PreferenceDisplayInfo instead.
    On Error Resume Next
    Dim pref As Object
    Set pref = PreferenceGet()
    If Err.Number <> 0 Then TestPrefDisplay = False: Exit Function
    On Error GoTo 0

    ' No preference property yet is a valid state (matches WPS behaviour).
    If pref Is Nothing Then
        TestPrefDisplay = True
        Exit Function
    End If

    TestPrefDisplay = pref.Exists("syncItems") And _
                      pref.Exists("refreshAll") And _
                      pref.Exists("style") And _
                      pref.Exists("extraSource") And _
                      pref.Exists("bibliographyTitleStyle") And _
                      pref.Exists("bibliographyEntryStyle")
End Function

Private Function PreferenceDisplayInfo() As String
    ' Human-readable summary of the current preference for the test report.
    ' Informational only - lines carry no [PASS]/[FAIL] marker, so they are
    ' not tallied. Never shows a dialog.
    On Error Resume Next
    Dim pref As Object
    Set pref = PreferenceGet()
    On Error GoTo 0

    If pref Is Nothing Then
        PreferenceDisplayInfo = "[INFO] no preference property set (valid state)"
        Exit Function
    End If

    Dim msg As String
    msg = "[INFO] === Global Preferences ===" & vbCrLf
    msg = msg & "[INFO]   Sync Items:      " & CStr(pref("syncItems")) & vbCrLf
    msg = msg & "[INFO]   Refresh All:     " & CStr(pref("refreshAll")) & vbCrLf

    Dim style As Object
    Set style = pref("style")
    If Not style Is Nothing Then
        msg = msg & "[INFO]   Style ID:        " & CStr(style("id")) & vbCrLf
        msg = msg & "[INFO]   Style Title:     " & CStr(style("title")) & vbCrLf
        msg = msg & "[INFO]   Citation Type:   " & CStr(style("citationType")) & vbCrLf
    End If

    msg = msg & "[INFO]   Bib Title Style:  " & CStr(pref("bibliographyTitleStyle")) & vbCrLf
    msg = msg & "[INFO]   Bib Entry Style:  " & CStr(pref("bibliographyEntryStyle")) & vbCrLf
    msg = msg & "[INFO]   Extra Source:     " & FormatOptionalJsonValue(pref("extraSource"))
    PreferenceDisplayInfo = msg
End Function


' --- Helpers ---

Private Function CreateFreshPreference() As Object
    ' Build a minimal valid preference for testing
    Dim pref As Object
    Set pref = New Dictionary
    pref("syncItems") = True
    pref("refreshAll") = False

    Dim style As Object
    Set style = GetPrefStyle("test-id", "Test Style", "intext-citation")
    Set pref("style") = style
    pref("extraSource") = Null
    pref("bibliographyTitleStyle") = "Test Bib Title"
    pref("bibliographyEntryStyle") = "Test Bib Entry"
    Set CreateFreshPreference = pref
End Function

Private Function FormatOptionalJsonValue(ByVal value As Variant) As String
    If IsNull(value) Or IsEmpty(value) Then
        FormatOptionalJsonValue = "(none)"
        Exit Function
    End If

    If IsObject(value) Then
        FormatOptionalJsonValue = JsonStringify(value)
        If Len(FormatOptionalJsonValue) = 0 Then
            FormatOptionalJsonValue = "(none)"
        End If
    Else
        FormatOptionalJsonValue = CStr(value)
    End If
End Function

Private Function TestResult(ByVal name As String, ByVal passed As Boolean) As String
    If passed Then
        TestResult = "[PASS] " & name
    Else
        TestResult = "[FAIL] " & name
    End If
End Function
