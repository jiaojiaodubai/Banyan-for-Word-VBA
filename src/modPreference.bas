Option Explicit

' ============================================================================
' Module  : modPreference
' Purpose : Read/write Banyan preferences stored in a CustomDocumentProperty.
'
'           Data structures match the WPS version exactly so that documents
'           edited in both Word and WPS remain fully compatible.
'
'           Global + Chapter 1 prefs -> CustomDocumentProperty "BANYAN_PREF"
'           Chapter 2+ overrides     -> ChapterBreak ADDIN fields (modChapterBreak)
'
' Public types (all Dictionary-based):
'   Preference        - { syncItems, refreshAll, style, extraSource,
'                         bibliographyTitleStyle, bibliographyEntryStyle }
'   PrefStyle         - { id, title, citationType }
'   CitationSource    - Dictionary
'
' Public API:
'   PreferenceGet()       -> Dictionary or Nothing
'   PreferenceEnsure()    -> Dictionary or Nothing; creates property if needed
'   PreferenceSave(pref)  -> Boolean success
'   PreferenceFetchStyle(currentStyle) -> Dictionary or Nothing
'   PreferenceInit()      -> no-op retained for API compatibility
'   RemovePreferencePart()-> deletes the property
' ============================================================================

Private Const PREFERENCE_PROPERTY As String = "BANYAN_PREF"
Private Const MSO_PROPERTY_TYPE_STRING As Long = 4

' --- Default bibliography style names (localized) ---
Private Const DEFAULT_BIB_TITLE_STYLE_ZH As String = "文献列表标题"
Private Const DEFAULT_BIB_ENTRY_STYLE_ZH As String = "文献列表题录"
Private Const DEFAULT_BIB_TITLE_STYLE_EN As String = "Bibliography Title"
Private Const DEFAULT_BIB_ENTRY_STYLE_EN As String = "Bibliography Entry"

Private m_lastSaveError As String


' --- PreferenceInit - retained for callers that initialize modules uniformly. ---

Public Sub PreferenceInit()
End Sub


' --- PreferenceGet - Return the current Preference object. ---
' Chapter-level fields are read from the nearest previous
' ChapterBreak; if none exists, the document property provides
' defaults for all fields.

Public Function PreferenceGet() As Object   ' Dictionary
    PreferenceInit
    Set PreferenceGet = GetPreferenceInternal()
End Function


' --- PreferenceEnsure - Return current preference, creating the property first ---
' if this document has not been initialized yet.

Public Function PreferenceEnsure() As Object
    PreferenceInit

    Dim pref As Object
    Set pref = GetPreferenceInternal()
    If Not pref Is Nothing Then
        Set PreferenceEnsure = pref
        Exit Function
    End If

    Set pref = CreatePreference()
    If pref Is Nothing Then
        Set PreferenceEnsure = Nothing
        Exit Function
    End If

    Set PreferenceEnsure = pref
End Function


' --- PreferenceSave - Persist a Preference object. ---
' If a previous ChapterBreak exists, chapter-level fields
' are written into that ADDIN field's JSON data.
' Otherwise all fields go into the document property.

Public Function PreferenceSave(ByVal pref As Object) As Boolean
    PreferenceInit
    m_lastSaveError = ""
    PreferenceSave = SavePreferenceInternal(pref)
End Function

Public Function PreferenceGetLastSaveError() As String
    PreferenceGetLastSaveError = m_lastSaveError
End Function


' --- PreferenceFetchStyle - Ask the backend for a Banyan citation style. ---
' currentStyle is passed so the backend can show or
' update the current selection when supported.

Public Function PreferenceFetchStyle(Optional ByVal currentStyle As Variant) As Object
    If IsMissing(currentStyle) Then
        Set PreferenceFetchStyle = FetchStyleFromBackend()
    Else
        Set PreferenceFetchStyle = FetchStyleFromBackend(currentStyle)
    End If
End Function


' --- RemovePreferencePart - Delete the Banyan preference document property. ---

Public Sub RemovePreferencePart()
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties(PREFERENCE_PROPERTY).Delete
    On Error GoTo 0
End Sub


' --- GetPrefStyle - Factory for a PrefStyle Dictionary. ---

Public Function GetPrefStyle(ByVal styleId As String, _
                             ByVal styleTitle As String, _
                             ByVal citationType As String) As Object
    Dim d As Object
    Set d = New Dictionary
    d("id") = styleId
    d("title") = styleTitle
    d("citationType") = citationType
    Set GetPrefStyle = d
End Function


' --- Internal - read preference from document property + ChapterBreak override ---

Private Function GetPreferenceInternal() As Object
    Dim partPref As Object
    Set partPref = GetPrefFromDocumentProperty()
    If partPref Is Nothing Then
        ' No preference property exists - preference has not been initialised.
        ' This matches WPS behaviour: getPreference() returns null
        ' when the document lacks a preference part and the backend
        ' style request fails / hasn't happened yet.
        Set GetPreferenceInternal = Nothing
        Exit Function
    End If

    ' Try to find a previous ChapterBreak for chapter-level overrides
    Dim chBreak As Collection
    Set chBreak = FindPreviousChapterBreak()
    If chBreak Is Nothing Then
        Set GetPreferenceInternal = partPref
        Exit Function
    End If

    ' ChapterBreak found - merge: global from property, chapter-level from field
    Dim chData As Object
    Set chData = chBreak("data")

    Dim merged As Object
    Set merged = New Dictionary
    merged("syncItems") = partPref("syncItems")
    merged("refreshAll") = partPref("refreshAll")
    DictCopyKey merged, "style", chData, "style"
    CopyOptionalDictionaryValue merged, "extraSource", chData, "extraSource", Null
    merged("bibliographyTitleStyle") = GetDictionaryString(chData, "bibliographyTitleStyle", DictKeyString(partPref, "bibliographyTitleStyle"))
    merged("bibliographyEntryStyle") = GetDictionaryString(chData, "bibliographyEntryStyle", DictKeyString(partPref, "bibliographyEntryStyle"))
    Set GetPreferenceInternal = merged
End Function


' --- Internal - persist preference ---

Private Function SavePreferenceInternal(ByVal pref As Object) As Boolean
    On Error GoTo ErrHandler

    Dim chBreak As Collection
    Set chBreak = FindPreviousChapterBreak()
    If Not chBreak Is Nothing Then
        ' Global fields go into the document property; chapter-level fields go
        ' into the ChapterBreak ADDIN field.
        Dim docPref As Object
        Set docPref = GetPrefFromDocumentProperty()
        If docPref Is Nothing Then
            Set docPref = pref
        Else
            docPref("syncItems") = pref("syncItems")
            docPref("refreshAll") = pref("refreshAll")
        End If
        If Not SaveDocumentProperty(docPref) Then Exit Function
        SavePreferenceInternal = SaveChapterBreakPrefs(chBreak, pref)
    Else
        ' All fields go into the document property
        SavePreferenceInternal = SavePartPrefs(pref)
    End If
    Exit Function
ErrHandler:
    DiagnosticsReraiseIfDev "modPreference.SavePreferenceInternal"
    SetLastSaveError "PreferenceSave", Err.Number, Err.Source, Err.Description, Erl
    SavePreferenceInternal = False
End Function


' --- SaveChapterBreakPrefs - Update chapter-level prefs in ADDIN field data. ---

Private Function SaveChapterBreakPrefs(ByVal chBreak As Collection, ByVal pref As Object) As Boolean
    On Error GoTo ErrHandler

    ' chBreak: Collection with keys "field" (Word.Field) and "data" (Dictionary)
    Dim fld As Field
    Set fld = chBreak("field")

    If fld Is Nothing Then
        SetLastSaveFailure "SaveChapterBreakPrefs", "Chapter break field is missing."
        Exit Function
    End If

    ' Read current data, update chapter-level fields
    Dim data As Object
    Set data = ReadFieldData(fld)
    If data Is Nothing Then
        SetLastSaveFailure "SaveChapterBreakPrefs", "Unable to read chapter break field data."
        Exit Function
    End If

    Set data("style") = pref("style")
    data("bibliographyTitleStyle") = pref("bibliographyTitleStyle")
    data("bibliographyEntryStyle") = pref("bibliographyEntryStyle")
    CopyOptionalExtraSourceToChapterBreak data, pref

    If Not WriteFieldData(fld, data) Then
        SetLastSaveFailure "SaveChapterBreakPrefs", "Unable to write chapter break field data."
        Exit Function
    End If

    SaveChapterBreakPrefs = True
    Exit Function
ErrHandler:
    DiagnosticsReraiseIfDev "modPreference.SaveChapterBreakPrefs"
    SetLastSaveError "SaveChapterBreakPrefs", Err.Number, Err.Source, Err.Description, Erl
    SaveChapterBreakPrefs = False
End Function


' --- Preference property helpers ---

Private Function CreatePreference() As Object
    Dim style As Object  ' { id, title, citationType }
    Set style = FetchStyleFromBackend()
    If style Is Nothing Then
        Set CreatePreference = Nothing
        Exit Function
    End If

    Dim pref As Object
    Set pref = New Dictionary
    pref("syncItems") = True
    pref("refreshAll") = True
    Set pref("style") = GetPrefStyle(DictKeyString(style, "id"), DictKeyString(style, "title"), DictKeyString(style, "citationType"))
    pref("extraSource") = Null
    pref("bibliographyTitleStyle") = DefaultBibTitleStyle()
    pref("bibliographyEntryStyle") = DefaultBibEntryStyle()

    If Not SaveDocumentProperty(pref) Then
        Set CreatePreference = Nothing
        Exit Function
    End If

    Set CreatePreference = pref
End Function


' --- FetchStyleFromBackend - POST /banyan/style to get the citation style. ---

Private Function FetchStyleFromBackend(Optional ByVal currentStyle As Variant) As Object
    Dim url As String
    url = HttpBuildUrl("style")

    Dim docId As String
    docId = GetDocumentId()

    Dim body As String
    body = "{""documentId"":""" & EscapeJson(docId) & """"
    Dim hasCurrentStyle As Boolean
    If Not IsMissing(currentStyle) Then
        hasCurrentStyle = IsObject(currentStyle)
    End If

    If hasCurrentStyle Then
        ' Read hints via modDict (safe accessor, see README "Dictionary safety").
        Dim stId As String
        Dim stTitle As String
        stId = DictKeyString(currentStyle, "id")
        stTitle = DictKeyString(currentStyle, "title")

        If Len(stId) > 0 Then
            body = body & ",""id"":""" & EscapeJson(stId) & """"
        End If
        If Len(stTitle) > 0 Then
            body = body & ",""title"":""" & EscapeJson(stTitle) & """"
        End If
    End If
    body = body & "}"

    Dim respText As String
    respText = HttpPost(url, body)
    If Len(respText) = 0 Then
        Set FetchStyleFromBackend = Nothing
        Exit Function
    End If

    On Error Resume Next
    Dim resp As Object
    Set resp = JsonParse(respText)
    If Err.Number <> 0 Then
        Set FetchStyleFromBackend = Nothing
        On Error GoTo 0
        Exit Function
    End If
    If resp Is Nothing Then
        Set FetchStyleFromBackend = Nothing
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    ' Response format: { "ok": true/false, "data": { "id":..., "title":..., "citationType":... } }
    On Error Resume Next
    If Not DictKeyBool(resp, "ok") Then
        Set FetchStyleFromBackend = Nothing
        On Error GoTo 0
        Exit Function
    End If

    Dim data As Object
    Set data = resp("data")
    If data Is Nothing Then
        Set FetchStyleFromBackend = Nothing
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Set FetchStyleFromBackend = data
End Function


' --- GetDocumentId - Unique identifier for the active document. ---

Public Function GetDocumentId() As String
    On Error Resume Next
    Dim doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then
        GetDocumentId = "__no_document__"
    ElseIf Len(doc.FullName) > 0 Then
        GetDocumentId = doc.FullName
    ElseIf Len(doc.Name) > 0 Then
        GetDocumentId = doc.Name
    Else
        GetDocumentId = "__unnamed__"
    End If
    On Error GoTo 0
End Function


' --- EscapeJson - Minimal JSON string escaping. ---

Private Function EscapeJson(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    EscapeJson = s
End Function

Private Function GetPrefFromDocumentProperty() As Object
    On Error GoTo ErrHandler

    ' The BANYAN_PREF property is optional - "not found" is a valid state that
    ' must return Nothing. Probe existence explicitly instead of letting the
    ' missing-property error fall into the ErrHandler (which is how the dev-mode
    ' DiagnosticsReraiseIfDev would otherwise turn an expected state into a throw).
    Dim hasProp As Boolean
    hasProp = False
    Dim prop As DocumentProperty
    For Each prop In ActiveDocument.CustomDocumentProperties
        If prop.Name = PREFERENCE_PROPERTY Then
            hasProp = True
            Exit For
        End If
    Next prop
    If Not hasProp Then
        Set GetPrefFromDocumentProperty = Nothing
        Exit Function
    End If

    Dim raw As String
    raw = CStr(ActiveDocument.CustomDocumentProperties(PREFERENCE_PROPERTY).Value)
    If Len(raw) = 0 Then
        Set GetPrefFromDocumentProperty = Nothing
        Exit Function
    End If

    Dim pref As Object
    Set pref = JsonParse(raw)
    If pref Is Nothing Then
        Set GetPrefFromDocumentProperty = Nothing
        Exit Function
    End If
    If Not NormalizePreference(pref) Then
        Set GetPrefFromDocumentProperty = Nothing
        Exit Function
    End If

    Set GetPrefFromDocumentProperty = pref
    Exit Function
ErrHandler:
    DiagnosticsReraiseIfDev "modPreference.GetPrefFromDocumentProperty"
    Set GetPrefFromDocumentProperty = Nothing
End Function

Private Function SavePartPrefs(ByVal pref As Object) As Boolean
    SavePartPrefs = SaveDocumentProperty(pref)
End Function

Private Function SaveDocumentProperty(ByVal pref As Object) As Boolean
    On Error GoTo ErrHandler

    If Not NormalizePreference(pref) Then
        SetLastSaveFailure "SaveDocumentProperty", "Invalid preference data."
        Exit Function
    End If

    Dim value As String
    value = JsonStringify(pref)
    If Len(value) = 0 Then
        SetLastSaveFailure "SaveDocumentProperty", "Unable to serialize preference JSON."
        Exit Function
    End If

    On Error Resume Next
    ActiveDocument.CustomDocumentProperties(PREFERENCE_PROPERTY).Value = value
    If Err.Number = 0 Then
        SaveDocumentProperty = True
        On Error GoTo 0
        Exit Function
    End If
    Err.Clear
    On Error GoTo ErrHandler

    ActiveDocument.CustomDocumentProperties.Add PREFERENCE_PROPERTY, _
                                                False, _
                                                MSO_PROPERTY_TYPE_STRING, _
                                                value
    SaveDocumentProperty = True
    Exit Function
ErrHandler:
    DiagnosticsReraiseIfDev "modPreference.SaveDocumentProperty"
    SetLastSaveError "SaveDocumentProperty", Err.Number, Err.Source, Err.Description, Erl
    SaveDocumentProperty = False
End Function

Private Function NormalizePreference(ByVal pref As Object) As Boolean
    On Error GoTo InvalidPreference

    If Not pref.Exists("syncItems") Then pref("syncItems") = True
    If Not pref.Exists("refreshAll") Then pref("refreshAll") = True
    If Not pref.Exists("style") Then GoTo InvalidPreference
    If Not DictKeyIsObject(pref, "style") Then GoTo InvalidPreference
    If Not DictKeyObject(pref, "style").Exists("id") Then GoTo InvalidPreference
    If Not pref("style").Exists("title") Then GoTo InvalidPreference
    If Not pref("style").Exists("citationType") Then GoTo InvalidPreference

    If Not pref.Exists("extraSource") Then pref("extraSource") = Null
    If Not pref.Exists("bibliographyTitleStyle") Then pref("bibliographyTitleStyle") = DefaultBibTitleStyle()
    If Len(DictKeyString(pref, "bibliographyTitleStyle")) = 0 Then pref("bibliographyTitleStyle") = DefaultBibTitleStyle()
    If Not pref.Exists("bibliographyEntryStyle") Then pref("bibliographyEntryStyle") = DefaultBibEntryStyle()
    If Len(DictKeyString(pref, "bibliographyEntryStyle")) = 0 Then pref("bibliographyEntryStyle") = DefaultBibEntryStyle()

    NormalizePreference = True
    Exit Function
InvalidPreference:
    NormalizePreference = False
End Function


' --- Dictionary value helpers ---

Private Sub CopyOptionalDictionaryValue(ByVal target As Object, _
                                        ByVal targetKey As String, _
                                        ByVal source As Object, _
                                        ByVal sourceKey As String, _
                                        ByVal fallback As Variant)
    On Error GoTo UseFallback

    If Not DictCopyKey(target, targetKey, source, sourceKey) Then GoTo UseFallback
    Exit Sub

UseFallback:
    If IsObject(fallback) Then
        Set target(targetKey) = fallback
    Else
        target(targetKey) = fallback
    End If
End Sub

Private Sub CopyOptionalExtraSourceToChapterBreak(ByVal data As Object, ByVal pref As Object)
    On Error GoTo ClearExtraSource

    If Not DictKeyIsObject(pref, "extraSource") Then GoTo ClearExtraSource
    If Not IsCitationSource(DictKeyObject(pref, "extraSource")) Then GoTo ClearExtraSource

    Set data("extraSource") = DictKeyObject(pref, "extraSource")
    Exit Sub

ClearExtraSource:
    On Error Resume Next
    If data.Exists("extraSource") Then data.Remove "extraSource"
    On Error GoTo 0
End Sub

Private Function GetDictionaryString(ByVal source As Object, _
                                     ByVal key As String, _
                                     ByVal fallback As String) As String
    On Error GoTo UseFallback

    If DictKeyIsScalar(source, key) Then
        GetDictionaryString = DictKeyString(source, key)
        Exit Function
    End If

UseFallback:
    GetDictionaryString = fallback
End Function

Private Function SerializeExtraSource(ByVal pref As Object) As String
    If Not DictKeyIsObject(pref, "extraSource") Then
        SerializeExtraSource = ""
        Exit Function
    End If
    If Not IsCitationSource(DictKeyObject(pref, "extraSource")) Then
        SerializeExtraSource = ""
        Exit Function
    End If
    SerializeExtraSource = JsonStringify(DictKeyObject(pref, "extraSource"))
End Function


' --- Save diagnostics ---

Private Sub SetLastSaveFailure(ByVal source As String, ByVal description As String)
    m_lastSaveError = "Source: " & source & vbCrLf & _
                      "Description: " & description
End Sub

Private Sub SetLastSaveError(ByVal scope As String, _
                             ByVal errNumber As Long, _
                             ByVal errSource As String, _
                             ByVal errDescription As String, _
                             ByVal errLine As Long)
    m_lastSaveError = DiagnosticErrorText(errNumber, _
                                          NonEmptyText(errSource, scope), _
                                          errDescription, _
                                          errLine)
End Sub

Private Function NonEmptyText(ByVal value As String, ByVal fallback As String) As String
    If Len(value) > 0 Then
        NonEmptyText = value
    Else
        NonEmptyText = fallback
    End If
End Function


' --- Localization helpers ---

Private Function DefaultBibTitleStyle() As String
    If I10nIsLangZH() Then
        DefaultBibTitleStyle = DEFAULT_BIB_TITLE_STYLE_ZH
    Else
        DefaultBibTitleStyle = DEFAULT_BIB_TITLE_STYLE_EN
    End If
End Function

Private Function DefaultBibEntryStyle() As String
    If I10nIsLangZH() Then
        DefaultBibEntryStyle = DEFAULT_BIB_ENTRY_STYLE_ZH
    Else
        DefaultBibEntryStyle = DEFAULT_BIB_ENTRY_STYLE_EN
    End If
End Function


