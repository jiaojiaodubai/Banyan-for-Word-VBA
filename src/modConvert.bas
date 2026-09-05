Option Explicit

' ============================================================================
' Module  : modConvert
' Purpose : Convert Zotero ADDIN citation fields into Banyan ADDIN fields.
'
'           Mirrors WPS moulds/convert.ts:
'             - read Zotero document preferences from ZOTERO_PREF fragments
'             - decide whether the document uses in-text citations, footnotes, or endnotes
'             - collect Zotero CSL_CITATION field JSON by citationID
'             - POST /banyan/convert
'             - replace Zotero fields with Banyan citation fields
'
'           The original CSL style is intentionally ignored; the backend only
'           needs the Zotero citation JSON and target Banyan citation type.
' ============================================================================

Private Const ZOTERO_PREF_PROPERTY As String = "ZOTERO_PREF"
Private Const ZOTERO_NOTE_TYPE_MISSING As Long = -1
Private Const ZOTERO_NOTE_INTEXT As Long = 0
Private Const ZOTERO_NOTE_FOOTNOTE As Long = 1
Private Const ZOTERO_NOTE_ENDNOTE As Long = 2

Private m_i18nReady As Boolean


' --- Ribbon action ---

Public Sub ConvertAction()
    EnsureConvertI10n

    Dim originalRange As Range
    On Error Resume Next
    Set originalRange = Selection.Range.Duplicate
    On Error GoTo ErrHandler

    Dim noteType As Long
    noteType = ResolveZoteroNoteType()
    If noteType = ZOTERO_NOTE_TYPE_MISSING Then
        MsgBox CText("noDocumentPreference", _
                     "Zotero document preferences were not found, so the document citation type could not be determined."), _
               vbExclamation, CText("dialogTitle", "Banyan Convert")
        RestoreConvertSelection originalRange
        Exit Sub
    End If

    ProgressOpen CText("progressReason", "Converting Zotero fields...")

    If noteType = ZOTERO_NOTE_ENDNOTE Then
        ConvertEndnoteCitationsToIntext
    ElseIf noteType = ZOTERO_NOTE_FOOTNOTE Then
        ConvertNoteCitations
    Else
        ConvertIntextCitations
    End If

CleanUp:
    ProgressClose
    RestoreConvertSelection originalRange
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.ConvertAction"
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim errLine As Long
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    errLine = Erl

    ProgressClose
    RestoreConvertSelection originalRange
    DiagnosticShowError CText("dialogTitle", "Banyan Convert"), _
                        Replace(CText("error", "An error occurred while converting Zotero fields: {message}"), _
                                "{message}", errDescription), _
                        errNumber, errSource, errDescription, errLine
End Sub

Private Sub RestoreConvertSelection(ByVal originalRange As Range)
    On Error Resume Next
    If Not originalRange Is Nothing Then originalRange.Select
    On Error GoTo 0
End Sub


' --- Conversion ---

Private Function ConvertIntextCitations() As Boolean
    Dim targets As Collection
    Set targets = CollectZoteroIntextFields()
    If targets.Count = 0 Then
        MsgBox CText("noFieldsFound", "No Zotero citation fields were found to convert."), _
               vbInformation, CText("dialogTitle", "Banyan Convert")
        Exit Function
    End If

    Dim response As Object
    Set response = RequestConvert("intext-citation", targets)
    If response Is Nothing Then Exit Function

    Dim convertedCount As Long
    Dim i As Long
    For i = targets.Count To 1 Step -1
        Dim target As Collection
        Set target = targets(i)

        Dim converted As Object
        Set converted = GetConvertedIntextCitation(response, CStr(target("fieldId")))
        If converted Is Nothing Then GoTo NextTarget

        Dim insertRange As Range
        Set insertRange = CreateCollapsedRange(CLng(target("field").Result.Start))
        FieldRemoveFieldSafely target("field")
        FieldCreateIntextCitationAtRange insertRange, converted
        convertedCount = convertedCount + 1

NextTarget:
    Next i

    If convertedCount = 0 Then
        MsgBox CText("invalidResponse", "The backend did not return any usable conversion results."), _
               vbExclamation, CText("dialogTitle", "Banyan Convert")
        Exit Function
    End If

    ConvertIntextCitations = True
End Function

Private Function ConvertEndnoteCitationsToIntext() As Boolean
    Dim targets As Collection
    Set targets = CollectZoteroEndnoteFields()
    If targets.Count = 0 Then
        MsgBox CText("noFieldsFound", "No Zotero citation fields were found to convert."), _
               vbInformation, CText("dialogTitle", "Banyan Convert")
        Exit Function
    End If

    Dim response As Object
    Set response = RequestConvert("intext-citation", targets)
    If response Is Nothing Then Exit Function

    Dim convertedCount As Long
    Dim i As Long
    For i = targets.Count To 1 Step -1
        Dim target As Collection
        Set target = targets(i)

        Dim converted As Object
        Set converted = GetConvertedIntextCitation(response, CStr(target("fieldId")))
        If converted Is Nothing Then GoTo NextTarget

        Dim insertRange As Range
        Set insertRange = CreateCollapsedRange(CLng(target("note").Reference.Start))
        FieldRemoveEndnoteSafely target("note")
        FieldCreateIntextCitationAtRange insertRange, converted
        convertedCount = convertedCount + 1

NextTarget:
    Next i

    If convertedCount = 0 Then
        MsgBox CText("invalidResponse", "The backend did not return any usable conversion results."), _
               vbExclamation, CText("dialogTitle", "Banyan Convert")
        Exit Function
    End If

    ConvertEndnoteCitationsToIntext = True
End Function

Private Function ConvertNoteCitations() As Boolean
    Dim targets As Collection
    Set targets = CollectZoteroNoteFields()
    If targets.Count = 0 Then
        MsgBox CText("noFieldsFound", "No Zotero citation fields were found to convert."), _
               vbInformation, CText("dialogTitle", "Banyan Convert")
        Exit Function
    End If

    Dim response As Object
    Set response = RequestConvert("note-citation", targets)
    If response Is Nothing Then Exit Function

    Dim convertedCount As Long
    Dim i As Long
    For i = targets.Count To 1 Step -1
        Dim target As Collection
        Set target = targets(i)

        Dim converted As Object
        Set converted = GetConvertedNoteCitation(response, CStr(target("fieldId")))
        If converted Is Nothing Then GoTo NextTarget

        Dim insertRange As Range
        Set insertRange = CreateCollapsedRange(CLng(target("note").Reference.Start))
        FieldRemoveFootnoteSafely target("note")
        FieldCreateNoteCitationAtRange insertRange, converted
        convertedCount = convertedCount + 1

NextTarget:
    Next i

    If convertedCount = 0 Then
        MsgBox CText("invalidResponse", "The backend did not return any usable conversion results."), _
               vbExclamation, CText("dialogTitle", "Banyan Convert")
        Exit Function
    End If

    ConvertNoteCitations = True
End Function


' --- Backend ---

Private Function RequestConvert(ByVal citationType As String, ByVal targets As Collection) As Object
    On Error GoTo ErrHandler

    Dim fieldCodesById As Object
    Set fieldCodesById = New Dictionary

    Dim target As Variant
    For Each target In targets
        Dim fieldId As String
        Dim fieldCode As String
        fieldId = CStr(target("fieldId"))
        fieldCode = CStr(target("fieldCode"))

        If Not fieldCodesById.Exists(fieldId) Then
            fieldCodesById.Add fieldId, fieldCode
        End If
    Next target

    Dim fields As Collection
    Set fields = New Collection

    Dim key As Variant
    For Each key In fieldCodesById.Keys
        Dim fieldInput As Object
        Set fieldInput = New Dictionary
        fieldInput("fieldId") = CStr(key)
        fieldInput("fieldCode") = DictKeyString(fieldCodesById, CStr(key))
        fields.Add fieldInput
    Next key

    Dim body As Object
    Set body = New Dictionary
    body("documentId") = GetDocumentId()
    body("citationType") = citationType
    Set body("fields") = fields

    Dim respText As String
    respText = HttpPost(HttpBuildUrl("convert"), JsonStringify(body))
    If Len(respText) = 0 Then Exit Function

    Dim envelope As Object
    Set envelope = JsonParse(respText)
    If envelope Is Nothing Then Exit Function
    If Not EnvelopeOk(envelope) Then
        ShowHttpEnvelopeError envelope
        Exit Function
    End If

    If Not HasDictionaryKey(envelope, "data") Then Exit Function
    If Not IsDictionaryRecord(envelope("data")) Then Exit Function

    Set RequestConvert = envelope("data")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.RequestConvert"
    Set RequestConvert = Nothing
End Function

Private Function EnvelopeOk(ByVal envelope As Object) As Boolean
    On Error GoTo ErrHandler
    If Not envelope.Exists("ok") Then Exit Function
    EnvelopeOk = DictKeyBool(envelope, "ok")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.EnvelopeOk"
    EnvelopeOk = False
End Function

Private Sub ShowHttpEnvelopeError(ByVal envelope As Object)
    On Error GoTo ErrHandler
    If envelope.Exists("error") Then
        Dim errObj As Object
        Set errObj = envelope("error")
        If Not errObj Is Nothing Then
            If errObj.Exists("code") Then
                If DictKeyString(errObj, "code") = "cancelled" Then Exit Sub
            End If
            MsgBox "HTTP error: [" & DictKeyString(errObj, "code") & "] " & DictKeyString(errObj, "message"), _
                   vbExclamation, CText("dialogTitle", "Banyan Convert")
        End If
    End If
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.ShowHttpEnvelopeError"
End Sub


' --- Zotero document preferences ---

Private Function ResolveZoteroNoteType() As Long
    ResolveZoteroNoteType = ZOTERO_NOTE_TYPE_MISSING

    Dim data As String
    data = ReadZoteroDocumentPreferenceData()
    If Len(data) = 0 Then Exit Function

    Dim parsed As Long
    parsed = ParseZoteroDocumentPreferenceJson(data)
    If parsed <> ZOTERO_NOTE_TYPE_MISSING Then
        ResolveZoteroNoteType = parsed
        Exit Function
    End If

    parsed = ParseZoteroDocumentPreferenceXml(data)
    If parsed <> ZOTERO_NOTE_TYPE_MISSING Then
        ResolveZoteroNoteType = parsed
        Exit Function
    End If

    ResolveZoteroNoteType = ParseLegacyZoteroDocumentPreference(data)
End Function

Private Function ReadZoteroDocumentPreferenceData() As String
    On Error GoTo ErrHandler

    Dim properties As Object
    Set properties = ActiveDocument.CustomDocumentProperties

    ReadZoteroDocumentPreferenceData = ReadCustomPropertyByFragments(properties, ZOTERO_PREF_PROPERTY)
    If Len(ReadZoteroDocumentPreferenceData) > 0 Then Exit Function

    ReadZoteroDocumentPreferenceData = ReadCustomPropertyByName(properties, ZOTERO_PREF_PROPERTY)
    If Len(ReadZoteroDocumentPreferenceData) > 0 Then Exit Function

    ReadZoteroDocumentPreferenceData = ReadCustomPropertyByIteration(properties, ZOTERO_PREF_PROPERTY)
    If Len(ReadZoteroDocumentPreferenceData) > 0 Then Exit Function

    Dim variables As Object
    Set variables = ActiveDocument.Variables

    ReadZoteroDocumentPreferenceData = ReadVariableByFragments(variables, ZOTERO_PREF_PROPERTY)
    If Len(ReadZoteroDocumentPreferenceData) > 0 Then Exit Function

    ReadZoteroDocumentPreferenceData = ReadVariableByName(variables, ZOTERO_PREF_PROPERTY)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.ReadZoteroDocumentPreferenceData"
    ReadZoteroDocumentPreferenceData = ""
End Function

Private Function ReadCustomPropertyByFragments(ByVal properties As Object, ByVal name As String) As String
    On Error GoTo ErrHandler
    If properties Is Nothing Then Exit Function

    Dim count As Long
    count = CLng(properties.Count)
    If count <= 0 Then Exit Function

    Dim fragments As Object
    Set fragments = New Dictionary

    Dim prefix As String
    prefix = UCase$(name & "_")

    Dim i As Long
    For i = 1 To count
        Dim property As Object
        Set property = properties.Item(i)

        Dim propertyName As String
        propertyName = UCase$(CStr(property.Name))
        If Left$(propertyName, Len(prefix)) <> prefix Then GoTo NextProperty

        Dim suffix As String
        suffix = Mid$(propertyName, Len(prefix) + 1)
        If Not IsPositiveIntegerString(suffix) Then GoTo NextProperty

        If Not fragments.Exists(CStr(CLng(suffix))) Then
            fragments.Add CStr(CLng(suffix)), VariantToText(property.Value)
        End If

NextProperty:
    Next i

    ReadCustomPropertyByFragments = JoinSequentialFragments(fragments)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.ReadCustomPropertyByFragments"
    ReadCustomPropertyByFragments = ""
End Function

Private Function ReadCustomPropertyByName(ByVal properties As Object, ByVal name As String) As String
    On Error GoTo ErrHandler
    If properties Is Nothing Then Exit Function

    Dim property As Object
    Set property = properties.Item(name)

    If Not StringEqualsIgnoreCase(CStr(property.Name), name) Then Exit Function
    ReadCustomPropertyByName = VariantToText(property.Value)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.ReadCustomPropertyByName"
    ReadCustomPropertyByName = ""
End Function

Private Function ReadCustomPropertyByIteration(ByVal properties As Object, ByVal name As String) As String
    On Error GoTo ErrHandler
    If properties Is Nothing Then Exit Function

    Dim count As Long
    count = CLng(properties.Count)
    If count <= 0 Then Exit Function

    Dim i As Long
    For i = 1 To count
        Dim property As Object
        Set property = properties.Item(i)
        If StringEqualsIgnoreCase(CStr(property.Name), name) Then
            ReadCustomPropertyByIteration = VariantToText(property.Value)
            Exit Function
        End If
    Next i
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.ReadCustomPropertyByIteration"
    ReadCustomPropertyByIteration = ""
End Function

Private Function ReadVariableByFragments(ByVal variables As Object, ByVal name As String) As String
    On Error GoTo ErrHandler
    If variables Is Nothing Then Exit Function

    Dim count As Long
    count = CLng(variables.Count)
    If count <= 0 Then Exit Function

    Dim fragments As Object
    Set fragments = New Dictionary

    Dim prefix As String
    prefix = UCase$(name & "_")

    Dim i As Long
    For i = 1 To count
        Dim variable As Object
        Set variable = variables.Item(i)

        Dim variableName As String
        variableName = UCase$(CStr(variable.Name))
        If Left$(variableName, Len(prefix)) <> prefix Then GoTo NextVariable

        Dim suffix As String
        suffix = Mid$(variableName, Len(prefix) + 1)
        If Not IsPositiveIntegerString(suffix) Then GoTo NextVariable

        If Not fragments.Exists(CStr(CLng(suffix))) Then
            fragments.Add CStr(CLng(suffix)), VariantToText(variable.Value)
        End If

NextVariable:
    Next i

    ReadVariableByFragments = JoinSequentialFragments(fragments)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.ReadVariableByFragments"
    ReadVariableByFragments = ""
End Function

Private Function ReadVariableByName(ByVal variables As Object, ByVal name As String) As String
    On Error GoTo IterationFallback
    If variables Is Nothing Then Exit Function

    Dim variable As Object
    Set variable = variables.Item(name)
    ReadVariableByName = VariantToText(variable.Value)
    Exit Function

IterationFallback:
    On Error GoTo ErrHandler

    Dim count As Long
    count = CLng(variables.Count)
    If count <= 0 Then Exit Function

    Dim i As Long
    For i = 1 To count
        Set variable = variables.Item(i)
        If StringEqualsIgnoreCase(CStr(variable.Name), name) Then
            ReadVariableByName = VariantToText(variable.Value)
            Exit Function
        End If
    Next i
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.ReadVariableByName"
    ReadVariableByName = ""
End Function

Private Function JoinSequentialFragments(ByVal fragments As Object) As String
    On Error GoTo ErrHandler
    If fragments Is Nothing Then Exit Function
    If Not fragments.Exists("1") Then Exit Function

    Dim combined As String
    Dim i As Long
    i = 1
    Do While fragments.Exists(CStr(i))
        combined = combined & DictKeyString(fragments, CStr(i))
        i = i + 1
    Loop

    JoinSequentialFragments = combined
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.JoinSequentialFragments"
    JoinSequentialFragments = ""
End Function


' --- Zotero preference parsing ---

Private Function ParseZoteroDocumentPreferenceJson(ByVal data As String) As Long
    ParseZoteroDocumentPreferenceJson = ZOTERO_NOTE_TYPE_MISSING

    Dim trimmed As String
    trimmed = Trim$(data)
    If Left$(trimmed, 1) <> "{" Then Exit Function

    On Error GoTo ErrHandler
    Dim parsed As Object
    Set parsed = JsonParse(trimmed)
    If parsed Is Nothing Then Exit Function
    If Not IsDictionaryRecord(parsed) Then Exit Function
    If Not HasDictionaryKey(parsed, "prefs") Then Exit Function
    If Not IsDictionaryRecord(parsed("prefs")) Then Exit Function

    Dim prefs As Object
    Set prefs = parsed("prefs")
    If Not HasDictionaryKey(prefs, "noteType") Then
        ParseZoteroDocumentPreferenceJson = ZOTERO_NOTE_INTEXT
        Exit Function
    End If

    ParseZoteroDocumentPreferenceJson = ParseNoteTypeValue(prefs("noteType"), ZOTERO_NOTE_INTEXT)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.ParseZoteroDocumentPreferenceJson"
    ParseZoteroDocumentPreferenceJson = ZOTERO_NOTE_TYPE_MISSING
End Function

Private Function ParseZoteroDocumentPreferenceXml(ByVal data As String) As Long
    ParseZoteroDocumentPreferenceXml = ZOTERO_NOTE_TYPE_MISSING

    Dim trimmed As String
    trimmed = Trim$(data)
    If Left$(trimmed, 1) <> "<" Then Exit Function

    On Error GoTo TextFallback
    Dim doc As Object
    Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    doc.async = False
    doc.validateOnParse = False
    doc.resolveExternals = False

    If Not doc.LoadXML(trimmed) Then GoTo TextFallback

    Dim prefs As Object
    Set prefs = doc.getElementsByTagName("pref")

    Dim i As Long
    For i = 0 To prefs.Length - 1
        Dim pref As Object
        Set pref = prefs.Item(i)
        If CStr(pref.getAttribute("name")) = "noteType" Then
            Dim value As String
            value = CStr(pref.getAttribute("value"))
            If Len(Trim$(value)) = 0 Then
                ParseZoteroDocumentPreferenceXml = ZOTERO_NOTE_INTEXT
            Else
                ParseZoteroDocumentPreferenceXml = ParseNoteTypeValue(value, ZOTERO_NOTE_TYPE_MISSING)
            End If
            Exit Function
        End If
    Next i

    ParseZoteroDocumentPreferenceXml = ZOTERO_NOTE_INTEXT
    Exit Function

TextFallback:
    ParseZoteroDocumentPreferenceXml = ParseZoteroNoteTypeFromXmlText(trimmed)
End Function

Private Function ParseZoteroNoteTypeFromXmlText(ByVal data As String) As Long
    ParseZoteroNoteTypeFromXmlText = ZOTERO_NOTE_TYPE_MISSING
    If Left$(Trim$(data), 1) <> "<" Then Exit Function

    Dim searchStart As Long
    searchStart = 1

    Do
        Dim tagStart As Long
        tagStart = InStr(searchStart, data, "<pref", vbTextCompare)
        If tagStart = 0 Then Exit Do

        Dim tagEnd As Long
        tagEnd = InStr(tagStart, data, ">")
        If tagEnd = 0 Then Exit Do

        Dim tagText As String
        tagText = Mid$(data, tagStart, tagEnd - tagStart + 1)
        If ExtractXmlAttribute(tagText, "name") = "noteType" Then
            Dim value As String
            value = ExtractXmlAttribute(tagText, "value")
            If Len(Trim$(value)) = 0 Then
                ParseZoteroNoteTypeFromXmlText = ZOTERO_NOTE_INTEXT
            Else
                ParseZoteroNoteTypeFromXmlText = ParseNoteTypeValue(value, ZOTERO_NOTE_TYPE_MISSING)
            End If
            Exit Function
        End If

        searchStart = tagEnd + 1
    Loop

    ParseZoteroNoteTypeFromXmlText = ZOTERO_NOTE_INTEXT
End Function

Private Function ParseLegacyZoteroDocumentPreference(ByVal data As String) As Long
    ParseLegacyZoteroDocumentPreference = ZOTERO_NOTE_TYPE_MISSING

    Dim parts() As String
    parts = Split(data, "::")
    If UBound(parts) < 4 Then Exit Function

    If CStr(parts(2)) <> "note" Then
        ParseLegacyZoteroDocumentPreference = ZOTERO_NOTE_INTEXT
    ElseIf CStr(parts(4)) = "1" Or CStr(parts(4)) = "True" Then
        ParseLegacyZoteroDocumentPreference = ZOTERO_NOTE_ENDNOTE
    Else
        ParseLegacyZoteroDocumentPreference = ZOTERO_NOTE_FOOTNOTE
    End If
End Function

Private Function ParseNoteTypeValue(ByVal value As Variant, ByVal fallback As Long) As Long
    On Error GoTo ErrHandler

    If IsNull(value) Or IsEmpty(value) Then
        ParseNoteTypeValue = fallback
        Exit Function
    End If

    Dim text As String
    text = Trim$(CStr(value))
    If Len(text) = 0 Then
        ParseNoteTypeValue = fallback
    Else
        ParseNoteTypeValue = CLng(text)
    End If
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.ParseNoteTypeValue"
    ParseNoteTypeValue = fallback
End Function

Private Function ExtractXmlAttribute(ByVal tagText As String, ByVal attrName As String) As String
    Dim pos As Long
    pos = 1

    Do
        pos = InStr(pos, tagText, attrName, vbTextCompare)
        If pos = 0 Then Exit Function

        Dim afterName As Long
        afterName = pos + Len(attrName)
        If Not IsAttributeNameBoundary(tagText, pos - 1) Then
            pos = afterName
            GoTo ContinueSearch
        End If
        If Not IsAttributeNameBoundary(tagText, afterName) Then
            pos = afterName
            GoTo ContinueSearch
        End If

        Dim i As Long
        i = afterName
        Do While i <= Len(tagText) And Mid$(tagText, i, 1) = " "
            i = i + 1
        Loop
        If i > Len(tagText) Or Mid$(tagText, i, 1) <> "=" Then
            pos = afterName
            GoTo ContinueSearch
        End If

        i = i + 1
        Do While i <= Len(tagText) And Mid$(tagText, i, 1) = " "
            i = i + 1
        Loop
        If i > Len(tagText) Then Exit Function

        Dim quoteChar As String
        quoteChar = Mid$(tagText, i, 1)
        If quoteChar <> """" And quoteChar <> "'" Then Exit Function

        Dim valueStart As Long
        Dim valueEnd As Long
        valueStart = i + 1
        valueEnd = InStr(valueStart, tagText, quoteChar, vbBinaryCompare)
        If valueEnd = 0 Then Exit Function

        ExtractXmlAttribute = Mid$(tagText, valueStart, valueEnd - valueStart)
        Exit Function

ContinueSearch:
    Loop
End Function

Private Function IsAttributeNameBoundary(ByVal text As String, ByVal position As Long) As Boolean
    If position < 1 Or position > Len(text) Then
        IsAttributeNameBoundary = True
        Exit Function
    End If

    Dim ch As String
    ch = Mid$(text, position, 1)
    IsAttributeNameBoundary = Not ((ch >= "A" And ch <= "Z") Or _
                                   (ch >= "a" And ch <= "z") Or _
                                   (ch >= "0" And ch <= "9") Or _
                                   ch = "_" Or ch = "-")
End Function


' --- Zotero field collection ---

Private Function CollectZoteroIntextFields() As Collection
    Dim targets As Collection
    Set targets = New Collection

    On Error GoTo ErrHandler
    Dim fld As Field
    For Each fld In ActiveDocument.Content.Fields
        Dim target As Collection
        Set target = ParseZoteroCitationField(fld)
        If Not target Is Nothing Then targets.Add target
    Next fld

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.CollectZoteroIntextFields"
    Set CollectZoteroIntextFields = targets
End Function

Private Function CollectZoteroNoteFields() As Collection
    Dim targets As Collection
    Set targets = New Collection

    On Error GoTo ErrHandler
    Dim note As Footnote
    For Each note In ActiveDocument.Footnotes
        Dim fld As Field
        For Each fld In note.Range.Fields
            Dim target As Collection
            Set target = ParseZoteroCitationField(fld)
            If Not target Is Nothing Then
                target.Add note, "note"
                targets.Add target
                Exit For
            End If
        Next fld
    Next note

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.CollectZoteroNoteFields"
    Set CollectZoteroNoteFields = targets
End Function

Private Function CollectZoteroEndnoteFields() As Collection
    Dim targets As Collection
    Set targets = New Collection

    On Error GoTo ErrHandler
    Dim note As Endnote
    For Each note In ActiveDocument.Endnotes
        Dim fld As Field
        For Each fld In note.Range.Fields
            Dim target As Collection
            Set target = ParseZoteroCitationField(fld)
            If Not target Is Nothing Then
                target.Add note, "note"
                targets.Add target
                Exit For
            End If
        Next fld
    Next note

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.CollectZoteroEndnoteFields"
    Set CollectZoteroEndnoteFields = targets
End Function

Private Function ParseZoteroCitationField(ByVal fld As Field) As Collection
    Dim fieldCode As String
    fieldCode = ReadZoteroCitationCode(fld)
    If Len(fieldCode) = 0 Then Exit Function

    Dim fieldId As String
    fieldId = ReadZoteroCitationId(fieldCode)
    If Len(fieldId) = 0 Then Exit Function

    Dim result As Collection
    Set result = New Collection
    result.Add fld, "field"
    result.Add fieldCode, "fieldCode"
    result.Add fieldId, "fieldId"
    Set ParseZoteroCitationField = result
End Function

Private Function ReadZoteroCitationCode(ByVal fld As Field) As String
    On Error GoTo ErrHandler
    If fld.Type <> wdFieldAddin Then Exit Function

    Dim codeText As String
    codeText = NormalizeFieldCodeText(CStr(fld.Code.Text))
    If Not IsZoteroCitationCode(CollapseWhitespace(codeText)) Then Exit Function

    Dim jsonStart As Long
    Dim jsonEnd As Long
    jsonStart = InStr(1, codeText, "{", vbBinaryCompare)
    jsonEnd = InStrRev(codeText, "}", -1, vbBinaryCompare)
    If jsonStart = 0 Or jsonEnd < jsonStart Then Exit Function

    ReadZoteroCitationCode = Trim$(Mid$(codeText, jsonStart, jsonEnd - jsonStart + 1))
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.ReadZoteroCitationCode"
    ReadZoteroCitationCode = ""
End Function

Private Function IsZoteroCitationCode(ByVal codeText As String) As Boolean
    Dim normalized As String
    normalized = UCase$(Trim$(codeText))

    If Left$(normalized, Len("ADDIN ")) = "ADDIN " Then
        normalized = Trim$(Mid$(normalized, Len("ADDIN ") + 1))
    End If

    If Left$(normalized, Len("ZOTERO_")) = "ZOTERO_" Then
        normalized = Mid$(normalized, Len("ZOTERO_") + 1)
    End If

    IsZoteroCitationCode = (Left$(normalized, Len("ITEM CSL_CITATION ")) = "ITEM CSL_CITATION " Or _
                            Left$(normalized, Len("CSL_CITATION ")) = "CSL_CITATION ")
End Function

Private Function ReadZoteroCitationId(ByVal fieldCode As String) As String
    On Error GoTo ErrHandler

    Dim parsed As Object
    Set parsed = JsonParse(fieldCode)
    If parsed Is Nothing Then Exit Function
    If Not DictIsDictionary(parsed) Then Exit Function

    ReadZoteroCitationId = DictKeyString(parsed, "citationID")
    If Len(ReadZoteroCitationId) > 0 Then Exit Function

    ReadZoteroCitationId = DictKeyString(parsed, "CITATIONID")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.ReadZoteroCitationId"
    ReadZoteroCitationId = ""
End Function


' --- Response lookup ---

Private Function GetConvertedIntextCitation(ByVal response As Object, ByVal fieldId As String) As Object
    On Error GoTo ErrHandler
    If Not response.Exists(fieldId) Then Exit Function
    If FieldIsIntextCitation(response(fieldId)) Then Set GetConvertedIntextCitation = response(fieldId)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.GetConvertedIntextCitation"
    Set GetConvertedIntextCitation = Nothing
End Function

Private Function GetConvertedNoteCitation(ByVal response As Object, ByVal fieldId As String) As Object
    On Error GoTo ErrHandler
    If Not response.Exists(fieldId) Then Exit Function
    If FieldIsNoteCitation(response(fieldId)) Then Set GetConvertedNoteCitation = response(fieldId)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.GetConvertedNoteCitation"
    Set GetConvertedNoteCitation = Nothing
End Function

Private Function CreateCollapsedRange(ByVal position As Long) As Range
    Dim result As Range
    Set result = ActiveDocument.Range(position, position)
    Set CreateCollapsedRange = result
End Function


' --- Type and text helpers ---

Private Function IsDictionaryRecord(ByVal value As Variant) As Boolean
    IsDictionaryRecord = DictIsDictionary(value)
End Function

Private Function HasDictionaryKey(ByVal dict As Object, ByVal key As String) As Boolean
    HasDictionaryKey = DictHasKey(dict, key)
End Function

Private Function VariantToText(ByVal value As Variant) As String
    On Error GoTo ErrHandler
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    VariantToText = CStr(value)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modConvert.VariantToText"
    VariantToText = ""
End Function

Private Function StringEqualsIgnoreCase(ByVal a As String, ByVal b As String) As Boolean
    StringEqualsIgnoreCase = (StrComp(a, b, vbTextCompare) = 0)
End Function

Private Function IsPositiveIntegerString(ByVal value As String) As Boolean
    If Len(value) = 0 Then Exit Function

    Dim i As Long
    Dim ch As String
    For i = 1 To Len(value)
        ch = Mid$(value, i, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next i

    IsPositiveIntegerString = (CLng(value) > 0)
End Function

Private Function CollapseWhitespace(ByVal value As String) As String
    value = NormalizeFieldCodeText(value)

    Do While InStr(value, "  ") > 0
        value = Replace(value, "  ", " ")
    Loop

    CollapseWhitespace = Trim$(value)
End Function

Private Function NormalizeFieldCodeText(ByVal value As String) As String
    value = Replace(value, vbCr, " ")
    value = Replace(value, vbLf, " ")
    value = Replace(value, vbTab, " ")

    NormalizeFieldCodeText = Trim$(value)
End Function


' --- Local i18n ---

Private Sub EnsureConvertI10n()
    If m_i18nReady Then Exit Sub

    I10nRegisterTable msoLanguageIDSimplifiedChinese, "convert", _
        "dialogTitle", "Banyan 转换", _
        "progressReason", "正在转换 Zotero 域...", _
        "noDocumentPreference", "未检测到 Zotero 文档偏好，无法判断当前文档是正文引注还是脚注引注。", _
        "noFieldsFound", "未检测到可转换的 Zotero 引注域。", _
        "invalidResponse", "后端没有返回可用的转换结果。", _
        "error", "转换 Zotero 域时发生错误：{message}"

    I10nRegisterTable msoLanguageIDEnglishUS, "convert", _
        "dialogTitle", "Banyan Convert", _
        "progressReason", "Converting Zotero fields...", _
        "noDocumentPreference", "Zotero document preferences were not found, so the document citation type could not be determined.", _
        "noFieldsFound", "No Zotero citation fields were found to convert.", _
        "invalidResponse", "The backend did not return any usable conversion results.", _
        "error", "An error occurred while converting Zotero fields: {message}"

    m_i18nReady = True
End Sub

Private Function CText(ByVal key As String, ByVal fallback As String) As String
    EnsureConvertI10n
    CText = T("convert." & key, fallback)
End Function
