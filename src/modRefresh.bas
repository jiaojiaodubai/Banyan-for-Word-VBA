Option Explicit

' ============================================================================
' Module  : modRefresh
' Purpose : Refresh Banyan citation and bibliography ADDIN fields.
'
'           Mirrors WPS moulds/refresh.ts:
'             - collect current chapter citation contexts
'             - POST /banyan/refresh
'             - write returned field JSON to Field.Data
'             - render returned RichText content into Field.Result
'             - refresh an existing bibliography block when present
' ============================================================================

Private m_i18nReady As Boolean


' --- Ribbon action ---

Public Sub RefreshAction()
    EnsureRefreshI10n

    Dim originalRange As Range
    On Error Resume Next
    Set originalRange = Selection.Range.Duplicate
    On Error GoTo ErrHandler

    ProgressOpen RText("progressReason", "Refreshing Banyan fields...")

    Dim pref As Object
    Set pref = PreferenceEnsure()
    If pref Is Nothing Then GoTo CleanUp

    If DictKeyBool(pref, "refreshAll") Then
        RefreshAll DictKeyBool(pref, "syncItems")
    Else
        Dim targetRange As Range
        Set targetRange = GetUpdateRange()
        If targetRange Is Nothing Then GoTo CleanUp
        RefreshInRange targetRange, DictKeyBool(pref, "syncItems")
    End If

CleanUp:
    ProgressClose
    RestoreRefreshSelection originalRange
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.RefreshAction"
    Dim errNumber As Long
    Dim errSource As String
    Dim errDescription As String
    Dim errLine As Long
    errNumber = Err.Number
    errSource = Err.Source
    errDescription = Err.Description
    errLine = Erl

    ProgressClose
    RestoreRefreshSelection originalRange
    DiagnosticShowError RText("dialogTitle", "Banyan Refresh"), _
                        Replace(RText("error", "Failed to refresh Banyan fields: {message}"), _
                                "{message}", errDescription), _
                        errNumber, errSource, errDescription, errLine
End Sub


' --- Public refresh API ---

Public Function RefreshInRange(ByVal targetRange As Range, _
                               Optional ByVal syncItemsOverride As Variant) As Boolean
    On Error GoTo ErrHandler

    If targetRange Is Nothing Then Exit Function

    Dim pref As Object
    Set pref = PreferenceEnsure()
    If pref Is Nothing Then
        RefreshLogWarn "No style found, stopping refresh."
        Exit Function
    End If

    Dim syncItems As Boolean
    If IsMissing(syncItemsOverride) Then
        syncItems = DictKeyBool(pref, "syncItems")
    Else
        syncItems = CBool(syncItemsOverride)
    End If

    Dim style As Object
    Set style = pref("style")

    Select Case DictKeyString(style, "citationType")
        Case "intext-citation"
            RefreshInRange = RefreshIntextRange(targetRange, pref, syncItems)
        Case "note-citation"
            RefreshInRange = RefreshNoteRange(targetRange, pref, syncItems)
    End Select
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.RefreshInRange"
    RefreshInRange = False
End Function

Public Sub RefreshAll(Optional ByVal syncItemsOverride As Variant)
    On Error GoTo CleanUp

    Dim originalRange As Range
    Set originalRange = Selection.Range.Duplicate

    Dim prefs As Object
    Set prefs = PreferenceEnsure()
    If prefs Is Nothing Then
        RefreshLogWarn "No style found, stopping refresh all."
        GoTo CleanUp
    End If

    ' Align with WPS: check style exists before proceeding
    Dim style As Object
    Set style = prefs("style")
    If style Is Nothing Then GoTo CleanUp

    Dim syncItems As Boolean
    If IsMissing(syncItemsOverride) Then
        syncItems = DictKeyBool(prefs, "syncItems")
    Else
        syncItems = CBool(syncItemsOverride)
    End If

    ' Start from document end, work backwards chapter by chapter
    ' to avoid index drift from forward modifications.
    Dim cursor As Range
    Set cursor = ActiveDocument.Content.Duplicate
    cursor.Collapse wdCollapseEnd
    cursor.Select

    Dim previousRangeKey As String
    Dim hasPreviousRangeKey As Boolean

    Do
        Dim targetRange As Range
        Set targetRange = GetUpdateRange()
        If targetRange Is Nothing Then Exit Do

        Dim currentRangeKey As String
        currentRangeKey = RefreshRangeKey(targetRange)
        If hasPreviousRangeKey Then
            If currentRangeKey = previousRangeKey Then
                RefreshLogWarn "RefreshAll detected repeated update range, aborting: " & currentRangeKey
                Exit Do
            End If
        End If
        previousRangeKey = currentRangeKey
        hasPreviousRangeKey = True

        RefreshInRange targetRange, syncItems

        Dim prevBreak As Collection
        Set prevBreak = FindPreviousChapterBreak()
        If prevBreak Is Nothing Then Exit Do

        Dim prevField As Field
        Set prevField = prevBreak("field")
        MoveCaretBeforeField prevField
    Loop

CleanUp:
    RestoreRefreshSelection originalRange
End Sub

Public Function RefreshForStyleChange(ByVal previousStyle As Object, ByVal nextStyle As Object) As Boolean
    On Error GoTo ErrHandler

    If SameStyle(previousStyle, nextStyle) Then Exit Function

    Dim originalRange As Range
    On Error Resume Next
    Set originalRange = Selection.Range.Duplicate
    On Error GoTo ErrHandler

    ProgressOpen RText("progressReason", "Refreshing Banyan fields...")

    Dim targetRange As Range
    Set targetRange = GetUpdateRange()
    If targetRange Is Nothing Then GoTo CleanUp

    If Not previousStyle Is Nothing Then
        If DictKeyString(previousStyle, "citationType") <> DictKeyString(nextStyle, "citationType") Then
            If DictKeyString(nextStyle, "citationType") = "note-citation" Then
                FieldMigrateIntextCitationsToNotes targetRange
            Else
                FieldMigrateNoteCitationsToIntext targetRange
            End If
        End If
    End If

    Set targetRange = GetUpdateRange()
    If targetRange Is Nothing Then GoTo CleanUp
    RefreshForStyleChange = RefreshInRange(targetRange)

CleanUp:
    ProgressClose
    RestoreRefreshSelection originalRange
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.RefreshForStyleChange"
    ProgressClose
    RestoreRefreshSelection originalRange
    RefreshForStyleChange = False
End Function

Private Sub RestoreRefreshSelection(ByVal originalRange As Range)
    On Error Resume Next
    If Not originalRange Is Nothing Then originalRange.Select
    On Error GoTo 0
End Sub

Private Function RefreshRangeKey(ByVal targetRange As Range) As String
    On Error GoTo ErrHandler
    RefreshRangeKey = CStr(targetRange.Start) & ":" & CStr(targetRange.End)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.RefreshRangeKey"
    RefreshRangeKey = ""
End Function

Private Sub MoveCaretBeforeField(ByVal fld As Field)
    On Error GoTo ErrHandler
    If fld Is Nothing Then Exit Sub

    Dim docStart As Long
    docStart = ActiveDocument.Content.Start

    Dim position As Long
    position = FieldNavigationStart(fld) - 1

    Dim caret As Range
    Do While position > docStart
        Set caret = ActiveDocument.Range(position, position)
        If caret.Fields.Count = 0 Then
            caret.Select
            Exit Sub
        End If
        position = position - 1
    Loop

    Set caret = ActiveDocument.Range(docStart, docStart)
    caret.Select
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.MoveCaretBeforeField"
End Sub

Private Function FieldNavigationStart(ByVal fld As Field) As Long
    On Error GoTo ResultOnly

    Dim resultStart As Long
    Dim codeStart As Long
    resultStart = fld.Result.Start
    codeStart = fld.Code.Start

    If codeStart < resultStart Then
        FieldNavigationStart = codeStart
    Else
        FieldNavigationStart = resultStart
    End If
    Exit Function

ResultOnly:
    On Error Resume Next
    FieldNavigationStart = fld.Result.Start
    On Error GoTo 0
End Function

Private Sub RefreshLogWarn(ByVal message As String)
    Debug.Print "[Banyan][Refresh] " & message
End Sub


' --- In-text refresh ---

Private Function RefreshIntextRange(ByVal targetRange As Range, _
                                    ByVal pref As Object, _
                                    ByVal syncItems As Boolean) As Boolean
    Dim pairs As Collection
    Set pairs = FieldCollectIntextCitationFieldsInRange(targetRange)
    If pairs.Count = 0 Then
        RefreshLogWarn "No in-text citations found; deleting existing bibliography and stopping refresh."
        RefreshIntextRange = DeleteExistingBibliography(targetRange)
        Exit Function
    End If

    Dim contexts As Collection
    Set contexts = New Collection

    Dim requestPairs As Collection
    Set requestPairs = New Collection

    Dim fd As Variant
    For Each fd In pairs
        Dim data As Object
        Dim fld As Field
        Set fld = fd("field")
        Set data = fd("data")

        Dim context As Object
        Set context = BuildCitationContext(fld, data)
        contexts.Add context
        requestPairs.Add MakeFieldContextPair(fld, context)
    Next fd

    Dim respond As Object
    Set respond = RequestRefresh(pref("style"), contexts, syncItems)
    If respond Is Nothing Then
        RefreshLogWarn "Could not get response from /refresh, skipping this chapter."
        Exit Function
    End If

    Dim didUpdateCitation As Boolean
    Dim pair As Variant
    For Each pair In requestPairs
        Dim updatedData As Object
        Set updatedData = FindCitationById(respond("citations"), DictKeyString(pair("context"), "id"))
        If updatedData Is Nothing Then
            RefreshLogWarn "No updated data found for citation with id " & DictKeyString(pair("context"), "id") & ", skipping."
        ElseIf Not FieldIsIntextCitation(updatedData) Then
            RefreshLogWarn "Updated data for citation with id " & DictKeyString(pair("context"), "id") & " is not a valid in-text citation, skipping."
        Else
            FieldWriteData pair("field"), updatedData
            FieldRenderStyledField pair("field"), updatedData("content")
            didUpdateCitation = True
        End If
    Next pair

    RefreshIntextRange = (didUpdateCitation Or RefreshBibliographyInRange(targetRange, respond, pref))
End Function


' --- Note refresh ---

Private Function RefreshNoteRange(ByVal targetRange As Range, _
                                  ByVal pref As Object, _
                                  ByVal syncItems As Boolean) As Boolean
    Dim pairs As Collection
    Set pairs = FieldCollectNoteCitationFootnotesInRange(targetRange)
    If pairs.Count = 0 Then
        RefreshLogWarn "No note citations found; deleting existing bibliography and stopping refresh."
        RefreshNoteRange = DeleteExistingBibliography(targetRange)
        Exit Function
    End If

    Dim contexts As Collection
    Set contexts = New Collection

    Dim requestPairs As Collection
    Set requestPairs = New Collection

    Dim fd As Variant
    For Each fd In pairs
        Dim data As Object
        Dim fld As Field
        Set fld = fd("field")
        Set data = fd("data")

        Dim context As Object
        Set context = BuildCitationContext(fld, data)
        contexts.Add context
        requestPairs.Add MakeNoteContextPair(fd("note"), fld, context)
    Next fd

    Dim respond As Object
    Set respond = RequestRefresh(pref("style"), contexts, syncItems)
    If respond Is Nothing Then
        RefreshLogWarn "Could not get response from /refresh, skipping this chapter."
        Exit Function
    End If

    Dim didUpdateCitation As Boolean
    Dim i As Long
    For i = requestPairs.Count To 1 Step -1
        Dim pair As Collection
        Set pair = requestPairs(i)

        Dim updatedData As Object
        Set updatedData = FindCitationById(respond("citations"), DictKeyString(pair("context"), "id"))
        If updatedData Is Nothing Then
            RefreshLogWarn "No updated data found for citation with id " & DictKeyString(pair("context"), "id") & ", skipping."
        ElseIf Not FieldIsNoteCitation(updatedData) Then
            RefreshLogWarn "Updated data for citation with id " & DictKeyString(pair("context"), "id") & " is not a valid note citation, skipping."
        Else
            ' Rebuild the footnote in place. VBA owns the custom footnote
            ' Reference (e.g. "[1]"), so the footnote is re-created to refresh
            ' the reference mark/number. The rebuild caches the footnote body
            ' first and replaces only the citation field, preserving any
            ' user-typed content around it.
            Dim rebuilt As Collection
            Set rebuilt = FieldRebuildNoteCitationAtRange(pair("note"), pair("field"), updatedData)
            If rebuilt Is Nothing Then
                RefreshLogWarn "Failed to rebuild note citation with id " & DictKeyString(pair("context"), "id") & ", skipping."
            Else
                didUpdateCitation = True
            End If
        End If
    Next i

    RefreshNoteRange = (didUpdateCitation Or RefreshBibliographyInRange(targetRange, respond, pref))
End Function


' --- Bibliography refresh ---

Private Function RefreshBibliographyInRange(ByVal targetRange As Range, _
                                            ByVal respond As Object, _
                                            ByVal pref As Object) As Boolean
    On Error GoTo ErrHandler

    If Not HasDictionaryKey(respond, "bibliography") Then Exit Function
    If Not IsCollectionObject(DictKeyObject(respond, "bibliography")) Then Exit Function
    If DictKeyObject(respond, "bibliography").Count = 0 Then Exit Function

    Dim bibliographyFields As Collection
    Set bibliographyFields = CollectBibliographyFieldsInRange(targetRange)
    If bibliographyFields.Count = 0 Then Exit Function

    Dim lines As Collection
    Set lines = New Collection

    Dim line As Variant
    For Each line In respond("bibliography")
        If FieldIsBibliographyTitle(line) Or FieldIsBibliographyEntry(line) Then
            lines.Add line
        End If
    Next line

    If lines.Count = 0 Then Exit Function

    Dim firstField As Field
    Set firstField = bibliographyFields(1)

    Dim caret As Range
    Set caret = firstField.Result.Duplicate
    caret.Collapse wdCollapseStart

    DeleteExistingBibliography targetRange

    Dim i As Long
    Dim data As Object
    Dim fld As Field
    For i = 1 To lines.Count
        Set data = lines(i)

        Dim fieldCode As String
        If FieldIsBibliographyEntry(data) Then
            fieldCode = "BANYAN_BIBLIOGRAPHY " & DictKeyString(data, "id")
        Else
            fieldCode = "BANYAN_BIBLIOGRAPHY"
        End If

        Set fld = FieldCreateRawAddinField(caret, fieldCode)
        If fld Is Nothing Then Exit Function

        FieldWriteData fld, data
        If FieldIsBibliographyTitle(data) Then
            FieldRenderStyledFieldWithStyle fld, DictKeyString(pref, "bibliographyTitleStyle"), wdStyleTypeParagraph, data("content")
        ElseIf FieldIsBibliographyEntry(data) Then
            FieldRenderStyledFieldWithStyle fld, DictKeyString(pref, "bibliographyEntryStyle"), wdStyleTypeParagraph, data("content")
            FieldAddBookmarkToField fld, FieldGetBibliographyBookmarkName(DictKeyString(data, "id"))
        End If

        ' Move caret to end of newly inserted field for next iteration
        Dim resultEnd As Range
        Set resultEnd = fld.Result.Duplicate
        resultEnd.Collapse wdCollapseEnd
        If i < lines.Count Then
            resultEnd.InsertParagraphAfter
            resultEnd.Collapse wdCollapseEnd
        End If
        caret.SetRange resultEnd.Start, resultEnd.End
    Next i
    RefreshBibliographyInRange = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.RefreshBibliographyInRange"
    RefreshBibliographyInRange = False
End Function

Private Function DeleteExistingBibliography(ByVal targetRange As Range) As Boolean
    On Error GoTo ErrHandler

    Dim bibliographyFields As Collection
    Set bibliographyFields = CollectBibliographyFieldsInRange(targetRange)
    If bibliographyFields.Count = 0 Then Exit Function

    Dim firstField As Field
    Set firstField = bibliographyFields(1)
    If firstField.Locked Then firstField.Locked = False
    firstField.Delete

    Dim i As Long
    Dim fld As Field
    For i = bibliographyFields.Count To 2 Step -1
        Set fld = bibliographyFields(i)
        If fld.Locked Then fld.Locked = False
        fld.Result.Paragraphs(1).Range.Delete
    Next i
    DeleteExistingBibliography = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.DeleteExistingBibliography"
    RefreshLogWarn "Failed to delete existing bibliography: " & Err.Description
    DeleteExistingBibliography = False
End Function

Private Function CollectBibliographyFieldsInRange(ByVal targetRange As Range) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim fld As Field
    Dim data As Object
    For Each fld In targetRange.Fields
        If fld.Type = wdFieldAddin Then
            Set data = FieldReadData(fld)
            If FieldIsBibliographyTitle(data) Or FieldIsBibliographyEntry(data) Then
                result.Add fld
            End If
        End If
    Next fld

    Set CollectBibliographyFieldsInRange = result
End Function

' --- HTTP ---

Private Function RequestRefresh(ByVal style As Object, _
                                ByVal contexts As Collection, _
                                ByVal syncItems As Boolean) As Object
    On Error GoTo ErrHandler

    Dim body As Object
    Set body = New Dictionary
    body("documentId") = GetDocumentId()
    Set body("style") = FieldAsStyleIdentifier(style)
    Set body("contexts") = contexts
    body("syncItems") = syncItems

    Dim respText As String
    respText = HttpPost(HttpBuildUrl("refresh"), JsonStringify(body))
    If Len(respText) = 0 Then Exit Function

    Dim envelope As Object
    Set envelope = JsonParse(respText)
    If envelope Is Nothing Then Exit Function
    If Not EnvelopeOk(envelope) Then Exit Function
    If Not HasDictionaryKey(envelope, "data") Then Exit Function
    If Not IsDictionaryRecord(envelope("data")) Then Exit Function

    Dim data As Object
    Set data = envelope("data")
    If Not HasDictionaryKey(data, "citations") Then Exit Function
    If Not HasDictionaryKey(data, "bibliography") Then Exit Function
    If Not IsCollectionObject(data("citations")) Then Exit Function
    If Not IsCollectionObject(data("bibliography")) Then Exit Function

    Set RequestRefresh = data
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.RequestRefresh"
    Set RequestRefresh = Nothing
End Function

Private Function EnvelopeOk(ByVal envelope As Object) As Boolean
    On Error GoTo ErrHandler
    If Not envelope.Exists("ok") Then Exit Function
    EnvelopeOk = DictKeyBool(envelope, "ok")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.EnvelopeOk"
    EnvelopeOk = False
End Function


' --- Context and lookup helpers ---

Private Function BuildCitationContext(ByVal fld As Field, ByVal data As Object) As Object
    Dim context As Object
    Set context = New Dictionary

    Dim source As Object
    Set source = data("source")

    Dim key As Variant
    For Each key In source.Keys
        DictCopyKey context, CStr(key), source, CStr(key)
    Next key

    context("id") = DictKeyString(data, "id")
    context("page") = FieldPageNumber(fld)
    Set BuildCitationContext = context
End Function

Private Function FieldPageNumber(ByVal fld As Field) As Long
    On Error GoTo ErrHandler
    FieldPageNumber = CLng(fld.Result.Information(wdActiveEndPageNumber))
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.FieldPageNumber"
    FieldPageNumber = 0
End Function

Private Function FindCitationById(ByVal citations As Collection, ByVal citationId As String) As Object
    On Error GoTo ErrHandler

    Dim item As Variant
    For Each item In citations
        If IsDictionaryRecord(item) Then
            If HasDictionaryKey(item, "id") Then
                If DictKeyString(item, "id") = citationId Then
                    Set FindCitationById = item
                    Exit Function
                End If
            End If
        End If
    Next item
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.FindCitationById"
    Set FindCitationById = Nothing
End Function

Private Function MakeFieldContextPair(ByVal fld As Field, ByVal context As Object) As Collection
    Dim result As Collection
    Set result = New Collection
    result.Add fld, "field"
    result.Add context, "context"
    Set MakeFieldContextPair = result
End Function

Private Function MakeNoteContextPair(ByVal note As Footnote, _
                                     ByVal fld As Field, _
                                     ByVal context As Object) As Collection
    Dim result As Collection
    Set result = New Collection
    result.Add note, "note"
    result.Add fld, "field"
    result.Add context, "context"
    Set MakeNoteContextPair = result
End Function

Private Function SameStyle(ByVal previousStyle As Object, ByVal nextStyle As Object) As Boolean
    On Error GoTo ErrHandler
    If previousStyle Is Nothing Then Exit Function
    If nextStyle Is Nothing Then Exit Function

    SameStyle = (DictKeyString(previousStyle, "id") = DictKeyString(nextStyle, "id") And _
                 DictKeyString(previousStyle, "title") = DictKeyString(nextStyle, "title") And _
                 DictKeyString(previousStyle, "citationType") = DictKeyString(nextStyle, "citationType"))
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.SameStyle"
    SameStyle = False
End Function


' --- Type helpers ---

Private Function IsDictionaryRecord(ByVal value As Variant) As Boolean
    IsDictionaryRecord = DictIsDictionary(value)
End Function

Private Function IsCollectionObject(ByVal value As Variant) As Boolean
    IsCollectionObject = DictIsCollection(value)
End Function

Private Function HasDictionaryKey(ByVal dict As Object, ByVal key As String) As Boolean
    HasDictionaryKey = DictHasKey(dict, key)
End Function


' --- Local i18n ---

Private Sub EnsureRefreshI10n()
    If m_i18nReady Then Exit Sub

    I10nRegisterTable msoLanguageIDSimplifiedChinese, "refresh", _
        "dialogTitle", "Banyan 刷新", _
        "error", "刷新 Banyan 内容失败：{message}", _
        "progressReason", "正在刷新 Banyan 内容..."

    I10nRegisterTable msoLanguageIDEnglishUS, "refresh", _
        "dialogTitle", "Banyan Refresh", _
        "error", "Failed to refresh Banyan fields: {message}", _
        "progressReason", "Refreshing Banyan fields..."

    m_i18nReady = True
End Sub

Private Function RText(ByVal key As String, ByVal fallback As String) As String
    EnsureRefreshI10n
    RText = T("refresh." & key, fallback)
End Function
