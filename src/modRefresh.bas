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

    Dim batchStarted As Boolean
    FieldBeginBatchUpdate
    batchStarted = True

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
    If batchStarted Then FieldEndBatchUpdate
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
    If batchStarted Then FieldEndBatchUpdate
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

    Dim batchStarted As Boolean
    FieldBeginBatchUpdate
    batchStarted = True

    Dim pref As Object
    Set pref = PreferenceEnsure()
    If pref Is Nothing Then
        RefreshLogWarn "No style found, stopping refresh."
        GoTo CleanUp
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
    GoTo CleanUp

CleanUp:
    If batchStarted Then FieldEndBatchUpdate
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.RefreshInRange"
    If batchStarted Then FieldEndBatchUpdate
    RefreshInRange = False
End Function

Public Sub RefreshAll(Optional ByVal syncItemsOverride As Variant)
    On Error GoTo CleanUp

    Dim originalRange As Range
    Set originalRange = Selection.Range.Duplicate

    Dim batchStarted As Boolean
    FieldBeginBatchUpdate
    batchStarted = True

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
    If batchStarted Then FieldEndBatchUpdate
End Sub

Public Function RefreshForStyleChange(ByVal previousStyle As Object, ByVal nextStyle As Object) As Boolean
    On Error GoTo ErrHandler

    If SameStyle(previousStyle, nextStyle) Then Exit Function

    Dim originalRange As Range
    On Error Resume Next
    Set originalRange = Selection.Range.Duplicate
    On Error GoTo ErrHandler

    Dim batchStarted As Boolean
    FieldBeginBatchUpdate
    batchStarted = True

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
    If batchStarted Then FieldEndBatchUpdate
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.RefreshForStyleChange"
    ProgressClose
    RestoreRefreshSelection originalRange
    If batchStarted Then FieldEndBatchUpdate
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

    ' Keep the initial field/data objects for this refresh call only.
    Dim fieldSnapshot As Object
    Set fieldSnapshot = BuildCitationFieldSnapshot(pairs)

    Dim fd As Variant
    For Each fd In pairs
        Dim data As Object
        Dim fld As Field
        Set fld = fd("field")
        Set data = fd("data")

        Dim context As Object
        Set context = BuildCitationContext(fld, data)
        contexts.Add context
        requestPairs.Add MakeFieldContextPair(fld, context, data)
    Next fd

    Dim respond As Object
    Set respond = RequestRefresh(pref("style"), contexts, syncItems)
    If respond Is Nothing Then
        RefreshLogWarn "Could not get response from /refresh, skipping this chapter."
        Set fieldSnapshot = Nothing
        Exit Function
    End If

    Dim responseIndex As Object
    Set responseIndex = BuildCitationResponseIndex(respond("citations"))

    Dim didUpdateCitation As Boolean
    Dim pair As Variant
    For Each pair In requestPairs
        Dim currentContent As Object
        Dim snapshotPair As Collection
        Set snapshotPair = FindCitationSnapshot(fieldSnapshot, pair)
        If snapshotPair Is Nothing Then
            Set currentContent = DictKeyObject(pair("data"), "content")
        Else
            Set currentContent = snapshotPair("content")
        End If
        Dim updatedData As Object
        Set updatedData = FindCitationById(responseIndex, DictKeyString(pair("context"), "id"))
        If updatedData Is Nothing Then
            RefreshLogWarn "No updated data found for citation with id " & DictKeyString(pair("context"), "id") & ", skipping."
        ElseIf Not FieldIsIntextCitation(updatedData) Then
            RefreshLogWarn "Updated data for citation with id " & DictKeyString(pair("context"), "id") & " is not a valid in-text citation, skipping."
        Else
            Dim contentChanged As Boolean
            Dim updatedContent As Object
            Set updatedContent = DictKeyObject(updatedData, "content")
            contentChanged = Not FieldRichTextEquals(currentContent, updatedContent)
            Dim targetField As Field
            If snapshotPair Is Nothing Then
                Set targetField = pair("field")
            Else
                Set targetField = snapshotPair("field")
            End If

            ' Source is authoritative response data but does not itself decide
            ' the Word result. Always persist it; render only when content differs.
            If FieldWriteData(targetField, updatedData) Then
                didUpdateCitation = True
            Else
                RefreshLogWarn "Failed to write updated data for citation with id " & DictKeyString(pair("context"), "id") & ", skipping render."
                GoTo NextIntextCitation
            End If
            If contentChanged Then
                FieldRenderStyledFieldWithData targetField, updatedData, DictKeyObject(updatedData, "content")
            End If
        End If
NextIntextCitation:
    Next pair

    RefreshIntextRange = (didUpdateCitation Or RefreshBibliographyInRange(targetRange, respond, pref))
    Set responseIndex = Nothing
    Set fieldSnapshot = Nothing
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

    ' Rebuilding notes changes live ranges; retain the initial field/data
    ' objects only for this refresh call.
    Dim fieldSnapshot As Object
    Set fieldSnapshot = BuildCitationFieldSnapshot(pairs)

    Dim fd As Variant
    For Each fd In pairs
        Dim data As Object
        Dim fld As Field
        Set fld = fd("field")
        Set data = fd("data")

        Dim context As Object
        Set context = BuildCitationContext(fld, data)
        contexts.Add context
        requestPairs.Add MakeNoteContextPair(fd("note"), fld, context, data)
    Next fd

    Dim respond As Object
    Set respond = RequestRefresh(pref("style"), contexts, syncItems)
    If respond Is Nothing Then
        RefreshLogWarn "Could not get response from /refresh, skipping this chapter."
        Set fieldSnapshot = Nothing
        Exit Function
    End If

    Dim responseIndex As Object
    Set responseIndex = BuildCitationResponseIndex(respond("citations"))

    Dim didUpdateCitation As Boolean
    Dim i As Long
    For i = requestPairs.Count To 1 Step -1
        Dim pair As Collection
        Set pair = requestPairs(i)

        Dim updatedData As Object
        Set updatedData = FindCitationById(responseIndex, DictKeyString(pair("context"), "id"))
        If updatedData Is Nothing Then
            RefreshLogWarn "No updated data found for citation with id " & DictKeyString(pair("context"), "id") & ", skipping."
        ElseIf Not FieldIsNoteCitation(updatedData) Then
            RefreshLogWarn "Updated data for citation with id " & DictKeyString(pair("context"), "id") & " is not a valid note citation, skipping."
        Else
            Dim currentContent As Object
            Dim currentReference As Object
            Dim snapshotPair As Collection
            Set snapshotPair = FindCitationSnapshot(fieldSnapshot, pair)
            If snapshotPair Is Nothing Then
                Set currentContent = DictKeyObject(pair("data"), "content")
                Set currentReference = DictKeyObject(pair("data"), "reference")
            Else
                Set currentContent = snapshotPair("content")
                Set currentReference = snapshotPair("reference")
            End If

            Dim targetNote As Footnote
            Dim targetField As Field
            Set targetNote = pair("note")
            Set targetField = pair("field")
            If Not snapshotPair Is Nothing Then
                Set targetNote = snapshotPair("note")
                Set targetField = snapshotPair("field")
            End If

            Dim presentationChanged As Boolean
            presentationChanged = Not FieldRichTextEquals(currentContent, DictKeyObject(updatedData, "content"))
            If Not presentationChanged Then
                presentationChanged = Not FieldRichTextEquals(currentReference, DictKeyObject(updatedData, "reference"))
            End If
            If presentationChanged Then
                ' Content changes require a clean field replacement; reference
                ' changes additionally require footnote recreation.
                Dim rebuilt As Collection
                Set rebuilt = FieldRebuildNoteCitationAtRange(targetNote, targetField, updatedData)
                If rebuilt Is Nothing Then
                    RefreshLogWarn "Failed to rebuild note citation with id " & DictKeyString(pair("context"), "id") & ", skipping."
                Else
                    didUpdateCitation = True
                End If
            ElseIf FieldWriteData(targetField, updatedData) Then
                ' A source-only change is persisted without touching the field
                ' result or footnote structure.
                didUpdateCitation = True
            Else
                RefreshLogWarn "Failed to write updated data for note citation with id " & DictKeyString(pair("context"), "id") & "."
            End If
        End If
    Next i

    RefreshNoteRange = (didUpdateCitation Or RefreshBibliographyInRange(targetRange, respond, pref))
    Set responseIndex = Nothing
    Set fieldSnapshot = Nothing
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

    Dim contentChanged As Boolean
    Dim metadataChanged As Boolean
    contentChanged = BibliographyContentChanged(bibliographyFields, lines, pref, metadataChanged)
    If Not contentChanged Then
        If metadataChanged Then
            UpdateBibliographyMetadata bibliographyFields, lines
            RefreshBibliographyInRange = True
        End If
        Exit Function
    End If

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

Private Function BibliographyContentChanged(ByVal fields As Collection, _
                                            ByVal lines As Collection, _
                                            ByVal pref As Object, _
                                            ByRef metadataChanged As Boolean) As Boolean
    On Error GoTo ErrHandler
    If fields.Count <> lines.Count Then
        BibliographyContentChanged = True
        Exit Function
    End If

    Dim i As Long
    Dim currentData As Object
    Dim nextData As Object
    For i = 1 To lines.Count
        Set currentData = FieldReadData(fields(i))
        Set nextData = lines(i)
        If currentData Is Nothing Then
            BibliographyContentChanged = True
            Exit Function
        End If
        If Not FieldContentEquals(currentData, nextData) Then
            BibliographyContentChanged = True
            Exit Function
        End If
        If Not BibliographyFieldIdentityEquals(currentData, nextData) Then
            BibliographyContentChanged = True
            Exit Function
        End If
        If Not BibliographyStyleEquals(fields(i), nextData, pref) Then
            BibliographyContentChanged = True
            Exit Function
        End If
        If Not FieldDataEquals(currentData, nextData) Then
            metadataChanged = True
        End If
    Next i
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.BibliographyContentChanged"
    BibliographyContentChanged = True
End Function

Private Function BibliographyStyleEquals(ByVal fld As Field, _
                                         ByVal data As Object, _
                                         ByVal pref As Object) As Boolean
    On Error GoTo ErrHandler

    Dim expectedStyle As String
    If FieldIsBibliographyTitle(data) Then
        expectedStyle = DictKeyString(pref, "bibliographyTitleStyle")
    Else
        expectedStyle = DictKeyString(pref, "bibliographyEntryStyle")
    End If
    If Len(expectedStyle) = 0 Then Exit Function

    BibliographyStyleEquals = (fld.Result.Style.NameLocal = expectedStyle)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.BibliographyStyleEquals"
    BibliographyStyleEquals = False
End Function

Private Function BibliographyFieldIdentityEquals(ByVal currentData As Object, _
                                                 ByVal nextData As Object) As Boolean
    If FieldIsBibliographyTitle(currentData) And FieldIsBibliographyTitle(nextData) Then
        BibliographyFieldIdentityEquals = True
    ElseIf FieldIsBibliographyEntry(currentData) And FieldIsBibliographyEntry(nextData) Then
        BibliographyFieldIdentityEquals = (DictKeyString(currentData, "id") = DictKeyString(nextData, "id"))
    End If
End Function

Private Sub UpdateBibliographyMetadata(ByVal fields As Collection, _
                                       ByVal lines As Collection)
    On Error GoTo ErrHandler

    Dim i As Long
    Dim currentData As Object
    Dim nextData As Object
    For i = 1 To lines.Count
        Set currentData = FieldReadData(fields(i))
        Set nextData = lines(i)
        If Not FieldDataEquals(currentData, nextData) Then
            FieldWriteData fields(i), nextData
            If FieldIsBibliographyEntry(nextData) Then
                FieldAddBookmarkToField fields(i), FieldGetBibliographyBookmarkName(DictKeyString(nextData, "id"))
            End If
        End If
    Next i
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.UpdateBibliographyMetadata"
End Sub

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
            If Not FieldCodeHasPrefix(fld, "BANYAN_BIBLIOGRAPHY") Then GoTo NextBibliographyField
            Set data = FieldReadData(fld)
            If FieldIsBibliographyTitle(data) Or FieldIsBibliographyEntry(data) Then
                result.Add fld
            End If
        End If
NextBibliographyField:
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

Private Function FindCitationById(ByVal citationIndex As Object, ByVal citationId As String) As Object
    On Error GoTo ErrHandler

    If citationIndex Is Nothing Then Exit Function
    If citationIndex.Exists(citationId) Then Set FindCitationById = citationIndex(citationId)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.FindCitationById"
    Set FindCitationById = Nothing
End Function

Private Function BuildCitationResponseIndex(ByVal citations As Collection) As Object
    On Error GoTo ErrHandler

    Dim result As Object
    Set result = New Dictionary

    Dim item As Variant
    For Each item In citations
        If IsDictionaryRecord(item) Then
            If HasDictionaryKey(item, "id") Then
                Dim citationId As String
                citationId = DictKeyString(item, "id")
                If Len(citationId) > 0 Then Set result(citationId) = item
            End If
        End If
    Next item

    Set BuildCitationResponseIndex = result
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.BuildCitationResponseIndex"
    Set BuildCitationResponseIndex = Nothing
End Function

Private Function BuildCitationFieldSnapshot(ByVal pairs As Collection) As Object
    On Error GoTo ErrHandler

    Dim result As Object
    Set result = New Dictionary

    Dim pair As Variant
    For Each pair In pairs
        Dim dataId As String
        dataId = DictKeyString(pair("data"), "id")
        If Len(dataId) > 0 Then Set result(dataId) = pair
    Next pair

    Set BuildCitationFieldSnapshot = result
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.BuildCitationFieldSnapshot"
    Set BuildCitationFieldSnapshot = Nothing
End Function

Private Function FindCitationSnapshot(ByVal snapshot As Object, _
                                      ByVal pair As Collection) As Collection
    On Error GoTo ErrHandler
    If snapshot Is Nothing Then Exit Function

    Dim citationId As String
    citationId = DictKeyString(pair("context"), "id")
    If Len(citationId) > 0 Then
        If snapshot.Exists(citationId) Then Set FindCitationSnapshot = snapshot(citationId)
    End If
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.FindCitationSnapshot"
    Set FindCitationSnapshot = Nothing
End Function

Private Function MakeFieldContextPair(ByVal fld As Field, _
                                      ByVal context As Object, _
                                      ByVal data As Object) As Collection
    Dim result As Collection
    Set result = New Collection
    result.Add fld, "field"
    result.Add context, "context"
    result.Add data, "data"
    Set MakeFieldContextPair = result
End Function

Private Function MakeNoteContextPair(ByVal note As Footnote, _
                                     ByVal fld As Field, _
                                     ByVal context As Object, _
                                     ByVal data As Object) As Collection
    Dim result As Collection
    Set result = New Collection
    result.Add note, "note"
    result.Add fld, "field"
    result.Add context, "context"
    result.Add data, "data"
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
