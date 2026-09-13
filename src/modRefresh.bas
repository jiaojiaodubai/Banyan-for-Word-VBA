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

    If Not DictHasKey(respond, "bibliography") Then Exit Function
    If Not DictIsCollection(DictKeyObject(respond, "bibliography")) Then Exit Function
    If DictKeyObject(respond, "bibliography").Count = 0 Then Exit Function

    Dim lines As Collection
    Set lines = New Collection

    Dim line As Variant
    For Each line In respond("bibliography")
        If FieldIsBibliographyTitle(line) Or FieldIsBibliographyEntry(line) Then
            lines.Add line
        End If
    Next line

    If lines.Count = 0 Then Exit Function

    RefreshBibliographyInRange = RefreshBibliographyLinesInRange(targetRange, lines, pref)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.RefreshBibliographyInRange"
    RefreshBibliographyInRange = False
End Function

Public Function RefreshBibliographyLinesInRange(ByVal targetRange As Range, _
                                                 ByVal lines As Collection, _
                                                 ByVal pref As Object) As Boolean
    On Error GoTo ErrHandler
    If targetRange Is Nothing Or lines Is Nothing Or pref Is Nothing Then Exit Function
    If lines.Count = 0 Then Exit Function

    Dim fields As Collection
    Set fields = CollectBibliographyFieldsInRange(targetRange)
    If fields.Count = 0 Then Exit Function

    Dim currentIds As Collection
    Dim nextPositions As Object
    If Not BuildBibliographyIndexes(fields, lines, currentIds, nextPositions) Then Exit Function

    Dim structureChanged As Boolean
    structureChanged = BibliographyStructureChanged(currentIds, lines)

    Dim matchedOld As Collection
    Dim matchedNext As Collection
    FindBibliographyMatches currentIds, nextPositions, matchedOld, matchedNext

    If structureChanged Then
        If matchedOld.Count = 0 Then
            RefreshBibliographyLinesInRange = ReplaceBibliography(fields, lines, pref)
            Exit Function
        End If
        If Not PatchBibliographyGaps(fields, lines, pref, matchedOld, matchedNext) Then Exit Function
    End If

    Dim didChange As Boolean
    Dim i As Long
    For i = 1 To matchedOld.Count
        Dim nextData As Object
        Set nextData = lines(CLng(matchedNext(i)))
        If UpdateBibliographyField(fields(CLng(matchedOld(i))), nextData, pref) Then
            didChange = True
            If FieldIsBibliographyEntry(nextData) Then
                FieldAddBookmarkToField fields(CLng(matchedOld(i))), _
                    FieldGetBibliographyBookmarkName(DictKeyString(nextData, "id"))
            End If
        End If
    Next i

    RefreshBibliographyLinesInRange = (structureChanged Or didChange)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.RefreshBibliographyLinesInRange"
    RefreshBibliographyLinesInRange = False
End Function

Private Function BuildBibliographyIndexes(ByVal fields As Collection, _
                                           ByVal lines As Collection, _
                                           ByRef currentIds As Collection, _
                                           ByRef nextPositions As Object) As Boolean
    On Error GoTo ErrHandler
    Set currentIds = New Collection
    Set nextPositions = New Dictionary

    Dim i As Long
    Dim data As Object
    Dim lineId As String
    For i = 1 To lines.Count
        Set data = lines(i)
        lineId = Trim$(DictKeyString(data, "id"))
        If Len(lineId) = 0 Or nextPositions.Exists(lineId) Then Exit Function
        nextPositions(lineId) = i
    Next i

    Dim currentIdIndex As Object
    Set currentIdIndex = New Dictionary
    For i = 1 To fields.Count
        lineId = BibliographyFieldId(fields(i))
        If Len(lineId) = 0 Or currentIdIndex.Exists(lineId) Then Exit Function
        currentIdIndex(lineId) = True
        currentIds.Add lineId
    Next i

    BuildBibliographyIndexes = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.BuildBibliographyIndexes"
    BuildBibliographyIndexes = False
End Function

Private Function BibliographyFieldId(ByVal fld As Field) As String
    On Error GoTo ErrHandler
    Const CODE_PREFIX As String = "ADDIN BANYAN_BIBLIOGRAPHY "
    Dim codeText As String
    codeText = Trim$(fld.Code.Text)
    If StrComp(Left$(codeText, Len(CODE_PREFIX)), CODE_PREFIX, vbTextCompare) <> 0 Then Exit Function
    BibliographyFieldId = Trim$(Mid$(codeText, Len(CODE_PREFIX) + 1))
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.BibliographyFieldId"
    BibliographyFieldId = ""
End Function

Private Sub FindBibliographyMatches(ByVal currentIds As Collection, _
                                    ByVal nextPositions As Object, _
                                    ByRef matchedOld As Collection, _
                                    ByRef matchedNext As Collection)
    Set matchedOld = New Collection
    Set matchedNext = New Collection

    Dim commonCount As Long
    Dim i As Long
    For i = 1 To currentIds.Count
        If nextPositions.Exists(CStr(currentIds(i))) Then commonCount = commonCount + 1
    Next i
    If commonCount = 0 Then Exit Sub

    Dim sequence() As Long
    Dim oldIndexes() As Long
    Dim previous() As Long
    Dim tails() As Long
    Dim tailSequenceIndexes() As Long
    ReDim sequence(1 To commonCount)
    ReDim oldIndexes(1 To commonCount)
    ReDim previous(1 To commonCount)
    ReDim tails(1 To commonCount)
    ReDim tailSequenceIndexes(1 To commonCount)

    Dim sequenceIndex As Long
    For i = 1 To currentIds.Count
        Dim lineId As String
        lineId = CStr(currentIds(i))
        If nextPositions.Exists(lineId) Then
            sequenceIndex = sequenceIndex + 1
            sequence(sequenceIndex) = CLng(nextPositions(lineId))
            oldIndexes(sequenceIndex) = i
        End If
    Next i

    Dim lisLength As Long
    For sequenceIndex = 1 To commonCount
        Dim low As Long
        Dim high As Long
        low = 1
        high = lisLength
        Do While low <= high
            Dim middle As Long
            middle = (low + high) \ 2
            If tails(middle) < sequence(sequenceIndex) Then
                low = middle + 1
            Else
                high = middle - 1
            End If
        Loop

        Dim lengthAtItem As Long
        lengthAtItem = low
        tails(lengthAtItem) = sequence(sequenceIndex)
        tailSequenceIndexes(lengthAtItem) = sequenceIndex
        If lengthAtItem > 1 Then previous(sequenceIndex) = tailSequenceIndexes(lengthAtItem - 1)
        If lengthAtItem > lisLength Then lisLength = lengthAtItem
    Next sequenceIndex

    Dim reverseOld() As Long
    Dim reverseNext() As Long
    ReDim reverseOld(1 To lisLength)
    ReDim reverseNext(1 To lisLength)
    sequenceIndex = tailSequenceIndexes(lisLength)
    For i = lisLength To 1 Step -1
        reverseOld(i) = oldIndexes(sequenceIndex)
        reverseNext(i) = sequence(sequenceIndex)
        sequenceIndex = previous(sequenceIndex)
    Next i
    For i = 1 To lisLength
        matchedOld.Add reverseOld(i)
        matchedNext.Add reverseNext(i)
    Next i
End Sub

Private Function BibliographyStructureChanged(ByVal currentIds As Collection, _
                                               ByVal lines As Collection) As Boolean
    If currentIds.Count <> lines.Count Then
        BibliographyStructureChanged = True
        Exit Function
    End If

    Dim i As Long
    For i = 1 To lines.Count
        If CStr(currentIds(i)) <> DictKeyString(lines(i), "id") Then
            BibliographyStructureChanged = True
            Exit Function
        End If
    Next i
End Function

Private Function UpdateBibliographyField(ByVal fld As Field, _
                                          ByVal nextData As Object, _
                                          ByVal pref As Object) As Boolean
    On Error GoTo ErrHandler
    Dim nextJson As String
    nextJson = JsonStringify(nextData)
    If Len(nextJson) = 0 Then Exit Function
    If fld.Data = nextJson Then Exit Function

    Dim currentData As Object
    Set currentData = FieldReadData(fld)
    If currentData Is Nothing Then Exit Function
    Dim renderChanged As Boolean
    renderChanged = Not FieldContentEquals(currentData, nextData)

    If Not FieldWriteData(fld, nextData) Then Exit Function
    UpdateBibliographyField = True

    If renderChanged Then
        If FieldIsBibliographyTitle(nextData) Then
            FieldRenderStyledFieldWithStyle fld, DictKeyString(pref, "bibliographyTitleStyle"), wdStyleTypeParagraph, nextData("content")
        Else
            FieldRenderStyledFieldWithStyle fld, DictKeyString(pref, "bibliographyEntryStyle"), wdStyleTypeParagraph, nextData("content")
        End If
        UpdateBibliographyField = True
    End If
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.UpdateBibliographyField"
    UpdateBibliographyField = False
End Function

Private Function CreateBibliographyFieldAtRange(ByVal targetRange As Range, _
                                                 ByVal data As Object) As Field
    Set CreateBibliographyFieldAtRange = FieldCreateRawAddinField( _
        targetRange, "BANYAN_BIBLIOGRAPHY " & DictKeyString(data, "id"))
End Function

Private Function BibliographyWholeFieldRange(ByVal fld As Field) As Range
    On Error GoTo ErrHandler
    Set BibliographyWholeFieldRange = fld.Result.Document.Range(fld.Code.Start - 1, fld.Result.End + 1)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.BibliographyWholeFieldRange"
    Set BibliographyWholeFieldRange = Nothing
End Function

Private Function PatchBibliographyGaps(ByVal fields As Collection, _
                                       ByVal lines As Collection, _
                                       ByVal pref As Object, _
                                       ByVal matchedOld As Collection, _
                                       ByVal matchedNext As Collection) As Boolean
    On Error GoTo ErrHandler
    Dim gap As Long
    For gap = matchedOld.Count To 0 Step -1
        Dim previousOld As Long
        Dim previousNext As Long
        Dim followingOld As Long
        Dim followingNext As Long
        If gap = 0 Then
            previousOld = 0
            previousNext = 0
        Else
            previousOld = CLng(matchedOld(gap))
            previousNext = CLng(matchedNext(gap))
        End If
        If gap = matchedOld.Count Then
            followingOld = fields.Count + 1
            followingNext = lines.Count + 1
        Else
            followingOld = CLng(matchedOld(gap + 1))
            followingNext = CLng(matchedNext(gap + 1))
        End If

        Dim oldGapCount As Long
        Dim nextGapCount As Long
        oldGapCount = followingOld - previousOld - 1
        nextGapCount = followingNext - previousNext - 1
        If oldGapCount = 0 And nextGapCount = 0 Then GoTo NextGap

        Dim gapStart As Long
        Dim gapEnd As Long
        If previousOld = 0 Then
            gapStart = BibliographyWholeFieldRange(fields(1)).Start
        Else
            gapStart = BibliographyWholeFieldRange(fields(previousOld)).End
        End If
        If followingOld = fields.Count + 1 Then
            gapEnd = BibliographyWholeFieldRange(fields(fields.Count)).End
        Else
            gapEnd = BibliographyWholeFieldRange(fields(followingOld)).Start
        End If

        Dim cursor As Range
        Set cursor = fields(1).Result.Document.Range(gapStart, gapEnd)
        cursor.Delete
        cursor.Collapse wdCollapseStart
        If Not InsertBibliographyLines(cursor, lines, previousNext + 1, followingNext - 1, pref, _
                                       previousOld > 0, _
                                       followingOld <= fields.Count And nextGapCount > 0) Then Exit Function
NextGap:
    Next gap

    PatchBibliographyGaps = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.PatchBibliographyGaps"
    PatchBibliographyGaps = False
End Function

Private Function ReplaceBibliography(ByVal fields As Collection, _
                                     ByVal lines As Collection, _
                                     ByVal pref As Object) As Boolean
    On Error GoTo ErrHandler
    Dim caret As Range
    Set caret = BibliographyWholeFieldRange(fields(1))
    caret.Collapse wdCollapseStart

    Dim blockEnd As Range
    Set blockEnd = BibliographyWholeFieldRange(fields(fields.Count))
    caret.Document.Range(caret.Start, blockEnd.End).Delete
    If Not InsertBibliographyLines(caret, lines, 1, lines.Count, pref, False, False) Then Exit Function
    ReplaceBibliography = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.ReplaceBibliography"
    ReplaceBibliography = False
End Function

Private Function InsertBibliographyLines(ByVal cursor As Range, _
                                         ByVal lines As Collection, _
                                         ByVal firstIndex As Long, _
                                         ByVal lastIndex As Long, _
                                         ByVal pref As Object, _
                                         ByVal separatorBeforeFirst As Boolean, _
                                         ByVal separatorAfterLast As Boolean) As Boolean
    On Error GoTo ErrHandler
    Dim i As Long
    For i = firstIndex To lastIndex
        If separatorBeforeFirst Or i > firstIndex Then
            cursor.InsertAfter vbCr
            cursor.Collapse wdCollapseEnd
        End If

        Dim data As Object
        Dim fld As Field
        Set data = lines(i)
        Set fld = CreateBibliographyFieldAtRange(cursor, data)
        If fld Is Nothing Then Exit Function
        If Not InitializeBibliographyField(fld, data, pref) Then Exit Function
        If FieldIsBibliographyEntry(data) Then
            FieldAddBookmarkToField fld, FieldGetBibliographyBookmarkName(DictKeyString(data, "id"))
        End If
        Set cursor = BibliographyWholeFieldRange(fld)
        cursor.Collapse wdCollapseEnd
    Next i
    If separatorAfterLast Then
        cursor.InsertAfter vbCr
        cursor.Collapse wdCollapseEnd
    End If
    InsertBibliographyLines = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.InsertBibliographyLines"
    InsertBibliographyLines = False
End Function

Private Function InitializeBibliographyField(ByVal fld As Field, _
                                              ByVal data As Object, _
                                              ByVal pref As Object) As Boolean
    On Error GoTo ErrHandler
    If Not FieldWriteData(fld, data) Then Exit Function
    If FieldIsBibliographyTitle(data) Then
        InitializeBibliographyField = FieldRenderStyledFieldWithStyle( _
            fld, DictKeyString(pref, "bibliographyTitleStyle"), wdStyleTypeParagraph, data("content"))
    Else
        InitializeBibliographyField = FieldRenderStyledFieldWithStyle( _
            fld, DictKeyString(pref, "bibliographyEntryStyle"), wdStyleTypeParagraph, data("content"))
    End If
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modRefresh.InitializeBibliographyField"
    InitializeBibliographyField = False
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
    For Each fld In targetRange.Fields
        If fld.Type = wdFieldAddin Then
            If Not FieldCodeHasPrefix(fld, "BANYAN_BIBLIOGRAPHY") Then GoTo NextBibliographyField
            result.Add fld
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
    If Not DictHasKey(envelope, "data") Then Exit Function
    If Not DictIsDictionary(envelope("data")) Then Exit Function

    Dim data As Object
    Set data = envelope("data")
    If Not DictHasKey(data, "citations") Then Exit Function
    If Not DictHasKey(data, "bibliography") Then Exit Function
    If Not DictIsCollection(data("citations")) Then Exit Function
    If Not DictIsCollection(data("bibliography")) Then Exit Function

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
        If DictIsDictionary(item) Then
            If DictHasKey(item, "id") Then
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
