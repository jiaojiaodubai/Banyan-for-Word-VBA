Option Explicit

' ============================================================================
' Module  : modField
' Purpose : Shared Banyan ADDIN field utilities.
'
'           This module intentionally prefixes public members with Field* to
'           avoid name collisions with the older public helpers in
'           modChapterBreak.
'
' Public API:
'   FieldReadData(field) / FieldWriteData(field, data)
'   FieldIsIntextCitation(data) / FieldIsNoteCitation(data)
'   FieldParseCode(codeText, kind, id) / FieldCitationCode(id)
'   FieldBibliographyCode(id) / FieldTextEquals(currentText, nextText)
'   FieldCreateIntextCitationAtRange(range, data)
'   FieldCreateNoteCitationAtRange(range, data) -> Collection {note, field}
'   FieldRebuildNoteCitationAtRange(note, field, data) -> Collection {note, field}
'     - the citation field is ALWAYS rebuilt fresh (FieldReplaceCitationInNote) so
'       the result never inherits the placeholder's red; user content around the
'       field is preserved untouched
'     - reference unchanged (incl. empty/auto): keeps the footnote, replaces only
'       the field in place
'     - reference changed: recreates the footnote just after the old reference,
'       copies the whole body via FormattedText (rich text + field structure),
'       deletes the old footnote, then replaces the copied field fresh
'   FieldRenderStyledField(field, [content])
'   FieldRenderStyledFieldWithStyle(field, styleName, [styleType], [content])
'   FieldCollectIntextCitationFieldsInRange(range) -> [{id, field}]
'   FieldCollectNoteCitationFootnotesInRange(range) -> [{id, note, field}]
'   FieldMigrateIntextCitationsToNotes(range)
'   FieldMigrateNoteCitationsToIntext(range)
'   FieldRemoveFieldSafely(field) / FieldRemoveFootnoteSafely(footnote)
'   FieldAddBookmarkToField(field, bookmarkName)
'   FieldNormalizeBookmarkName(rawName)
'   FieldGetBibliographyBookmarkName(entryId)
' ============================================================================

Private Const FIELD_PLACEHOLDER_COLOR As String = "#ff0000"
Private Const FIELD_RAW_PLACEHOLDER As String = "{Citation}"

' Word bookmark rules enforced by FieldNormalizeBookmarkName.
Private Const FIELD_BOOKMARK_MAX_LENGTH As Long = 40
Private Const FIELD_BIBLIOGRAPHY_BOOKMARK_PREFIX As String = "Banyan_Entry_"

Private Const INTEXT_STYLE_ZH As String = "Banyan 引注"
Private Const INTEXT_STYLE_EN As String = "Banyan Citation"

' --- Field code contract ----------------------------------------------------
'   BANYAN_CITATION <id>      in-text citation or note citation
'   BANYAN_BIBLIOGRAPHY <id>  bibliography title line or entry line
'   <chapter prompt>          chapter break
' Every code carries the data id, pending placeholders included; a matching code
' is trusted, so collecting never parses Field.Data. A typed code WITHOUT an id
' is a broken field and the collectors drop it.
' Keep these declarations above the first procedure: the VBE does not register a
' module-level Const declared below one.
Private Const FIELD_CODE_CITATION As String = "BANYAN_CITATION"
Private Const FIELD_CODE_BIBLIOGRAPHY As String = "BANYAN_BIBLIOGRAPHY"
Private Const FIELD_CODE_CHAPTER_ZH As String = "Banyan章节分隔符"
Private Const FIELD_CODE_CHAPTER_EN As String = "Banyan chapter break"

' Kinds returned by FieldParseCode.
Public Const FIELD_KIND_CITATION As String = "citation"
Public Const FIELD_KIND_BIBLIOGRAPHY As String = "bibliography"
Public Const FIELD_KIND_CHAPTER As String = "chapter"

' VBA-only process state. The cache is scoped to the active Document and the
' batch depth makes nested refresh calls restore ScreenUpdating exactly once.
Private m_styleCache As Object
Private m_styleCacheDocument As Document
Private m_batchUpdateDepth As Long
Private m_batchScreenUpdating As Boolean
Private m_batchUndoRecord As Object
Private m_batchUndoStarted As Boolean


' --- VBA global performance state ------------------------------------------

Public Sub FieldBeginBatchUpdate()
    On Error GoTo ErrHandler

    If m_batchUpdateDepth = 0 Then
        m_batchScreenUpdating = Application.ScreenUpdating
        Application.ScreenUpdating = False
        FieldBeginCustomUndoRecord
    End If
    m_batchUpdateDepth = m_batchUpdateDepth + 1
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldBeginBatchUpdate"
End Sub

Public Sub FieldEndBatchUpdate()
    On Error Resume Next
    If m_batchUpdateDepth <= 0 Then Exit Sub

    m_batchUpdateDepth = m_batchUpdateDepth - 1
    If m_batchUpdateDepth = 0 Then
        FieldEndCustomUndoRecord
        Application.ScreenUpdating = m_batchScreenUpdating
    End If
    On Error GoTo 0
End Sub

Private Sub FieldBeginCustomUndoRecord()
    ' Late binding keeps older Word and Word for Mac builds from acquiring a
    ' hard UndoRecord dependency. Unsupported hosts simply retain their native
    ' undo behavior while the batch itself continues normally.
    On Error GoTo Unsupported
    Set m_batchUndoRecord = CallByName(Application, "UndoRecord", VbGet)
    If m_batchUndoRecord Is Nothing Then Exit Sub
    CallByName m_batchUndoRecord, "StartCustomRecord", VbMethod, "Banyan Update"
    m_batchUndoStarted = True
    Exit Sub

Unsupported:
    Err.Clear
    m_batchUndoStarted = False
    Set m_batchUndoRecord = Nothing
End Sub

Private Sub FieldEndCustomUndoRecord()
    On Error Resume Next
    If m_batchUndoStarted Then
        CallByName m_batchUndoRecord, "EndCustomRecord", VbMethod
    End If
    m_batchUndoStarted = False
    Set m_batchUndoRecord = Nothing
    Err.Clear
    On Error GoTo 0
End Sub


Public Function FieldCitationCode(ByVal id As String) As String
    FieldCitationCode = FIELD_CODE_CITATION & " " & id
End Function

Public Function FieldBibliographyCode(ByVal id As String) As String
    FieldBibliographyCode = FIELD_CODE_BIBLIOGRAPHY & " " & id
End Function

' Parse a field code into its kind and data id; False when it is not a Banyan
' field. A typed code without an id parses with an empty id, which the collectors
' treat as broken.
Public Function FieldParseCode(ByVal codeText As String, _
                               ByRef outKind As String, _
                               ByRef outId As String) As Boolean
    On Error GoTo ErrHandler

    Dim code As String
    code = Trim$(codeText)
    If Len(code) = 0 Then Exit Function
    If StrComp(Left$(code, 6), "ADDIN ", vbTextCompare) = 0 Then code = Trim$(Mid$(code, 7))

    ' A chapter break keeps its readable prompt as the code.
    If InStr(1, code, FIELD_CODE_CHAPTER_ZH, vbTextCompare) > 0 Then
        outKind = FIELD_KIND_CHAPTER
        FieldParseCode = True
        Exit Function
    End If
    If InStr(1, code, FIELD_CODE_CHAPTER_EN, vbTextCompare) > 0 Then
        outKind = FIELD_KIND_CHAPTER
        FieldParseCode = True
        Exit Function
    End If

    If StrComp(Left$(code, Len(FIELD_CODE_BIBLIOGRAPHY)), FIELD_CODE_BIBLIOGRAPHY, vbTextCompare) = 0 Then
        outKind = FIELD_KIND_BIBLIOGRAPHY
    ElseIf StrComp(Left$(code, Len(FIELD_CODE_CITATION)), FIELD_CODE_CITATION, vbTextCompare) = 0 Then
        outKind = FIELD_KIND_CITATION
    Else
        Exit Function
    End If

    outId = FieldRemainder(code)
    FieldParseCode = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldParseCode"
    FieldParseCode = False
End Function

Private Function FieldRemainder(ByVal value As String) As String
    Dim text As String
    text = Trim$(value)
    Dim spacePos As Long
    spacePos = InStr(text, " ")
    If spacePos > 0 Then FieldRemainder = Trim$(Mid$(text, spacePos + 1))
End Function

' True when the field's code parses to the given kind.
Public Function FieldHasCodeKind(ByVal fld As Field, ByVal codeKind As String) As Boolean
    If fld Is Nothing Then Exit Function

    Dim kind As String
    Dim id As String
    If Not FieldParseCode(fld.Code.Text, kind, id) Then Exit Function
    FieldHasCodeKind = (kind = codeKind)
End Function

' String equality with a length short-circuit (VBA's `=` scans first).
Public Function FieldTextEquals(ByVal currentText As String, ByVal nextText As String) As Boolean
    If Len(currentText) <> Len(nextText) Then Exit Function
    FieldTextEquals = (currentText = nextText)
End Function


' --- Data comparison --------------------------------------------------------

Public Function FieldContentEquals(ByVal currentData As Object, _
                                   ByVal nextData As Object) As Boolean
    On Error GoTo ErrHandler
    If currentData Is Nothing Or nextData Is Nothing Then Exit Function
    If Not DictKeyIsDict(currentData, "content") Then Exit Function
    If Not DictKeyIsDict(nextData, "content") Then Exit Function

    Dim currentContent As Object
    Dim nextContent As Object
    Set currentContent = DictKeyObject(currentData, "content")
    Set nextContent = DictKeyObject(nextData, "content")
    If currentContent Is Nothing Or nextContent Is Nothing Then Exit Function
    FieldContentEquals = FieldRichTextEquals(currentContent, nextContent)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldContentEquals"
    FieldContentEquals = False
End Function

Public Function FieldRichTextEquals(ByVal currentContent As Object, _
                                    ByVal nextContent As Object) As Boolean
    On Error GoTo ErrHandler
    If currentContent Is Nothing Or nextContent Is Nothing Then Exit Function
    If TypeName(currentContent) <> "Dictionary" Then Exit Function
    If TypeName(nextContent) <> "Dictionary" Then Exit Function

    ' Both values have already passed FieldIsRichText in the collector/response
    ' validators. Read the fixed schema directly here: generic Dict helpers add
    ' measurable overhead for every property of every mark.
    Dim currentText As String
    Dim nextText As String
    currentText = CStr(currentContent("text"))
    nextText = CStr(nextContent("text"))
    If Not FieldTextEquals(currentText, nextText) Then Exit Function

    Dim currentMarks As Object
    Dim nextMarks As Object
    Set currentMarks = currentContent("marks")
    Set nextMarks = nextContent("marks")
    If TypeName(currentMarks) <> "Collection" Then Exit Function
    If TypeName(nextMarks) <> "Collection" Then Exit Function
    If currentMarks.Count <> nextMarks.Count Then Exit Function

    Dim i As Long
    For i = 1 To currentMarks.Count
        If Not InlineMarksEqual(currentMarks(i), nextMarks(i)) Then Exit Function
    Next i

    FieldRichTextEquals = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldRichTextEquals"
    FieldRichTextEquals = False
End Function

Private Function InlineMarksEqual(ByVal currentMark As Variant, _
                                  ByVal nextMark As Variant) As Boolean
    On Error GoTo ErrHandler
    If Not IsObject(currentMark) Or Not IsObject(nextMark) Then Exit Function

    Dim currentObject As Object
    Dim nextObject As Object
    Set currentObject = currentMark
    Set nextObject = nextMark
    If TypeName(currentObject) <> "Dictionary" Then Exit Function
    If TypeName(nextObject) <> "Dictionary" Then Exit Function

    Dim markType As String
    markType = CStr(currentObject("type"))
    If markType <> CStr(nextObject("type")) Then Exit Function
    If CLng(currentObject("start")) <> CLng(nextObject("start")) Then Exit Function
    If CLng(currentObject("end")) <> CLng(nextObject("end")) Then Exit Function

    Select Case markType
        Case "bold", "italic"
            If VarType(currentObject("value")) <> vbBoolean Then Exit Function
            If VarType(nextObject("value")) <> vbBoolean Then Exit Function
            If CBool(currentObject("value")) <> CBool(nextObject("value")) Then Exit Function
        Case "script", "color", "backgroundColor", "link"
            If VarType(currentObject("value")) <> vbString Then Exit Function
            If VarType(nextObject("value")) <> vbString Then Exit Function
            If CStr(currentObject("value")) <> CStr(nextObject("value")) Then Exit Function
        Case Else
            Exit Function
    End Select

    InlineMarksEqual = True
    Exit Function

ErrHandler:
    InlineMarksEqual = False
End Function


' --- JSON data ---

Public Function FieldReadData(ByVal fld As Field) As Object
    Dim jsonText As String
    Set FieldReadData = FieldReadDataText(fld, jsonText)
End Function

' The stored JSON as text, without parsing it.
Public Function FieldDataText(ByVal fld As Field) As String
    On Error GoTo ErrHandler

    FieldDataText = CStr(fld.Data)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldDataText"
    FieldDataText = ""
End Function

' Read the stored JSON once: both the parsed data and the text it came from.
Public Function FieldReadDataText(ByVal fld As Field, ByRef outJsonText As String) As Object
    On Error GoTo ErrHandler

    Dim jsonText As String
    jsonText = CStr(fld.Data)
    outJsonText = jsonText
    If Len(jsonText) = 0 Then Exit Function

    Dim data As Object
    Set data = JsonParse(jsonText)
    If data Is Nothing Then Exit Function

    Set FieldReadDataText = data
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldReadDataText"
    Set FieldReadDataText = Nothing
End Function

Public Function FieldWriteData(ByVal fld As Field, ByVal data As Object) As Boolean
    On Error GoTo ErrHandler

    Dim jsonText As String
    jsonText = JsonStringify(data)
    If Len(jsonText) = 0 Then Exit Function

    fld.Data = jsonText
    FieldWriteData = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldWriteData"
    FieldWriteData = False
End Function


' --- Citation data factories ---

Public Function FieldCreateId() As String
    Randomize
    FieldCreateId = RandomHex(8) & "-" & _
                    RandomHex(4) & "-4" & RandomHex(3) & "-" & _
                    Hex$(8 + Int(Rnd() * 4)) & RandomHex(3) & "-" & _
                    RandomHex(12)
    FieldCreateId = LCase$(FieldCreateId)
End Function

Public Function FieldCreateEmptyCitationSource() As Object
    Dim source As Object
    Set source = New Dictionary
    Set source("cites") = New Collection
    Set source("params") = New Dictionary
    Set FieldCreateEmptyCitationSource = source
End Function

Public Function FieldCreatePlaceholderIntextCitationData(ByVal id As String, _
                                                         Optional ByVal source As Object) As Object
    If source Is Nothing Then Set source = FieldCreateEmptyCitationSource()

    Dim data As Object
    Set data = New Dictionary
    data("id") = id
    data("type") = "intext-citation"
    Set data("source") = source
    Set data("content") = FieldCreateRichText("{ INTEXT_CITATION }", FIELD_PLACEHOLDER_COLOR)
    Set FieldCreatePlaceholderIntextCitationData = data
End Function

Public Function FieldCreatePlaceholderNoteCitationData(ByVal id As String, _
                                                       Optional ByVal source As Object) As Object
    If source Is Nothing Then Set source = FieldCreateEmptyCitationSource()

    Dim data As Object
    Set data = New Dictionary
    data("id") = id
    data("type") = "note-citation"
    Set data("source") = source
    Set data("content") = FieldCreateRichText("{ NOTE_CITATION }", FIELD_PLACEHOLDER_COLOR)
    Set data("reference") = FieldCreateEmptyRichText()
    Set FieldCreatePlaceholderNoteCitationData = data
End Function

Public Function FieldAsStyleIdentifier(ByVal style As Object) As Object
    Dim d As Object
    Set d = New Dictionary
    ' Read via modDict (safe accessor, see README "Dictionary safety").
    d("id") = DictKeyString(style, "id")
    d("title") = DictKeyString(style, "title")
    Set FieldAsStyleIdentifier = d
End Function


' --- Field creation ---

Public Function FieldCreateIntextCitationAtRange(ByVal targetRange As Range, _
                                                 ByVal data As Object) As Field
    On Error GoTo ErrHandler

    targetRange.Collapse wdCollapseEnd

    Dim fld As Field
    Set fld = FieldCreateRawAddinField(targetRange, FieldCitationCode(DictKeyString(data, "id")))
    FieldWriteData fld, data
    FieldRenderStyledFieldWithData fld, data, DictKeyObject(data, "content")

    Set FieldCreateIntextCitationAtRange = fld
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldCreateIntextCitationAtRange"
    Set FieldCreateIntextCitationAtRange = Nothing
End Function

Public Function FieldCreateNoteCitationAtRange(ByVal targetRange As Range, _
                                               ByVal data As Object) As Collection
    On Error GoTo ErrHandler

    targetRange.Collapse wdCollapseEnd

    Dim note As Footnote
    Set note = FieldCreateNoteWithReference(targetRange, data)
    If note Is Nothing Then Exit Function

    Dim noteRange As Range
    Set noteRange = note.Range.Duplicate
    noteRange.Collapse wdCollapseStart

    Dim fld As Field
    Set fld = FieldCreateRawAddinField(noteRange, FieldCitationCode(DictKeyString(data, "id")))
    FieldWriteData fld, data
    FieldRenderStyledFieldWithData fld, data, DictKeyObject(data, "content")
    FieldRestoreCaretAfterNote note

    Dim result As Collection
    Set result = New Collection
    result.Add note, "note"
    result.Add fld, "field"
    Set FieldCreateNoteCitationAtRange = result
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldCreateNoteCitationAtRange"
    Set FieldCreateNoteCitationAtRange = Nothing
End Function

Private Function FieldCreateNoteWithReference(ByVal targetRange As Range, _
                                              ByVal data As Object) As Footnote
    ' Create a footnote at targetRange with the custom reference mark carried in
    ' data("reference") (rendered as rich text + links). This is what lets the
    ' VBA add-in support references Word itself cannot produce (e.g. "[1]").
    On Error GoTo ErrHandler

    Dim referenceContent As Object
    Dim referenceText As String
    Set referenceContent = GetOptionalRichText(data, "reference")
    referenceText = FieldPlainTextFromContent(referenceContent)

    Dim note As Footnote
    If Len(referenceText) > 0 Then
        Set note = ActiveDocument.Footnotes.Add(Range:=targetRange, Reference:=referenceText)
        FieldApplyRichTextStylesToRange note.Reference, referenceContent
        FieldApplyRichTextLinksToRange note.Reference, referenceContent
    Else
        Set note = ActiveDocument.Footnotes.Add(Range:=targetRange)
    End If

    Set FieldCreateNoteWithReference = note
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldCreateNoteWithReference"
    Set FieldCreateNoteWithReference = Nothing
End Function

Public Function FieldRebuildNoteCitationAtRange(ByVal note As Footnote, _
                                                ByVal fld As Field, _
                                                ByVal data As Object) As Collection
    ' Refresh a note citation while preserving the rich text of the footnote
    ' body (fonts, bold, ... around the citation field).
    '
    ' VBA owns the custom footnote Reference (e.g. "[1]"); the reference mark
    ' cannot be edited in place (setting note.Reference.Text corrupts the
    ' footnote), so a CHANGED reference requires recreating the footnote.
    '
    ' In BOTH cases the citation field itself is REBUILT fresh (FieldReplace-
    ' CitationInNote): a reused result range would inherit the placeholder's
    ' direct formatting (the red "{ NOTE_CITATION }"), and a fresh ADDIN field
    ' renders cleanly from data("content").
    '
    '   * Reference unchanged (including empty = Word-managed auto numbering,
    '     which is stable across refresh): keep the footnote, replace only the
    '     field in place - the surrounding user content is never touched.
    '   * Reference changed: create the NEW footnote just AFTER the old reference
    '     mark (so the two references never overlap), copy the ENTIRE body across
    '     with FormattedText (preserves formatting AND the field structure) while
    '     both footnotes exist, then delete the old footnote - the new reference
    '     slides into the old position - and replace the copied field fresh.
    On Error GoTo ErrHandler

    If note Is Nothing Then Exit Function
    If fld Is Nothing Then Exit Function

    ' --- Does the footnote reference need rebuilding? ---
    Dim newRefText As String
    Dim newRefContent As Object
    Set newRefContent = GetOptionalRichText(data, "reference")
    newRefText = FieldPlainTextFromContent(newRefContent)

    Dim oldRefContent As Object
    Dim oldData As Object
    Set oldData = FieldReadData(fld)
    If Not oldData Is Nothing Then
        Set oldRefContent = GetOptionalRichText(oldData, "reference")
    End If

    If Not oldData Is Nothing Then
        If FieldContentEquals(oldData, data) And _
           FieldRichTextEquals(oldRefContent, newRefContent) Then
            ' Source and other metadata do not affect the Word result. Keep the
            ' existing field and footnote, but persist the latest response data.
            If Not FieldWriteData(fld, data) Then Exit Function
            Dim unchangedResult As Collection
            Set unchangedResult = New Collection
            unchangedResult.Add note, "note"
            unchangedResult.Add fld, "field"
            Set FieldRebuildNoteCitationAtRange = unchangedResult
            Exit Function
        End If
    End If

    If FieldRichTextEquals(oldRefContent, newRefContent) Then
        ' Reference unchanged - keep the footnote, replace the field in place.
        Dim sameField As Field
        Set sameField = FieldReplaceCitationInNote(note, fld, data)
        If sameField Is Nothing Then Exit Function

        Dim sameResult As Collection
        Set sameResult = New Collection
        sameResult.Add note, "note"
        sameResult.Add sameField, "field"
        Set FieldRebuildNoteCitationAtRange = sameResult
        Exit Function
    End If

    ' --- Reference changed: rebuild preserving the whole (rich) body. ---
    Dim insertPosition As Long
    insertPosition = note.Reference.End
    insertPosition = ClampInsertPosition(insertPosition)

    Dim insertRange As Range
    Set insertRange = ActiveDocument.Range(insertPosition, insertPosition)

    Dim newNote As Footnote
    Set newNote = FieldCreateNoteWithReference(insertRange, data)
    If newNote Is Nothing Then Exit Function

    ' Copy the whole body (rich text + field structure) into the new footnote
    ' while both footnotes exist - FormattedText ranges are live document ranges
    ' and must not survive the old footnote's deletion.
    newNote.Range.FormattedText = note.Range.FormattedText

    ' Remove the old footnote; the new reference settles into the old spot.
    FieldRemoveFootnoteSafely note

    ' Replace the copied field with a fresh one (clean formatting); if the copy
    ' did not preserve a field (e.g. the body was empty), insert one at the start.
    Dim newField As Field
    Dim copiedField As Field
    Set copiedField = FieldFindFirstCitationInNote(newNote)
    If copiedField Is Nothing Then
        Dim noteRange As Range
        Set noteRange = newNote.Range.Duplicate
        noteRange.Collapse wdCollapseStart
        Set newField = FieldCreateRawAddinField(noteRange, FieldCitationCode(DictKeyString(data, "id")))
        If Not newField Is Nothing Then
            FieldWriteData newField, data
            FieldRenderStyledFieldWithData newField, data, DictKeyObject(data, "content")
        End If
    Else
        Set newField = FieldReplaceCitationInNote(newNote, copiedField, data)
    End If
    If newField Is Nothing Then Exit Function

    Dim result As Collection
    Set result = New Collection
    result.Add newNote, "note"
    result.Add newField, "field"
    Set FieldRebuildNoteCitationAtRange = result
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldRebuildNoteCitationAtRange"
    Set FieldRebuildNoteCitationAtRange = Nothing
End Function

Private Function FieldReplaceCitationInNote(ByVal note As Footnote, _
                                            ByVal fld As Field, _
                                            ByVal data As Object) As Field
    ' Replace the citation field in the footnote body with a FRESH ADDIN field
    ' at the same spot. The user content before/after the field is preserved
    ' untouched (we only remove the field and insert the new one at its former
    ' boundary), and the fresh field renders cleanly - it does NOT inherit the
    ' placeholder's direct formatting (the red "{ NOTE_CITATION }").
    On Error GoTo ErrHandler

    If fld Is Nothing Then Exit Function

    ' Capture the live range of the user content BEFORE the field; after the
    ' field is removed this range's End tracks the spot where it used to start.
    Dim beforeRange As Range
    Set beforeRange = note.Range.Duplicate
    beforeRange.End = fld.Result.Start

    ' Remove the old field entirely (code + result text).
    FieldRemoveFieldSafely fld

    ' Insert the fresh field right after the before-content. Positions derive
    ' from the footnote body range (footnote story - ActiveDocument.Range cannot
    ' address it).
    Dim insertAt As Range
    Set insertAt = note.Range.Duplicate
    insertAt.SetRange beforeRange.End, beforeRange.End

    Dim newField As Field
    Set newField = FieldCreateRawAddinField(insertAt, FieldCitationCode(DictKeyString(data, "id")))
    If newField Is Nothing Then Exit Function

    FieldWriteData newField, data
    FieldRenderStyledFieldWithData newField, data, DictKeyObject(data, "content")

    Set FieldReplaceCitationInNote = newField
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldReplaceCitationInNote"
    Set FieldReplaceCitationInNote = Nothing
End Function

Private Function FieldFindFirstCitationInNote(ByVal note As Footnote) As Field
    On Error GoTo ErrHandler
    Dim f As Field
    Dim kind As String
    Dim id As String
    For Each f In note.Range.Fields
        If FieldParseCode(f.Code.Text, kind, id) Then
            If kind = FIELD_KIND_CITATION Then
                Set FieldFindFirstCitationInNote = f
                Exit Function
            End If
        End If
    Next f
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldFindFirstCitationInNote"
    Set FieldFindFirstCitationInNote = Nothing
End Function


' --- Render ---

Public Function FieldRenderStyledField(ByVal fld As Field, _
                                       Optional ByVal content As Variant) As Boolean
    On Error GoTo ErrHandler

    Dim resolvedContent As Object
    Set resolvedContent = ResolveFieldContent(fld, content)
    If resolvedContent Is Nothing Then Exit Function

    Dim data As Object
    Set data = FieldReadData(fld)
    FieldRenderStyledField = RenderStyledFieldCore(fld, resolvedContent, data)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldRenderStyledField"
    FieldRenderStyledField = False
End Function

' Use this entry point when the caller already has the parsed field data.
' The legacy FieldRenderStyledField above intentionally keeps its read-from-
' Field.Data behavior for callers that only have a Field object. Both paths
' converge on RenderStyledFieldCore so rendering behavior cannot diverge.
Public Function FieldRenderStyledFieldWithData(ByVal fld As Field, _
                                               ByVal data As Object, _
                                               Optional ByVal content As Variant) As Boolean
    On Error GoTo ErrHandler

    If data Is Nothing Then Exit Function

    Dim resolvedContent As Object
    Set resolvedContent = ResolveFieldContentFromData(data, content)
    If resolvedContent Is Nothing Then Exit Function

    FieldRenderStyledFieldWithData = RenderStyledFieldCore(fld, resolvedContent, data)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldRenderStyledFieldWithData"
    FieldRenderStyledFieldWithData = False
End Function

Private Function RenderStyledFieldCore(ByVal fld As Field, _
                                       ByVal resolvedContent As Object, _
                                       ByVal data As Object) As Boolean
    On Error GoTo ErrHandler

    WriteContentTextToRange fld.Result, resolvedContent
    fld.ShowCodes = False

    If Not data Is Nothing Then
        If DictHasKey(data, "type") Then
            If DictKeyString(data, "type") = "intext-citation" Then
                ' A CHARACTER style assignment clears the direct formatting the text
                ' write just inherited (Word does this itself), so no reset is needed.
                FieldApplyIntextCitationStyle fld
            ElseIf DictKeyString(data, "type") = "note-citation" Then
                ' The note citation is styled with the built-in PARAGRAPH style, and a
                ' paragraph style assignment does NOT clear direct formatting, so the
                ' inherited formatting (e.g. the red placeholder) is reset explicitly.
                FieldClearDirectCharacterFormatting fld.Result
                FieldApplyNoteCitationStyle fld
            End If
        End If
    End If

    FieldApplyRichTextStylesToRange fld.Result, resolvedContent
    FieldApplyRichTextLinksToRange fld.Result, resolvedContent
    RenderStyledFieldCore = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.RenderStyledFieldCore"
    RenderStyledFieldCore = False
End Function

' Clear the direct character formatting a field result inherited.
'
' Reason: assigning a CHARACTER style clears direct character formatting by itself,
' but assigning a PARAGRAPH style does not (measured in WPS; Word not compared) - it
' only touches paragraph-level properties. The bibliography lines (title and entry)
' and the note citation are styled with paragraph styles only, so without this reset
' an entry keeps whatever the insertion point or the previous render (e.g. the red
' "{ BIBLIOGRAPHY }" placeholder) left on the result range. The rich-text marks are
' applied afterwards and win over the style.
'
' Keep in sync with the WPS port: clearDirectCharacterFormatting() in field.ts.
Private Sub FieldClearDirectCharacterFormatting(ByVal targetRange As Range)
    On Error GoTo ErrHandler

    Dim f As Font
    Set f = targetRange.Font
    f.Color = wdColorAutomatic
    f.Bold = 0
    f.Italic = 0
    f.Subscript = 0
    f.Superscript = 0
    f.SmallCaps = 0
    f.AllCaps = 0
    targetRange.Shading.BackgroundPatternColor = wdColorAutomatic
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldClearDirectCharacterFormatting"
End Sub

Public Function FieldRenderStyledFieldWithStyle(ByVal fld As Field, _
                                                ByVal styleName As String, _
                                                Optional ByVal styleType As WdStyleType = wdStyleTypeCharacter, _
                                                Optional ByVal content As Variant) As Boolean
    On Error GoTo ErrHandler

    Dim resolvedContent As Object
    Set resolvedContent = ResolveFieldContent(fld, content)
    If resolvedContent Is Nothing Then Exit Function

    WriteContentTextToRange fld.Result, resolvedContent
    fld.ShowCodes = False

    ' Same reason as in RenderStyledFieldCore: a paragraph style assignment does not
    ' clear direct formatting, and this entry point is the one used for the
    ' bibliography lines (paragraph styles).
    If styleType = wdStyleTypeParagraph Then FieldClearDirectCharacterFormatting fld.Result

    FieldApplyStyleToField fld, styleName, styleType
    FieldApplyRichTextStylesToRange fld.Result, resolvedContent
    FieldApplyRichTextLinksToRange fld.Result, resolvedContent

    FieldRenderStyledFieldWithStyle = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldRenderStyledFieldWithStyle"
    FieldRenderStyledFieldWithStyle = False
End Function

Public Sub FieldApplyIntextCitationStyle(ByVal fld As Field)
    FieldApplyStyleToField fld, FieldIntextCitationStyleName(), wdStyleTypeCharacter
End Sub

Public Sub FieldApplyNoteCitationStyle(ByVal fld As Field)
    On Error Resume Next
    ActiveDocument.Styles(wdStyleFootnoteReference).UnhideWhenUsed = True
    ActiveDocument.Styles(wdStyleFootnoteReference).QuickStyle = True
    ActiveDocument.Styles(wdStyleFootnoteText).UnhideWhenUsed = True
    ActiveDocument.Styles(wdStyleFootnoteText).QuickStyle = True
    fld.Result.Style = wdStyleFootnoteText
    On Error GoTo 0
End Sub

Public Sub FieldApplyStyleToField(ByVal fld As Field, _
                                  ByVal styleName As String, _
                                  Optional ByVal styleType As WdStyleType = wdStyleTypeCharacter)
    On Error GoTo ErrHandler
    If Len(Trim$(styleName)) = 0 Then Exit Sub

    Dim style As Style
    Set style = CachedWordStyle(styleName, styleType)
    If style Is Nothing Then
        Exit Sub
    End If

    style.UnhideWhenUsed = True
    style.QuickStyle = True
    fld.Result.Style = style
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldApplyStyleToField"
End Sub

Private Function CachedWordStyle(ByVal styleName As String, _
                                 ByVal styleType As WdStyleType) As Style
    On Error GoTo ErrHandler
    If Len(Trim$(styleName)) = 0 Then Exit Function

    If m_styleCache Is Nothing Then Set m_styleCache = New Dictionary

    Dim cacheMatchesDocument As Boolean
    Dim hasCachedDocument As Boolean
    hasCachedDocument = Not (m_styleCacheDocument Is Nothing)
    If Not hasCachedDocument Then
        Set m_styleCacheDocument = ActiveDocument
        cacheMatchesDocument = True
    Else
        cacheMatchesDocument = (m_styleCacheDocument Is ActiveDocument)
    End If
    If Not cacheMatchesDocument Then
        Set m_styleCache = New Dictionary
        Set m_styleCacheDocument = ActiveDocument
    End If

    If m_styleCache.Exists(styleName) Then
        Set CachedWordStyle = m_styleCache(styleName)
        Exit Function
    End If

    Dim style As Style
    Set style = FindWordStyle(styleName)
    If style Is Nothing Then
        Set style = ActiveDocument.Styles.Add(Name:=styleName, Type:=styleType)
    End If

    Set m_styleCache(styleName) = style
    Set CachedWordStyle = style
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.CachedWordStyle"
    Set CachedWordStyle = Nothing
End Function


' --- Collectors ---

' Ids are unique per document, but Word reports no paste: a copied citation
' arrives with a duplicated id and the refresh would match two fields to one
' response. Re-key the copy here, where its code is parsed. The stored data is
' deliberately left alone: the request is built from the new id, so the response
' data differs from the stored text and replaces it at the next write.
Private Function FieldUniqueCitationId(ByVal fld As Field, ByVal id As String, ByVal seenIds As Object) As String
    If Not seenIds.Exists(id) Then
        seenIds(id) = True
        FieldUniqueCitationId = id
        Exit Function
    End If

    Dim candidate As String
    Do
        candidate = FieldCreateId()
    Loop While seenIds.Exists(candidate)

    If Not FieldWriteCitationCodeId(fld, candidate) Then
        ' The code could not be rewritten, so it still spells the identity.
        FieldUniqueCitationId = id
        Exit Function
    End If

    seenIds(candidate) = True
    FieldUniqueCitationId = candidate
End Function

' Swap the id inside an existing citation code, keeping the " ADDIN ... " shape
' FieldCreateRawAddinField writes. Word drops Field.Data when an ADDIN code is
' rewritten, so the stored payload is read as text before the swap and put back
' after it (a string copy, only for a re-keyed copy).
Private Function FieldWriteCitationCodeId(ByVal fld As Field, ByVal id As String) As Boolean
    On Error GoTo ErrHandler
    If fld Is Nothing Then Exit Function

    Dim storedText As String
    storedText = FieldDataText(fld)

    fld.Code.Text = " ADDIN " & FieldCitationCode(id) & " "
    If Len(storedText) > 0 Then fld.Data = storedText
    FieldWriteCitationCodeId = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldWriteCitationCodeId"
    FieldWriteCitationCodeId = False
End Function

' Delete a typed field whose code carries no id. Without an id it cannot be
' matched to a response, so it is broken rather than a citation.
Private Sub FieldDeleteBrokenFields(ByVal fields As Collection)
    On Error GoTo ErrHandler

    Dim i As Long
    For i = fields.Count To 1 Step -1
        FieldRemoveFieldSafely fields(i)
    Next i
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldDeleteBrokenFields"
End Sub

' Note citations leave with their whole footnote, not just the field.
Private Sub FieldDeleteBrokenNotes(ByVal notes As Collection)
    On Error GoTo ErrHandler

    Dim i As Long
    For i = notes.Count To 1 Step -1
        FieldRemoveFootnoteSafely notes(i)
    Next i
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldDeleteBrokenNotes"
End Sub

Public Function FieldCollectIntextCitationFieldsInRange(ByVal targetRange As Range) As Collection
    Dim result As Collection
    Set result = New Collection

    ' Code-only classification: no Field.Data access.
    Dim seenIds As Object
    Set seenIds = New Dictionary
    Dim broken As Collection

    Dim fld As Field
    Dim kind As String
    Dim id As String
    For Each fld In targetRange.Fields
        If fld.Type = wdFieldAddin Then
            If FieldParseCode(fld.Code.Text, kind, id) Then
                If kind = FIELD_KIND_CITATION Then
                    If Len(id) = 0 Then
                        If broken Is Nothing Then Set broken = New Collection
                        broken.Add fld
                    Else
                        result.Add MakeFieldAndId(fld, FieldUniqueCitationId(fld, id, seenIds))
                    End If
                End If
            End If
        End If
    Next fld

    ' Deleting while the field collection is being walked would be unsafe.
    If Not broken Is Nothing Then FieldDeleteBrokenFields broken

    Set FieldCollectIntextCitationFieldsInRange = result
End Function

Public Function FieldCollectNoteCitationFootnotesInRange(ByVal targetRange As Range) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim seenIds As Object
    Set seenIds = New Dictionary
    Dim broken As Collection

    Dim note As Footnote
    Dim fld As Field
    Dim kind As String
    Dim id As String
    For Each note In targetRange.Footnotes
        If note.Range.Fields.Count > 0 Then
            Set fld = note.Range.Fields(1)
            If Not fld Is Nothing Then
                If fld.Type = wdFieldAddin Then
                    If FieldParseCode(fld.Code.Text, kind, id) Then
                        If kind = FIELD_KIND_CITATION Then
                            If Len(id) = 0 Then
                                If broken Is Nothing Then Set broken = New Collection
                                broken.Add note
                            Else
                                result.Add MakeNoteFieldAndId(note, fld, FieldUniqueCitationId(fld, id, seenIds))
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next note

    If Not broken Is Nothing Then FieldDeleteBrokenNotes broken

    Set FieldCollectNoteCitationFootnotesInRange = result
End Function

' --- Migration helpers ---

Public Sub FieldMigrateIntextCitationsToNotes(ByVal targetRange As Range)
    On Error GoTo ErrHandler

    Dim citations As Collection
    Set citations = FieldCollectIntextCitationFieldsInRange(targetRange)

    Dim i As Long
    For i = citations.Count To 1 Step -1
        Dim fd As Collection
        Set fd = citations(i)

        Dim fld As Field
        Dim data As Object
        Set fld = fd("field")
        Set data = FieldReadData(fld)
        If data Is Nothing Then GoTo NextCitation

        ' Record the field's first position (the code range starts one character
        ' after the field's begin marker, so Code.Start - 1 is it) BEFORE the
        ' removal: removing the field deletes the CODE too, so any position inside
        ' the field is stale afterwards and the footnote reference would land a
        ' whole code length to the right. A plain number is enough because the loop
        ' walks the citations BACKWARDS: migrating a later citation never moves an
        ' earlier one, and removing a field does not move its own start.
        Dim insertPosition As Long
        insertPosition = fld.Code.Start - 1

        Dim convertedData As Object
        Set convertedData = FieldCreatePlaceholderNoteCitationData(CStr(fd("id")), DictKeyObject(data, "source"))

        FieldRemoveFieldSafely fld

        Dim insertRange As Range
        Set insertRange = ActiveDocument.Range(insertPosition, insertPosition)
        FieldCreateNoteCitationAtRange insertRange, convertedData
NextCitation:
    Next i
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldMigrateIntextCitationsToNotes"
End Sub

Public Sub FieldMigrateNoteCitationsToIntext(ByVal targetRange As Range)
    On Error GoTo ErrHandler

    Dim citations As Collection
    Set citations = FieldCollectNoteCitationFootnotesInRange(targetRange)

    Dim i As Long
    For i = citations.Count To 1 Step -1
        Dim fd As Collection
        Set fd = citations(i)

        Dim note As Footnote
        Dim data As Object
        Set note = fd("note")
        Set data = FieldReadData(fd("field"))
        If data Is Nothing Then GoTo NextNote

        ' No live anchor is needed in this direction: the reference mark holds a
        ' single character, so its own start does not move when the note is
        ' removed (only the text after the mark shifts left).
        Dim insertPosition As Long
        insertPosition = note.Reference.Start

        Dim convertedData As Object
        Set convertedData = FieldCreatePlaceholderIntextCitationData(CStr(fd("id")), DictKeyObject(data, "source"))

        FieldRemoveFootnoteSafely note

        ' Removing the note can shrink the document when it is the last content,
        ' leaving the captured position out of range (error 4608). Clamp it.
        insertPosition = ClampInsertPosition(insertPosition)

        Dim insertRange As Range
        Set insertRange = ActiveDocument.Range(insertPosition, insertPosition)
        FieldCreateIntextCitationAtRange insertRange, convertedData
NextNote:
    Next i
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldMigrateNoteCitationsToIntext"
End Sub

Private Function ClampInsertPosition(ByVal position As Long) As Long
    ' Keep a captured document position inside [Content.Start, Content.End - 1].
    ' After a field/footnote at the end of the document is removed the document
    ' shrinks and the old position becomes out of range (ActiveDocument.Range
    ' then raises error 4608 "value out of range").
    On Error GoTo ErrHandler

    Dim startPos As Long
    Dim endPos As Long
    startPos = ActiveDocument.Content.Start
    endPos = ActiveDocument.Content.End

    If endPos > startPos Then
        If position > endPos - 1 Then position = endPos - 1
    End If
    If position < startPos Then position = startPos

    ClampInsertPosition = position
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.ClampInsertPosition"
    ClampInsertPosition = position
End Function


' --- Raw ADDIN field creation ---

Public Function FieldCreateRawAddinField(ByVal targetRange As Range, ByVal fieldCode As String) As Field
    On Error GoTo ErrHandler

    targetRange.Collapse wdCollapseEnd

    Dim fld As Field
    ' Create the ADDIN field through an intermediate QUOTE field (wdFieldQuote),
    ' exactly as Zotero's Windows integration does in insertFieldRaw
    ' (zotero-word-for-windows-integration). A QUOTE field materializes a real
    ' (separator + result) range, whereas creating an ADDIN field directly
    ' produces a field with no separator whose Result is an empty collapsed
    ' range - writing to Field.Result.Text then would not land inside the field.
    Set fld = targetRange.Fields.Add(targetRange, wdFieldQuote, FIELD_RAW_PLACEHOLDER, True)
    ' Replace the QUOTE code with the real ADDIN code; Word keeps the result.
    fld.Code.Text = " ADDIN " & fieldCode & " "
    ' Update flips the field type to wdFieldAddin (collectors check
    ' fld.Type = wdFieldAddin) while preserving the existing result range.
    fld.Update

    Set FieldCreateRawAddinField = fld
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldCreateRawAddinField"
    Set FieldCreateRawAddinField = Nothing
End Function


' --- Type guards ---

Public Function FieldIsIntextCitation(ByVal data As Variant) As Boolean
    If Not FieldIsCitation(data) Then Exit Function
    FieldIsIntextCitation = (DictKeyString(data, "type") = "intext-citation")
End Function

Public Function FieldIsNoteCitation(ByVal data As Variant) As Boolean
    If Not FieldIsCitation(data) Then Exit Function
    If DictKeyString(data, "type") <> "note-citation" Then Exit Function
    If Not DictKeyIsDict(data, "reference") Then Exit Function
    If Not FieldIsRichText(DictKeyObject(data, "reference")) Then Exit Function
    FieldIsNoteCitation = True
End Function

Public Function FieldIsBibliographyTitle(ByVal data As Variant) As Boolean
    If Not FieldIsBanyanFieldData(data) Then Exit Function
    FieldIsBibliographyTitle = (DictKeyString(data, "type") = "bibliography-title")
End Function

Public Function FieldIsBibliographyEntry(ByVal data As Variant) As Boolean
    If Not FieldIsBanyanFieldData(data) Then Exit Function
    FieldIsBibliographyEntry = (DictKeyString(data, "type") = "bibliography-entry")
End Function

Public Function FieldIsCitationSource(ByVal data As Variant) As Boolean
    If Not DictIsDictionary(data) Then Exit Function
    If Not DictKeyIsCollection(data, "cites") Then Exit Function
    If Not DictKeyIsDict(data, "params") Then Exit Function
    FieldIsCitationSource = True
End Function

Public Function FieldIsRichText(ByVal value As Variant) As Boolean
    On Error GoTo ErrHandler

    Dim d As Object
    Set d = value
    If d Is Nothing Then Exit Function
    If TypeName(d) <> "Dictionary" Then Exit Function
    If Not DictKeyIsString(d, "text") Then Exit Function
    If Not DictHasKey(d, "marks") Then Exit Function
    If Not DictIsCollection(DictKeyObject(d, "marks")) Then Exit Function

    Dim textLength As Long
    textLength = Len(DictKeyString(d, "text"))

    Dim item As Variant
    For Each item In DictKeyObject(d, "marks")
        If Not FieldIsInlineMark(item, textLength) Then Exit Function
    Next item

    FieldIsRichText = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldIsRichText"
    FieldIsRichText = False
End Function


' --- Cleanup helpers ---

Public Sub FieldRemoveFieldSafely(ByVal fld As Field)
    On Error Resume Next
    If Not fld Is Nothing Then
        If fld.Locked Then fld.Locked = False
        fld.Delete
    End If
    On Error GoTo 0
End Sub

Public Sub FieldRemoveFootnoteSafely(ByVal note As Footnote)
    On Error Resume Next
    If Not note Is Nothing Then note.Delete
    On Error GoTo 0
End Sub

' Put the caret back in the MAIN TEXT right after a footnote reference.
'
' Both hosts move the insertion point INTO the new footnote when the insertion point
' sits at the insertion position (the normal "insert a citation at the cursor" flow;
' WPS does it unconditionally, Word when the caret is at the insertion position).
' The parked caret then breaks the next action: Footnotes.Add with a selection-derived
' range raises "footnotes can only be added to the main body of the document", and
' Selection.GoTo cannot escape a footnote story in Word. Selecting the end of the
' reference mark puts the caret back into the main text in both hosts.
'
' Keep in sync with the WPS port: restoreMainTextAfterNote() in field.ts.
Public Sub FieldRestoreCaretAfterNote(ByVal note As Footnote)
    On Error GoTo ErrHandler
    If note Is Nothing Then Exit Sub

    Dim caret As Range
    Set caret = note.Reference.Duplicate
    caret.Collapse wdCollapseEnd
    caret.Select
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldRestoreCaretAfterNote"
End Sub

Public Sub FieldRemoveEndnoteSafely(ByVal note As Endnote)
    On Error Resume Next
    If Not note Is Nothing Then note.Delete
    On Error GoTo 0
End Sub

' Word bookmark naming rules - authoritative sources:
'   - Office.js Word.BookmarkCollection.add
'     (learn.microsoft.com/en-us/javascript/api/word/word.bookmarkcollection): the
'     name "cannot be more than 40 characters or include more than one word. Also,
'     the name must begin with a letter. It can include both numbers and letters,
'     but not spaces. If you need to separate words, use an underscore."
'   - Word VBA Bookmarks.Add (learn.microsoft.com/en-us/office/vba/api/word.bookmarks.add):
'     "The name cannot be more than 40 characters or include more than one word."
' Verified in Word 16 (zh-CN): spaces and punctuation (. - ,) and a digit-first
' name are rejected, an underscore start creates a hidden bookmark, and a name
' longer than 40 characters is accepted but silently truncated to 40 - so the
' truncation here matches what Word would store.
Public Function FieldNormalizeBookmarkName(ByVal rawName As String) As String
    Dim normalized As String
    Dim ch As String
    Dim code As Long
    Dim i As Long

    rawName = Trim$(rawName)
    If Len(rawName) = 0 Then Exit Function

    For i = 1 To Len(rawName)
        ch = Mid$(rawName, i, 1)
        Select Case AscW(ch)
            Case 48 To 57, 65 To 90, 95, 97 To 122
                normalized = normalized & ch
            Case Else
                normalized = normalized & "_"
        End Select
    Next i

    code = AscW(Left$(normalized, 1))
    If Not ((code >= 65 And code <= 90) Or (code >= 97 And code <= 122)) Then
        normalized = "B" & normalized
    End If

    If Len(normalized) > FIELD_BOOKMARK_MAX_LENGTH Then
        normalized = Left$(normalized, FIELD_BOOKMARK_MAX_LENGTH)
    End If
    FieldNormalizeBookmarkName = normalized
End Function

Public Sub FieldAddBookmarkToField(ByVal fld As Field, ByVal bookmarkName As String)
    On Error GoTo ErrHandler
    If fld Is Nothing Then Exit Sub

    bookmarkName = FieldNormalizeBookmarkName(bookmarkName)
    If Len(bookmarkName) = 0 Then Exit Sub

    Dim bookmarks As Bookmarks
    Set bookmarks = ActiveDocument.Bookmarks

    If bookmarks.Exists(bookmarkName) Then
        bookmarks(bookmarkName).Delete
    End If

    bookmarks.Add bookmarkName, fld.Result
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldAddBookmarkToField"
End Sub

' Canonical bibliography bookmark name. ApplyRichTextLink builds the hyperlink
' SubAddress from this same function, so bookmark and link always carry the
' identical normalized name.
Public Function FieldGetBibliographyBookmarkName(ByVal entryId As String) As String
    FieldGetBibliographyBookmarkName = _
        FieldNormalizeBookmarkName(FIELD_BIBLIOGRAPHY_BOOKMARK_PREFIX & entryId)
End Function


' --- Internal type helpers ---

Private Function FieldIsBanyanFieldData(ByVal data As Variant) As Boolean
    If Not DictIsDictionary(data) Then Exit Function
    If Not DictKeyIsString(data, "type") Then Exit Function
    If Not DictKeyIsDict(data, "content") Then Exit Function
    If Not FieldIsRichText(DictKeyObject(data, "content")) Then Exit Function
    FieldIsBanyanFieldData = True
End Function

Private Function FieldIsCitation(ByVal data As Variant) As Boolean
    If Not FieldIsBanyanFieldData(data) Then Exit Function
    If Not DictKeyIsString(data, "id") Then Exit Function
    If Not DictKeyIsDict(data, "source") Then Exit Function
    FieldIsCitation = True
End Function


' --- Render internals ---

Private Function ResolveFieldContent(ByVal fld As Field, Optional ByVal content As Variant) As Object
    On Error GoTo ErrHandler

    If Not IsMissing(content) Then
        If IsObject(content) Then
            If FieldIsRichText(content) Then
                Set ResolveFieldContent = content
                Exit Function
            End If
        End If
    End If

    Dim data As Object
    Set data = FieldReadData(fld)
    If data Is Nothing Then Exit Function
    If Not DictHasKey(data, "content") Then Exit Function
    If Not FieldIsRichText(DictKeyObject(data, "content")) Then Exit Function

    Set ResolveFieldContent = DictKeyObject(data, "content")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.ResolveFieldContent"
    Set ResolveFieldContent = Nothing
End Function

Private Function ResolveFieldContentFromData(ByVal data As Object, Optional ByVal content As Variant) As Object
    On Error GoTo ErrHandler

    If Not IsMissing(content) Then
        If IsObject(content) Then
            If FieldIsRichText(content) Then
                Set ResolveFieldContentFromData = content
                Exit Function
            End If
        End If
    End If

    If data Is Nothing Then Exit Function
    If Not DictHasKey(data, "content") Then Exit Function
    If Not FieldIsRichText(DictKeyObject(data, "content")) Then Exit Function
    Set ResolveFieldContentFromData = DictKeyObject(data, "content")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.ResolveFieldContentFromData"
    Set ResolveFieldContentFromData = Nothing
End Function

Private Function GetOptionalRichText(ByVal data As Object, ByVal key As String) As Object
    On Error GoTo ErrHandler
    If data Is Nothing Then Exit Function
    If Not DictHasKey(data, key) Then Exit Function
    If Not FieldIsRichText(DictKeyObject(data, key)) Then Exit Function
    Set GetOptionalRichText = DictKeyObject(data, key)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.GetOptionalRichText"
    Set GetOptionalRichText = Nothing
End Function

Private Sub WriteContentTextToRange(ByVal targetRange As Range, ByVal content As Object)
    targetRange.Text = FieldPlainTextFromContent(content)
End Sub

Public Function FieldPlainTextFromContent(ByVal content As Object) As String
    On Error GoTo ErrHandler
    If content Is Nothing Then Exit Function
    If Not FieldIsRichText(content) Then Exit Function
    FieldPlainTextFromContent = DictKeyString(content, "text")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldPlainTextFromContent"
    FieldPlainTextFromContent = ""
End Function

Private Sub FieldApplyRichTextStylesToRange(ByVal targetRange As Range, ByVal content As Object)
    On Error GoTo ErrHandler
    If Not FieldIsRichText(content) Then Exit Sub

    Dim mark As Variant
    Dim markObj As Object
    Dim segment As Range

    For Each mark In DictKeyObject(content, "marks")
        Set markObj = mark
        If DictKeyString(markObj, "type") <> "link" Then
            Set segment = targetRange.Duplicate
            segment.SetRange targetRange.Start + DictKeyLong(markObj, "start"), targetRange.Start + DictKeyLong(markObj, "end")
            ApplyInlineMarkStyle segment, markObj
        End If
    Next mark
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldApplyRichTextStylesToRange"
End Sub

Public Function FieldApplyRichTextLinksToRange(ByVal targetRange As Range, _
                                               ByVal content As Object) As Boolean
    On Error GoTo ErrHandler
    If Not FieldIsRichText(content) Then Exit Function

    Dim segments As Collection
    Set segments = New Collection

    Dim mark As Variant
    Dim markObj As Object
    For Each mark In DictKeyObject(content, "marks")
        Set markObj = mark
        If DictKeyString(markObj, "type") = "link" Then
            segments.Add markObj
        End If
    Next mark

    Dim i As Long
    Dim linkMark As Object
    Dim segment As Range
    For i = segments.Count To 1 Step -1
        Set linkMark = segments(i)
        Set segment = targetRange.Duplicate
        segment.SetRange targetRange.Start + DictKeyLong(linkMark, "start"), targetRange.Start + DictKeyLong(linkMark, "end")
        ApplyRichTextLink segment, DictKeyString(linkMark, "value")
    Next i
    FieldApplyRichTextLinksToRange = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldApplyRichTextLinksToRange"
    FieldApplyRichTextLinksToRange = False
End Function

Private Function FieldIsInlineMark(ByVal value As Variant, ByVal textLength As Long) As Boolean
    On Error GoTo ErrHandler

    Dim d As Object
    Set d = value
    If d Is Nothing Then Exit Function
    If TypeName(d) <> "Dictionary" Then Exit Function
    If Not DictKeyIsString(d, "type") Then Exit Function
    If Not DictHasKey(d, "start") Then Exit Function
    If Not DictHasKey(d, "end") Then Exit Function
    If Not DictHasKey(d, "value") Then Exit Function
    If DictKeyIsNull(d, "start") Or DictKeyIsEmpty(d, "start") Then Exit Function
    If DictKeyIsNull(d, "end") Or DictKeyIsEmpty(d, "end") Then Exit Function
    If DictKeyIsObject(d, "start") Or DictKeyIsObject(d, "end") Then Exit Function

    Dim startPos As Long
    Dim endPos As Long
    startPos = DictKeyLong(d, "start")
    endPos = DictKeyLong(d, "end")
    If startPos < 0 Then Exit Function
    If endPos <= startPos Then Exit Function
    If endPos > textLength Then Exit Function

    Dim markType As String
    markType = DictKeyString(d, "type")

    Select Case markType
        Case "bold", "italic"
            If DictKeyIsNull(d, "value") Or DictKeyIsEmpty(d, "value") Then Exit Function
            If Not DictKeyIsBoolean(d, "value") Then Exit Function
        Case "script"
            If Not DictKeyIsString(d, "value") Then Exit Function
            If DictKeyString(d, "value") <> "superscript" And DictKeyString(d, "value") <> "subscript" Then Exit Function
        Case "color", "backgroundColor"
            If Not DictKeyIsString(d, "value") Then Exit Function
            If Not FieldIsHexColorString(DictKeyString(d, "value")) Then Exit Function
        Case "link"
            If Not DictKeyIsString(d, "value") Then Exit Function
            If Len(DictKeyString(d, "value")) = 0 Then Exit Function
        Case Else
            Exit Function
    End Select

    FieldIsInlineMark = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldIsInlineMark"
    FieldIsInlineMark = False
End Function

Private Sub ApplyInlineMarkStyle(ByVal targetRange As Range, ByVal mark As Object)
    On Error Resume Next

    Select Case DictKeyString(mark, "type")
        Case "bold"
            targetRange.Font.Bold = DictKeyBool(mark, "value")
        Case "italic"
            targetRange.Font.Italic = DictKeyBool(mark, "value")
        Case "script"
            If DictKeyString(mark, "value") = "superscript" Then
                targetRange.Font.Superscript = True
                targetRange.Font.Subscript = False
            ElseIf DictKeyString(mark, "value") = "subscript" Then
                targetRange.Font.Subscript = True
                targetRange.Font.Superscript = False
            End If
        Case "color"
            targetRange.Font.Color = HexColorIndex(DictKeyString(mark, "value"))
        Case "backgroundColor"
            targetRange.Shading.BackgroundPatternColor = HexColorIndex(DictKeyString(mark, "value"))
    End Select
    On Error GoTo 0
End Sub

Private Sub ApplyRichTextLink(ByVal targetRange As Range, ByVal link As String)
    On Error GoTo ErrHandler
    If Len(link) = 0 Then Exit Sub

    If Left$(LCase$(link), Len("banyan://entry/")) = "banyan://entry/" Then
        ActiveDocument.Hyperlinks.Add Anchor:=targetRange, _
                                      Address:="", _
                                      SubAddress:=FieldGetBibliographyBookmarkName(Mid$(link, Len("banyan://entry/") + 1))
    ElseIf Left$(LCase$(link), 7) = "http://" Or Left$(LCase$(link), 8) = "https://" Then
        ActiveDocument.Hyperlinks.Add Anchor:=targetRange, Address:=link
    End If
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modField.ApplyRichTextLink"
End Sub

' --- Misc helpers ---

Public Function FieldCreateEmptyRichText() As Object
    Dim content As Object
    Set content = New Dictionary
    content("text") = ""
    Set content("marks") = New Collection
    Set FieldCreateEmptyRichText = content
End Function

Public Function FieldCreateRichText(ByVal text As String, Optional ByVal color As String = "") As Object
    Dim content As Object
    Set content = New Dictionary
    content("text") = text

    Dim marks As Collection
    Set marks = New Collection
    If Len(text) > 0 And Len(color) > 0 Then
        marks.Add CreateInlineMark("color", 0, Len(text), color)
    End If
    Set content("marks") = marks

    Set FieldCreateRichText = content
End Function

Private Function CreateInlineMark(ByVal markType As String, _
                                  ByVal startPos As Long, _
                                  ByVal endPos As Long, _
                                  ByVal value As Variant) As Object
    Dim mark As Object
    Set mark = New Dictionary
    mark("type") = markType
    mark("start") = startPos
    mark("end") = endPos
    If IsObject(value) Then
        Set mark("value") = value
    Else
        mark("value") = value
    End If
    Set CreateInlineMark = mark
End Function

Private Function MakeFieldAndId(ByVal fld As Field, ByVal id As String) As Collection
    Dim result As Collection
    Set result = New Collection
    result.Add id, "id"
    result.Add fld, "field"
    Set MakeFieldAndId = result
End Function

Private Function MakeNoteFieldAndId(ByVal note As Footnote, ByVal fld As Field, ByVal id As String) As Collection
    Dim result As Collection
    Set result = New Collection
    result.Add id, "id"
    result.Add note, "note"
    result.Add fld, "field"
    Set MakeNoteFieldAndId = result
End Function

Private Function FindWordStyle(ByVal styleName As String) As Style
    On Error GoTo ErrHandler

    Dim style As Style
    For Each style In ActiveDocument.Styles
        If style.NameLocal = styleName Then
            Set FindWordStyle = style
            Exit Function
        End If
    Next style
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FindWordStyle"
    Set FindWordStyle = Nothing
End Function

Private Function FieldIntextCitationStyleName() As String
    If I10nIsLangZH() Then
        FieldIntextCitationStyleName = INTEXT_STYLE_ZH
    Else
        FieldIntextCitationStyleName = INTEXT_STYLE_EN
    End If
End Function

Private Function HexColorIndex(ByVal hexColor As String) As Long
    If Not FieldIsHexColorString(hexColor) Then
        HexColorIndex = wdColorAutomatic
        Exit Function
    End If

    Dim normalized As String
    normalized = Trim$(hexColor)
    If Left$(normalized, 1) = "#" Then normalized = Mid$(normalized, 2)

    If Len(normalized) = 3 Then
        normalized = Mid$(normalized, 1, 1) & Mid$(normalized, 1, 1) & _
                     Mid$(normalized, 2, 1) & Mid$(normalized, 2, 1) & _
                     Mid$(normalized, 3, 1) & Mid$(normalized, 3, 1)
    End If

    If Len(normalized) <> 6 Then
        HexColorIndex = wdColorAutomatic
        Exit Function
    End If

    Dim r As Long
    Dim g As Long
    Dim b As Long
    r = CLng("&H" & Mid$(normalized, 1, 2))
    g = CLng("&H" & Mid$(normalized, 3, 2))
    b = CLng("&H" & Mid$(normalized, 5, 2))
    HexColorIndex = RGB(r, g, b)
End Function

Private Function FieldIsHexColorString(ByVal hexColor As String) As Boolean
    On Error GoTo ErrHandler

    Dim normalized As String
    normalized = Trim$(hexColor)
    If Left$(normalized, 1) = "#" Then normalized = Mid$(normalized, 2)

    If Len(normalized) = 3 Then
        normalized = Mid$(normalized, 1, 1) & Mid$(normalized, 1, 1) & _
                     Mid$(normalized, 2, 1) & Mid$(normalized, 2, 1) & _
                     Mid$(normalized, 3, 1) & Mid$(normalized, 3, 1)
    End If

    If Len(normalized) <> 6 Then Exit Function

    Dim i As Long
    Dim ch As String
    For i = 1 To 6
        ch = Mid$(normalized, i, 1)
        If InStr(1, "0123456789abcdefABCDEF", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i

    FieldIsHexColorString = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldIsHexColorString"
    FieldIsHexColorString = False
End Function

Private Function RandomHex(ByVal length As Long) As String
    Dim i As Long
    Dim s As String
    For i = 1 To length
        s = s & Hex$(Int(Rnd() * 16))
    Next i
    RandomHex = s
End Function
