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
'   FieldApplyRichTextLinksToFieldResult(field)
'   FieldCollectIntextCitationFieldsInRange(range)
'   FieldCollectNoteCitationFootnotesInRange(range)
'   FieldMigrateIntextCitationsToNotes(range)
'   FieldMigrateNoteCitationsToIntext(range)
'   FieldRemoveFieldSafely(field) / FieldRemoveFootnoteSafely(footnote)
'   FieldAddBookmarkToField(field, bookmarkName)
'   FieldGetBibliographyBookmarkName(entryId)
' ============================================================================

Private Const FIELD_PLACEHOLDER_COLOR As String = "#ff0000"
Private Const FIELD_RAW_PLACEHOLDER As String = "{Citation}"

Private Const INTEXT_STYLE_ZH As String = "Banyan 引注"
Private Const INTEXT_STYLE_EN As String = "Banyan Citation"


' --- JSON data ---

Public Function FieldReadData(ByVal fld As Field) As Object
    On Error GoTo ErrHandler

    Dim jsonText As String
    jsonText = fld.Data
    If Len(jsonText) = 0 Then Exit Function

    Dim data As Object
    Set data = JsonParse(jsonText)
    If data Is Nothing Then Exit Function

    Set FieldReadData = data
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldReadData"
    Set FieldReadData = Nothing
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
    Set fld = FieldCreateRawAddinField(targetRange, "BANYAN_CITATION " & DictKeyString(data, "id"))
    FieldWriteData fld, data
    FieldRenderStyledField fld

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
    Set fld = FieldCreateRawAddinField(noteRange, "BANYAN_CITATION " & DictKeyString(data, "id"))
    FieldWriteData fld, data
    FieldRenderStyledField fld

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

    Dim oldRefText As String
    Dim oldData As Object
    Set oldData = FieldReadData(fld)
    If Not oldData Is Nothing Then
        oldRefText = FieldPlainTextFromContent(GetOptionalRichText(oldData, "reference"))
    End If

    If newRefText = oldRefText Then
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
        Set newField = FieldCreateRawAddinField(noteRange, "BANYAN_CITATION " & DictKeyString(data, "id"))
        If Not newField Is Nothing Then
            FieldWriteData newField, data
            FieldRenderStyledField newField
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
    Set newField = FieldCreateRawAddinField(insertAt, "BANYAN_CITATION " & DictKeyString(data, "id"))
    If newField Is Nothing Then Exit Function

    FieldWriteData newField, data
    FieldRenderStyledField newField

    Set FieldReplaceCitationInNote = newField
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldReplaceCitationInNote"
    Set FieldReplaceCitationInNote = Nothing
End Function

Private Function FieldFindFirstCitationInNote(ByVal note As Footnote) As Field
    On Error GoTo ErrHandler
    Dim f As Field
    For Each f In note.Range.Fields
        If InStr(1, f.Code.Text, "BANYAN_CITATION", vbTextCompare) > 0 Then
            Set FieldFindFirstCitationInNote = f
            Exit Function
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

    WriteContentTextToRange fld.Result, resolvedContent
    fld.ShowCodes = False

    Dim data As Object
    Set data = FieldReadData(fld)
    If Not data Is Nothing Then
        If HasDictionaryKey(data, "type") Then
            If DictKeyString(data, "type") = "intext-citation" Then
                FieldApplyIntextCitationStyle fld
            ElseIf DictKeyString(data, "type") = "note-citation" Then
                FieldApplyNoteCitationStyle fld
            End If
        End If
    End If

    FieldApplyRichTextStylesToRange fld.Result, resolvedContent
    FieldApplyRichTextLinksToRange fld.Result, resolvedContent

    FieldRenderStyledField = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldRenderStyledField"
    FieldRenderStyledField = False
End Function

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
    Set style = FindWordStyle(styleName)
    If style Is Nothing Then
        Set style = ActiveDocument.Styles.Add(Name:=styleName, Type:=styleType)
    End If

    style.UnhideWhenUsed = True
    style.QuickStyle = True
    fld.Result.Style = style
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldApplyStyleToField"
End Sub


' --- Collectors ---

Public Function FieldCollectIntextCitationFieldsInRange(ByVal targetRange As Range) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim fld As Field
    Dim data As Object
    For Each fld In targetRange.Fields
        If fld.Type = wdFieldAddin Then
            Set data = FieldReadData(fld)
            If FieldIsIntextCitation(data) Then
                result.Add MakeFieldAndData(fld, data)
            End If
        End If
    Next fld

    Set FieldCollectIntextCitationFieldsInRange = result
End Function

Public Function FieldCollectNoteCitationFootnotesInRange(ByVal targetRange As Range) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim note As Footnote
    Dim fld As Field
    Dim data As Object
    For Each note In targetRange.Footnotes
        If note.Range.Fields.Count > 0 Then
            Set fld = note.Range.Fields(1)
            If Not fld Is Nothing Then
                If fld.Type = wdFieldAddin Then
                    Set data = FieldReadData(fld)
                    If FieldIsNoteCitation(data) Then
                        result.Add MakeNoteFieldAndData(note, fld, data)
                    End If
                End If
            End If
        End If
    Next note

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
        Set data = fd("data")

        Dim insertPosition As Long
        insertPosition = fld.Result.Start

        Dim convertedData As Object
        Set convertedData = FieldCreatePlaceholderNoteCitationData(DictKeyString(data, "id"), DictKeyObject(data, "source"))

        FieldRemoveFieldSafely fld

        ' Removing the field can shrink the document when it is the last content,
        ' leaving the captured position out of range (error 4608). Clamp it.
        insertPosition = ClampInsertPosition(insertPosition)

        Dim insertRange As Range
        Set insertRange = ActiveDocument.Range(insertPosition, insertPosition)
        FieldCreateNoteCitationAtRange insertRange, convertedData
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
        Set data = fd("data")

        Dim insertPosition As Long
        insertPosition = note.Reference.Start

        Dim convertedData As Object
        Set convertedData = FieldCreatePlaceholderIntextCitationData(DictKeyString(data, "id"), DictKeyObject(data, "source"))

        FieldRemoveFootnoteSafely note

        ' Removing the note can shrink the document when it is the last content,
        ' leaving the captured position out of range (error 4608). Clamp it.
        insertPosition = ClampInsertPosition(insertPosition)

        Dim insertRange As Range
        Set insertRange = ActiveDocument.Range(insertPosition, insertPosition)
        FieldCreateIntextCitationAtRange insertRange, convertedData
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
    If Not IsStringMember(d, "text") Then Exit Function
    If Not HasDictionaryKey(d, "marks") Then Exit Function
    If Not IsCollectionObject(DictKeyObject(d, "marks")) Then Exit Function

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

Public Sub FieldRemoveEndnoteSafely(ByVal note As Endnote)
    On Error Resume Next
    If Not note Is Nothing Then note.Delete
    On Error GoTo 0
End Sub

Public Sub FieldAddBookmarkToField(ByVal fld As Field, ByVal bookmarkName As String)
    On Error GoTo ErrHandler
    If fld Is Nothing Then Exit Sub
    If Len(Trim$(bookmarkName)) = 0 Then Exit Sub

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

Public Function FieldGetBibliographyBookmarkName(ByVal entryId As String) As String
    FieldGetBibliographyBookmarkName = "Banyan_Entry_" & entryId
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

Private Function IsDictionaryRecord(ByVal value As Variant) As Boolean
    IsDictionaryRecord = DictIsDictionary(value)
End Function

Private Function IsCollectionObject(ByVal value As Variant) As Boolean
    IsCollectionObject = DictIsCollection(value)
End Function

Private Function HasDictionaryKey(ByVal dict As Object, ByVal key As String) As Boolean
    HasDictionaryKey = DictHasKey(dict, key)
End Function

Private Function IsStringMember(ByVal dict As Object, ByVal key As String) As Boolean
    IsStringMember = DictKeyIsString(dict, key)
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
    If Not HasDictionaryKey(data, "content") Then Exit Function
    If Not FieldIsRichText(DictKeyObject(data, "content")) Then Exit Function

    Set ResolveFieldContent = DictKeyObject(data, "content")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.ResolveFieldContent"
    Set ResolveFieldContent = Nothing
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

Public Function FieldApplyRichTextLinksToFieldResult(ByVal fld As Field, _
                                                     Optional ByVal content As Variant) As Boolean
    On Error GoTo ErrHandler

    Dim resolvedContent As Object
    Set resolvedContent = ResolveFieldContent(fld, content)
    If resolvedContent Is Nothing Then Exit Function

    FieldApplyRichTextLinksToFieldResult = FieldApplyRichTextLinksToRange(fld.Result, resolvedContent)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modField.FieldApplyRichTextLinksToFieldResult"
    FieldApplyRichTextLinksToFieldResult = False
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
    If Not IsStringMember(d, "type") Then Exit Function
    If Not HasDictionaryKey(d, "start") Then Exit Function
    If Not HasDictionaryKey(d, "end") Then Exit Function
    If Not HasDictionaryKey(d, "value") Then Exit Function
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
            If Not IsStringMember(d, "value") Then Exit Function
            If DictKeyString(d, "value") <> "superscript" And DictKeyString(d, "value") <> "subscript" Then Exit Function
        Case "color", "backgroundColor"
            If Not IsStringMember(d, "value") Then Exit Function
            If Not FieldIsHexColorString(DictKeyString(d, "value")) Then Exit Function
        Case "link"
            If Not IsStringMember(d, "value") Then Exit Function
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

Private Function MakeFieldAndData(ByVal fld As Field, ByVal data As Object) As Collection
    Dim result As Collection
    Set result = New Collection
    result.Add fld, "field"
    result.Add data, "data"
    Set MakeFieldAndData = result
End Function

Private Function MakeNoteFieldAndData(ByVal note As Footnote, ByVal fld As Field, ByVal data As Object) As Collection
    Dim result As Collection
    Set result = New Collection
    result.Add note, "note"
    result.Add fld, "field"
    result.Add data, "data"
    Set MakeNoteFieldAndData = result
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
