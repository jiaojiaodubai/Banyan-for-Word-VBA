Option Explicit

Private m_failure As String

' ============================================================================
' Module  : testField
' Purpose : Tests for modField against the active document.
'
'           Part of the dev-only test suite under test/ (module name test*).
'           Run the FULL suite via test/Run-Tests.ps1.
'
'           Each test constructs its own citation data locally (no backend
'           required), appends a small area at the end of the document,
'           verifies, then cleans up everything it created so runs are
'           repeatable.
'
' Public API:
'   RunTests() -> String report of [PASS]/[FAIL] lines (no dialogs)
' ============================================================================


Public Function RunTests() As String
    Dim report As String
    report = "testField" & vbCrLf & String(40, "-") & vbCrLf
    m_failure = ""
    If ActiveDocument Is Nothing Then
        RunTests = report & "[FAIL] no active document"
        Exit Function
    End If
    report = report & TestResult("create intext + data round-trip", TestFieldCreateIntext()) & vbCrLf
    report = report & TestResult("render rich text (bold/color/link)", TestFieldRenderRichText()) & vbCrLf
    report = report & TestResult("render from Field.Data (no content)", TestFieldRenderFromData()) & vbCrLf
    report = report & TestResult("note citation create + reference", TestFieldNoteCitation()) & vbCrLf
    report = report & TestResult("note source-only refresh writes data (no rebuild)", TestFieldRebuildNoteCitationNoop()) & vbCrLf
    report = report & TestResult("note refresh same reference (no rebuild)", TestFieldRebuildNoteCitation()) & vbCrLf
    report = report & TestResult("note rebuild on reference change (rich copy)", TestFieldRebuildNoteCitationRefChange()) & vbCrLf
    report = report & TestResult("note renumber via custom refs (rebuild)", TestFieldNoteRenumberRebuild()) & vbCrLf
    report = report & TestResult("migrate intext -> note", TestFieldMigrateIntextToNotes()) & vbCrLf
    report = report & TestResult("migrate note -> intext", TestFieldMigrateNotesToIntext()) & vbCrLf
    report = report & TestResult("collectors in range", TestFieldCollectors()) & vbCrLf
    report = report & TestResult("collectors re-key duplicated ids", TestFieldDuplicateCitationIds()) & vbCrLf
    report = report & TestResult("broken citation fields are deleted", TestFieldBrokenCitationFields()) & vbCrLf
    report = report & TestResult("field code contract", TestFieldCodeContract()) & vbCrLf
    report = report & TestResult("data text read (lazy parsing)", TestFieldDataTextRead()) & vbCrLf
    report = report & TestResult("validators", TestFieldValidators()) & vbCrLf
    report = report & TestResult("style identifier", TestFieldStyleIdentifier()) & vbCrLf
    report = report & TestResult("content comparison", TestFieldContentComparison()) & vbCrLf
    report = report & TestResult("targeted rich-text comparison", TestFieldRichTextComparison()) & vbCrLf
    report = report & TestResult("bookmark name normalization", TestFieldBookmarkNameNormalization()) & vbCrLf
    report = report & TestResult("batch screen updating state", TestFieldBatchScreenUpdating()) & vbCrLf
    report = report & TestResult("batch custom undo record", TestFieldBatchUndoRecord()) & vbCrLf
    RunTests = report
End Function

Private Function TestFieldRebuildNoteCitationNoop() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim data As Object
    Set data = FieldCreatePlaceholderNoteCitationData("test-note-noop")
    Set data("reference") = FieldCreateRichText("[1]")

    Dim created As Collection
    Set created = FieldCreateNoteCitationAtRange(TestDocEndRange(ActiveDocument), data)
    If created Is Nothing Then Exit Function

    ' A local result edit makes an unnecessary render/rebuild observable.
    created("field").Result.Text = "LOCAL RESULT"

    Dim updatedData As Object
    Set updatedData = FieldCreatePlaceholderNoteCitationData("test-note-noop")
    Set updatedData("content") = DictKeyObject(data, "content")
    Set updatedData("reference") = DictKeyObject(data, "reference")
    Set updatedData("source") = TestSource("updated-source")

    Dim rebuilt As Collection
    Set rebuilt = FieldRebuildNoteCitationAtRange(created("note"), created("field"), updatedData)
    If rebuilt Is Nothing Then Exit Function

    Dim ok As Boolean
    ok = (rebuilt("note") Is created("note"))
    ok = ok And (rebuilt("field") Is created("field"))
    ok = ok And (rebuilt("field").Result.Text = "LOCAL RESULT")

    Dim storedData As Object
    Set storedData = FieldReadData(rebuilt("field"))
    If storedData Is Nothing Then
        ok = False
    Else
        ok = ok And (CStr(storedData("source")("cites")(1)) = "updated-source")
    End If
    FieldRemoveFootnoteSafely created("note")
    TestFieldRebuildNoteCitationNoop = ok
    Exit Function

ErrHandler:
    TestFieldRebuildNoteCitationNoop = False
End Function

Private Function TestFieldRichTextComparison() As Boolean
    On Error GoTo ErrHandler

    Dim currentContent As Object
    Dim nextContent As Object
    Set currentContent = TestComparisonRichText()
    Set nextContent = TestComparisonRichText()

    Dim ok As Boolean
    ok = FieldRichTextEquals(currentContent, nextContent)

    nextContent("text") = "Different"
    ok = ok And (Not FieldRichTextEquals(currentContent, nextContent))

    Set nextContent = TestComparisonRichText()
    nextContent("marks").Remove 2
    ok = ok And (Not FieldRichTextEquals(currentContent, nextContent))

    Set nextContent = TestComparisonRichText()
    nextContent("marks")(1)("type") = "italic"
    ok = ok And (Not FieldRichTextEquals(currentContent, nextContent))

    Set nextContent = TestComparisonRichText()
    nextContent("marks")(1)("start") = 1
    ok = ok And (Not FieldRichTextEquals(currentContent, nextContent))

    Set nextContent = TestComparisonRichText()
    nextContent("marks")(1)("end") = 3
    ok = ok And (Not FieldRichTextEquals(currentContent, nextContent))

    Set nextContent = TestComparisonRichText()
    nextContent("marks")(1)("value") = False
    ok = ok And (Not FieldRichTextEquals(currentContent, nextContent))

    Set nextContent = TestComparisonRichText()
    nextContent("marks")(2)("value") = "banyan://entry/other"
    ok = ok And (Not FieldRichTextEquals(currentContent, nextContent))

    TestFieldRichTextComparison = ok
    Exit Function

ErrHandler:
    TestFieldRichTextComparison = False
End Function

Private Function TestFieldBookmarkNameNormalization() As Boolean
    On Error GoTo ErrHandler

    ' Valid names pass through unchanged (Zotero keys are 8-char alphanumerics).
    Dim ok As Boolean
    ok = (FieldGetBibliographyBookmarkName("JZGANGMD") = "Banyan_Entry_JZGANGMD")

    ' Spaces and punctuation become underscores, a non-letter start gets a
    ' letter prefix, and the result never exceeds 40 characters.
    ok = ok And (FieldNormalizeBookmarkName("supp fig 1.2") = "supp_fig_1_2")
    ok = ok And (FieldNormalizeBookmarkName("1abc") = "B1abc")
    ok = ok And (FieldNormalizeBookmarkName("_hidden") = "B_hidden")
    ok = ok And (FieldNormalizeBookmarkName("") = "")
    ok = ok And (Len(FieldNormalizeBookmarkName(String(60, "a"))) = 40)
    ok = ok And (Len(FieldGetBibliographyBookmarkName(String(60, "a"))) = 40)
    ok = ok And TestIsWordBookmarkName(FieldGetBibliographyBookmarkName("a b.c-d"))

    ' The document gate normalizes whatever name a caller passes.
    Dim fld As Field
    Set fld = FieldCreateRawAddinField(TestDocEndRange(ActiveDocument), "BANYAN_TEST bookmark")
    If fld Is Nothing Then Exit Function
    FieldAddBookmarkToField fld, "Banyan Entry bad.name"
    ok = ok And ActiveDocument.Bookmarks.Exists("Banyan_Entry_bad_name")

    If ActiveDocument.Bookmarks.Exists("Banyan_Entry_bad_name") Then _
        ActiveDocument.Bookmarks("Banyan_Entry_bad_name").Delete
    FieldRemoveFieldSafely fld

    TestFieldBookmarkNameNormalization = ok
    Exit Function

ErrHandler:
    On Error Resume Next
    If ActiveDocument.Bookmarks.Exists("Banyan_Entry_bad_name") Then _
        ActiveDocument.Bookmarks("Banyan_Entry_bad_name").Delete
    If Not fld Is Nothing Then FieldRemoveFieldSafely fld
    On Error GoTo 0
    TestFieldBookmarkNameNormalization = False
End Function

Private Function TestFieldContentComparison() As Boolean
    On Error GoTo ErrHandler

    Dim currentData As Object
    Dim nextData As Object
    Set currentData = FieldCreatePlaceholderIntextCitationData("test-compare")
    Set nextData = FieldCreatePlaceholderIntextCitationData("test-compare")

    Dim ok As Boolean
    ok = FieldContentEquals(currentData, nextData)

    Set nextData("content") = FieldCreateRichText("[UPDATED]")
    ok = ok And (Not FieldContentEquals(currentData, nextData))

    ' A source-only change must not be treated as a render change.
    Set nextData("content") = DictKeyObject(currentData, "content")
    Set nextData("source") = TestSource("changed-source")
    ok = ok And FieldContentEquals(currentData, nextData)

    TestFieldContentComparison = ok
    Exit Function

ErrHandler:
    TestFieldContentComparison = False
End Function

Private Function TestFieldBatchScreenUpdating() As Boolean
    On Error GoTo ErrHandler

    Dim originalValue As Boolean
    originalValue = Application.ScreenUpdating

    FieldBeginBatchUpdate
    If Application.ScreenUpdating Then GoTo Failed
    FieldBeginBatchUpdate
    If Application.ScreenUpdating Then GoTo Failed
    FieldEndBatchUpdate
    If Application.ScreenUpdating Then GoTo Failed
    FieldEndBatchUpdate
    If Application.ScreenUpdating <> originalValue Then GoTo Failed

    TestFieldBatchScreenUpdating = True
    Exit Function

Failed:
    FieldEndBatchUpdate
    FieldEndBatchUpdate
    Application.ScreenUpdating = originalValue
    Exit Function

ErrHandler:
    On Error Resume Next
    FieldEndBatchUpdate
    FieldEndBatchUpdate
    Application.ScreenUpdating = originalValue
    On Error GoTo 0
    TestFieldBatchScreenUpdating = False
End Function

Private Function TestFieldBatchUndoRecord() As Boolean
    On Error GoTo ErrHandler

    Dim undoRecord As Object
    Set undoRecord = CallByName(Application, "UndoRecord", VbGet)
    If undoRecord Is Nothing Then Exit Function

    Dim originalText As String
    originalText = ActiveDocument.Content.Text

    FieldBeginBatchUpdate
    Dim activeAtOuterDepth As Boolean
    activeAtOuterDepth = CBool(CallByName(undoRecord, "IsRecordingCustomRecord", VbGet))
    TestDocEndRange(ActiveDocument).InsertBefore "undo-outer"

    FieldBeginBatchUpdate
    Dim activeAtNestedDepth As Boolean
    activeAtNestedDepth = CBool(CallByName(undoRecord, "IsRecordingCustomRecord", VbGet))
    TestDocEndRange(ActiveDocument).InsertBefore "undo-nested"
    FieldEndBatchUpdate
    Dim activeAfterNestedEnd As Boolean
    activeAfterNestedEnd = CBool(CallByName(undoRecord, "IsRecordingCustomRecord", VbGet))

    FieldEndBatchUpdate
    Dim activeAfterOuterEnd As Boolean
    activeAfterOuterEnd = CBool(CallByName(undoRecord, "IsRecordingCustomRecord", VbGet))

    Dim undoSucceeded As Boolean
    undoSucceeded = ActiveDocument.Undo
    Dim changesMerged As Boolean
    changesMerged = undoSucceeded And (ActiveDocument.Content.Text = originalText)

    TestFieldBatchUndoRecord = activeAtOuterDepth And activeAtNestedDepth And _
                               activeAfterNestedEnd And Not activeAfterOuterEnd And _
                               changesMerged
    Exit Function

ErrHandler:
    On Error Resume Next
    FieldEndBatchUpdate
    FieldEndBatchUpdate
    If Len(originalText) > 0 Then ActiveDocument.Content.Text = originalText
    On Error GoTo 0
    TestFieldBatchUndoRecord = False
End Function


Private Function TestFieldCreateIntext() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = ActiveDocument
    Dim startPos As Long
    startPos = doc.Content.End
    Dim rng As Range
    Set rng = TestDocEndRange(doc)

    Dim data As Object
    Set data = FieldCreatePlaceholderIntextCitationData("test-int-1")

    Dim fld As Field
    Set fld = FieldCreateIntextCitationAtRange(rng, data)
    If fld Is Nothing Then Exit Function

    Dim ok As Boolean
    ok = (fld.Type = wdFieldAddin)
    ok = ok And (fld.Result.Text = "{ INTEXT_CITATION }")

    Dim rd As Object
    Set rd = FieldReadData(fld)
    If rd Is Nothing Then
        ok = False
    Else
        ok = ok And (CStr(rd("id")) = "test-int-1")
        ok = ok And (CStr(rd("type")) = "intext-citation")
    End If

    FieldRemoveFieldSafely fld
    TestFieldCreateIntext = ok
    Exit Function

ErrHandler:
    TestFieldCreateIntext = False
End Function

Private Function TestFieldRenderRichText() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = ActiveDocument
    Dim startPos As Long
    startPos = doc.Content.End
    Dim rng As Range
    Set rng = TestDocEndRange(doc)

    Dim data As Object
    Set data = FieldCreatePlaceholderIntextCitationData("test-render-1")

    Dim fld As Field
    Set fld = FieldCreateIntextCitationAtRange(rng, data)
    If fld Is Nothing Then Exit Function

    ' Real rich content: "(Zhang, 2020)" with bold(1..6), color(0..13), link(0..13)
    Dim content As Object
    Set content = New Dictionary
    content("text") = "(Zhang, 2020)"
    Dim marks As Collection
    Set marks = New Collection
    marks.Add TestMark("bold", 1, 6, True)
    marks.Add TestMark("color", 0, 13, "#ff0000")
    marks.Add TestMark("link", 0, 13, "banyan://entry/test-render-1")
    Set content("marks") = marks

    Set data("content") = content
    FieldWriteData fld, data

    Dim ok As Boolean
    ok = FieldRenderStyledFieldWithData(fld, data, content)
    ok = ok And (fld.Result.Text = "(Zhang, 2020)")

    Dim res As Range
    Set res = fld.Result
    ok = ok And (res.Hyperlinks.Count >= 1)
    ok = ok And (res.Characters(2).Font.Bold = True)   ' 'Z' of "Zhang" is bold

    FieldRemoveFieldSafely fld
    TestFieldRenderRichText = ok
    Exit Function

ErrHandler:
    TestFieldRenderRichText = False
End Function

Private Function TestFieldRenderFromData() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = ActiveDocument
    Dim startPos As Long
    startPos = doc.Content.End
    Dim rng As Range
    Set rng = TestDocEndRange(doc)

    ' Build data with real content directly (as refresh would write it)
    Dim content As Object
    Set content = New Dictionary
    content("text") = "(Wang 2019, 5)"
    Dim marks As Collection
    Set marks = New Collection
    marks.Add TestMark("bold", 0, 4, True)
    Set content("marks") = marks

    Dim data As Object
    Set data = New Dictionary
    data("id") = "test-data-1"
    data("type") = "intext-citation"
    Set data("source") = TestSource("Q-1")
    Set data("content") = content

    Dim fld As Field
    Set fld = FieldCreateRawAddinField(rng, "BANYAN_CITATION test-data-1")
    If fld Is Nothing Then Exit Function
    FieldWriteData fld, data

    Dim ok As Boolean
    ok = FieldRenderStyledField(fld)              ' no content arg -> reads data("content")
    ok = ok And (fld.Result.Text = "(Wang 2019, 5)")
    ' Bold mark covers 0..4; the first field-result char can report wdUndefined,
    ' so assert on char index 1 ('W') which reliably reflects the direct bold.
    ok = ok And (fld.Result.Characters(2).Font.Bold = True)

    FieldRemoveFieldSafely fld
    TestFieldRenderFromData = ok
    Exit Function

ErrHandler:
    TestFieldRenderFromData = False
End Function

Private Function TestFieldNoteCitation() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = ActiveDocument
    Dim startPos As Long
    startPos = doc.Content.End
    Dim rng As Range
    Set rng = TestDocEndRange(doc)

    Dim data As Object
    Set data = FieldCreatePlaceholderNoteCitationData("test-note-1")

    ' Add a reference rich text (the custom footnote reference mark)
    Dim refContent As Object
    Set refContent = New Dictionary
    refContent("text") = "Note"
    Set refContent("marks") = New Collection
    Set data("reference") = refContent

    Dim created As Collection
    Set created = FieldCreateNoteCitationAtRange(rng, data)
    If created Is Nothing Then Exit Function

    Dim note As Footnote
    Dim fld As Field
    Set note = created("note")
    Set fld = created("field")

    Dim ok As Boolean
    ok = (fld.Type = wdFieldAddin)
    ok = ok And (note.Reference.Text = "Note")
    ok = ok And (fld.Result.Text = "{ NOTE_CITATION }")

    FieldRemoveFootnoteSafely note
    TestFieldNoteCitation = ok
    Exit Function

ErrHandler:
    TestFieldNoteCitation = False
End Function

Private Function TestFieldRebuildNoteCitation() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = ActiveDocument
    Dim rng As Range
    Set rng = TestDocEndRange(doc)

    ' Custom reference "[1]". The rebuild data carries the SAME reference, so
    ' FieldRebuildNoteCitationAtRange must take the no-rebuild path: the footnote
    ' is kept and only the field data + result are refreshed - the body (and its
    ' rich text) is never touched.
    Dim refContent As Object
    Set refContent = New Dictionary
    refContent("text") = "[1]"
    Set refContent("marks") = New Collection

    Dim data As Object
    Set data = FieldCreatePlaceholderNoteCitationData("test-note-rebuild")
    Set data("reference") = refContent

    Dim created As Collection
    Set created = FieldCreateNoteCitationAtRange(rng, data)
    If created Is Nothing Then Exit Function

    Dim note As Footnote
    Dim fld As Field
    Set note = created("note")
    Set fld = created("field")

    ' The footnote-area reference mark must survive creation (regression: the
    ' old code cleared the whole footnote paragraph and deleted the mark).
    Dim para As Range
    Set para = note.Range.Duplicate
    para.Expand wdParagraph
    If Left$(para.Text, 3) <> "[1]" Then Exit Function

    ' User-typed content around the field (outside the field result).
    Dim t1 As Range
    Set t1 = note.Range.Duplicate
    t1.Collapse wdCollapseStart
    t1.InsertBefore " SUFFIX"
    Dim t2 As Range
    Set t2 = note.Range.Duplicate
    t2.Collapse wdCollapseEnd
    t2.InsertAfter "PREFIX "

    ' Updated data with new content (as /refresh would return).
    Dim updatedData As Object
    Set updatedData = FieldCreatePlaceholderNoteCitationData("test-note-rebuild")
    Set updatedData("content") = FieldCreateRichText("[UPDATED]", "#0000ff")
    Set updatedData("reference") = refContent

    Dim rebuilt As Collection
    Set rebuilt = FieldRebuildNoteCitationAtRange(note, fld, updatedData)
    If rebuilt Is Nothing Then Exit Function

    Dim note2 As Footnote
    Dim fld2 As Field
    Set note2 = rebuilt("note")
    Set fld2 = rebuilt("field")

    Dim ok As Boolean
    ok = True

    ' 0. no-rebuild path: the same footnote is returned (nothing re-created)
    ok = ok And (note2 Is note)

    ' 1. footnote-area mark still present
    Set para = note2.Range.Duplicate
    para.Expand wdParagraph
    ok = ok And (Left$(para.Text, 3) = "[1]")

    ' 2. user-typed content preserved around the refreshed field
    '    (placement varies by Word build, so check membership, not order)
    ok = ok And (InStr(note2.Range.Text, "[UPDATED]") > 0)
    ok = ok And (InStr(note2.Range.Text, "PREFIX") > 0)
    ok = ok And (InStr(note2.Range.Text, "SUFFIX") > 0)

    ' 3. field data updated
    Dim fldData As Object
    Set fldData = FieldReadData(fld2)
    If fldData Is Nothing Then
        ok = False
    Else
        ok = ok And (CStr(fldData("content")("text")) = "[UPDATED]")
    End If

    ' 4. field result rendered with the new rich text, exactly one field
    ok = ok And (fld2.Result.Text = "[UPDATED]")
    ok = ok And (note2.Range.Fields.Count = 1)

    ' 5. the fresh field must NOT inherit the placeholder's red color
    ok = ok And (fld2.Result.Font.Color <> 255)

    FieldRemoveFootnoteSafely note2
    TestFieldRebuildNoteCitation = ok
    Exit Function

ErrHandler:
    TestFieldRebuildNoteCitation = False
End Function

Private Function TestFieldRebuildNoteCitationRefChange() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = ActiveDocument
    Dim rng As Range
    Set rng = TestDocEndRange(doc)

    ' Create with reference "[1]".
    Dim ref1 As Object
    Set ref1 = New Dictionary
    ref1("text") = "[1]"
    Set ref1("marks") = New Collection

    Dim data As Object
    Set data = FieldCreatePlaceholderNoteCitationData("test-note-rebuild2")
    Set data("reference") = ref1

    Dim created As Collection
    Set created = FieldCreateNoteCitationAtRange(rng, data)
    If created Is Nothing Then Exit Function

    Dim note As Footnote
    Dim fld As Field
    Set note = created("note")
    Set fld = created("field")

    ' User-typed content around the field (outside the field result).
    Dim t1 As Range
    Set t1 = note.Range.Duplicate
    t1.Collapse wdCollapseStart
    t1.InsertBefore " SUFFIX"
    Dim t2 As Range
    Set t2 = note.Range.Duplicate
    t2.Collapse wdCollapseEnd
    t2.InsertAfter "PREFIX "

    ' Updated data: NEW content + CHANGED reference "[2]" - this must trigger
    ' the rebuild path, which recreates the footnote and copies the rich body.
    Dim ref2 As Object
    Set ref2 = New Dictionary
    ref2("text") = "[2]"
    Set ref2("marks") = New Collection

    Dim updatedData As Object
    Set updatedData = FieldCreatePlaceholderNoteCitationData("test-note-rebuild2")
    Set updatedData("content") = FieldCreateRichText("[UPDATED]", "#0000ff")
    Set updatedData("reference") = ref2

    Dim rebuilt As Collection
    Set rebuilt = FieldRebuildNoteCitationAtRange(note, fld, updatedData)
    If rebuilt Is Nothing Then Exit Function

    Dim note2 As Footnote
    Dim fld2 As Field
    Set note2 = rebuilt("note")
    Set fld2 = rebuilt("field")

    Dim ok As Boolean
    ok = True

    ' 1. reference mark updated to "[2]" in the main text
    ok = ok And (note2.Reference.Text = "[2]")

    ' 2. footnote-area mark present and shows the new reference
    Dim para As Range
    Set para = note2.Range.Duplicate
    para.Expand wdParagraph
    ok = ok And (Left$(para.Text, 3) = "[2]")

    ' 3. user-typed content preserved around the refreshed field
    ok = ok And (InStr(note2.Range.Text, "[UPDATED]") > 0)
    ok = ok And (InStr(note2.Range.Text, "PREFIX") > 0)
    ok = ok And (InStr(note2.Range.Text, "SUFFIX") > 0)

    ' 4. field data updated (content + reference)
    Dim fldData As Object
    Set fldData = FieldReadData(fld2)
    If fldData Is Nothing Then
        ok = False
    Else
        ok = ok And (CStr(fldData("content")("text")) = "[UPDATED]")
        ok = ok And (CStr(fldData("reference")("text")) = "[2]")
    End If

    ' 5. field result rendered with the new rich text, exactly one field
    ok = ok And (fld2.Result.Text = "[UPDATED]")
    ok = ok And (note2.Range.Fields.Count = 1)

    ' 6. the fresh field must NOT inherit the placeholder's red color
    ok = ok And (fld2.Result.Font.Color <> 255)

    FieldRemoveFootnoteSafely note2
    TestFieldRebuildNoteCitationRefChange = ok
    Exit Function

ErrHandler:
    TestFieldRebuildNoteCitationRefChange = False
End Function

Private Function TestFieldNoteRenumberRebuild() As Boolean
    ' Smoke test that simulates a backend /refresh renumbering with CUSTOM
    ' references (bracketed numbers), proving the rebuild strategy genuinely
    ' recreates the footnote (it is not just Word auto-numbering):
    '   1) Insert note A with custom reference "[1]" - renders as "[1]".
    '   2) Insert note B BEFORE A with reference "[1]".
    '   3) Simulate the refresh response assigning references in document order:
    '      B -> "[1]" (unchanged -> no rebuild), A -> "[2]" (CHANGED -> the old
    '      footnote is deleted and recreated; rich user content must survive).
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = ActiveDocument

    ' ---- Phase 1: insert note A with custom reference "[1]" ----
    Dim refA As Object
    Set refA = New Dictionary
    refA("text") = "[1]"
    Set refA("marks") = New Collection

    Dim dataA As Object
    Set dataA = FieldCreatePlaceholderNoteCitationData("renumber-A")
    Set dataA("content") = FieldCreateRichText("ContentA", "")
    Set dataA("reference") = refA

    Dim rngA As Range
    Set rngA = TestDocEndRange(doc)
    Dim createdA As Collection
    Set createdA = FieldCreateNoteCitationAtRange(rngA, dataA)
    If createdA Is Nothing Then Exit Function

    Dim noteA As Footnote
    Dim fldA As Field
    Set noteA = createdA("note")
    Set fldA = createdA("field")

    ' user-typed rich content around A's field (must survive the rebuild)
    Dim ta1 As Range
    Set ta1 = noteA.Range.Duplicate
    ta1.Collapse wdCollapseStart
    ta1.InsertBefore " SUFFIXA"
    Dim ta2 As Range
    Set ta2 = noteA.Range.Duplicate
    ta2.Collapse wdCollapseEnd
    ta2.InsertAfter "PREFIXA "

    Dim ok As Boolean
    ok = True

    ' A renders with custom reference "[1]"
    ok = ok And (noteA.Reference.Text = "[1]")
    Dim paraA0 As Range
    Set paraA0 = noteA.Range.Duplicate
    paraA0.Expand wdParagraph
    ok = ok And (Left$(paraA0.Text, 3) = "[1]")
    ok = ok And (InStr(noteA.Range.Text, "ContentA") > 0)

    ' ---- Phase 2: insert note B BEFORE A with reference "[1]" ----
    Dim refB As Object
    Set refB = New Dictionary
    refB("text") = "[1]"
    Set refB("marks") = New Collection

    Dim dataB As Object
    Set dataB = FieldCreatePlaceholderNoteCitationData("renumber-B")
    Set dataB("content") = FieldCreateRichText("ContentB", "")
    Set dataB("reference") = refB

    Dim rngB As Range
    Set rngB = doc.Content.Duplicate
    rngB.Collapse wdCollapseStart
    Dim createdB As Collection
    Set createdB = FieldCreateNoteCitationAtRange(rngB, dataB)
    If createdB Is Nothing Then Exit Function

    Dim noteB As Footnote
    Dim fldB As Field
    Set noteB = createdB("note")
    Set fldB = createdB("field")

    ok = ok And (doc.Footnotes.Count = 2)
    ok = ok And (noteB.Reference.Text = "[1]")

    ' ---- Phase 3: simulate /refresh renumbering (B first, A second) ----
    ' B keeps "[1]" (unchanged -> no rebuild).
    Dim updatedB As Object
    Set updatedB = FieldCreatePlaceholderNoteCitationData("renumber-B")
    Set updatedB("content") = FieldCreateRichText("ContentB2", "")
    Set updatedB("reference") = refB

    Dim rebuiltB As Collection
    Set rebuiltB = FieldRebuildNoteCitationAtRange(noteB, fldB, updatedB)
    If rebuiltB Is Nothing Then Exit Function
    Dim noteB2 As Footnote
    Dim fldB2 As Field
    Set noteB2 = rebuiltB("note")
    Set fldB2 = rebuiltB("field")

    ok = ok And (noteB2 Is noteB)             ' B: not rebuilt
    ok = ok And (noteB2.Reference.Text = "[1]")
    ok = ok And (fldB2.Result.Text = "ContentB2")

    ' A is renumbered to "[2]" (CHANGED -> footnote genuinely rebuilt).
    Dim refA2 As Object
    Set refA2 = New Dictionary
    refA2("text") = "[2]"
    Set refA2("marks") = New Collection

    Dim updatedA As Object
    Set updatedA = FieldCreatePlaceholderNoteCitationData("renumber-A")
    Set updatedA("content") = FieldCreateRichText("ContentA2", "")
    Set updatedA("reference") = refA2

    Dim rebuiltA As Collection
    Set rebuiltA = FieldRebuildNoteCitationAtRange(noteA, fldA, updatedA)
    If rebuiltA Is Nothing Then Exit Function
    Dim noteA2 As Footnote
    Dim fldA2 As Field
    Set noteA2 = rebuiltA("note")
    Set fldA2 = rebuiltA("field")

    ok = ok And (Not noteA2 Is noteA)         ' A WAS rebuilt (new footnote object)
    ok = ok And (noteA2.Reference.Text = "[2]")
    Dim paraA2 As Range
    Set paraA2 = noteA2.Range.Duplicate
    paraA2.Expand wdParagraph
    ok = ok And (Left$(paraA2.Text, 3) = "[2]")

    ' rich user content preserved around A's rebuilt field
    ok = ok And (InStr(noteA2.Range.Text, "ContentA2") > 0)
    ok = ok And (InStr(noteA2.Range.Text, "PREFIXA") > 0)
    ok = ok And (InStr(noteA2.Range.Text, "SUFFIXA") > 0)

    ' document order after refresh: B = "[1]", A = "[2]"
    ok = ok And (doc.Footnotes.Count = 2)
    ok = ok And (doc.Footnotes.Item(1).Reference.Text = "[1]")
    ok = ok And (doc.Footnotes.Item(2).Reference.Text = "[2]")

    ' cleanup
    FieldRemoveFootnoteSafely noteB2
    FieldRemoveFootnoteSafely noteA2
    TestFieldNoteRenumberRebuild = ok
    Exit Function

ErrHandler:
    TestFieldNoteRenumberRebuild = False
End Function

Private Function TestFieldMigrateIntextToNotes() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = ActiveDocument
    Dim startPos As Long
    startPos = doc.Content.End
    Dim rng As Range
    Set rng = TestDocEndRange(doc)
    rng.Text = "Migrate I2N "
    Dim endPos As Long
    endPos = doc.Content.End
    Set rng = TestDocEndRange(doc)

    Dim data As Object
    Set data = FieldCreatePlaceholderIntextCitationData("test-mig-i2n")
    Dim fld As Field
    Set fld = FieldCreateIntextCitationAtRange(rng, data)
    If fld Is Nothing Then Exit Function

    ' Scope the migration to the test area only
    Dim targetRange As Range
    Set targetRange = doc.Range(startPos, doc.Content.End)
    FieldMigrateIntextCitationsToNotes targetRange

    ' Verify: a footnote carrying our id exists
    Dim foundNote As Footnote
    Set foundNote = Nothing
    Dim note As Footnote
    Dim nfld As Field
    Dim ndata As Object
    For Each note In doc.Footnotes
        If note.Range.Fields.Count > 0 Then
            Set nfld = note.Range.Fields(1)
            Set ndata = FieldReadData(nfld)
            If Not ndata Is Nothing Then
                If CStr(ndata("id")) = "test-mig-i2n" Then
                    Set foundNote = note
                    Exit For
                End If
            End If
        End If
    Next note

    Dim ok As Boolean
    ok = (Not foundNote Is Nothing)

    If Not foundNote Is Nothing Then FieldRemoveFootnoteSafely foundNote
    On Error Resume Next
    doc.Range(startPos, doc.Content.End).Delete
    On Error GoTo 0
    TestFieldMigrateIntextToNotes = ok
    Exit Function

ErrHandler:
    TestFieldMigrateIntextToNotes = False
End Function

Private Function TestFieldMigrateNotesToIntext() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = ActiveDocument
    Dim startPos As Long
    startPos = doc.Content.End
    Dim rng As Range
    Set rng = TestDocEndRange(doc)
    rng.Text = "Migrate N2I "
    Dim endPos As Long
    endPos = doc.Content.End
    Set rng = TestDocEndRange(doc)

    Dim data As Object
    Set data = FieldCreatePlaceholderNoteCitationData("test-mig-n2i")
    Dim created As Collection
    Set created = FieldCreateNoteCitationAtRange(rng, data)
    If created Is Nothing Then Exit Function

    Dim targetRange As Range
    Set targetRange = doc.Range(startPos, doc.Content.End)
    FieldMigrateNoteCitationsToIntext targetRange

    ' Verify: an intext field carrying our id exists
    Dim foundField As Field
    Set foundField = Nothing
    Dim fld As Field
    Dim fdata As Object
    For Each fld In doc.Fields
        If fld.Type = wdFieldAddin Then
            Set fdata = FieldReadData(fld)
            If Not fdata Is Nothing Then
                If CStr(fdata("id")) = "test-mig-n2i" Then
                    Set foundField = fld
                    Exit For
                End If
            End If
        End If
    Next fld

    Dim ok As Boolean
    ok = (Not foundField Is Nothing)

    If Not foundField Is Nothing Then FieldRemoveFieldSafely foundField
    On Error Resume Next
    doc.Range(startPos, doc.Content.End).Delete
    On Error GoTo 0
    TestFieldMigrateNotesToIntext = ok
    Exit Function

ErrHandler:
    TestFieldMigrateNotesToIntext = False
End Function

Private Function TestFieldCollectors() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = ActiveDocument
    Dim startPos As Long
    startPos = doc.Content.End
    Dim rng As Range
    Set rng = TestDocEndRange(doc)
    rng.Text = "Collect "
    Dim endPos As Long
    endPos = doc.Content.End

    ' 2 intext citations
    Dim f1 As Field
    Dim f2 As Field
    Dim data1 As Object
    Dim data2 As Object
    Set data1 = FieldCreatePlaceholderIntextCitationData("test-col-1")
    Set data2 = FieldCreatePlaceholderIntextCitationData("test-col-2")
    Set rng = TestDocEndRange(doc)
    Set f1 = FieldCreateIntextCitationAtRange(rng, data1)
    endPos = doc.Content.End
    Set rng = TestDocEndRange(doc)
    Set f2 = FieldCreateIntextCitationAtRange(rng, data2)

    ' 1 note citation
    Dim data3 As Object
    Set data3 = FieldCreatePlaceholderNoteCitationData("test-col-3")
    Dim created As Collection
    endPos = doc.Content.End
    Set rng = TestDocEndRange(doc)
    Set created = FieldCreateNoteCitationAtRange(rng, data3)
    If created Is Nothing Then Exit Function

    Dim targetRange As Range
    Set targetRange = doc.Range(startPos, doc.Content.End)

    Dim intextCol As Collection
    Set intextCol = FieldCollectIntextCitationFieldsInRange(targetRange)
    Dim noteCol As Collection
    Set noteCol = FieldCollectNoteCitationFootnotesInRange(targetRange)

    Dim ok As Boolean
    ok = (intextCol.Count = 2)
    ok = ok And (noteCol.Count = 1)
    If ok Then
        ' Pairs are {id, field}; the id comes from the code.
        ok = (CStr(intextCol(1)("id")) = "test-col-1")
        ok = ok And (CStr(intextCol(2)("id")) = "test-col-2")
        ok = ok And (CStr(noteCol(1)("id")) = "test-col-3")

        Dim intextField As Field
        Dim noteField As Field
        Set intextField = intextCol(1)("field")
        Set noteField = noteCol(1)("field")
        ok = ok And FieldHasCodeKind(intextField, FIELD_KIND_CITATION)
        ok = ok And FieldHasCodeKind(noteField, FIELD_KIND_CITATION)
        ok = ok And (CStr(FieldReadData(intextField)("id")) = CStr(intextCol(1)("id")))
        ok = ok And (CStr(FieldReadData(noteField)("id")) = CStr(noteCol(1)("id")))
    End If

    FieldRemoveFieldSafely f1
    FieldRemoveFieldSafely f2
    FieldRemoveFootnoteSafely created("note")
    On Error Resume Next
    doc.Range(startPos, doc.Content.End).Delete
    On Error GoTo 0
    TestFieldCollectors = ok
    Exit Function

ErrHandler:
    TestFieldCollectors = False
End Function

' A pasted citation arrives with a duplicated id (Word reports no paste); the
' collector re-keys the field code as it parses it and reports the new id, while
' the stored data keeps the pasted id until the refresh writes the response.
Private Function TestFieldDuplicateCitationIds() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = ActiveDocument
    Dim startPos As Long
    startPos = doc.Content.End
    Dim rng As Range
    Set rng = TestDocEndRange(doc)
    rng.Text = "Dup "

    Dim data As Object
    Set data = FieldCreatePlaceholderIntextCitationData("dup-intext")
    Dim first As Field
    Dim copy As Field
    Set first = FieldCreateIntextCitationAtRange(TestDocEndRange(doc), data)
    Set copy = FieldCreateIntextCitationAtRange(TestDocEndRange(doc), data)
    If first Is Nothing Or copy Is Nothing Then
        m_failure = "in-text field creation failed"
        GoTo CleanUp
    End If

    Dim noteData As Object
    Set noteData = FieldCreatePlaceholderNoteCitationData("dup-note")
    Dim created As Collection
    Dim noteCopy As Collection
    Set created = FieldCreateNoteCitationAtRange(TestDocEndRange(doc), noteData)
    Set noteCopy = FieldCreateNoteCitationAtRange(TestDocEndRange(doc), noteData)
    If created Is Nothing Or noteCopy Is Nothing Then
        m_failure = "note field creation failed"
        GoTo CleanUp
    End If

    Dim targetRange As Range
    Set targetRange = doc.Range(startPos, doc.Content.End)

    Dim ok As Boolean
    Dim intextCol As Collection
    Set intextCol = FieldCollectIntextCitationFieldsInRange(targetRange)
    ok = (intextCol.Count = 2)
    If Not ok Then m_failure = "intext count=" & CStr(intextCol.Count)
    Dim keptId As String
    Dim newId As String
    If intextCol.Count = 2 Then
        Dim copyField As Field
        Dim copyData As Object
        Dim copyText As String
        keptId = CStr(intextCol(1)("id"))
        newId = CStr(intextCol(2)("id"))
        Set copyField = intextCol(2)("field")
        copyText = FieldDataText(copyField)
        Set copyData = FieldReadData(copyField)

        If keptId <> "dup-intext" Then m_failure = "kept id=" & keptId
        If Len(newId) <> 36 Then m_failure = "new id=" & newId
        If newId = keptId Then m_failure = "copy was not re-keyed"
        If Not FieldHasCodeKind(copyField, FIELD_KIND_CITATION) Then m_failure = "code kind lost"
        If TestFieldCodeId(copyField) <> newId Then m_failure = "code does not carry the new id"

        ok = ok And (keptId = "dup-intext")
        ok = ok And (Len(newId) = 36)
        ok = ok And (newId <> keptId)
        ok = ok And FieldHasCodeKind(copyField, FIELD_KIND_CITATION)
        ' The code carries the new id; the pasted payload stays untouched.
        ok = ok And (TestFieldCodeId(copyField) = newId)

        ' The re-key touches the code only: the pasted payload must survive.
        If copyData Is Nothing Then
            m_failure = "copy data unreadable, text len=" & CStr(Len(copyText))
            ok = False
        Else
            If CStr(copyData("id")) <> keptId Then m_failure = "data id now=" & CStr(copyData("id"))
            ok = ok And (CStr(copyData("id")) = keptId)
        End If
        If Len(copyText) = 0 Or InStr(copyText, keptId) = 0 Then
            m_failure = "stored text lost: len=" & CStr(Len(copyText))
            ok = False
        End If
    End If

    ' The next collection is stable: a re-keyed code is not touched again.
    Set intextCol = FieldCollectIntextCitationFieldsInRange(targetRange)
    ok = ok And (intextCol.Count = 2)
    If intextCol.Count <> 2 Then m_failure = "second pass count=" & CStr(intextCol.Count)
    If intextCol.Count = 2 Then
        If CStr(intextCol(1)("id")) <> "dup-intext" Or CStr(intextCol(2)("id")) <> newId Then _
            m_failure = "unstable second pass=" & CStr(intextCol(1)("id")) & "," & CStr(intextCol(2)("id"))
        ok = ok And (CStr(intextCol(1)("id")) = "dup-intext")
        ok = ok And (CStr(intextCol(2)("id")) = newId)
    End If

    Dim noteCol As Collection
    Set noteCol = FieldCollectNoteCitationFootnotesInRange(targetRange)
    ok = ok And (noteCol.Count = 2)
    If noteCol.Count <> 2 Then m_failure = "note count=" & CStr(noteCol.Count)
    If noteCol.Count = 2 Then
        If CStr(noteCol(1)("id")) <> "dup-note" Then m_failure = "note kept id=" & CStr(noteCol(1)("id"))
        If Len(CStr(noteCol(2)("id"))) <> 36 Or CStr(noteCol(2)("id")) = "dup-note" Then _
            m_failure = "note copy id=" & CStr(noteCol(2)("id"))
        ok = ok And (CStr(noteCol(1)("id")) = "dup-note")
        ok = ok And (Len(CStr(noteCol(2)("id"))) = 36)
        ok = ok And (CStr(noteCol(2)("id")) <> "dup-note")
    End If

CleanUp:
    FieldRemoveFieldSafely first
    FieldRemoveFieldSafely copy
    If Not created Is Nothing Then FieldRemoveFootnoteSafely created("note")
    If Not noteCopy Is Nothing Then FieldRemoveFootnoteSafely noteCopy("note")
    On Error Resume Next
    doc.Range(startPos, doc.Content.End).Delete
    On Error GoTo 0
    TestFieldDuplicateCitationIds = ok
    Exit Function

ErrHandler:
    m_failure = "error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    TestFieldDuplicateCitationIds = False
End Function

' A typed code without an id is not a citation: the collectors delete the broken
' field, and a broken note citation takes its whole footnote with it.
Private Function TestFieldBrokenCitationFields() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = ActiveDocument
    Dim startPos As Long
    startPos = doc.Content.End
    Dim rng As Range
    Set rng = TestDocEndRange(doc)
    rng.Text = "Broken "

    Dim data As Object
    Set data = FieldCreatePlaceholderIntextCitationData("broken-valid")
    Dim valid As Field
    Set valid = FieldCreateIntextCitationAtRange(TestDocEndRange(doc), data)
    Dim broken As Field
    Set broken = FieldCreateRawAddinField(TestDocEndRange(doc), FieldCitationCode(""))
    If valid Is Nothing Or broken Is Nothing Then
        m_failure = "in-text field creation failed"
        GoTo CleanUp
    End If

    Dim noteData As Object
    Set noteData = FieldCreatePlaceholderNoteCitationData("broken-valid-note")
    Dim validNote As Collection
    Set validNote = FieldCreateNoteCitationAtRange(TestDocEndRange(doc), noteData)
    If validNote Is Nothing Then
        m_failure = "note creation failed"
        GoTo CleanUp
    End If

    Dim brokenNote As Footnote
    Set brokenNote = doc.Footnotes.Add(Range:=TestDocEndRange(doc))
    If brokenNote Is Nothing Then
        m_failure = "broken footnote creation failed"
        GoTo CleanUp
    End If
    Dim noteRange As Range
    Set noteRange = brokenNote.Range.Duplicate
    noteRange.Collapse wdCollapseStart
    Dim brokenNoteField As Field
    Set brokenNoteField = FieldCreateRawAddinField(noteRange, FieldCitationCode(""))
    If brokenNoteField Is Nothing Then
        m_failure = "broken note field creation failed"
        GoTo CleanUp
    End If
    Dim notesBefore As Long
    notesBefore = doc.Footnotes.Count

    Dim targetRange As Range
    Set targetRange = doc.Range(startPos, doc.Content.End)

    Dim ok As Boolean
    Dim intextCol As Collection
    Set intextCol = FieldCollectIntextCitationFieldsInRange(targetRange)
    ok = (intextCol.Count = 1)
    If intextCol.Count <> 1 Then m_failure = "in-text count=" & CStr(intextCol.Count)
    If intextCol.Count = 1 Then ok = ok And (CStr(intextCol(1)("id")) = "broken-valid")

    Dim noteCol As Collection
    Set targetRange = doc.Range(startPos, doc.Content.End)
    Set noteCol = FieldCollectNoteCitationFootnotesInRange(targetRange)
    ok = ok And (noteCol.Count = 1)
    If noteCol.Count <> 1 Then m_failure = "note count=" & CStr(noteCol.Count)
    If noteCol.Count = 1 Then ok = ok And (CStr(noteCol(1)("id")) = "broken-valid-note")

    ' The broken field is gone from the body, the broken footnote as a whole.
    Set targetRange = doc.Range(startPos, doc.Content.End)
    Dim remaining As Long
    Dim fld As Field
    For Each fld In targetRange.Fields
        If FieldHasCodeKind(fld, FIELD_KIND_CITATION) Then remaining = remaining + 1
    Next fld
    ok = ok And (remaining = 1)
    If remaining <> 1 Then m_failure = "citation fields left=" & CStr(remaining)
    ok = ok And (doc.Footnotes.Count = notesBefore - 1)
    If doc.Footnotes.Count <> notesBefore - 1 Then m_failure = "footnote count=" & CStr(doc.Footnotes.Count)

CleanUp:
    FieldRemoveFieldSafely valid
    If Not validNote Is Nothing Then FieldRemoveFootnoteSafely validNote("note")
    On Error Resume Next
    doc.Range(startPos, doc.Content.End).Delete
    On Error GoTo 0
    TestFieldBrokenCitationFields = ok
    Exit Function

ErrHandler:
    m_failure = "error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    TestFieldBrokenCitationFields = False
End Function

' The field code carries the field type and the data id.
Private Function TestFieldCodeContract() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim kind As String
    Dim id As String

    ' Code builders
    If FieldCitationCode("abc-1") <> "BANYAN_CITATION abc-1" Then Exit Function
    If FieldBibliographyCode("entry-1") <> "BANYAN_BIBLIOGRAPHY entry-1" Then Exit Function

    ' A real Word code carries the ADDIN prefix and padding spaces.
    If Not FieldParseCode(" ADDIN BANYAN_CITATION abc-1 ", kind, id) Then Exit Function
    If kind <> FIELD_KIND_CITATION Or id <> "abc-1" Then Exit Function
    If Not FieldParseCode("ADDIN BANYAN_BIBLIOGRAPHY entry-1", kind, id) Then Exit Function
    If kind <> FIELD_KIND_BIBLIOGRAPHY Or id <> "entry-1" Then Exit Function

    ' The chapter-break prompt is the code (no id).
    If Not FieldParseCode(" ADDIN  ==========Banyan chapter break (Do not edit)==========  ", kind, id) Then Exit Function
    If kind <> FIELD_KIND_CHAPTER Then Exit Function

    ' A field written before the contract has no id in its code.
    If Not FieldParseCode("ADDIN BANYAN_BIBLIOGRAPHY", kind, id) Then Exit Function
    If kind <> FIELD_KIND_BIBLIOGRAPHY Or Len(id) <> 0 Then Exit Function

    ' Foreign and malformed codes are not Banyan fields.
    If FieldParseCode("ADDIN REF _Ref123 \h", kind, id) Then Exit Function
    If FieldParseCode("", kind, id) Then Exit Function
    If FieldParseCode("BANYAN_SOMETHING else", kind, id) Then Exit Function

    ' Length short-circuit plus equality.
    If Not FieldTextEquals("", "") Then Exit Function
    If Not FieldTextEquals("same", "same") Then Exit Function
    If FieldTextEquals("same", "diff") Then Exit Function
    If FieldTextEquals("same", "same-longer") Then Exit Function

    ' A created field carries its id in the code.
    Dim doc As Document
    Set doc = ActiveDocument
    Dim startPos As Long
    startPos = doc.Content.End

    Dim data As Object
    Set data = FieldCreatePlaceholderIntextCitationData("contract-1")
    Dim fld As Field
    Set fld = FieldCreateIntextCitationAtRange(TestDocEndRange(doc), data)
    If fld Is Nothing Then Exit Function

    If Not FieldParseCode(fld.Code.Text, kind, id) Then GoTo CleanUp
    If kind <> FIELD_KIND_CITATION Or id <> "contract-1" Then GoTo CleanUp
    If Not FieldHasCodeKind(fld, FIELD_KIND_CITATION) Then GoTo CleanUp

    TestFieldCodeContract = True

CleanUp:
    FieldRemoveFieldSafely fld
    On Error Resume Next
    doc.Range(startPos, doc.Content.End).Delete
    On Error GoTo 0
    Exit Function

ErrHandler:
    TestFieldCodeContract = False
End Function

' FieldDataText / FieldReadDataText must match the stored text.
Private Function TestFieldDataTextRead() As Boolean
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then Exit Function

    Dim doc As Document
    Set doc = ActiveDocument
    Dim startPos As Long
    startPos = doc.Content.End

    Dim data As Object
    Set data = FieldCreatePlaceholderIntextCitationData("readtext-1")
    Dim fld As Field
    Set fld = FieldCreateIntextCitationAtRange(TestDocEndRange(doc), data)
    If fld Is Nothing Then Exit Function

    Dim storedText As String
    storedText = FieldDataText(fld)
    If Len(storedText) = 0 Then GoTo CleanUp

    Dim parsed As Object
    Dim parsedText As String
    Set parsed = FieldReadDataText(fld, parsedText)
    If parsed Is Nothing Then GoTo CleanUp
    If parsedText <> storedText Then GoTo CleanUp
    If CStr(parsed("id")) <> "readtext-1" Then GoTo CleanUp

    Dim viaWrapper As Object
    Set viaWrapper = FieldReadData(fld)
    If viaWrapper Is Nothing Then GoTo CleanUp
    If CStr(viaWrapper("id")) <> "readtext-1" Then GoTo CleanUp

    ' A field without data reports empty text and no parsed object.
    Dim raw As Field
    Set raw = FieldCreateRawAddinField(TestDocEndRange(doc), "BANYAN_CITATION readtext-2")
    If raw Is Nothing Then GoTo CleanUp
    If Len(FieldDataText(raw)) <> 0 Then GoTo CleanUp
    Dim rawParsed As Object
    Dim rawText As String
    Set rawParsed = FieldReadDataText(raw, rawText)
    If Not rawParsed Is Nothing Then GoTo CleanUp
    If Len(rawText) <> 0 Then GoTo CleanUp
    FieldRemoveFieldSafely raw

    TestFieldDataTextRead = True

CleanUp:
    If Not fld Is Nothing Then FieldRemoveFieldSafely fld
    On Error Resume Next
    doc.Range(startPos, doc.Content.End).Delete
    On Error GoTo 0
    Exit Function

ErrHandler:
    TestFieldDataTextRead = False
End Function

Private Function TestFieldValidators() As Boolean
    On Error GoTo ErrHandler

    Dim ok As Boolean
    ok = True

    ' --- FieldIsRichText: valid plain + rich, invalid out-of-range mark ---
    Dim goodContent As Object
    Set goodContent = New Dictionary
    goodContent("text") = "Plain"
    Set goodContent("marks") = New Collection
    ok = ok And FieldIsRichText(goodContent)

    Dim richContent As Object
    Set richContent = New Dictionary
    richContent("text") = "(A, 2020)"
    Dim rmarks As Collection
    Set rmarks = New Collection
    rmarks.Add TestMark("bold", 1, 2, True)
    rmarks.Add TestMark("link", 0, 8, "banyan://entry/abc")
    Set richContent("marks") = rmarks
    ok = ok And FieldIsRichText(richContent)

    Dim badContent As Object
    Set badContent = New Dictionary
    badContent("text") = "AB"
    Dim bmarks As Collection
    Set bmarks = New Collection
    bmarks.Add TestMark("bold", 1, 5, True)   ' end 5 > text length 2 -> invalid
    Set badContent("marks") = bmarks
    ok = ok And (Not FieldIsRichText(badContent))

    ' --- citation type guards ---
    Dim intData As Object
    Set intData = FieldCreatePlaceholderIntextCitationData("test-val-1")
    ok = ok And FieldIsIntextCitation(intData)
    ok = ok And (Not FieldIsNoteCitation(intData))

    Dim noteData As Object
    Set noteData = FieldCreatePlaceholderNoteCitationData("test-val-2")
    ok = ok And FieldIsNoteCitation(noteData)
    ok = ok And (Not FieldIsIntextCitation(noteData))

    ' --- citation source guard ---
    Dim src As Object
    Set src = TestSource("Q-1")
    ok = ok And FieldIsCitationSource(src)

    ' --- bibliography guards ---
    Dim bibTitle As Object
    Set bibTitle = New Dictionary
    bibTitle("id") = "bib-title"
    bibTitle("type") = "bibliography-title"
    Set bibTitle("content") = goodContent
    ok = ok And FieldIsBibliographyTitle(bibTitle)
    ok = ok And (Not FieldIsBibliographyEntry(bibTitle))

    Dim bibEntry As Object
    Set bibEntry = New Dictionary
    bibEntry("id") = "bib-entry-1"
    bibEntry("type") = "bibliography-entry"
    Set bibEntry("content") = goodContent
    ok = ok And FieldIsBibliographyEntry(bibEntry)
    ok = ok And (Not FieldIsBibliographyTitle(bibEntry))

    TestFieldValidators = ok
    Exit Function

ErrHandler:
    TestFieldValidators = False
End Function

Private Function TestFieldStyleIdentifier() As Boolean
    ' FieldAsStyleIdentifier reads the style via modDict and always includes
    ' id + title (no empty-id special-casing - the id is guaranteed present).
    On Error GoTo ErrHandler

    Dim style As Object
    Set style = GetPrefStyle("the-journal-of-international-studies", "国际政治研究", "note-citation")
    Dim ident As Object
    Set ident = FieldAsStyleIdentifier(style)

    Dim ok As Boolean
    ok = ident.Exists("id")
    ok = ok And (CStr(ident("id")) = "the-journal-of-international-studies")
    ok = ok And ident.Exists("title")
    ok = ok And (CStr(ident("title")) = "国际政治研究")

    TestFieldStyleIdentifier = ok
    Exit Function

ErrHandler:
    TestFieldStyleIdentifier = False
End Function


' --- Field test data builders (local, no backend) ---

Private Function TestSource(ByVal citeId As String) As Object
    Dim source As Object
    Set source = New Dictionary
    Dim cites As Collection
    Set cites = New Collection
    cites.Add citeId
    Set source("cites") = cites
    Dim params As Object
    Set params = New Dictionary
    params("page") = "12-34"
    Set source("params") = params
    Set TestSource = source
End Function

Private Function TestMark(ByVal markType As String, _
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
    Set TestMark = mark
End Function

Private Function TestComparisonRichText() As Object
    Dim content As Object
    Set content = New Dictionary
    content("text") = "Compare"

    Dim marks As Collection
    Set marks = New Collection
    marks.Add TestMark("bold", 0, 2, True)
    marks.Add TestMark("link", 0, 7, "banyan://entry/compare")
    Set content("marks") = marks
    Set TestComparisonRichText = content
End Function

Private Function TestIsWordBookmarkName(ByVal name As String) As Boolean
    ' Word rules (see FieldNormalizeBookmarkName for the authoritative sources):
    ' start with a letter, letters/digits/underscore only, max 40 characters.
    If Len(name) = 0 Or Len(name) > 40 Then Exit Function

    Dim code As Long
    code = AscW(Left$(name, 1))
    If Not ((code >= 65 And code <= 90) Or (code >= 97 And code <= 122)) Then Exit Function

    Dim i As Long
    For i = 1 To Len(name)
        code = AscW(Mid$(name, i, 1))
        If Not ((code >= 48 And code <= 57) Or (code >= 65 And code <= 90) _
            Or (code >= 97 And code <= 122) Or code = 95) Then Exit Function
    Next i
    TestIsWordBookmarkName = True
End Function

Private Function TestDocEndRange(ByVal doc As Document) As Range
    ' Collapsed range at the very end of the document. Do NOT use
    ' doc.Range(doc.Content.End, doc.Content.End) - that raises error 4608
    ' (value out of range) in some Word builds; content + collapse is safe.
    Set TestDocEndRange = doc.Content.Duplicate
    TestDocEndRange.Collapse wdCollapseEnd
End Function

Private Function TestFieldCodeId(ByVal fld As Field) As String
    Dim kind As String
    Dim id As String
    If FieldParseCode(fld.Code.Text, kind, id) Then TestFieldCodeId = id
End Function

Private Function TestResult(ByVal name As String, ByVal passed As Boolean) As String
    If passed Then
        TestResult = "[PASS] " & name
    Else
        TestResult = "[FAIL] " & name & IIf(Len(m_failure) > 0, ": " & m_failure, "")
    End If
    m_failure = ""
End Function
