Option Explicit

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
    If ActiveDocument Is Nothing Then
        RunTests = report & "[FAIL] no active document"
        Exit Function
    End If
    report = report & TestResult("create intext + data round-trip", TestFieldCreateIntext()) & vbCrLf
    report = report & TestResult("render rich text (bold/color/link)", TestFieldRenderRichText()) & vbCrLf
    report = report & TestResult("render from Field.Data (no content)", TestFieldRenderFromData()) & vbCrLf
    report = report & TestResult("note citation create + reference", TestFieldNoteCitation()) & vbCrLf
    report = report & TestResult("note refresh same reference (no rebuild)", TestFieldRebuildNoteCitation()) & vbCrLf
    report = report & TestResult("note rebuild on reference change (rich copy)", TestFieldRebuildNoteCitationRefChange()) & vbCrLf
    report = report & TestResult("note renumber via custom refs (rebuild)", TestFieldNoteRenumberRebuild()) & vbCrLf
    report = report & TestResult("migrate intext -> note", TestFieldMigrateIntextToNotes()) & vbCrLf
    report = report & TestResult("migrate note -> intext", TestFieldMigrateNotesToIntext()) & vbCrLf
    report = report & TestResult("collectors in range", TestFieldCollectors()) & vbCrLf
    report = report & TestResult("validators", TestFieldValidators()) & vbCrLf
    report = report & TestResult("style identifier", TestFieldStyleIdentifier()) & vbCrLf
    RunTests = report
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
    ok = FieldRenderStyledField(fld, content)
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

Private Function TestDocEndRange(ByVal doc As Document) As Range
    ' Collapsed range at the very end of the document. Do NOT use
    ' doc.Range(doc.Content.End, doc.Content.End) - that raises error 4608
    ' (value out of range) in some Word builds; content + collapse is safe.
    Set TestDocEndRange = doc.Content.Duplicate
    TestDocEndRange.Collapse wdCollapseEnd
End Function

Private Function TestResult(ByVal name As String, ByVal passed As Boolean) As String
    If passed Then
        TestResult = "[PASS] " & name
    Else
        TestResult = "[FAIL] " & name
    End If
End Function
