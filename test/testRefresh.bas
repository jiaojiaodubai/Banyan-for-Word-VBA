Option Explicit

' Regression tests for bibliography id-diff refresh. No backend calls are made.

Private m_failure As String

Public Function RunTests() As String
    Dim report As String
    report = "testRefresh" & vbCrLf & String(40, "-") & vbCrLf
    m_failure = ""
    report = report & TestResult("bibliography id diff add/delete/reorder/update", TestBibliographyIdDiff()) & vbCrLf
    m_failure = ""
    report = report & TestResult("bibliography tail append preserves prefix", TestBibliographyTailAppend()) & vbCrLf
    m_failure = ""
    report = report & TestResult("bibliography local insert preserves existing sequence", TestBibliographyLocalInsert()) & vbCrLf
    m_failure = ""
    report = report & TestResult("bibliography LCS preserves multiple segments", TestBibliographyMultipleSegments()) & vbCrLf
    m_failure = ""
    report = report & TestResult("bibliography rejects missing ids", TestBibliographyRejectsMissingIds()) & vbCrLf
    m_failure = ""
    report = report & TestResult("bibliography leaves empty-data field alone", TestBibliographyEmptyData()) & vbCrLf
    RunTests = report
End Function

Private Function TestBibliographyMultipleSegments() As Boolean
    On Error GoTo ErrHandler
    Dim startPos As Long
    startPos = BeginTestArea()
    Dim pref As Object
    Set pref = TestBibliographyPreference()

    Dim initial As Collection
    Set initial = New Collection
    initial.Add TestBibliographyLine("title", "bibliography-title", "References")
    initial.Add TestBibliographyLine("A", "bibliography-entry", "Entry A")
    initial.Add TestBibliographyLine("B", "bibliography-entry", "Entry B")
    initial.Add TestBibliographyLine("C", "bibliography-entry", "Entry C")
    initial.Add TestBibliographyLine("D", "bibliography-entry", "Entry D")
    InsertTestBibliography initial, pref
    Dim fields As Collection
    Set fields = TestBibliographyFields(startPos)
    fields(2).Result.Text = "LOCAL A"
    fields(5).Result.Text = "LOCAL D"

    Dim updated As Collection
    Set updated = New Collection
    updated.Add TestBibliographyLine("title", "bibliography-title", "References")
    updated.Add TestBibliographyLine("A", "bibliography-entry", "Entry A")
    updated.Add TestBibliographyLine("X", "bibliography-entry", "Entry X")
    updated.Add TestBibliographyLine("B", "bibliography-entry", "Entry B")
    updated.Add TestBibliographyLine("C", "bibliography-entry", "Entry C")
    updated.Add TestBibliographyLine("Y", "bibliography-entry", "Entry Y")
    updated.Add TestBibliographyLine("D", "bibliography-entry", "Entry D")

    Dim target As Range
    Set target = ActiveDocument.Range(startPos, ActiveDocument.Content.End)
    Dim ok As Boolean
    ok = RefreshBibliographyLinesInRange(target, updated, pref)
    Set fields = TestBibliographyFields(startPos)
    ok = ok And (fields.Count = 7)
    If fields.Count = 7 Then
        Dim orderText As String
        orderText = TestFieldId(fields(1)) & "," & TestFieldId(fields(2)) & "," & _
                    TestFieldId(fields(3)) & "," & TestFieldId(fields(4)) & "," & _
                    TestFieldId(fields(5)) & "," & TestFieldId(fields(6)) & "," & TestFieldId(fields(7))
        ok = ok And (orderText = "title,A,X,B,C,Y,D")
        ok = ok And (fields(2).Result.Text = "LOCAL A")
        ok = ok And (fields(7).Result.Text = "LOCAL D")
        If orderText <> "title,A,X,B,C,Y,D" Then m_failure = "multi-segment order=" & orderText
        If fields(2).Result.Text <> "LOCAL A" Or fields(7).Result.Text <> "LOCAL D" Then _
            m_failure = "a retained segment was rebuilt"
    End If

    EndTestArea startPos
    TestBibliographyMultipleSegments = ok
    Exit Function

ErrHandler:
    m_failure = "multi-segment error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    EndTestArea startPos
    TestBibliographyMultipleSegments = False
End Function

Private Function TestBibliographyIdDiff() As Boolean
    On Error GoTo ErrHandler
    Dim startPos As Long
    startPos = BeginTestArea()

    Dim pref As Object
    Set pref = TestBibliographyPreference()
    Dim initial As Collection
    Set initial = New Collection
    initial.Add TestBibliographyLine("title", "bibliography-title", "References")
    initial.Add TestBibliographyLine("A", "bibliography-entry", "Entry A")
    initial.Add TestBibliographyLine("B", "bibliography-entry", "Entry B")
    initial.Add TestBibliographyLine("C", "bibliography-entry", "Entry C")
    InsertTestBibliography initial, pref

    Dim fields As Collection
    Set fields = TestBibliographyFields(startPos)
    fields(1).Result.Text = "LOCAL TITLE"

    Dim updated As Collection
    Set updated = New Collection
    updated.Add TestBibliographyLine("title", "bibliography-title", "References")
    updated.Add TestBibliographyLine("C", "bibliography-entry", "Entry C updated")
    Dim lineA As Object
    Set lineA = TestBibliographyLine("A", "bibliography-entry", "Entry A")
    lineA("revision") = 2
    updated.Add lineA
    updated.Add TestBibliographyLine("D", "bibliography-entry", "Entry D")

    Dim target As Range
    Set target = ActiveDocument.Range(startPos, ActiveDocument.Content.End)
    Dim ok As Boolean
    ok = RefreshBibliographyLinesInRange(target, updated, pref)
    If Not ok Then
        m_failure = "apply returned false"
        GoTo Finish
    End If

    Set fields = TestBibliographyFields(startPos)
    If fields.Count <> 4 Then m_failure = "field count=" & CStr(fields.Count)
    ok = ok And (fields.Count = 4)
    If fields.Count = 4 Then
        Dim orderText As String
        orderText = TestFieldId(fields(1)) & "," & TestFieldId(fields(2)) & "," & TestFieldId(fields(3)) & "," & TestFieldId(fields(4))
        If orderText <> "title,C,A,D" Then m_failure = "order=" & orderText
        ok = ok And (orderText = "title,C,A,D")
        If fields(1).Result.Text <> "LOCAL TITLE" Then m_failure = "retained LCS field was rendered"
        ok = ok And (fields(1).Result.Text = "LOCAL TITLE")
        If fields(2).Result.Text <> "Entry C updated" Then m_failure = "C result=" & fields(2).Result.Text
        ok = ok And (fields(2).Result.Text = "Entry C updated")
        If fields(3).Result.Text <> "Entry A" Then m_failure = "A result=" & fields(3).Result.Text
        ok = ok And (fields(3).Result.Text = "Entry A")
        If fields(4).Result.Text <> "Entry D" Then m_failure = "D result=" & fields(4).Result.Text
        ok = ok And (fields(4).Result.Text = "Entry D")
        Dim storedA As Object
        Set storedA = FieldReadData(fields(3))
        If storedA Is Nothing Then
            m_failure = "A data missing"
            ok = False
        Else
            If Not storedA.Exists("revision") Then m_failure = "A revision missing"
            ok = ok And storedA.Exists("revision")
            If storedA.Exists("revision") Then ok = ok And (CLng(storedA("revision")) = 2)
        End If
    End If
    ok = ok And (Not ActiveDocument.Bookmarks.Exists(FieldGetBibliographyBookmarkName("B")))
    ok = ok And ActiveDocument.Bookmarks.Exists(FieldGetBibliographyBookmarkName("A"))
    ok = ok And ActiveDocument.Bookmarks.Exists(FieldGetBibliographyBookmarkName("C"))
    ok = ok And ActiveDocument.Bookmarks.Exists(FieldGetBibliographyBookmarkName("D"))

    ' Data is now identical. The locally edited title proves a no-op refresh
    ' does not render a retained LCS line.
    ok = ok And (Not RefreshBibliographyLinesInRange(target, updated, pref))
    Set fields = TestBibliographyFields(startPos)
    If fields.Count = 4 Then ok = ok And (fields(1).Result.Text = "LOCAL TITLE")

Finish:
    EndTestArea startPos
    TestBibliographyIdDiff = ok
    Exit Function

ErrHandler:
    m_failure = "error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    EndTestArea startPos
    TestBibliographyIdDiff = False
End Function

Private Function TestBibliographyTailAppend() As Boolean
    On Error GoTo ErrHandler
    Dim startPos As Long
    startPos = BeginTestArea()
    Dim pref As Object
    Set pref = TestBibliographyPreference()

    Dim initial As Collection
    Set initial = New Collection
    initial.Add TestBibliographyLine("title", "bibliography-title", "References")
    initial.Add TestBibliographyLine("A", "bibliography-entry", "Entry A")
    initial.Add TestBibliographyLine("B", "bibliography-entry", "Entry B")
    InsertTestBibliography initial, pref
    Dim fields As Collection
    Set fields = TestBibliographyFields(startPos)
    fields(2).Result.Text = "LOCAL A"

    Dim updated As Collection
    Set updated = New Collection
    updated.Add TestBibliographyLine("title", "bibliography-title", "References")
    updated.Add TestBibliographyLine("A", "bibliography-entry", "Entry A")
    updated.Add TestBibliographyLine("B", "bibliography-entry", "Entry B updated")
    updated.Add TestBibliographyLine("C", "bibliography-entry", "Entry C")

    Dim target As Range
    Set target = ActiveDocument.Range(startPos, ActiveDocument.Content.End)
    Dim ok As Boolean
    ok = RefreshBibliographyLinesInRange(target, updated, pref)
    Set fields = TestBibliographyFields(startPos)
    ok = ok And (fields.Count = 4)
    If fields.Count = 4 Then
        ok = ok And (TestFieldId(fields(1)) = "title")
        ok = ok And (TestFieldId(fields(2)) = "A")
        ok = ok And (TestFieldId(fields(3)) = "B")
        ok = ok And (TestFieldId(fields(4)) = "C")
        ok = ok And (fields(2).Result.Text = "LOCAL A")
        ok = ok And (fields(3).Result.Text = "Entry B updated")
        If fields(2).Result.Text <> "LOCAL A" Then m_failure = "unchanged prefix was rebuilt"
    End If

    EndTestArea startPos
    TestBibliographyTailAppend = ok
    Exit Function

ErrHandler:
    m_failure = "append error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    EndTestArea startPos
    TestBibliographyTailAppend = False
End Function

Private Function TestBibliographyLocalInsert() As Boolean
    On Error GoTo ErrHandler
    Dim startPos As Long
    startPos = BeginTestArea()
    Dim pref As Object
    Set pref = TestBibliographyPreference()

    Dim initial As Collection
    Set initial = New Collection
    initial.Add TestBibliographyLine("title", "bibliography-title", "References")
    initial.Add TestBibliographyLine("A", "bibliography-entry", "Entry A")
    initial.Add TestBibliographyLine("B", "bibliography-entry", "Entry B")
    initial.Add TestBibliographyLine("C", "bibliography-entry", "Entry C")
    InsertTestBibliography initial, pref
    Dim fields As Collection
    Set fields = TestBibliographyFields(startPos)
    fields(2).Result.Text = "LOCAL A"

    Dim updated As Collection
    Set updated = New Collection
    updated.Add TestBibliographyLine("title", "bibliography-title", "References")
    updated.Add TestBibliographyLine("X", "bibliography-entry", "Entry X")
    updated.Add TestBibliographyLine("A", "bibliography-entry", "Entry A")
    updated.Add TestBibliographyLine("B", "bibliography-entry", "Entry B")
    updated.Add TestBibliographyLine("C", "bibliography-entry", "Entry C")

    Dim target As Range
    Set target = ActiveDocument.Range(startPos, ActiveDocument.Content.End)
    Dim ok As Boolean
    ok = RefreshBibliographyLinesInRange(target, updated, pref)
    Set fields = TestBibliographyFields(startPos)
    ok = ok And (fields.Count = 5)
    If fields.Count = 5 Then
        Dim orderText As String
        orderText = TestFieldId(fields(1)) & "," & TestFieldId(fields(2)) & "," & _
                    TestFieldId(fields(3)) & "," & TestFieldId(fields(4)) & "," & TestFieldId(fields(5))
        ok = ok And (orderText = "title,X,A,B,C")
        ok = ok And (fields(3).Result.Text = "LOCAL A")
        If orderText <> "title,X,A,B,C" Then m_failure = "local insert order=" & orderText
        If fields(3).Result.Text <> "LOCAL A" Then m_failure = "retained sequence was rebuilt"
    End If

    EndTestArea startPos
    TestBibliographyLocalInsert = ok
    Exit Function

ErrHandler:
    m_failure = "local insert error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    EndTestArea startPos
    TestBibliographyLocalInsert = False
End Function

Private Function TestBibliographyRejectsMissingIds() As Boolean
    On Error GoTo ErrHandler
    Dim startPos As Long
    startPos = BeginTestArea()

    Dim pref As Object
    Set pref = TestBibliographyPreference()
    Dim missingIdLine As Object
    Set missingIdLine = TestBibliographyLine("missing", "bibliography-title", "Old title")
    missingIdLine.Remove "id"
    Dim initial As Collection
    Set initial = New Collection
    initial.Add missingIdLine
    InsertTestBibliography initial, pref

    Dim updated As Collection
    Set updated = New Collection
    updated.Add TestBibliographyLine("title-new", "bibliography-title", "New title")

    Dim target As Range
    Set target = ActiveDocument.Range(startPos, ActiveDocument.Content.End)
    Dim ok As Boolean
    ok = Not RefreshBibliographyLinesInRange(target, updated, pref)
    If Not ok Then m_failure = "missing id was accepted"
    Dim fields As Collection
    Set fields = TestBibliographyFields(startPos)
    If fields.Count <> 1 Then m_failure = "field count changed=" & CStr(fields.Count)
    ok = ok And (fields.Count = 1)
    If fields.Count = 1 Then
        ok = ok And (fields(1).Result.Text = "Old title")
        If fields(1).Result.Text <> "Old title" Then m_failure = "missing-id field was modified"
    End If

Finish:
    EndTestArea startPos
    TestBibliographyRejectsMissingIds = ok
    Exit Function

ErrHandler:
    m_failure = "missing-id error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    EndTestArea startPos
    TestBibliographyRejectsMissingIds = False
End Function

Private Sub InsertTestBibliography(ByVal lines As Collection, ByVal pref As Object)
    Dim cursor As Range
    Set cursor = ActiveDocument.Content.Duplicate
    cursor.Collapse wdCollapseEnd

    Dim i As Long
    For i = 1 To lines.Count
        Dim data As Object
        Dim fld As Field
        Set data = lines(i)
        Set fld = FieldCreateRawAddinField(cursor, "BANYAN_BIBLIOGRAPHY " & DictKeyString(data, "id"))
        FieldWriteData fld, data
        If FieldIsBibliographyTitle(data) Then
            FieldRenderStyledFieldWithStyle fld, DictKeyString(pref, "bibliographyTitleStyle"), wdStyleTypeParagraph, data("content")
        Else
            FieldRenderStyledFieldWithStyle fld, DictKeyString(pref, "bibliographyEntryStyle"), wdStyleTypeParagraph, data("content")
            FieldAddBookmarkToField fld, FieldGetBibliographyBookmarkName(DictKeyString(data, "id"))
        End If
        Set cursor = fld.Result.Duplicate
        cursor.Collapse wdCollapseEnd
        If i < lines.Count Then
            cursor.InsertParagraphAfter
            cursor.Collapse wdCollapseEnd
        End If
    Next i
End Sub

Private Function TestFieldId(ByVal fld As Field) As String
    Dim data As Object
    Set data = FieldReadData(fld)
    If Not data Is Nothing Then TestFieldId = DictKeyString(data, "id")
End Function

Private Function TestBibliographyFields(ByVal startPos As Long) As Collection
    Dim result As Collection
    Set result = New Collection
    Dim fld As Field
    For Each fld In ActiveDocument.Range(startPos, ActiveDocument.Content.End).Fields
        If FieldHasCodeKind(fld, FIELD_KIND_BIBLIOGRAPHY) Then result.Add fld
    Next fld
    Set TestBibliographyFields = result
End Function

' A field without stored data must be left untouched by the refresh.
Private Function TestBibliographyEmptyData() As Boolean
    On Error GoTo ErrHandler
    Dim startPos As Long
    startPos = BeginTestArea()

    Dim pref As Object
    Set pref = TestBibliographyPreference()

    Dim cursor As Range
    Set cursor = ActiveDocument.Content.Duplicate
    cursor.Collapse wdCollapseEnd
    Dim fld As Field
    Set fld = FieldCreateRawAddinField(cursor, "BANYAN_BIBLIOGRAPHY empty-data")
    If fld Is Nothing Then GoTo Finish

    Dim updated As Collection
    Set updated = New Collection
    updated.Add TestBibliographyLine("empty-data", "bibliography-entry", "Entry")

    Dim target As Range
    Set target = ActiveDocument.Range(startPos, ActiveDocument.Content.End)
    Dim ok As Boolean
    ok = Not RefreshBibliographyLinesInRange(target, updated, pref)
    ok = ok And (Len(FieldDataText(fld)) = 0)

Finish:
    EndTestArea startPos
    TestBibliographyEmptyData = ok
    Exit Function

ErrHandler:
    m_failure = "empty-data error " & CStr(Err.Number) & " from " & Err.Source & ": " & Err.Description
    EndTestArea startPos
    TestBibliographyEmptyData = False
End Function

Private Function TestBibliographyLine(ByVal id As String, _
                                      ByVal lineType As String, _
                                      ByVal text As String) As Object
    Dim data As Object
    Set data = New Dictionary
    data("id") = id
    data("type") = lineType
    Set data("content") = FieldCreateRichText(text)
    Set TestBibliographyLine = data
End Function

Private Function TestBibliographyPreference() As Object
    Const TITLE_STYLE As String = "Banyan Test Bibliography Title"
    Const ENTRY_STYLE As String = "Banyan Test Bibliography Entry"
    EnsureTestStyle TITLE_STYLE
    EnsureTestStyle ENTRY_STYLE
    Dim pref As Object
    Set pref = New Dictionary
    pref("bibliographyTitleStyle") = TITLE_STYLE
    pref("bibliographyEntryStyle") = ENTRY_STYLE
    Set TestBibliographyPreference = pref
End Function

Private Sub EnsureTestStyle(ByVal styleName As String)
    On Error Resume Next
    Dim style As Style
    Set style = ActiveDocument.Styles(styleName)
    On Error GoTo 0
    If style Is Nothing Then ActiveDocument.Styles.Add Name:=styleName, Type:=wdStyleTypeParagraph
End Sub

Private Function BeginTestArea() As Long
    Dim cursor As Range
    Set cursor = ActiveDocument.Content.Duplicate
    cursor.Collapse wdCollapseEnd
    cursor.InsertAfter vbCr
    cursor.Collapse wdCollapseEnd
    BeginTestArea = cursor.Start
End Function

Private Sub EndTestArea(ByVal startPos As Long)
    On Error Resume Next
    If startPos > 0 Then ActiveDocument.Range(startPos, ActiveDocument.Content.End).Delete
    On Error GoTo 0
End Sub

Private Function TestResult(ByVal name As String, ByVal passed As Boolean) As String
    If passed Then
        TestResult = "[PASS] " & name
    Else
        TestResult = "[FAIL] " & name & IIf(Len(m_failure) > 0, ": " & m_failure, "")
    End If
End Function
