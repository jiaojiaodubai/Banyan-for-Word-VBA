Option Explicit

' ============================================================================
' Module  : testPerf
' Purpose : Diagnostic timing probe for the local Word/VBA portion of citation
'           insertion and refresh. No backend calls are made.
' ============================================================================

#If Mac Then
#ElseIf VBA7 Then
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#Else
Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef value As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef value As Currency) As Long
#End If

Private m_report As String

Private Const INTEXT_TEXT As String = "(Zhang, 2020, p. 37)"
Private Const NOTE_BODY As String = "1 Zhou, J.: Self-Efficacy, Shanghai: ECUST Press, 2003, pp. 37-38."


Public Function RunPerf(Optional ByVal sizesCsv As String = "10,50,100", _
                        Optional ByVal repetitions As Long = 4) As String
    On Error GoTo ErrHandler
    If ActiveDocument Is Nothing Then
        RunPerf = "[FAIL] no active document"
        Exit Function
    End If
    If repetitions < 1 Then repetitions = 1

    Dim sizes As Collection
    Set sizes = ParseSizes(sizesCsv)
    If sizes.Count = 0 Then Err.Raise 5, "testPerf.RunPerf", "No valid positive scale sizes were supplied."

    m_report = "testPerf" & vbCrLf & String(72, "-") & vbCrLf
    Out "[INFO] clock=" & PerfClockName() & "; repetitions=" & CStr(repetitions) & _
        "; sizes=" & sizesCsv
    Out "[INFO] Word=" & Application.Version & " build=" & Application.Build & _
        "; visible=" & BoolText(Application.Visible) & _
        "; screenUpdating=" & BoolText(Application.ScreenUpdating) & _
        "; pagination=" & BoolText(Options.Pagination) & _
        "; styles=" & CStr(ActiveDocument.Styles.Count)
    Out "[INFO] scope=local Word object model and JSON only; HTTP/backend/dialog time excluded"
    Out ""

    ResetPerfDocument
    WarmUp
    InsertLikeOps repetitions
    RenderVariants repetitions
    StyleLookupOps repetitions
    ScaleOps sizes
    PaginationOps sizes
    ProofingOps sizes
    BibliographyOps sizes
    JsonAndLookupOps sizes
    ComparisonOps sizes
    NoteOps repetitions
    ResetPerfDocument

    RunPerf = m_report
    Exit Function

ErrHandler:
    Dim failureDescription As String
    failureDescription = Err.Description
    On Error Resume Next
    ResetPerfDocument
    On Error GoTo 0
    RunPerf = m_report & "[FAIL] testPerf.RunPerf: " & failureDescription
End Function

Public Function RunComparisonPerf(Optional ByVal sizesCsv As String = "100,1000", _
                                  Optional ByVal repetitions As Long = 4) As String
    On Error GoTo ErrHandler
    If repetitions < 1 Then repetitions = 1

    Dim sizes As Collection
    Set sizes = ParseSizes(sizesCsv)
    If sizes.Count = 0 Then Err.Raise 5, "testPerf.RunComparisonPerf", "No valid positive scale sizes were supplied."

    m_report = "testPerf comparison-only" & vbCrLf & String(72, "-") & vbCrLf
    Out "[INFO] clock=" & PerfClockName() & "; repetitions=" & CStr(repetitions) & _
        "; sizes=" & sizesCsv

    Dim repetition As Long
    For repetition = 1 To repetitions
        Out "[INFO] repetition=" & CStr(repetition)
        ComparisonOps sizes
    Next repetition

    RunComparisonPerf = m_report
    Exit Function

ErrHandler:
    RunComparisonPerf = m_report & "[FAIL] testPerf.RunComparisonPerf: " & Err.Description
End Function


' --- Production-shaped insert path -----------------------------------------

Private Sub InsertLikeOps(ByVal reps As Long)
    Out "[SECTION] insert-like local path (averages over reps)"

    Dim tPlaceholder As Double
    Dim tRepaint As Double
    Dim tWrite As Double
    Dim tRenderImplicit As Double
    Dim tRead As Double
    Dim r As Long

    For r = 1 To reps
        ResetPerfDocument

        Dim placeholder As Object
        Set placeholder = FieldCreatePlaceholderIntextCitationData("insert-" & CStr(r))

        Dim target As Range
        Set target = DocEnd()
        Dim fld As Field
        Dim t0 As Double

        t0 = PerfNow()
        Set fld = FieldCreateIntextCitationAtRange(target, placeholder)
        tPlaceholder = tPlaceholder + PerfElapsed(t0)
        If fld Is Nothing Then Err.Raise 5, "testPerf.InsertLikeOps", "Could not create placeholder field."

        t0 = PerfNow()
        Application.ScreenRefresh
        DoEvents
        tRepaint = tRepaint + PerfElapsed(t0)

        Dim data As Object
        Set data = BuildIntextData("insert-" & CStr(r), "rich")

        t0 = PerfNow()
        FieldWriteData fld, data
        tWrite = tWrite + PerfElapsed(t0)

        ' Citation insertion omits the content argument. ResolveFieldContent
        ' parses Field.Data, followed by another parse for the citation style.
        t0 = PerfNow()
        FieldRenderStyledField fld
        tRenderImplicit = tRenderImplicit + PerfElapsed(t0)

        Dim readBack As Object
        t0 = PerfNow()
        Set readBack = FieldReadData(fld)
        tRead = tRead + PerfElapsed(t0)
    Next r

    OutTime "placeholder create+write+render", tPlaceholder / reps
    OutTime "ScreenRefresh+DoEvents", tRepaint / reps
    OutTime "selected data write", tWrite / reps
    OutTime "selected rich render (implicit)", tRenderImplicit / reps
    OutTime "single Field.Data read", tRead / reps
    Out ""
End Sub


' --- Pagination impact ----------------------------------------------------

Private Sub PaginationOps(ByVal sizes As Collection)
    Out "[SECTION] pagination impact on rich rendering"
    Out "[INFO] ScreenUpdating is forced off; only Options.Pagination is varied"

    Dim item As Variant
    For Each item In sizes
        RunPaginationScale CLng(item), True
        RunPaginationScale CLng(item), False
    Next item
    Out ""
End Sub

Private Sub RunPaginationScale(ByVal fieldCount As Long, ByVal paginationEnabled As Boolean)
    ResetPerfDocument

    Dim originalScreenUpdating As Boolean
    Dim originalPagination As Boolean
    originalScreenUpdating = Application.ScreenUpdating
    originalPagination = Options.Pagination
    Application.ScreenUpdating = False
    Options.Pagination = paginationEnabled

    Dim fields As Collection
    Set fields = New Collection
    Dim i As Long
    For i = 1 To fieldCount
        Dim data As Object
        Set data = BuildIntextData("pagination-" & CStr(i), "rich")
        Dim fld As Field
        Set fld = CreatePreparedField(data)
        fields.Add fld
        If i < fieldCount Then DocEnd().InsertAfter " "
    Next i

    Dim t0 As Double
    t0 = PerfNow()
    Dim item As Variant
    Dim renderData As Object
    Set renderData = BuildIntextData("pagination-render", "rich")
    For Each item In fields
        FieldRenderStyledField item, renderData("content")
    Next item
    Dim elapsed As Double
    elapsed = PerfElapsed(t0)

    OutScale "render rich", fieldCount, IIf(paginationEnabled, "paginationOn", "paginationOff"), elapsed

    Options.Pagination = originalPagination
    Application.ScreenUpdating = originalScreenUpdating
    ResetPerfDocument
End Sub


' --- Background proofing impact -------------------------------------------

Private Sub ProofingOps(ByVal sizes As Collection)
    Out "[SECTION] background proofing impact on rich rendering"
    Out "[INFO] ScreenUpdating is forced off; spelling and grammar are restored after each run"

    Dim item As Variant
    For Each item In sizes
        RunProofingScale CLng(item), True
        RunProofingScale CLng(item), False
    Next item
    Out ""
End Sub

Private Sub RunProofingScale(ByVal fieldCount As Long, ByVal proofingEnabled As Boolean)
    On Error GoTo CleanUp

    ResetPerfDocument

    Dim originalScreenUpdating As Boolean
    Dim originalSpelling As Boolean
    Dim originalGrammar As Boolean
    originalScreenUpdating = Application.ScreenUpdating
    originalSpelling = Options.CheckSpellingAsYouType
    originalGrammar = Options.CheckGrammarAsYouType
    Dim settingsCaptured As Boolean
    settingsCaptured = True
    Application.ScreenUpdating = False
    Options.CheckSpellingAsYouType = proofingEnabled
    Options.CheckGrammarAsYouType = proofingEnabled

    Dim fields As Collection
    Set fields = New Collection
    Dim i As Long
    For i = 1 To fieldCount
        Dim data As Object
        Set data = BuildIntextData("proofing-" & CStr(i), "rich")
        Dim fld As Field
        Set fld = CreatePreparedField(data)
        fields.Add fld
        If i < fieldCount Then DocEnd().InsertAfter " "
    Next i

    Dim renderData As Object
    Set renderData = BuildIntextData("proofing-render", "rich")
    Dim t0 As Double
    t0 = PerfNow()
    Dim item As Variant
    For Each item In fields
        FieldRenderStyledField item, renderData("content")
    Next item
    OutScale "render rich", fieldCount, IIf(proofingEnabled, "proofingOn", "proofingOff"), PerfElapsed(t0)

CleanUp:
    On Error Resume Next
    If settingsCaptured Then
        Options.CheckSpellingAsYouType = originalSpelling
        Options.CheckGrammarAsYouType = originalGrammar
        Application.ScreenUpdating = originalScreenUpdating
    End If
    ResetPerfDocument
    On Error GoTo 0
End Sub


' --- Data/content comparison cost ----------------------------------------

Private Sub ComparisonOps(ByVal sizes As Collection)
    Out "[SECTION] targeted rich-text comparison cost"
    Out "[INFO] Measures the content equality check used to skip unchanged citation renders"

    Dim item As Variant
    For Each item In sizes
        RunComparisonScale CLng(item)
    Next item
    Out ""
End Sub

Private Sub RunComparisonScale(ByVal itemCount As Long)
    Dim currentItems As Collection
    Dim sameItems As Collection
    Dim changedItems As Collection
    Set currentItems = New Collection
    Set sameItems = New Collection
    Set changedItems = New Collection

    Dim i As Long
    For i = 1 To itemCount
        Dim currentData As Object
        Set currentData = BuildIntextData("compare-" & CStr(i), "rich")
        currentItems.Add currentData
        sameItems.Add BuildIntextData("compare-" & CStr(i), "rich")
        changedItems.Add BuildIntextData("compare-" & CStr(i), "plain")
    Next i

    Dim t0 As Double
    Dim leftDataObject As Object
    Dim rightDataObject As Object
    Dim leftContentObject As Object
    Dim rightContentObject As Object

    Dim sameContent As Long
    t0 = PerfNow()
    For i = 1 To itemCount
        Set leftDataObject = currentItems(i)
        Set rightDataObject = sameItems(i)
        Set leftContentObject = DictKeyObject(leftDataObject, "content")
        Set rightContentObject = DictKeyObject(rightDataObject, "content")
        If FieldRichTextEquals(leftContentObject, rightContentObject) Then sameContent = sameContent + 1
    Next i
    OutScale "cached content equals (same)", itemCount, "refreshShape", PerfElapsed(t0)

    Dim changedContent As Long
    t0 = PerfNow()
    For i = 1 To itemCount
        Set leftDataObject = currentItems(i)
        Set rightDataObject = changedItems(i)
        Set leftContentObject = DictKeyObject(leftDataObject, "content")
        Set rightContentObject = DictKeyObject(rightDataObject, "content")
        If FieldRichTextEquals(leftContentObject, rightContentObject) Then changedContent = changedContent + 1
    Next i
    OutScale "cached content equals (changed)", itemCount, "refreshShape", PerfElapsed(t0)

End Sub


' --- Render cost by feature ------------------------------------------------

Private Sub RenderVariants(ByVal reps As Long)
    Out "[SECTION] render feature isolation (explicit content; averages over reps)"
    RenderVariant "plain", reps
    RenderVariant "format", reps
    RenderVariant "link", reps
    RenderVariant "rich", reps

    Dim explicitTotal As Double
    Dim implicitTotal As Double
    Dim r As Long
    For r = 1 To reps
        Dim data As Object
        Set data = BuildIntextData("arg-" & CStr(r), "rich")

        Dim fld As Field
        Set fld = CreatePreparedField(data)
        Dim t0 As Double
        t0 = PerfNow()
        FieldRenderStyledField fld, data("content")
        explicitTotal = explicitTotal + PerfElapsed(t0)
        FieldRemoveFieldSafely fld

        Set fld = CreatePreparedField(data)
        t0 = PerfNow()
        FieldRenderStyledField fld
        implicitTotal = implicitTotal + PerfElapsed(t0)
        FieldRemoveFieldSafely fld
    Next r

    OutTime "rich render explicit content", explicitTotal / reps
    OutTime "rich render implicit content", implicitTotal / reps
    OutTime "implicit argument overhead", (implicitTotal - explicitTotal) / reps
    Out ""
End Sub

Private Sub RenderVariant(ByVal kind As String, ByVal reps As Long)
    Dim total As Double
    Dim r As Long
    For r = 1 To reps
        Dim data As Object
        Set data = BuildIntextData("render-" & kind & "-" & CStr(r), kind)

        Dim fld As Field
        Set fld = CreatePreparedField(data)

        Dim t0 As Double
        t0 = PerfNow()
        FieldRenderStyledField fld, data("content")
        total = total + PerfElapsed(t0)
        FieldRemoveFieldSafely fld
    Next r

    OutTime "render " & kind, total / reps
End Sub


' --- Style lookup ----------------------------------------------------------

Private Sub StyleLookupOps(ByVal reps As Long)
    Out "[SECTION] citation style lookup isolation"
    Out "[INFO] current path scans every document style for every rendered in-text citation"

    Dim data As Object
    Set data = BuildIntextData("style-probe", "plain")
    Dim fld As Field
    Set fld = CreatePreparedField(data)
    FieldApplyIntextCitationStyle fld

    Dim cachedStyle As Style
    Set cachedStyle = fld.Result.Style

    Dim calls As Long
    calls = reps * 5
    Dim i As Long
    Dim t0 As Double
    Dim currentTotal As Double
    t0 = PerfNow()
    For i = 1 To calls
        FieldClearPerformanceCache
        FieldApplyIntextCitationStyle fld
    Next i
    currentTotal = PerfElapsed(t0)

    Dim cachedTotal As Double
    t0 = PerfNow()
    For i = 1 To calls
        cachedStyle.UnhideWhenUsed = True
        cachedStyle.QuickStyle = True
        fld.Result.Style = cachedStyle
    Next i
    cachedTotal = PerfElapsed(t0)

    OutTime "current scan per call", currentTotal / calls
    OutTime "cached style per call", cachedTotal / calls
    OutTime "avoidable style lookup per call", (currentTotal - cachedTotal) / calls
    FieldRemoveFieldSafely fld
    Out ""
End Sub


' --- Refresh-like batch scaling -------------------------------------------

Private Sub ScaleOps(ByVal sizes As Collection)
    Out "[SECTION] refresh-like batch scaling"
    Out "[INFO] setup and cleanup are excluded; update render uses explicit content like modRefresh"

    Dim originalScreenUpdating As Boolean
    originalScreenUpdating = Application.ScreenUpdating

    Dim item As Variant
    For Each item In sizes
        RunScale CLng(item), False
        RunScale CLng(item), True
    Next item

    Application.ScreenUpdating = originalScreenUpdating
    Out ""
End Sub

Private Sub RunScale(ByVal fieldCount As Long, ByVal suppressScreen As Boolean)
    ResetPerfDocument

    Dim originalScreenUpdating As Boolean
    originalScreenUpdating = Application.ScreenUpdating
    If suppressScreen Then Application.ScreenUpdating = False

    Dim i As Long
    Dim fld As Field
    Dim data As Object
    For i = 1 To fieldCount
        Set data = BuildIntextData("scale-" & CStr(i), "plain")
        Set fld = FieldCreateRawAddinField(DocEnd(), "BANYAN_CITATION scale-" & CStr(i))
        FieldWriteData fld, data
        fld.Result.Text = INTEXT_TEXT
        If i < fieldCount Then DocEnd().InsertAfter " "
    Next i

    Dim mode As String
    If suppressScreen Then
        mode = "screenOff"
    Else
        mode = "screenOn"
    End If

    Dim t0 As Double
    Dim elapsed As Double
    Dim pairs As Collection
    t0 = PerfNow()
    Set pairs = FieldCollectIntextCitationFieldsInRange(ActiveDocument.Content)
    elapsed = PerfElapsed(t0)
    OutScale "collect+parse", fieldCount, mode, elapsed

    Dim pair As Variant
    Dim pageNumber As Long
    t0 = PerfNow()
    For Each pair In pairs
        pageNumber = CLng(pair("field").Result.Information(wdActiveEndPageNumber))
    Next pair
    elapsed = PerfElapsed(t0)
    OutScale "page lookup", fieldCount, mode, elapsed

    Dim updates As Collection
    Set updates = New Collection
    t0 = PerfNow()
    For i = 1 To pairs.Count
        updates.Add BuildIntextData("scale-" & CStr(i), "rich")
    Next i
    elapsed = PerfElapsed(t0)
    OutScale "build response objects", fieldCount, mode, elapsed

    t0 = PerfNow()
    For i = 1 To pairs.Count
        FieldWriteData pairs(i)("field"), updates(i)
    Next i
    elapsed = PerfElapsed(t0)
    OutScale "write JSON", fieldCount, mode, elapsed

    t0 = PerfNow()
    For i = 1 To pairs.Count
        FieldRenderStyledFieldWithData pairs(i)("field"), updates(i), updates(i)("content")
    Next i
    elapsed = PerfElapsed(t0)
    OutScale "render rich", fieldCount, mode, elapsed

    Application.ScreenUpdating = originalScreenUpdating
    ResetPerfDocument
End Sub


' --- Bibliography rebuild --------------------------------------------------

Private Sub BibliographyOps(ByVal sizes As Collection)
    Out "[SECTION] bibliography rebuild scaling"
    Out "[INFO] mirrors full delete/recreate; one formatted entry and bookmark per item"

    Dim item As Variant
    For Each item In sizes
        RunBibliographyScale CLng(item)
    Next item
    Out ""
End Sub

Private Sub RunBibliographyScale(ByVal itemCount As Long)
    ResetPerfDocument

    Dim styleName As String
    styleName = "Banyan Perf Bibliography Entry"
    EnsureParagraphStyle styleName

    Dim tCreate As Double
    Dim tWrite As Double
    Dim tRender As Double
    Dim tBookmark As Double
    Dim tParagraph As Double
    Dim i As Long

    Dim cursor As Range
    Set cursor = DocEnd()
    For i = 1 To itemCount
        Dim data As Object
        Set data = BuildBibliographyData(i)
        Dim fld As Field
        Dim t0 As Double

        t0 = PerfNow()
        Set fld = FieldCreateRawAddinField(cursor, "BANYAN_BIBLIOGRAPHY bib_" & CStr(i))
        tCreate = tCreate + PerfElapsed(t0)

        t0 = PerfNow()
        FieldWriteData fld, data
        tWrite = tWrite + PerfElapsed(t0)

        t0 = PerfNow()
        FieldRenderStyledFieldWithStyle fld, styleName, wdStyleTypeParagraph, data("content")
        tRender = tRender + PerfElapsed(t0)

        t0 = PerfNow()
        FieldAddBookmarkToField fld, FieldGetBibliographyBookmarkName("bib_" & CStr(i))
        tBookmark = tBookmark + PerfElapsed(t0)

        Set cursor = fld.Result.Duplicate
        cursor.Collapse wdCollapseEnd
        If i < itemCount Then
            t0 = PerfNow()
            cursor.InsertParagraphAfter
            cursor.Collapse wdCollapseEnd
            tParagraph = tParagraph + PerfElapsed(t0)
        End If
    Next i

    OutScale "bibliography create field", itemCount, "rebuild", tCreate
    OutScale "bibliography write JSON", itemCount, "rebuild", tWrite
    OutScale "bibliography render", itemCount, "rebuild", tRender
    OutScale "bibliography bookmark", itemCount, "rebuild", tBookmark
    OutScale "bibliography paragraph", itemCount, "rebuild", tParagraph

    Dim bibliographyFields As Collection
    Set bibliographyFields = New Collection
    Dim candidate As Field
    Dim readData As Object
    t0 = PerfNow()
    For Each candidate In ActiveDocument.Content.Fields
        If candidate.Type = wdFieldAddin Then
            Set readData = FieldReadData(candidate)
            If FieldIsBibliographyEntry(readData) Then bibliographyFields.Add candidate
        End If
    Next candidate
    OutScale "bibliography collect", itemCount, "delete", PerfElapsed(t0)

    Dim deleteElapsed As Double
    t0 = PerfNow()
    If bibliographyFields.Count > 0 Then
        Set candidate = bibliographyFields(1)
        If candidate.Locked Then candidate.Locked = False
        candidate.Delete
        For i = bibliographyFields.Count To 2 Step -1
            Set candidate = bibliographyFields(i)
            If candidate.Locked Then candidate.Locked = False
            candidate.Result.Paragraphs(1).Range.Delete
        Next i
    End If
    deleteElapsed = PerfElapsed(t0)
    OutScale "bibliography delete", itemCount, "delete", deleteElapsed
    ResetPerfDocument
End Sub


' --- JSON payload and response matching -----------------------------------

Private Sub JsonAndLookupOps(ByVal sizes As Collection)
    Out "[SECTION] JSON payload and response matching scaling"

    Dim item As Variant
    For Each item In sizes
        Dim itemCount As Long
        itemCount = CLng(item)

        Dim citations As Collection
        Set citations = New Collection
        Dim i As Long
        For i = 1 To itemCount
            citations.Add BuildIntextData("lookup-" & CStr(i), "rich")
        Next i

        Dim t0 As Double
        Dim elapsed As Double
        Dim payload As String
        t0 = PerfNow()
        payload = JsonStringify(citations)
        elapsed = PerfElapsed(t0)
        OutScale "JSON stringify", itemCount, "memory", elapsed

        Dim parsed As Variant
        t0 = PerfNow()
        Set parsed = JsonParse(payload)
        elapsed = PerfElapsed(t0)
        OutScale "JSON parse", itemCount, "memory", elapsed

        Dim found As Variant
        Dim foundObject As Object
        t0 = PerfNow()
        For i = itemCount To 1 Step -1
            Set foundObject = LinearFindById(citations, "lookup-" & CStr(i))
        Next i
        elapsed = PerfElapsed(t0)
        OutScale "linear id matching", itemCount, "worstOrder", elapsed

        Dim index As Object
        Set index = New Dictionary
        t0 = PerfNow()
        For Each found In citations
            Set index(DictKeyString(found, "id")) = found
        Next found
        For i = itemCount To 1 Step -1
            Set foundObject = index("lookup-" & CStr(i))
        Next i
        elapsed = PerfElapsed(t0)
        OutScale "indexed id matching", itemCount, "build+lookup", elapsed
        Out "[INFO] JSON chars size=" & CStr(itemCount) & " value=" & CStr(Len(payload))
    Next item
    Out ""
End Sub

Private Function LinearFindById(ByVal citations As Collection, ByVal citationId As String) As Object
    Dim item As Variant
    For Each item In citations
        If DictKeyString(item, "id") = citationId Then
            Set LinearFindById = item
            Exit Function
        End If
    Next item
End Function


' --- Footnote create/rebuild ----------------------------------------------

Private Sub NoteOps(ByVal reps As Long)
    Out "[SECTION] note citation operations (averages over reps)"

    Dim tCreate As Double
    Dim tSame As Double
    Dim tChanged As Double
    Dim r As Long
    For r = 1 To reps
        ResetPerfDocument
        Dim created As Collection
        Dim t0 As Double
        t0 = PerfNow()
        Set created = FieldCreateNoteCitationAtRange(DocEnd(), BuildNoteData("note-" & CStr(r), "[1]"))
        tCreate = tCreate + PerfElapsed(t0)
        If created Is Nothing Then Err.Raise 5, "testPerf.NoteOps", "Could not create note citation."

        Dim rebuilt As Collection
        t0 = PerfNow()
        Set rebuilt = FieldRebuildNoteCitationAtRange(created("note"), created("field"), _
                                                      BuildNoteData("note-" & CStr(r), "[1]"))
        tSame = tSame + PerfElapsed(t0)

        ResetPerfDocument
        Set created = FieldCreateNoteCitationAtRange(DocEnd(), BuildNoteData("note-" & CStr(r), "[1]"))
        t0 = PerfNow()
        Set rebuilt = FieldRebuildNoteCitationAtRange(created("note"), created("field"), _
                                                      BuildNoteData("note-" & CStr(r), "[2]"))
        tChanged = tChanged + PerfElapsed(t0)
    Next r

    OutTime "note create", tCreate / reps
    OutTime "note rebuild same reference", tSame / reps
    OutTime "note rebuild changed reference", tChanged / reps
    Out ""
End Sub


' --- Data builders ---------------------------------------------------------

Private Function BuildIntextData(ByVal id As String, Optional ByVal kind As String = "rich") As Object
    Dim data As Object
    Set data = FieldCreatePlaceholderIntextCitationData(id)
    Set data("source") = EmptySource()

    Dim content As Object
    Set content = New Dictionary
    content("text") = INTEXT_TEXT

    Dim marks As Collection
    Set marks = New Collection
    Select Case kind
        Case "format"
            AddFormatMarks marks
        Case "link"
            marks.Add Mark("link", 0, Len(INTEXT_TEXT), "banyan://entry/" & id)
        Case "rich"
            AddFormatMarks marks
            marks.Add Mark("link", 0, Len(INTEXT_TEXT), "banyan://entry/" & id)
    End Select
    Set content("marks") = marks
    Set data("content") = content
    Set BuildIntextData = data
End Function

Private Sub AddFormatMarks(ByVal marks As Collection)
    marks.Add Mark("bold", 1, 6, True)
    marks.Add Mark("italic", 8, 12, True)
    marks.Add Mark("color", 0, Len(INTEXT_TEXT), "#ff0000")
End Sub

Private Function BuildNoteData(ByVal id As String, ByVal refText As String) As Object
    Dim data As Object
    Set data = FieldCreatePlaceholderNoteCitationData(id)
    Set data("source") = EmptySource()
    Set data("content") = RichTextPlain(NOTE_BODY)
    Set data("reference") = RichTextPlain(refText)
    Set BuildNoteData = data
End Function

Private Function BuildBibliographyData(ByVal itemNumber As Long) As Object
    Dim data As Object
    Set data = New Dictionary
    data("id") = "bib_" & CStr(itemNumber)
    data("type") = "bibliography-entry"

    Dim content As Object
    Set content = New Dictionary
    content("text") = CStr(itemNumber) & ". Zhang. A representative bibliography entry for performance testing."
    Dim marks As Collection
    Set marks = New Collection
    marks.Add Mark("italic", 3, 8, True)
    Set content("marks") = marks
    Set data("content") = content
    Set BuildBibliographyData = data
End Function

Private Function EmptySource() As Object
    Dim source As Object
    Set source = New Dictionary
    Set source("cites") = New Collection
    Set source("params") = New Dictionary
    Set EmptySource = source
End Function

Private Function RichTextPlain(ByVal text As String) As Object
    Dim content As Object
    Set content = New Dictionary
    content("text") = text
    Set content("marks") = New Collection
    Set RichTextPlain = content
End Function

Private Function Mark(ByVal markType As String, _
                      ByVal startPos As Long, _
                      ByVal endPos As Long, _
                      ByVal value As Variant) As Object
    Dim item As Object
    Set item = New Dictionary
    item("type") = markType
    item("start") = startPos
    item("end") = endPos
    If IsObject(value) Then
        Set item("value") = value
    Else
        item("value") = value
    End If
    Set Mark = item
End Function


' --- Helpers ---------------------------------------------------------------

Private Sub WarmUp()
    Dim fld As Field
    Dim data As Object
    Set data = BuildIntextData("warmup", "rich")
    Set fld = CreatePreparedField(data)
    FieldRenderStyledField fld, data("content")
    FieldRemoveFieldSafely fld

    Dim json As String
    json = JsonStringify(data)
    Dim parsed As Variant
    Set parsed = JsonParse(json)
End Sub

Private Function CreatePreparedField(ByVal data As Object) As Field
    Dim fld As Field
    Set fld = FieldCreateRawAddinField(DocEnd(), "BANYAN_CITATION " & DictKeyString(data, "id"))
    If fld Is Nothing Then Err.Raise 5, "testPerf.CreatePreparedField", "Could not create raw field."
    FieldWriteData fld, data
    Set CreatePreparedField = fld
End Function

Private Sub EnsureParagraphStyle(ByVal styleName As String)
    On Error Resume Next
    Dim style As Style
    Set style = ActiveDocument.Styles(styleName)
    On Error GoTo 0
    If style Is Nothing Then ActiveDocument.Styles.Add Name:=styleName, Type:=wdStyleTypeParagraph
End Sub

Private Function DocEnd() As Range
    Dim target As Range
    Set target = ActiveDocument.Content.Duplicate
    target.Collapse wdCollapseEnd
    Set DocEnd = target
End Function

Private Sub ResetPerfDocument()
    On Error Resume Next
    ActiveDocument.Content.Delete
    ActiveDocument.Content.Text = "Banyan perf doc. "
    On Error GoTo 0
End Sub

Private Function ParseSizes(ByVal csv As String) As Collection
    Dim result As Collection
    Set result = New Collection

    Dim parts As Variant
    parts = Split(csv, ",")
    Dim part As Variant
    For Each part In parts
        If IsNumeric(Trim$(CStr(part))) Then
            Dim value As Long
            value = CLng(Trim$(CStr(part)))
            If value > 0 And value <= 1000 Then result.Add value
        End If
    Next part
    Set ParseSizes = result
End Function

Private Function PerfNow() As Double
#If Mac Then
    PerfNow = CDbl(Timer)
#Else
    Dim counter As Currency
    Dim frequency As Currency
    If QueryPerformanceCounter(counter) <> 0 And QueryPerformanceFrequency(frequency) <> 0 Then
        PerfNow = CDbl(counter / frequency)
    Else
        PerfNow = CDbl(Timer)
    End If
#End If
End Function

Private Function PerfElapsed(ByVal startedAt As Double) As Double
    Dim finishedAt As Double
    finishedAt = PerfNow()
    PerfElapsed = finishedAt - startedAt
    If PerfElapsed < 0 Then PerfElapsed = PerfElapsed + 86400#
End Function

Private Function PerfClockName() As String
#If Mac Then
    PerfClockName = "Timer"
#Else
    PerfClockName = "QueryPerformanceCounter"
#End If
End Function

Private Function BoolText(ByVal value As Boolean) As String
    If value Then
        BoolText = "true"
    Else
        BoolText = "false"
    End If
End Function

Private Function FmtMs(ByVal seconds As Double) As String
    If seconds >= 1# Then
        FmtMs = Format$(seconds, "0.000") & " s"
    Else
        FmtMs = Format$(seconds * 1000#, "0.00") & " ms"
    End If
End Function

Private Sub OutTime(ByVal name As String, ByVal seconds As Double)
    Out "[TIME] " & name & " total=" & FmtMs(seconds)
End Sub

Private Sub OutScale(ByVal name As String, ByVal itemCount As Long, _
                     ByVal mode As String, ByVal seconds As Double)
    Out "[TIME] " & name & " size=" & CStr(itemCount) & " mode=" & mode & _
        " total=" & FmtMs(seconds) & " perItem=" & FmtMs(seconds / itemCount)
End Sub

Private Sub Out(ByVal text As String)
    m_report = m_report & text & vbCrLf
End Sub
