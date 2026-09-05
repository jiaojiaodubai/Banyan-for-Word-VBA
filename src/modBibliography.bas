Option Explicit

' ============================================================================
' Module  : modBibliography
' Purpose : Insert/edit Banyan bibliography blocks, based on WPS bibliography.ts.
'
' Public API:
'   BibliographyAction()  - Ribbon command entry point
'
' Behaviour:
'   - Empty selection/caret: insert a bibliography for the current chapter.
'   - One selected bibliography-entry field: edit that entry via backend.
'   - Existing bibliography fields in the current chapter are replaced on insert.
' ============================================================================

Private Const FIELD_PLACEHOLDER_COLOR As String = "#ff0000"

Private m_i18nReady As Boolean


' --- Ribbon action ---

Public Sub BibliographyAction()
    EnsureBibliographyI10n

    On Error GoTo ErrHandler

    If Selection.StoryType <> wdMainTextStory Then
        MsgBox BText("notInMainText", "Bibliography can only be inserted in the main text story."), vbExclamation
        Exit Sub
    End If

    Dim pref As Object
    Set pref = PreferenceEnsure()
    If pref Is Nothing Then
        MsgBox BText("noStyle", "Please set citation style in Preferences first."), vbExclamation
        Exit Sub
    End If

    Dim mode As Collection
    Set mode = GetActionMode()
    If mode Is Nothing Then Exit Sub

    Dim updateRange As Range
    Set updateRange = GetUpdateRange()
    If updateRange Is Nothing Then Exit Sub

    Dim contexts As Collection
    Set contexts = CollectCitationContexts(updateRange)
    If contexts.Count = 0 Then
        MsgBox BText("noCitationInRange", "No citation was found in the current section. Please add citations first."), vbExclamation
        Exit Sub
    End If

    ProgressOpen BText("progressReason", "Processing bibliography...")
    ' NOTE: `mode` is a VBA Collection (see GetActionMode/MakeAddMode/
    ' MakeEditMode), NOT a Dictionary - read it with Collection key syntax.
    ' (DictKeyString on a Collection raises error 438: Collection has no
    ' .Exists.) mode("range") / mode("field") below are also Collection reads.
    If CStr(mode("type")) = "add" Then
        AddBibliography mode("range"), updateRange, contexts, pref
    ElseIf CStr(mode("type")) = "edit" Then
        EditBibliographyEntry mode("field"), pref
    End If
    ProgressClose
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.BibliographyAction"
    ProgressClose
    DiagnosticShowError BText("dialogTitle", "Banyan Bibliography"), _
                        Replace(BText("error", "An error occurred while handling bibliography: {message}"), _
                                "{message}", Err.Description), _
                        Err.Number, Err.Source, Err.Description, Erl
End Sub


' --- Action mode ---

Private Function GetActionMode() As Collection
    Dim selectedRange As Range
    Set selectedRange = Selection.Range.Duplicate

    If selectedRange.Fields.Count = 0 Then
        Dim caretField As Field
        Set caretField = FindAddinFieldAtCaret(selectedRange)
        If Not caretField Is Nothing Then
            Dim caretData As Object
            Set caretData = FieldReadData(caretField)
            If FieldIsBibliographyEntry(caretData) Then
                Set GetActionMode = MakeEditMode(caretField)
                Exit Function
            End If

            MsgBox BText("notBanyanBibliographyEntry", "Please select a Banyan created bibliography entry field."), vbExclamation
            Exit Function
        End If

        selectedRange.Collapse wdCollapseEnd
        Set GetActionMode = MakeAddMode(selectedRange)
        Exit Function
    End If

    If selectedRange.Fields.Count > 1 Then
        MsgBox BText("multipleFields", _
                     "Multiple fields detected. Select only one bibliography entry field or place the cursor where you want to insert bibliography."), _
               vbExclamation
        Exit Function
    End If

    Dim fld As Field
    Set fld = selectedRange.Fields(1)
    If fld.Type <> wdFieldAddin Then
        MsgBox BText("notBanyanBibliographyEntry", "Please select a Banyan created bibliography entry field."), vbExclamation
        Exit Function
    End If

    Dim data As Object
    Set data = FieldReadData(fld)
    If Not FieldIsBibliographyEntry(data) Then
        MsgBox BText("notBanyanBibliographyEntry", "Please select a Banyan created bibliography entry field."), vbExclamation
        Exit Function
    End If

    Set GetActionMode = MakeEditMode(fld)
End Function

Private Function FindAddinFieldAtCaret(ByVal selectedRange As Range) As Field
    On Error GoTo ErrHandler

    Dim caret As Long
    caret = selectedRange.End

    Dim paragraph As Range
    Set paragraph = selectedRange.Paragraphs(selectedRange.Paragraphs.Count).Range.Duplicate

    Dim fld As Field
    Dim data As Object
    For Each fld In paragraph.Fields
        If fld.Type = wdFieldAddin Then
            If RangeContainsPosition(fld.Result, caret) Then
                Set FindAddinFieldAtCaret = fld
                Exit Function
            End If
        End If
    Next fld
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.FindAddinFieldAtCaret"
    Set FindAddinFieldAtCaret = Nothing
End Function

Private Function MakeAddMode(ByVal targetRange As Range) As Collection
    Dim result As Collection
    Set result = New Collection
    result.Add "add", "type"
    result.Add targetRange, "range"
    Set MakeAddMode = result
End Function

Private Function MakeEditMode(ByVal fld As Field) As Collection
    Dim result As Collection
    Set result = New Collection
    result.Add "edit", "type"
    result.Add fld, "field"
    Set MakeEditMode = result
End Function

Private Function RangeContainsPosition(ByVal targetRange As Range, ByVal position As Long) As Boolean
    RangeContainsPosition = (targetRange.Start <= position And targetRange.End >= position)
End Function


' --- Add bibliography ---

Private Sub AddBibliography(ByVal insertRange As Range, _
                            ByVal updateRange As Range, _
                            ByVal contexts As Collection, _
                            ByVal pref As Object)
    On Error GoTo ErrHandler

    Dim pendingField As Field
    Set pendingField = CreatePendingBibliographyField(insertRange, pref)
    If pendingField Is Nothing Then Exit Sub

    Dim response As Object
    Set response = RequestRefresh(pref("style"), contexts, DictKeyBool(pref, "syncItems"))
    If response Is Nothing Then
        FieldRemoveFieldSafely pendingField
        MsgBox BText("refreshFailed", "Failed to fetch bibliography data."), vbExclamation
        Exit Sub
    End If

    Dim lines As Collection
    Set lines = ExtractBibliographyLines(response)
    If lines.Count = 0 Then
        FieldRemoveFieldSafely pendingField
        MsgBox BText("invalidBibliographyLine", "Invalid bibliography data returned by server."), vbExclamation
        Exit Sub
    End If

    DeleteExistingBibliography updateRange
    InsertBibliography insertRange, lines, pref
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.AddBibliography"
    On Error Resume Next
    FieldRemoveFieldSafely pendingField
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Private Function CreatePendingBibliographyField(ByVal targetRange As Range, ByVal pref As Object) As Field
    On Error GoTo ErrHandler

    Dim data As Object
    Set data = PendingBibliographyLine()

    Dim cursor As Range
    Set cursor = targetRange.Duplicate
    cursor.Collapse wdCollapseEnd

    Dim fld As Field
    Set fld = FieldCreateRawAddinField(cursor, "BANYAN_BIBLIOGRAPHY")
    If fld Is Nothing Then Exit Function

    FieldWriteData fld, data
    FieldRenderStyledFieldWithStyle fld, DictKeyString(pref, "bibliographyTitleStyle"), wdStyleTypeParagraph, data("content")
    Set CreatePendingBibliographyField = fld
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.CreatePendingBibliographyField"
    Set CreatePendingBibliographyField = Nothing
End Function

Private Function PendingBibliographyLine() As Object
    Dim data As Object
    Set data = New Dictionary
    data("type") = "bibliography-title"
    Set data("content") = FieldCreateRichText("{ BIBLIOGRAPHY }", FIELD_PLACEHOLDER_COLOR)

    Set PendingBibliographyLine = data
End Function

Private Sub DeleteExistingBibliography(ByVal targetRange As Range)
    On Error GoTo ErrHandler

    Dim bibliographyFields As Collection
    Set bibliographyFields = CollectBibliographyFieldsInRange(targetRange)
    If bibliographyFields.Count = 0 Then Exit Sub

    Dim firstField As Field
    Set firstField = bibliographyFields(1)
    FieldRemoveFieldSafely firstField

    Dim i As Long
    Dim fld As Field
    For i = bibliographyFields.Count To 2 Step -1
        Set fld = bibliographyFields(i)
        If fld.Locked Then fld.Locked = False
        fld.Result.Paragraphs(1).Range.Delete
    Next i
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.DeleteExistingBibliography"
End Sub

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

Private Sub InsertBibliography(ByVal targetRange As Range, ByVal lines As Collection, ByVal pref As Object)
    On Error GoTo ErrHandler

    Dim cursor As Range
    Set cursor = targetRange.Duplicate
    cursor.Collapse wdCollapseEnd

    Dim i As Long
    Dim line As Object
    Dim fld As Field
    Dim fieldCode As String

    For i = 1 To lines.Count
        Set line = lines(i)

        If FieldIsBibliographyEntry(line) Then
            fieldCode = "BANYAN_BIBLIOGRAPHY " & DictKeyString(line, "id")
        Else
            fieldCode = "BANYAN_BIBLIOGRAPHY"
        End If

        Set fld = FieldCreateRawAddinField(cursor, fieldCode)
        If fld Is Nothing Then Exit Sub

        FieldWriteData fld, line
        If FieldIsBibliographyTitle(line) Then
            FieldRenderStyledFieldWithStyle fld, DictKeyString(pref, "bibliographyTitleStyle"), wdStyleTypeParagraph, line("content")
        ElseIf FieldIsBibliographyEntry(line) Then
            FieldRenderStyledFieldWithStyle fld, DictKeyString(pref, "bibliographyEntryStyle"), wdStyleTypeParagraph, line("content")
            FieldAddBookmarkToField fld, FieldGetBibliographyBookmarkName(DictKeyString(line, "id"))
        End If

        Dim resultEnd As Range
        Set resultEnd = fld.Result.Duplicate
        resultEnd.Collapse wdCollapseEnd
        If i < lines.Count Then
            resultEnd.InsertParagraphAfter
            resultEnd.Collapse wdCollapseEnd
        End If
        cursor.SetRange resultEnd.Start, resultEnd.End
    Next i
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.InsertBibliography"
End Sub

Private Function ExtractBibliographyLines(ByVal response As Object) As Collection
    Dim result As Collection
    Set result = New Collection

    If Not HasDictionaryKey(response, "bibliography") Then
        Set ExtractBibliographyLines = result
        Exit Function
    End If
    If Not IsCollectionObject(response("bibliography")) Then
        Set ExtractBibliographyLines = result
        Exit Function
    End If

    Dim line As Variant
    For Each line In response("bibliography")
        If FieldIsBibliographyTitle(line) Or FieldIsBibliographyEntry(line) Then
            result.Add line
        End If
    Next line

    Set ExtractBibliographyLines = result
End Function


' --- Edit bibliography entry ---

Private Sub EditBibliographyEntry(ByVal fld As Field, ByVal pref As Object)
    Dim currentLine As Object
    Set currentLine = FieldReadData(fld)
    If Not FieldIsBibliographyEntry(currentLine) Then
        MsgBox BText("notBanyanBibliographyEntry", "Please select a Banyan created bibliography entry field."), vbExclamation
        Exit Sub
    End If

    Dim response As Object
    Set response = RequestBibliography(pref("style"), currentLine, pref)
    If response Is Nothing Then
        MsgBox BText("bibliographyFailed", "Failed to edit bibliography entry."), vbExclamation
        Exit Sub
    End If

    If Not HasDictionaryKey(response, "line") Then
        MsgBox BText("invalidBibliographyLine", "Invalid bibliography data returned by server."), vbExclamation
        Exit Sub
    End If
    If Not FieldIsBibliographyEntry(response("line")) Then
        MsgBox BText("invalidBibliographyLine", "Invalid bibliography data returned by server."), vbExclamation
        Exit Sub
    End If

    FieldWriteData fld, response("line")
    FieldRenderStyledFieldWithStyle fld, DictKeyString(pref, "bibliographyEntryStyle"), wdStyleTypeParagraph, DictKeyObject(DictKeyObject(response, "line"), "content")
    FieldAddBookmarkToField fld, FieldGetBibliographyBookmarkName(DictKeyString(DictKeyObject(response, "line"), "id"))

    SaveReturnedExtraSource pref, response
End Sub

Private Sub SaveReturnedExtraSource(ByVal pref As Object, ByVal response As Object)
    On Error GoTo ErrHandler
    If Not DictHasKey(response, "extraSource") Then Exit Sub

    If DictKeyIsNull(response, "extraSource") Then
        pref("extraSource") = Null
    ElseIf DictKeyIsEmpty(response, "extraSource") Then
        pref("extraSource") = Null
    ElseIf DictKeyIsObject(response, "extraSource") Then
        If FieldIsCitationSource(DictKeyObject(response, "extraSource")) Then
            Set pref("extraSource") = DictKeyObject(response, "extraSource")
        Else
            pref("extraSource") = Null
        End If
    Else
        pref("extraSource") = Null
    End If

    PreferenceSave pref
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.SaveReturnedExtraSource"
End Sub


' --- Citation contexts ---

Private Function CollectCitationContexts(ByVal targetRange As Range) As Collection
    Dim byId As Object
    Set byId = New Dictionary

    AddIntextContexts targetRange, byId
    AddNoteContexts targetRange, byId

    Dim result As Collection
    Set result = New Collection

    Dim key As Variant
    For Each key In byId.Keys
        result.Add byId(key)
    Next key

    Set CollectCitationContexts = result
End Function

Private Sub AddIntextContexts(ByVal targetRange As Range, ByVal byId As Object)
    On Error GoTo ErrHandler

    Dim fd As Variant
    For Each fd In FieldCollectIntextCitationFieldsInRange(targetRange)
        Dim context As Object
        Set context = BuildCitationContext(fd("field"), fd("data"))
        AddContextById byId, context
    Next fd
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.AddIntextContexts"
End Sub

Private Sub AddNoteContexts(ByVal targetRange As Range, ByVal byId As Object)
    On Error GoTo ErrHandler

    Dim fd As Variant
    For Each fd In FieldCollectNoteCitationFootnotesInRange(targetRange)
        Dim context As Object
        Set context = BuildCitationContext(fd("field"), fd("data"))
        AddContextById byId, context
    Next fd
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.AddNoteContexts"
End Sub

Private Sub AddContextById(ByVal byId As Object, ByVal context As Object)
    If context Is Nothing Then Exit Sub
    If Not HasDictionaryKey(context, "id") Then Exit Sub

    Dim contextId As String
    contextId = DictKeyString(context, "id")
    If Len(contextId) = 0 Then Exit Sub

    If byId.Exists(contextId) Then byId.Remove contextId
    byId.Add contextId, context
End Sub

Private Function BuildCitationContext(ByVal fld As Field, ByVal data As Object) As Object
    On Error GoTo ErrHandler

    Dim context As Object
    Set context = New Dictionary

    If Not HasDictionaryKey(data, "source") Then Exit Function
    If Not IsDictionaryRecord(data("source")) Then Exit Function

    Dim source As Object
    Set source = data("source")

    Dim key As Variant
    For Each key In source.Keys
        DictCopyKey context, CStr(key), source, CStr(key)
    Next key

    context("id") = DictKeyString(data, "id")
    context("page") = FieldPageNumber(fld)
    Set BuildCitationContext = context
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.BuildCitationContext"
    Set BuildCitationContext = Nothing
End Function

Private Function FieldPageNumber(ByVal fld As Field) As Long
    On Error GoTo ErrHandler
    FieldPageNumber = CLng(fld.Result.Information(wdActiveEndPageNumber))
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.FieldPageNumber"
    FieldPageNumber = 0
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
    If Not EnvelopeOk(envelope) Then
        ShowHttpEnvelopeError envelope
        Exit Function
    End If

    If Not HasDictionaryKey(envelope, "data") Then Exit Function
    If Not IsDictionaryRecord(envelope("data")) Then Exit Function

    Dim data As Object
    Set data = envelope("data")
    If Not HasDictionaryKey(data, "bibliography") Then Exit Function
    If Not IsCollectionObject(data("bibliography")) Then Exit Function

    Set RequestRefresh = data
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.RequestRefresh"
    Set RequestRefresh = Nothing
End Function

Private Function RequestBibliography(ByVal style As Object, _
                                     ByVal line As Object, _
                                     ByVal pref As Object) As Object
    On Error GoTo ErrHandler

    Dim body As Object
    Set body = New Dictionary
    body("documentId") = GetDocumentId()
    Set body("style") = FieldAsStyleIdentifier(style)
    Set body("line") = line
    CopyOptionalExtraSource body, pref

    Dim respText As String
    respText = HttpPost(HttpBuildUrl("bibliography"), JsonStringify(body))
    If Len(respText) = 0 Then Exit Function

    Dim envelope As Object
    Set envelope = JsonParse(respText)
    If envelope Is Nothing Then Exit Function
    If Not EnvelopeOk(envelope) Then
        ShowHttpEnvelopeError envelope
        Exit Function
    End If

    If Not DictKeyIsDict(envelope, "data") Then Exit Function

    Set RequestBibliography = DictKeyObject(envelope, "data")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.RequestBibliography"
    Set RequestBibliography = Nothing
End Function

Private Sub CopyOptionalExtraSource(ByVal body As Object, ByVal pref As Object)
    On Error GoTo ErrHandler

    If Not DictKeyIsObject(pref, "extraSource") Then Exit Sub
    If Not FieldIsCitationSource(DictKeyObject(pref, "extraSource")) Then Exit Sub

    Set body("extraSource") = DictKeyObject(pref, "extraSource")
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.CopyOptionalExtraSource"
End Sub

Private Function EnvelopeOk(ByVal envelope As Object) As Boolean
    On Error GoTo ErrHandler
    If Not envelope.Exists("ok") Then Exit Function
    EnvelopeOk = DictKeyBool(envelope, "ok")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.EnvelopeOk"
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
            MsgBox "HTTP error: [" & DictKeyString(errObj, "code") & "] " & DictKeyString(errObj, "message"), vbExclamation
        End If
    End If
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modBibliography.ShowHttpEnvelopeError"
End Sub


' --- Local i18n ---

Private Sub EnsureBibliographyI10n()
    If m_i18nReady Then Exit Sub

    I10nRegisterTable msoLanguageIDSimplifiedChinese, "bibliography", _
        "dialogTitle", "Banyan 参考文献表", _
        "notInMainText", "参考文献表只能插入在主文档正文中。", _
        "multipleFields", "检测到多个域，请只选择一个书目条目域或将光标放在要插入书目的位置。", _
        "notBanyanBibliographyEntry", "请选择 Banyan 创建的书目条目域。", _
        "noStyle", "请先在设置中选择引用样式。", _
        "noCitationInRange", "当前章节没有检测到引注，请先添加引注。", _
        "refreshFailed", "获取参考文献表数据失败。", _
        "bibliographyFailed", "编辑参考文献条目失败。", _
        "invalidBibliographyLine", "服务器返回的书目数据无效。", _
        "error", "操作参考文献表时发生错误：{message}", _
        "progressReason", "正在处理参考文献表..."

    I10nRegisterTable msoLanguageIDEnglishUS, "bibliography", _
        "dialogTitle", "Banyan Bibliography", _
        "notInMainText", "Bibliography can only be inserted in the main text story.", _
        "multipleFields", "Multiple fields detected. Select only one bibliography entry field or place the cursor where you want to insert bibliography.", _
        "notBanyanBibliographyEntry", "Please select a Banyan created bibliography entry field.", _
        "noStyle", "Please set citation style in Preferences first.", _
        "noCitationInRange", "No citation was found in the current section. Please add citations first.", _
        "refreshFailed", "Failed to fetch bibliography data.", _
        "bibliographyFailed", "Failed to edit bibliography entry.", _
        "invalidBibliographyLine", "Invalid bibliography data returned by server.", _
        "error", "An error occurred while handling bibliography: {message}", _
        "progressReason", "Processing bibliography..."

    m_i18nReady = True
End Sub

Private Function BText(ByVal key As String, ByVal fallback As String) As String
    EnsureBibliographyI10n
    BText = T("bibliography." & key, fallback)
End Function


' --- Small helpers ---

Private Function IsDictionaryRecord(ByVal value As Variant) As Boolean
    IsDictionaryRecord = DictIsDictionary(value)
End Function

Private Function IsCollectionObject(ByVal value As Variant) As Boolean
    IsCollectionObject = DictIsCollection(value)
End Function

Private Function HasDictionaryKey(ByVal dict As Object, ByVal key As String) As Boolean
    HasDictionaryKey = DictHasKey(dict, key)
End Function
