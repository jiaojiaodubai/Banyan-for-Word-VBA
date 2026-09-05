Option Explicit

' ============================================================================
' Module  : modCitation
' Purpose : Insert/edit Banyan citations, based on WPS citation.ts.
'
' Public API:
'   CitationAction()  - Ribbon command entry point
' ============================================================================

Private m_i18nReady As Boolean


' --- Ribbon action ---

Public Sub CitationAction()
    EnsureCitationI10n

    On Error GoTo ErrHandler

    If Selection.StoryType <> wdMainTextStory Then
        MsgBox CText("notInMainText", "Citations can only be inserted into the main text of the document."), vbExclamation
        Exit Sub
    End If

    Dim pref As Object
    Set pref = PreferenceEnsure()
    If pref Is Nothing Then
        MsgBox CText("noStyle", "Please set citation style in Preferences first."), vbExclamation
        Exit Sub
    End If

    Dim style As Object
    Set style = pref("style")

    Dim citationType As String
    citationType = DictKeyString(style, "citationType")

    Dim updated As Boolean
    Select Case citationType
        Case "intext-citation"
            updated = HandleIntextCitation(Selection.Range.Duplicate, style)
        Case "note-citation"
            updated = HandleNoteCitation(Selection.Range.Duplicate, style)
        Case Else
            MsgBox CText("noStyle", "Please set citation style in Preferences first."), vbExclamation
            Exit Sub
    End Select

    If updated Then
        ProgressOpen CText("progressReason", "Processing citation...")
        RefreshAfterCitation pref
        ProgressClose
    End If
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modCitation.CitationAction"
    ProgressClose
    DiagnosticShowError CText("dialogTitle", "Banyan Citation"), _
                        Replace(CText("error", "An error occurred while handling the citation: {message}"), _
                                "{message}", Err.Description), _
                        Err.Number, Err.Source, Err.Description, Erl
End Sub

' --- In-text citation flow ---

Private Function HandleIntextCitation(ByVal targetRange As Range, ByVal style As Object) As Boolean
    Dim fd As Collection
    Set fd = FindSelectedIntextCitation(targetRange)

    If fd Is Nothing Then
        HandleIntextCitation = AddIntextCitation(targetRange, style)
    Else
        HandleIntextCitation = EditIntextCitation(fd, style)
    End If
End Function

Private Function AddIntextCitation(ByVal targetRange As Range, ByVal style As Object) As Boolean
    On Error GoTo ErrHandler

    Dim data As Object
    Set data = FieldCreatePlaceholderIntextCitationData(FieldCreateId(), FieldCreateEmptyCitationSource())

    Dim fld As Field
    Set fld = FieldCreateIntextCitationAtRange(targetRange, data)
    If fld Is Nothing Then Exit Function
    RepaintCitationPlaceholder

    Dim source As Object
    Set source = RequestCitation(style)
    If source Is Nothing Then
        FieldRemoveFieldSafely fld
        Exit Function
    End If

    Set data("source") = source
    FieldWriteData fld, data
    FieldRenderStyledField fld
    AddIntextCitation = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modCitation.AddIntextCitation"
    On Error Resume Next
    FieldRemoveFieldSafely fld
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function EditIntextCitation(ByVal fd As Collection, ByVal style As Object) As Boolean
    Dim fld As Field
    Dim data As Object
    Set fld = fd("field")
    Set data = fd("data")

    Dim source As Object
    Set source = RequestCitation(style, data("source"))
    If source Is Nothing Then Exit Function

    Set data("source") = source
    FieldWriteData fld, data
    FieldRenderStyledField fld
    EditIntextCitation = True
End Function


' --- Note citation flow ---

Private Function HandleNoteCitation(ByVal targetRange As Range, ByVal style As Object) As Boolean
    Dim fd As Collection
    Set fd = FindSelectedNoteCitation(targetRange)

    If fd Is Nothing Then
        HandleNoteCitation = AddNoteCitation(targetRange, style)
    Else
        HandleNoteCitation = EditNoteCitation(fd, style)
    End If
End Function

Private Function AddNoteCitation(ByVal targetRange As Range, ByVal style As Object) As Boolean
    On Error GoTo ErrHandler

    Dim data As Object
    Set data = FieldCreatePlaceholderNoteCitationData(FieldCreateId(), FieldCreateEmptyCitationSource())

    Dim created As Collection
    Set created = FieldCreateNoteCitationAtRange(targetRange, data)
    If created Is Nothing Then Exit Function

    Dim note As Footnote
    Dim fld As Field
    Set note = created("note")
    Set fld = created("field")
    RepaintCitationPlaceholder

    Dim source As Object
    Set source = RequestCitation(style)
    If source Is Nothing Then
        FieldRemoveFootnoteSafely note
        Exit Function
    End If

    Set data("source") = source
    FieldWriteData fld, data
    FieldRenderStyledField fld
    AddNoteCitation = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modCitation.AddNoteCitation"
    On Error Resume Next
    FieldRemoveFootnoteSafely note
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function EditNoteCitation(ByVal fd As Collection, ByVal style As Object) As Boolean
    Dim fld As Field
    Dim data As Object
    Set fld = fd("field")
    Set data = fd("data")

    Dim source As Object
    Set source = RequestCitation(style, data("source"))
    If source Is Nothing Then Exit Function

    Set data("source") = source
    FieldWriteData fld, data
    EditNoteCitation = True
End Function


' --- Selection detection ---

Private Function FindSelectedIntextCitation(ByVal targetRange As Range) As Collection
    On Error GoTo ErrHandler

    Dim caret As Long
    caret = targetRange.End

    Dim paragraph As Range
    Set paragraph = targetRange.Paragraphs(targetRange.Paragraphs.Count).Range.Duplicate

    Dim fld As Field
    Dim data As Object
    For Each fld In paragraph.Fields
        If fld.Type = wdFieldAddin Then
            If RangeContainsCaret(fld.Result, caret) Then
                Set data = FieldReadData(fld)
                If FieldIsIntextCitation(data) Then
                    Set FindSelectedIntextCitation = MakeFieldAndData(fld, data)
                    Exit Function
                End If

                MsgBox CText("notBanyanIntextCitation", "Please select a Banyan created citation."), vbExclamation
                Exit Function
            End If
        End If
    Next fld
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modCitation.FindSelectedIntextCitation"
    Set FindSelectedIntextCitation = Nothing
End Function

Private Function FindSelectedNoteCitation(ByVal targetRange As Range) As Collection
    On Error GoTo ErrHandler

    Dim caret As Long
    caret = targetRange.End

    Dim paragraph As Range
    Set paragraph = targetRange.Paragraphs(targetRange.Paragraphs.Count).Range.Duplicate

    Dim note As Footnote
    Dim fld As Field
    Dim data As Object

    For Each note In paragraph.Footnotes
        If RangeContainsCaret(note.Reference, caret) Then
            For Each fld In note.Range.Fields
                If fld.Type = wdFieldAddin Then
                    Set data = FieldReadData(fld)
                    If FieldIsNoteCitation(data) Then
                        Set FindSelectedNoteCitation = MakeNoteFieldAndData(note, fld, data)
                        Exit Function
                    End If

                    MsgBox CText("notBanyanNoteCitation", "Please select a Banyan created note citation."), vbExclamation
                    Exit Function
                End If
            Next fld
        End If
    Next note
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modCitation.FindSelectedNoteCitation"
    Set FindSelectedNoteCitation = Nothing
End Function

Private Function RangeContainsCaret(ByVal targetRange As Range, ByVal caret As Long) As Boolean
    RangeContainsCaret = (targetRange.Start <= caret And targetRange.End >= caret)
End Function


' --- Backend ---

Private Function RequestCitation(ByVal style As Object, Optional ByVal currentSource As Variant) As Object
    Dim body As Object
    Set body = New Dictionary
    body("documentId") = GetDocumentId()
    Set body("style") = FieldAsStyleIdentifier(style)

    If Not IsMissing(currentSource) Then
        If IsObject(currentSource) Then
            If Not currentSource Is Nothing Then
                Set body("source") = currentSource
            End If
        End If
    End If

    Dim respText As String
    respText = HttpPost(HttpBuildUrl("citation"), JsonStringify(body))
    If Len(respText) = 0 Then Exit Function

    Dim envelope As Object
    Set envelope = JsonParse(respText)
    If envelope Is Nothing Then Exit Function

    If Not HttpEnvelopeOk(envelope) Then
        ShowHttpEnvelopeError envelope
        Exit Function
    End If

    If Not DictHasKey(envelope, "data") Then Exit Function
    If DictKeyIsNull(envelope, "data") Or DictKeyIsEmpty(envelope, "data") Then Exit Function
    If Not DictKeyIsObject(envelope, "data") Then Exit Function
    If Not FieldIsCitationSource(DictKeyObject(envelope, "data")) Then Exit Function

    Set RequestCitation = DictKeyObject(envelope, "data")
End Function

Private Function HttpEnvelopeOk(ByVal envelope As Object) As Boolean
    On Error GoTo ErrHandler
    If Not envelope.Exists("ok") Then Exit Function
    HttpEnvelopeOk = DictKeyBool(envelope, "ok")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modCitation.HttpEnvelopeOk"
    HttpEnvelopeOk = False
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
    DiagnosticsReraiseIfDev "modCitation.ShowHttpEnvelopeError"
End Sub


' --- Refresh hook ---

Private Sub RefreshAfterCitation(ByVal pref As Object)
    On Error Resume Next
    RefreshAction
    On Error GoTo 0
End Sub

Private Sub RepaintCitationPlaceholder()
    On Error Resume Next
    Application.ScreenRefresh
    DoEvents
    On Error GoTo 0
End Sub


' --- Local i18n ---

Private Sub EnsureCitationI10n()
    If m_i18nReady Then Exit Sub

    I10nRegisterTable msoLanguageIDSimplifiedChinese, "citation", _
        "dialogTitle", "Banyan 引注", _
        "notInMainText", "引注只能插入在主文档正文中。", _
        "notBanyanIntextCitation", "请选择 Banyan 创建的引注。", _
        "notBanyanNoteCitation", "请选择 Banyan 创建的脚注引用。", _
        "noStyle", "请先在设置中选择引用样式。", _
        "error", "操作引注时发生错误：{message}", _
        "progressReason", "正在处理引注..."

    I10nRegisterTable msoLanguageIDEnglishUS, "citation", _
        "dialogTitle", "Banyan Citation", _
        "notInMainText", "Citations can only be inserted into the main text of the document.", _
        "notBanyanIntextCitation", "Please select a Banyan created citation.", _
        "notBanyanNoteCitation", "Please select a Banyan created note citation.", _
        "noStyle", "Please set citation style in Preferences first.", _
        "error", "An error occurred while handling the citation: {message}", _
        "progressReason", "Processing citation..."

    m_i18nReady = True
End Sub

Private Function CText(ByVal key As String, ByVal fallback As String) As String
    EnsureCitationI10n
    CText = T("citation." & key, fallback)
End Function


' --- Small collection helpers ---

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
