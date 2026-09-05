Option Explicit

' ============================================================================
' Module  : modChapterBreak
' Purpose : Chapter-break ADDIN field management for Banyan.
'
'           ChapterBreak fields partition the document into chapters.
'           Each field stores JSON data matching the WPS format:
'
'           {
'             "type": "chapter-break",
'             "style": { "id":"...", "title":"...", "citationType":"..." },
'             "bibliographyTitleStyle": "...",
'             "bibliographyEntryStyle": "...",
'             "extraSource": { ... },
'             "content": { "text": "", "marks": [] }
'           }
'
'           The field Code shows a visible prompt when ShowCodes = True;
'           the JSON data is stored in the Data property.
'
' Public API:
'   InsertChapterBreak(style, bibTitleStyle, bibEntryStyle, extraSource)
'   InsertChapterBreakFromPreference()
'   FindPreviousChapterBreak() -> Collection { "field", "data" } or Nothing
'   FindNextChapterBreak()     -> Collection { "field", "data" } or Nothing
'   GetUpdateRange()           -> Word.Range between adjacent chapter breaks
'   IsChapterBreak(data)       -> Boolean
'   IsCitationSource(data)     -> Boolean
'   CanInsertChapterBreakAtSelection() -> Boolean
'   ReadFieldData(field)       -> Dictionary (parsed JSON) or Nothing
'   WriteFieldData(field, dict)-> Boolean
'
' Dependencies: modJson, modI10n, modPreference
' ============================================================================

' --- Localization strings (match WPS) ---
Private Const CB_PROMPT_ZH As String = " ==========Banyan章节分隔符（请勿编辑）========== "
Private Const CB_PROMPT_EN As String = " ==========Banyan chapter break (Do not edit)========== "

Private Const CB_NOT_MAIN_ZH As String = "章节分隔符只能插入在主文档正文中。"
Private Const CB_NOT_MAIN_EN As String = "Chapter breaks can only be inserted into the main text of the document."

Private Const CB_NOT_START_ZH As String = "章节分隔符只能插入在段落开头。"
Private Const CB_NOT_START_EN As String = "Chapter breaks can only be inserted at the beginning of a paragraph."

Private Const CB_IN_FIELD_ZH As String = "章节分隔符不能插入在域代码中，请将光标移到普通正文位置。"
Private Const CB_IN_FIELD_EN As String = "Chapter breaks cannot be inserted inside fields. Move the cursor to normal text."

Private Const CB_NO_STYLE_ZH As String = "请先在设置中选择引用样式。"
Private Const CB_NO_STYLE_EN As String = "Please set citation style in Preferences first."


' --- InsertChapterBreak ---
' Creates a new ChapterBreak ADDIN field at the current selection.

Public Sub InsertChapterBreak(ByVal style As Object, _
                               ByVal bibliographyTitleStyle As String, _
                               ByVal bibliographyEntryStyle As String, _
    Optional ByVal extraSource As Variant = Null)
    If Not IsPrefStyle(style) Then Exit Sub

    Dim insertRange As Range
    Set insertRange = Selection.Range.Duplicate
    insertRange.Collapse wdCollapseEnd

    Dim paragraphEnd As Long
    paragraphEnd = insertRange.Paragraphs(1).Range.End

    Dim breakData As Object
    Set breakData = New Dictionary
    breakData("type") = "chapter-break"
    Set breakData("style") = style
    breakData("bibliographyTitleStyle") = bibliographyTitleStyle
    breakData("bibliographyEntryStyle") = bibliographyEntryStyle
    CopyOptionalExtraSource breakData, extraSource
    Set breakData("content") = FieldCreateEmptyRichText()

    Dim jsonStr As String
    jsonStr = JsonStringify(breakData)
    If Len(jsonStr) = 0 Then Exit Sub

    Dim fld As Field
    Set fld = FieldCreateRawAddinField(insertRange, PromptText)
    If fld Is Nothing Then Exit Sub

    fld.Data = jsonStr
    fld.Result.Text = ""
    fld.ShowCodes = True
    fld.Locked = True
    fld.Code.Font.Color = wdColorRed

    On Error Resume Next
    Dim afterField As Range
    Set afterField = fld.Code.Duplicate
    afterField.Collapse wdCollapseEnd
    If afterField.End + 1 < paragraphEnd Then
        afterField.InsertParagraphAfter
    End If
    On Error GoTo 0
End Sub


' --- InsertChapterBreakFromPreference ---
' Ribbon action: validate the insertion point, ensure Preference exists,
' then insert a ChapterBreak from the current chapter preference.

Public Sub InsertChapterBreakFromPreference()
    On Error GoTo ErrHandler

    If Not ValidateChapterBreakInsertionPoint() Then Exit Sub

    Dim pref As Object
    Set pref = PreferenceEnsure()
    If pref Is Nothing Then
        MsgBox PromptTextNoStyle, vbExclamation
        Exit Sub
    End If

    Dim style As Object
    Set style = pref("style")
    If Not IsPrefStyle(style) Then
        MsgBox PromptTextNoStyle, vbExclamation
        Exit Sub
    End If

    InsertChapterBreak style, _
        DictKeyString(pref, "bibliographyTitleStyle"), _
        DictKeyString(pref, "bibliographyEntryStyle"), _
        pref("extraSource")
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.InsertChapterBreakFromPreference"
    MsgBox PromptTextNoStyle, vbExclamation
End Sub


' --- FindPreviousChapterBreak ---
' Returns the nearest ChapterBreak before the normalized caret position.
' Mirrors WPS: start with Selection.PreviousField, then Field.Previous.
' The user's original selection is restored before returning.

Public Function FindPreviousChapterBreak() As Collection
    Dim originalRange As Range
    On Error Resume Next
    Set originalRange = Selection.Range.Duplicate
    On Error GoTo CleanUp

    EnsureCaretInMainText

    Set FindPreviousChapterBreak = FindPreviousChapterBreakByPosition()

CleanUp:
    RestoreSelection originalRange
    Exit Function
End Function

' --- FindNextChapterBreak ---
' Returns the nearest ChapterBreak after the normalized caret position.
' Mirrors WPS: start with Selection.NextField, then Field.Next.
' The user's original selection is restored before returning.

Public Function FindNextChapterBreak() As Collection
    Dim originalRange As Range
    On Error Resume Next
    Set originalRange = Selection.Range.Duplicate
    On Error GoTo CleanUp

    EnsureCaretInMainText

    Set FindNextChapterBreak = FindNextChapterBreakByPosition()

CleanUp:
    RestoreSelection originalRange
    Exit Function
End Function


' --- GetUpdateRange ---
' Returns the range for the current chapter:
' previous ChapterBreak start or document start
' next ChapterBreak start or document end

Public Function GetUpdateRange() As Range
    On Error GoTo ErrHandler

    Dim startPos As Long
    Dim endPos As Long
    Dim prevBreak As Collection
    Dim nextBreak As Collection

    Set prevBreak = FindPreviousChapterBreak()
    If prevBreak Is Nothing Then
        startPos = ActiveDocument.Content.Start
    Else
        startPos = FieldStart(prevBreak("field"))
    End If

    Set nextBreak = FindNextChapterBreak()
    If nextBreak Is Nothing Then
        endPos = ActiveDocument.Content.End
    Else
        endPos = FieldStart(nextBreak("field"))
    End If

    If endPos < startPos Then endPos = startPos
    Set GetUpdateRange = ActiveDocument.Range(startPos, endPos)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.GetUpdateRange"
    Set GetUpdateRange = Nothing
End Function


' --- IsChapterBreak / IsCitationSource ---

Public Function IsChapterBreak(ByVal data As Variant) As Boolean
    If Not DictIsDictionary(data) Then Exit Function
    If DictKeyString(data, "type") <> "chapter-break" Then Exit Function
    If Not DictKeyIsDict(data, "style") Then Exit Function
    If Not IsPrefStyle(DictKeyObject(data, "style")) Then Exit Function
    If Not DictKeyIsDict(data, "content") Then Exit Function
    If Not FieldIsRichText(DictKeyObject(data, "content")) Then Exit Function

    If DictHasKey(data, "extraSource") Then
        If Not IsOptionalCitationSource(DictKeyValue(data, "extraSource")) Then Exit Function
    End If

    IsChapterBreak = True
End Function

Public Function IsCitationSource(ByVal data As Variant) As Boolean
    If Not DictIsDictionary(data) Then Exit Function
    If Not DictKeyIsCollection(data, "cites") Then Exit Function
    If Not DictKeyIsDict(data, "params") Then Exit Function
    IsCitationSource = True
End Function


' --- CanInsertChapterBreakAtSelection ---
' Silent predicate matching WPS canInsertChapterBreakAtSelection().

Public Function CanInsertChapterBreakAtSelection() As Boolean
    On Error GoTo ErrHandler

    If Selection.StoryType <> wdMainTextStory Then Exit Function
    If Selection.Range.Fields.Count > 0 Then Exit Function
    If Not IsSelectionAtParagraphStart(Selection.Range) Then Exit Function

    CanInsertChapterBreakAtSelection = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.CanInsertChapterBreakAtSelection"
    CanInsertChapterBreakAtSelection = False
End Function


' --- ReadFieldData / WriteFieldData ---

Public Function ReadFieldData(ByVal fld As Field) As Object
    On Error GoTo ErrHandler

    Dim jsonStr As String
    jsonStr = fld.Data
    If Len(jsonStr) = 0 Then Exit Function

    Dim obj As Object
    Set obj = JsonParse(jsonStr)
    If obj Is Nothing Then Exit Function

    Set ReadFieldData = obj
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.ReadFieldData"
    Set ReadFieldData = Nothing
End Function

Public Function WriteFieldData(ByVal fld As Field, ByVal data As Object) As Boolean
    On Error GoTo ErrHandler

    Dim jsonStr As String
    jsonStr = JsonStringify(data)
    If Len(jsonStr) = 0 Then Exit Function

    fld.Data = jsonStr
    WriteFieldData = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.WriteFieldData"
    WriteFieldData = False
End Function


' --- Selection helpers ---

Private Function ValidateChapterBreakInsertionPoint() As Boolean
    On Error GoTo ErrHandler

    If Selection.StoryType <> wdMainTextStory Then
        MsgBox PromptTextNotMain, vbExclamation
        Exit Function
    End If

    If Selection.Range.Fields.Count > 0 Then
        MsgBox PromptTextInField, vbExclamation
        Exit Function
    End If

    If Not IsSelectionAtParagraphStart(Selection.Range) Then
        MsgBox PromptTextNotStart, vbExclamation
        Exit Function
    End If

    ValidateChapterBreakInsertionPoint = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.ValidateChapterBreakInsertionPoint"
    MsgBox PromptTextNotMain, vbExclamation
End Function

Private Sub EnsureCaretInMainText()
    If Selection.StoryType <> wdMainTextStory Then
        On Error Resume Next
        Dim currentPage As Long
        currentPage = Selection.Information(wdActiveEndPageNumber)
        Dim pageRange As Range
        Set pageRange = Selection.GoTo(wdGoToPage, wdGoToAbsolute, currentPage)
        pageRange.Collapse wdCollapseEnd
        pageRange.Select
        On Error GoTo 0
    End If

    If Selection.Range.Fields.Count > 0 Then
        Dim chkField As Field
        Set chkField = Selection.Range.Fields(1)
        If chkField.Type = wdFieldAddin Then
            If IsChapterBreak(ReadFieldData(chkField)) Then
                Dim afterRange As Range
                Set afterRange = chkField.Result.Duplicate
                afterRange.Collapse wdCollapseEnd
                afterRange.Select
                Exit Sub
            End If
        End If
    End If

    Selection.Collapse wdCollapseEnd
End Sub

Private Function IsSelectionAtParagraphStart(ByVal rng As Range) As Boolean
    On Error GoTo ErrHandler

    Dim caret As Range
    Set caret = rng.Duplicate
    caret.Collapse wdCollapseStart

    If caret.Paragraphs.Count = 0 Then Exit Function
    IsSelectionAtParagraphStart = (caret.Start = caret.Paragraphs(1).Range.Start)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.IsSelectionAtParagraphStart"
    IsSelectionAtParagraphStart = False
End Function

Private Sub RestoreSelection(ByVal originalRange As Range)
    On Error Resume Next
    If Not originalRange Is Nothing Then originalRange.Select
    On Error GoTo 0
End Sub


' --- Field helpers ---

Private Function MakeFieldAndData(ByVal fld As Field, ByVal data As Object) As Collection
    Dim result As Collection
    Set result = New Collection
    result.Add fld, "field"
    result.Add data, "data"
    Set MakeFieldAndData = result
End Function

Private Function FindPreviousChapterBreakByPosition() As Collection
    On Error GoTo ErrHandler

    Dim fld As Field
    Dim data As Object
    Set fld = Selection.PreviousField

    Do While Not fld Is Nothing
        If IsMainTextAddInField(fld) Then
            Set data = ReadFieldData(fld)
            If IsChapterBreak(data) Then
                Set FindPreviousChapterBreakByPosition = MakeFieldAndData(fld, data)
                Exit Function
            End If
        End If
        Set fld = fld.Previous
    Loop
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.FindPreviousChapterBreakByPosition"
    Set FindPreviousChapterBreakByPosition = Nothing
End Function

Private Function FindNextChapterBreakByPosition() As Collection
    On Error GoTo ErrHandler

    Dim fld As Field
    Dim data As Object
    Set fld = Selection.NextField

    Do While Not fld Is Nothing
        If IsMainTextAddInField(fld) Then
            Set data = ReadFieldData(fld)
            If IsChapterBreak(data) Then
                Set FindNextChapterBreakByPosition = MakeFieldAndData(fld, data)
                Exit Function
            End If
        End If
        Set fld = fld.Next
    Loop
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.FindNextChapterBreakByPosition"
    Set FindNextChapterBreakByPosition = Nothing
End Function

Private Function IsMainTextAddInField(ByVal fld As Field) As Boolean
    On Error GoTo ErrHandler
    If fld.Type <> wdFieldAddin Then Exit Function
    If fld.Result.StoryType <> wdMainTextStory Then Exit Function
    IsMainTextAddInField = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.IsMainTextAddInField"
    IsMainTextAddInField = False
End Function

Private Function FieldStart(ByVal fld As Field) As Long
    On Error GoTo ErrHandler
    FieldStart = fld.Result.Start
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.FieldStart"
    FieldStart = -1
End Function

Private Function FieldEnd(ByVal fld As Field) As Long
    On Error GoTo ErrHandler
    FieldEnd = fld.Result.End
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.FieldEnd"
    FieldEnd = -1
End Function


' --- Type guard helpers ---

Private Function IsDictionaryRecord(ByVal value As Variant) As Boolean
    IsDictionaryRecord = DictIsDictionary(value)
End Function

Private Function HasDictionaryKey(ByVal dict As Object, ByVal key As String) As Boolean
    HasDictionaryKey = DictHasKey(dict, key)
End Function

Private Function IsStringMember(ByVal dict As Object, ByVal key As String) As Boolean
    IsStringMember = DictKeyIsString(dict, key)
End Function

Private Function IsPrefStyle(ByVal value As Variant) As Boolean
    If Not DictIsDictionary(value) Then Exit Function
    If Not DictKeyIsString(value, "id") Then Exit Function
    If Not DictKeyIsString(value, "title") Then Exit Function
    If Not DictKeyIsString(value, "citationType") Then Exit Function
    IsPrefStyle = True
End Function

Private Function IsCollectionObject(ByVal value As Variant) As Boolean
    IsCollectionObject = DictIsCollection(value)
End Function

Private Function IsOptionalCitationSource(ByVal value As Variant) As Boolean
    On Error GoTo ErrHandler

    If DictIsNull(value) Then
        IsOptionalCitationSource = True
        Exit Function
    End If

    If DictIsEmpty(value) Then
        IsOptionalCitationSource = True
        Exit Function
    End If

    IsOptionalCitationSource = IsCitationSource(value)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.IsOptionalCitationSource"
    IsOptionalCitationSource = False
End Function

Private Sub CopyOptionalExtraSource(ByVal target As Object, ByVal value As Variant)
    On Error GoTo ErrHandler

    If DictIsNull(value) Then Exit Sub
    If DictIsEmpty(value) Then Exit Sub
    If Not IsCitationSource(value) Then Exit Sub

    Set target("extraSource") = value
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modChapterBreak.CopyOptionalExtraSource"
End Sub


' --- Localization helpers ---

Private Function PromptText() As String
    If I10nIsLangZH() Then PromptText = CB_PROMPT_ZH Else PromptText = CB_PROMPT_EN
End Function

Private Function PromptTextNotMain() As String
    If I10nIsLangZH() Then PromptTextNotMain = CB_NOT_MAIN_ZH Else PromptTextNotMain = CB_NOT_MAIN_EN
End Function

Private Function PromptTextNotStart() As String
    If I10nIsLangZH() Then PromptTextNotStart = CB_NOT_START_ZH Else PromptTextNotStart = CB_NOT_START_EN
End Function

Private Function PromptTextInField() As String
    If I10nIsLangZH() Then PromptTextInField = CB_IN_FIELD_ZH Else PromptTextInField = CB_IN_FIELD_EN
End Function

Private Function PromptTextNoStyle() As String
    If I10nIsLangZH() Then PromptTextNoStyle = CB_NO_STYLE_ZH Else PromptTextNoStyle = CB_NO_STYLE_EN
End Function
