Option Explicit

' ============================================================================
' Module  : modFinalize
' Purpose : Finalize a Banyan document.
'
'           Mirrors WPS moulds/finalize.ts:
'             - ask for confirmation
'             - create a backup in the same folder
'             - refresh Banyan citations before removing chapter metadata
'             - walk fields in reverse order
'             - unlink citation/bibliography ADDIN fields
'             - delete ChapterBreak ADDIN fields
'
'           Finalize is intentionally irreversible for the active document.
' ============================================================================

Private m_i18nReady As Boolean


' --- Public entry point ---

Public Sub FinalizeAction()
    EnsureFinalizeI10n

    Dim answer As VbMsgBoxResult
    answer = MsgBox(FText("confirmBody", _
                          "Finalize will replace Banyan fields with their rendered results and this action cannot be undone."), _
                    vbQuestion + vbYesNo + vbDefaultButton2, _
                    FText("confirmTitle", "Finalize Confirmation"))
    If answer <> vbYes Then Exit Sub

    Dim backupPath As String
    Dim backupError As String

    On Error GoTo ErrHandler
    If Not TryBackupCurrentDocument(backupPath, backupError) Then
        If Not ConfirmContinueWithoutAutoBackup(backupError) Then Exit Sub
        backupPath = FText("manualBackupRequired", "No automatic backup was created. Please keep your manual backup.")
    End If

    ProgressOpen FText("progressReason", "Finalizing document...")

    RefreshBeforeFinalize

    Dim stats As Object
    Set stats = UnlinkBanyanFields()

    ProgressClose
    If DictKeyLong(stats, "failed") > 0 Then
        MsgBox FormatMessage(FText("partial", _
                                   "Finalize partially completed." & vbCrLf & vbCrLf & _
                                   "Backup: {backupPath}" & vbCrLf & _
                                   "Fields processed: {success}" & vbCrLf & _
                                   "Failed: {failed}"), _
                             backupPath, DictKeyLong(stats, "success"), DictKeyLong(stats, "failed")), _
               vbExclamation, FText("confirmTitle", "Finalize Confirmation")
    Else
        MsgBox FormatMessage(FText("success", _
                                   "Finalize completed." & vbCrLf & vbCrLf & _
                                   "Backup: {backupPath}" & vbCrLf & _
                                   "Fields processed: {success}" & vbCrLf & _
                                   "Failed: {failed}"), _
                             backupPath, DictKeyLong(stats, "success"), DictKeyLong(stats, "failed")), _
               vbInformation, FText("confirmTitle", "Finalize Confirmation")
    End If
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modFinalize.FinalizeAction"
    Dim errMessage As String
    errMessage = Err.Description

    ProgressClose
    MsgBox Replace(FText("failed", "Finalize failed: {message}"), "{message}", errMessage), _
           vbExclamation, FText("confirmTitle", "Finalize Confirmation")
End Sub


' --- Backup ---

Private Function TryBackupCurrentDocument(ByRef backupPath As String, _
                                          ByRef failureMessage As String) As Boolean
    backupPath = ""
    failureMessage = ""

    Dim sourcePath As String
    sourcePath = CurrentDocumentFullPath()
    If Len(sourcePath) = 0 Then
        failureMessage = FText("noFilePath", "The current document is not saved. Please save it first so a backup can be created.")
        Exit Function
    End If

    On Error GoTo BackupErr
    backupPath = CreateBackupPath(sourcePath)

    On Error GoTo SaveCopyErr
    ActiveDocument.SaveCopyAs backupPath
    TryBackupCurrentDocument = True
    Exit Function

SaveCopyErr:
    Dim saveCopyMessage As String
    saveCopyMessage = Err.Description

    On Error GoTo CopyErr
    If FileExists(backupPath) Then backupPath = CreateBackupPath(sourcePath)
    If Not ActiveDocument.Saved Then
        ActiveDocument.Save
    End If
    CopyDocumentFile sourcePath, backupPath
    TryBackupCurrentDocument = True
    Exit Function

CopyErr:
    failureMessage = BuildBackupFailureMessage(saveCopyMessage, Err.Description)
    backupPath = ""
    Exit Function

BackupErr:
    failureMessage = Err.Description
    backupPath = ""
End Function

Private Sub CopyDocumentFile(ByVal sourcePath As String, ByVal backupPath As String)
    FileCopy sourcePath, backupPath
End Sub

Private Function BuildBackupFailureMessage(ByVal saveCopyMessage As String, _
                                           ByVal copyMessage As String) As String
    If Len(saveCopyMessage) > 0 And Len(copyMessage) > 0 Then
        BuildBackupFailureMessage = FText("backupFailedDetails", _
                                          "SaveCopyAs failed: {saveCopyMessage}; file copy failed: {copyMessage}")
        BuildBackupFailureMessage = Replace(BuildBackupFailureMessage, "{saveCopyMessage}", saveCopyMessage)
        BuildBackupFailureMessage = Replace(BuildBackupFailureMessage, "{copyMessage}", copyMessage)
    ElseIf Len(copyMessage) > 0 Then
        BuildBackupFailureMessage = copyMessage
    Else
        BuildBackupFailureMessage = saveCopyMessage
    End If
End Function

Private Function ConfirmContinueWithoutAutoBackup(ByVal failureMessage As String) As Boolean
    Dim prompt As String
    prompt = FText("manualBackupPrompt", _
                   "Unable to create an automatic backup: {message}" & vbCrLf & vbCrLf & _
                   "Please create a manual backup before continuing." & vbCrLf & vbCrLf & _
                   "Continue finalize without an automatic backup?")
    prompt = Replace(prompt, "{message}", failureMessage)

    ConfirmContinueWithoutAutoBackup = (MsgBox(prompt, _
                                               vbExclamation + vbYesNo + vbDefaultButton2, _
                                               FText("confirmTitle", "Finalize Confirmation")) = vbYes)
End Function

Private Function CurrentDocumentFullPath() As String
    On Error GoTo ErrHandler

    Dim fullName As String
    fullName = Trim$(ActiveDocument.FullName)
    If HasPathSeparator(fullName) Then
        CurrentDocumentFullPath = fullName
        Exit Function
    End If

    Dim dirPath As String
    Dim fileName As String
    dirPath = Trim$(ActiveDocument.Path)
    fileName = Trim$(ActiveDocument.Name)
    If Len(dirPath) = 0 Or Len(fileName) = 0 Then Exit Function

    CurrentDocumentFullPath = JoinPath(dirPath, fileName)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modFinalize.CurrentDocumentFullPath"
    CurrentDocumentFullPath = ""
End Function

Private Function CreateBackupPath(ByVal sourcePath As String) As String
    Dim dirPath As String
    Dim fileName As String
    SplitPath sourcePath, dirPath, fileName
    If Len(dirPath) = 0 Or Len(fileName) = 0 Then Err.Raise vbObjectError + 4301, "Finalize", FText("noFilePath", "The current document is not saved. Please save it first so a backup can be created.")

    Dim baseName As String
    Dim ext As String
    SplitFileName fileName, baseName, ext

    Dim candidate As String
    Dim i As Long
    Dim stamp As String
    stamp = Format$(Now, "mm-dd hh-nn-ss")

    candidate = JoinPath(dirPath, baseName & "-" & stamp & ext)
    If Not FileExists(candidate) Then
        CreateBackupPath = candidate
        Exit Function
    End If

    For i = 2 To 99
        candidate = JoinPath(dirPath, baseName & "-" & stamp & "-" & CStr(i) & ext)
        If Not FileExists(candidate) Then
            CreateBackupPath = candidate
            Exit Function
        End If
    Next i

    Err.Raise vbObjectError + 4303, "Finalize", FText("backupNameUnavailable", "Unable to create backup: no available backup file name.")
End Function

Private Sub SplitPath(ByVal fullPath As String, ByRef dirPath As String, ByRef fileName As String)
    Dim slashIndex As Long
    slashIndex = LastPathSeparatorIndex(fullPath)
    If slashIndex <= 0 Or slashIndex >= Len(fullPath) Then Exit Sub

    dirPath = Left$(fullPath, slashIndex - 1)
    fileName = Mid$(fullPath, slashIndex + 1)
End Sub

Private Sub SplitFileName(ByVal fileName As String, ByRef baseName As String, ByRef ext As String)
    Dim dotIndex As Long
    dotIndex = InStrRev(fileName, ".")
    If dotIndex <= 1 Or dotIndex = Len(fileName) Then
        baseName = fileName
        ext = ""
    Else
        baseName = Left$(fileName, dotIndex - 1)
        ext = Mid$(fileName, dotIndex)
    End If
End Sub

Private Function JoinPath(ByVal dirPath As String, ByVal fileName As String) As String
    Dim sep As String
    sep = "\"
    If InStr(1, dirPath, "/", vbBinaryCompare) > 0 Then sep = "/"

    If Right$(dirPath, 1) = "\" Or Right$(dirPath, 1) = "/" Then
        JoinPath = dirPath & fileName
    Else
        JoinPath = dirPath & sep & fileName
    End If
End Function

Private Function HasPathSeparator(ByVal pathValue As String) As Boolean
    HasPathSeparator = (InStr(1, pathValue, "\", vbBinaryCompare) > 0 Or _
                        InStr(1, pathValue, "/", vbBinaryCompare) > 0)
End Function

Private Function LastPathSeparatorIndex(ByVal pathValue As String) As Long
    Dim backslashIndex As Long
    Dim slashIndex As Long
    backslashIndex = InStrRev(pathValue, "\")
    slashIndex = InStrRev(pathValue, "/")
    If backslashIndex > slashIndex Then
        LastPathSeparatorIndex = backslashIndex
    Else
        LastPathSeparatorIndex = slashIndex
    End If
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(filePath, vbNormal + vbHidden + vbSystem)) > 0)
    On Error GoTo 0
End Function


' --- Refresh and unlink ---

Private Sub RefreshBeforeFinalize()
    On Error Resume Next
    RefreshAll
    On Error GoTo 0
End Sub

Private Function UnlinkBanyanFields() As Object
    Dim stats As Object
    Set stats = New Dictionary
    stats("total") = 0
    stats("success") = 0
    stats("failed") = 0

    UnlinkMainTextFields stats
    UnlinkFootnoteFields stats

    Set UnlinkBanyanFields = stats
End Function

Private Sub UnlinkMainTextFields(ByVal stats As Object)
    On Error GoTo ErrHandler

    Dim fields As Fields
    Set fields = ActiveDocument.Content.Fields

    Dim i As Long
    For i = fields.Count To 1 Step -1
        FinalizeOneField fields(i), stats
    Next i
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modFinalize.UnlinkMainTextFields"
End Sub

Private Sub UnlinkFootnoteFields(ByVal stats As Object)
    On Error GoTo ErrHandler

    Dim i As Long
    Dim j As Long
    Dim note As Footnote
    Dim fields As Fields

    For i = ActiveDocument.Footnotes.Count To 1 Step -1
        Set note = ActiveDocument.Footnotes(i)
        Set fields = note.Range.Fields
        For j = fields.Count To 1 Step -1
            FinalizeOneField fields(j), stats
        Next j
    Next i
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modFinalize.UnlinkFootnoteFields"
End Sub

Private Sub FinalizeOneField(ByVal fld As Field, ByVal stats As Object)
    Dim action As String
    action = GetFinalizeAction(fld)
    If Len(action) = 0 Then Exit Sub

    Dim data As Object
    Dim resultRange As Range
    If action = "unlink" Then
        Set data = FieldReadData(fld)
        Set resultRange = fld.Result.Duplicate
    End If

    stats("total") = DictKeyLong(stats, "total") + 1
    If ApplyFinalizeAction(fld, action, data, resultRange) Then
        stats("success") = DictKeyLong(stats, "success") + 1
    Else
        stats("failed") = DictKeyLong(stats, "failed") + 1
    End If
End Sub

Private Function GetFinalizeAction(ByVal fld As Field) As String
    On Error GoTo ErrHandler
    If fld.Type <> wdFieldAddin Then Exit Function

    Dim data As Object
    Set data = FieldReadData(fld)
    If data Is Nothing Then Exit Function

    If FieldIsIntextCitation(data) Or _
       FieldIsNoteCitation(data) Or _
       FieldIsBibliographyTitle(data) Or _
       FieldIsBibliographyEntry(data) Then
        GetFinalizeAction = "unlink"
    ElseIf IsChapterBreak(data) Then
        GetFinalizeAction = "delete"
    End If
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modFinalize.GetFinalizeAction"
    GetFinalizeAction = ""
End Function

Private Function ApplyFinalizeAction(ByVal fld As Field, _
                                     ByVal action As String, _
                                     Optional ByVal data As Object, _
                                     Optional ByVal resultRange As Range) As Boolean
    On Error GoTo ErrHandler

    If fld.Locked Then fld.Locked = False

    If action = "unlink" Then
        fld.Unlink
        RestoreFinalizeResultLinks resultRange, data
    ElseIf action = "delete" Then
        fld.Delete
    Else
        Exit Function
    End If

    ApplyFinalizeAction = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modFinalize.ApplyFinalizeAction"
    ApplyFinalizeAction = False
End Function

Private Sub RestoreFinalizeResultLinks(ByVal resultRange As Range, ByVal data As Object)
    On Error GoTo ErrHandler
    If resultRange Is Nothing Then Exit Sub
    If data Is Nothing Then Exit Sub
    If Not FinalizeDataHasContent(data) Then Exit Sub

    FieldApplyRichTextLinksToRange resultRange, data("content")
    Exit Sub

ErrHandler:
    DiagnosticsReraiseIfDev "modFinalize.RestoreFinalizeResultLinks"
End Sub

Private Function FinalizeDataHasContent(ByVal data As Object) As Boolean
    On Error GoTo ErrHandler
    If data Is Nothing Then Exit Function
    If Not data.Exists("content") Then Exit Function
    If Not FieldIsRichText(data("content")) Then Exit Function
    FinalizeDataHasContent = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modFinalize.FinalizeDataHasContent"
    FinalizeDataHasContent = False
End Function


' --- Local i18n ---

Private Sub EnsureFinalizeI10n()
    If m_i18nReady Then Exit Sub

    I10nRegisterTable msoLanguageIDSimplifiedChinese, "finalize", _
        "confirmTitle", "定稿确认", _
        "confirmBody", "执行“定稿”会先刷新文档，然后将 Banyan 域代码替换为当前显示结果，且不可撤销。" & vbCrLf & vbCrLf & "系统会先自动备份当前文档，再继续执行。" & vbCrLf & vbCrLf & "是否继续？", _
        "progressReason", "正在执行文档定稿...", _
        "noFilePath", "当前文档尚未保存，无法在原路径创建备份。请先保存文档后再执行定稿。", _
        "copyFailed", "创建备份失败：{message}", _
        "backupNameUnavailable", "创建备份失败：没有可用的备份文件名。", _
        "backupFailedDetails", "SaveCopyAs 失败：{saveCopyMessage}；文件复制失败：{copyMessage}", _
        "manualBackupPrompt", "无法自动创建备份：{message}" & vbCrLf & vbCrLf & "请先手动另存一份备份，确认已备份后再继续。" & vbCrLf & vbCrLf & "是否在没有自动备份文件的情况下继续定稿？", _
        "manualBackupRequired", "未创建自动备份，请确认已手动备份。", _
        "success", "定稿完成。" & vbCrLf & vbCrLf & "备份文件：{backupPath}" & vbCrLf & "已处理域：{success}" & vbCrLf & "处理失败：{failed}", _
        "failed", "定稿失败：{message}", _
        "partial", "定稿部分完成。" & vbCrLf & vbCrLf & "备份文件：{backupPath}" & vbCrLf & "已处理域：{success}" & vbCrLf & "处理失败：{failed}"

    I10nRegisterTable msoLanguageIDEnglishUS, "finalize", _
        "confirmTitle", "Finalize Confirmation", _
        "confirmBody", "Finalize will refresh the document, then replace Banyan fields with their rendered results. This action cannot be undone." & vbCrLf & vbCrLf & "A backup will be created in the same folder before processing." & vbCrLf & vbCrLf & "Do you want to continue?", _
        "progressReason", "Finalizing document...", _
        "noFilePath", "The current document is not saved. Please save it first so a backup can be created.", _
        "copyFailed", "Unable to create backup: {message}", _
        "backupNameUnavailable", "Unable to create backup: no available backup file name.", _
        "backupFailedDetails", "SaveCopyAs failed: {saveCopyMessage}; file copy failed: {copyMessage}", _
        "manualBackupPrompt", "Unable to create an automatic backup: {message}" & vbCrLf & vbCrLf & "Please create a manual backup before continuing." & vbCrLf & vbCrLf & "Continue finalize without an automatic backup?", _
        "manualBackupRequired", "No automatic backup was created. Please keep your manual backup.", _
        "success", "Finalize completed." & vbCrLf & vbCrLf & "Backup: {backupPath}" & vbCrLf & "Fields processed: {success}" & vbCrLf & "Failed: {failed}", _
        "failed", "Finalize failed: {message}", _
        "partial", "Finalize partially completed." & vbCrLf & vbCrLf & "Backup: {backupPath}" & vbCrLf & "Fields processed: {success}" & vbCrLf & "Failed: {failed}"

    m_i18nReady = True
End Sub

Private Function FText(ByVal key As String, ByVal fallback As String) As String
    EnsureFinalizeI10n
    FText = T("finalize." & key, fallback)
End Function

Private Function FormatMessage(ByVal template As String, _
                               ByVal backupPath As String, _
                               ByVal successCount As Long, _
                               ByVal failedCount As Long) As String
    FormatMessage = template
    FormatMessage = Replace(FormatMessage, "{backupPath}", backupPath)
    FormatMessage = Replace(FormatMessage, "{success}", CStr(successCount))
    FormatMessage = Replace(FormatMessage, "{failed}", CStr(failedCount))
End Function