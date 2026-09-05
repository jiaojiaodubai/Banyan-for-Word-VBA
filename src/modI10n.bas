Option Explicit

' ============================================================================
' Module  : modI10n
' Purpose : Lightweight localization engine - pure VBA, zero dependencies.
'           Works on Windows & macOS without any external references.
'
' Internal storage: built-in Collection  (key = "langID|module.key")
'
' Pattern (mirrors WPS useI10n):
'   1. Module declares tables via  I10nRegisterTable  (ParamArray API)
'   2. Callers resolve text via     T("module.key", fallback)
'   3. Fallback: current language -> en-US -> fallback -> key itself
'
' Ref: https://learn.microsoft.com/en-us/office/vba/api/office/mso.languageid
' ============================================================================

' --- Module-level storage ---
' Flat Collection keyed by "langID|module.key" -> translated text
Private m_translations As Collection

' --- Well-known language IDs (MsoLanguageID enum) ---
Private Const LCID_EN_US As Long = 1033   ' msoLanguageIDEnglishUS
Private Const LCID_ZH_CN As Long = 2052   ' msoLanguageIDSimplifiedChinese
Private Const PRIMARY_LANG_CHINESE As Long = &H4
Private Const PRIMARY_LANG_MASK As Long = &H3FF


' --- I10nRegisterTable ---
' Register key-value pairs under a language + module namespace.
'
' langID        - MSO language ID (use built-in msoLanguageID* enums)
' moduleName    - Logical namespace, e.g. "ribbon", "settings"
' keyValuePairs - Alternating: key1, value1, key2, value2, …
'
' Example:
' I10nRegisterTable msoLanguageIDEnglishUS, "ribbon", _
' "btnCitation", "Insert/Edit Citation", _
' "btnRefresh",  "Refresh"

Public Sub I10nRegisterTable(ByVal langID As Long, _
                              ByVal moduleName As String, _
                              ParamArray keyValuePairs() As Variant)
    If m_translations Is Nothing Then
        Set m_translations = New Collection
    End If

    Dim i As Long
    Dim fullKey As String
    Dim value As String

    For i = LBound(keyValuePairs) To UBound(keyValuePairs) Step 2
        If i + 1 <= UBound(keyValuePairs) Then
            fullKey = CStr(langID) & "|" & moduleName & "." & CStr(keyValuePairs(i))
            value = CStr(keyValuePairs(i + 1))
            CollectionAddOrReplace m_translations, fullKey, value
        End If
    Next i
End Sub


' --- T - Resolve a localization key. ---
'
' key       - Dot-separated path, e.g. "ribbon.btnCitation"
' fallback  - Returned when no translation is found (default: key itself)

Public Function T(ByVal key As String, Optional ByVal fallback As String = "") As String
    If m_translations Is Nothing Then
        Set m_translations = New Collection
    End If

    Dim curLang As Long
    curLang = I10nGetCurrentLangID()

    ' 1) Try current UI language
    T = Lookup(curLang, key)
    If Len(T) > 0 Then Exit Function

    ' 2) For any Chinese UI locale, fall back to Simplified Chinese.
    If IsChineseLangID(curLang) And curLang <> LCID_ZH_CN Then
        T = Lookup(LCID_ZH_CN, key)
        If Len(T) > 0 Then Exit Function
    End If

    ' 3) Fallback to en-US (if not already the current language)
    If curLang <> LCID_EN_US Then
        T = Lookup(LCID_EN_US, key)
        If Len(T) > 0 Then Exit Function
    End If

    ' 4) Ultimate fallback
    If Len(fallback) > 0 Then
        T = fallback
    Else
        T = key
    End If
End Function


' --- Lookup - Internal: retrieve a single key from a specific language. ---

Private Function Lookup(ByVal langID As Long, ByVal key As String) As String
    Dim compositeKey As String
    compositeKey = CStr(langID) & "|" & key

    On Error Resume Next
    Lookup = CStr(m_translations(compositeKey))
    If Err.Number <> 0 Then Lookup = ""
    On Error GoTo 0
End Function


' --- CollectionAddOrReplace - Add or overwrite a keyed item in a Collection. ---

Private Sub CollectionAddOrReplace(ByRef col As Collection, _
                                    ByVal key As String, _
                                    ByVal value As String)
    On Error Resume Next
    col.Remove key     ' try to drop existing key (ignore if not found)
    On Error GoTo 0
    col.Add value, key
End Sub


' --- I10nGetCurrentLangID - Detect Word's UI language (returns MSO LCID). ---

Public Function I10nGetCurrentLangID() As Long
    On Error Resume Next
    I10nGetCurrentLangID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    If Err.Number <> 0 Or I10nGetCurrentLangID = 0 Then
        I10nGetCurrentLangID = Application.Language
    End If
    If Err.Number <> 0 Or I10nGetCurrentLangID = 0 Then
        I10nGetCurrentLangID = LCID_EN_US
    End If
    On Error GoTo 0
End Function


' --- I10nIsLangZH - Convenience: is current UI language Chinese (simplified)? ---

Public Function I10nIsLangZH() As Boolean
    I10nIsLangZH = IsChineseLangID(I10nGetCurrentLangID())
End Function

Private Function IsChineseLangID(ByVal langID As Long) As Boolean
    IsChineseLangID = ((langID And PRIMARY_LANG_MASK) = PRIMARY_LANG_CHINESE)
End Function
