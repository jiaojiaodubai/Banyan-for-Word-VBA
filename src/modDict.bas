Option Explicit

' ============================================================================
' Module  : modDict
' Purpose : Safe access helpers for Banyan dictionaries (Dictionary.cls).
'
' Background
' ----------
' Banyan data objects (citation data, citation sources, rich-text content and
' marks) are instances of Dictionary.cls (vba-tools/vba-dictionary), which is
' used for macOS compatibility. That class is declared VB_Exposed = True (an
' Automation-exposed VBA class). When such an object is held in a VARIANT and
' its default Item property result is passed straight into VBA intrinsic
' functions (IsNull / IsEmpty / IsObject / VarType), the Automation wrapper
' raises runtime error 458 ("Automation type not supported") or misbehaves.
' This is the trap behind the field rich-text rendering bug.
'
' The proven-safe pattern is to coerce the dictionary to an Object local first
' (Set d = value) and only then read properties and call intrinsics on the
' resulting local Variant. Every helper in this module does exactly that, so
' callers can read values and type-check them without reasoning about the trap.
'
' Usage notes
' -----------
' - All helpers tolerate Nothing / non-dictionary / missing-key input and
'   return safe defaults (Empty, "", 0, False, Nothing).
' - Prefer DictKey* helpers over raw `dict("key")` + IsNull/IsEmpty/IsObject/
'   VarType. Plain conversion reads (CStr/CLng/CBool on `dict("key")`) are safe
'   without this module, but the typed getters are recommended for uniformity
'   and fallback handling.
' - Object self checks (TypeName(x), x.Exists(k), IsObject(x)) are safe without
'   this module too; the DictIs* / DictKeyIs* helpers exist so the safe pattern
'   lives in one place.
' - SET vs LET for OBJECT-valued members (the second Dictionary.cls trap,
'   verified on Word 16): reading an object-valued member through an As Object
'   variable with Let (`v = dict("objKey")`, `v = DictKeyValue(dict, "objKey")`)
'   raises runtime error 450 ("Wrong number of arguments / invalid property
'   assignment"). Set works. Therefore:
'     * `DictKeyValue` is a SCALAR read (also used by the DictKeyIs* helpers,
'       which pass its result BY VALUE into DictIs* - that is safe).
'     * To obtain an object value, use `Set x = DictKeyObject(dict, key)` or
'       `Set x = dict("objKey")` (object-typed dict).
'     * To copy a member between dictionaries, use DictCopyKey (it picks Set
'       for objects and Let for scalars). Never write `v = dict("objKey")`.
'
' Public API:
'   Self checks (is `value` itself of type X?):
'     DictIsDictionary(value) / DictIsCollection(value) / DictIsObject(value)
'     DictIsString(value) / DictIsNumber(value) / DictIsBoolean(value)
'     DictIsNull(value) / DictIsEmpty(value) / DictIsScalar(value)
'   Key checks / reads (safe access into `dict`):
'     DictHasKey(dict, key)
'     DictKeyValue(dict, key)  -> Variant (Empty when absent; SCALAR use only)
'     DictKeyObject(dict, key) -> Object (Nothing when absent/not an object)
'     DictKeyString(dict, key, [fallback]) -> String
'     DictKeyLong(dict, key, [fallback])   -> Long
'     DictKeyBool(dict, key, [fallback])   -> Boolean
'     DictKeyIsString(dict, key) / DictKeyIsNumber(dict, key)
'     DictKeyIsBoolean(dict, key) / DictKeyIsScalar(dict, key)
'     DictKeyIsNull(dict, key) / DictKeyIsEmpty(dict, key)
'     DictKeyIsDict(dict, key) / DictKeyIsCollection(dict, key)
'     DictKeyIsObject(dict, key)
'   Copy (safe Set/Let member copy):
'     DictCopyKey(target, targetKey, source, sourceKey) -> Boolean
' ============================================================================


' --- Self checks ------------------------------------------------------------
' These inspect the passed `value` itself (not a key inside a dictionary).
' They are the safe replacements for the IsDictionaryRecord / IsCollectionObject
' helpers found in modField and modChapterBreak.

Public Function DictIsDictionary(ByVal value As Variant) As Boolean
    On Error GoTo ErrHandler
    If Not IsObject(value) Then Exit Function
    If value Is Nothing Then Exit Function
    DictIsDictionary = (TypeName(value) = "Dictionary")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictIsDictionary"
    DictIsDictionary = False
End Function

Public Function DictIsCollection(ByVal value As Variant) As Boolean
    On Error GoTo ErrHandler
    If Not IsObject(value) Then Exit Function
    If value Is Nothing Then Exit Function
    DictIsCollection = (TypeName(value) = "Collection")
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictIsCollection"
    DictIsCollection = False
End Function

Public Function DictIsObject(ByVal value As Variant) As Boolean
    On Error GoTo ErrHandler
    If Not IsObject(value) Then Exit Function
    If value Is Nothing Then Exit Function
    DictIsObject = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictIsObject"
    DictIsObject = False
End Function

Public Function DictIsString(ByVal value As Variant) As Boolean
    On Error GoTo ErrHandler
    If IsObject(value) Then Exit Function
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    DictIsString = (VarType(value) = vbString)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictIsString"
    DictIsString = False
End Function

Public Function DictIsNumber(ByVal value As Variant) As Boolean
    On Error GoTo ErrHandler
    If IsObject(value) Then Exit Function
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    Select Case VarType(value)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            DictIsNumber = True
    End Select
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictIsNumber"
    DictIsNumber = False
End Function

Public Function DictIsBoolean(ByVal value As Variant) As Boolean
    On Error GoTo ErrHandler
    If IsObject(value) Then Exit Function
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    DictIsBoolean = (VarType(value) = vbBoolean)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictIsBoolean"
    DictIsBoolean = False
End Function

Public Function DictIsNull(ByVal value As Variant) As Boolean
    On Error GoTo ErrHandler
    DictIsNull = IsNull(value)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictIsNull"
    DictIsNull = False
End Function

Public Function DictIsEmpty(ByVal value As Variant) As Boolean
    On Error GoTo ErrHandler
    DictIsEmpty = IsEmpty(value)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictIsEmpty"
    DictIsEmpty = False
End Function

Public Function DictIsScalar(ByVal value As Variant) As Boolean
    On Error GoTo ErrHandler
    If IsObject(value) Then Exit Function
    If IsNull(value) Or IsEmpty(value) Then Exit Function
    DictIsScalar = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictIsScalar"
    DictIsScalar = False
End Function


' --- Key checks and reads ---------------------------------------------------
' Every function coerces the dictionary to an Object local first (AsDict), then
' reads properties through that Object reference. Values read this way are safe
' to feed to VBA intrinsics.

Public Function DictHasKey(ByVal dict As Variant, ByVal key As String) As Boolean
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = AsDict(dict)
    If d Is Nothing Then Exit Function
    DictHasKey = d.Exists(key)
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictHasKey"
    DictHasKey = False
End Function

Public Function DictKeyValue(ByVal dict As Variant, ByVal key As String) As Variant
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = AsDict(dict)
    If d Is Nothing Then Exit Function
    If d.Exists(key) Then
        If IsObject(d(key)) Then
            Set DictKeyValue = d(key)
        Else
            DictKeyValue = d(key)
        End If
    End If
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictKeyValue"
    ' leave Empty
End Function

Public Function DictKeyObject(ByVal dict As Variant, ByVal key As String) As Object
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = AsDict(dict)
    If d Is Nothing Then Exit Function
    If d.Exists(key) Then
        If IsObject(d(key)) Then
            Set DictKeyObject = d(key)
        End If
    End If
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictKeyObject"
    Set DictKeyObject = Nothing
End Function

Public Function DictKeyString(ByVal dict As Variant, ByVal key As String, _
                              Optional ByVal fallback As String = "") As String
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = AsDict(dict)
    If d Is Nothing Then GoTo UseFallback
    If Not d.Exists(key) Then GoTo UseFallback
    Dim v As Variant
    v = d(key)
    If IsNull(v) Or IsEmpty(v) Then GoTo UseFallback
    If IsObject(v) Then GoTo UseFallback
    DictKeyString = CStr(v)
    Exit Function

UseFallback:
    DictKeyString = fallback
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictKeyString"
    DictKeyString = fallback
End Function

Public Function DictKeyLong(ByVal dict As Variant, ByVal key As String, _
                            Optional ByVal fallback As Long = 0) As Long
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = AsDict(dict)
    If d Is Nothing Then GoTo UseFallback
    If Not d.Exists(key) Then GoTo UseFallback
    Dim v As Variant
    v = d(key)
    If IsNull(v) Or IsEmpty(v) Then GoTo UseFallback
    If IsObject(v) Then GoTo UseFallback
    If IsNumeric(v) Then
        DictKeyLong = CLng(v)
    Else
        GoTo UseFallback
    End If
    Exit Function

UseFallback:
    DictKeyLong = fallback
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictKeyLong"
    DictKeyLong = fallback
End Function

Public Function DictKeyBool(ByVal dict As Variant, ByVal key As String, _
                            Optional ByVal fallback As Boolean = False) As Boolean
    On Error GoTo ErrHandler
    Dim d As Object
    Set d = AsDict(dict)
    If d Is Nothing Then GoTo UseFallback
    If Not d.Exists(key) Then GoTo UseFallback
    Dim v As Variant
    v = d(key)
    If IsNull(v) Or IsEmpty(v) Then GoTo UseFallback
    If IsObject(v) Then GoTo UseFallback
    If VarType(v) = vbBoolean Then
        DictKeyBool = CBool(v)
    ElseIf IsNumeric(v) Then
        DictKeyBool = CBool(v)
    Else
        GoTo UseFallback
    End If
    Exit Function

UseFallback:
    DictKeyBool = fallback
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictKeyBool"
    DictKeyBool = fallback
End Function

Public Function DictKeyIsString(ByVal dict As Variant, ByVal key As String) As Boolean
    DictKeyIsString = DictIsString(DictKeyValue(dict, key))
End Function

Public Function DictKeyIsNumber(ByVal dict As Variant, ByVal key As String) As Boolean
    DictKeyIsNumber = DictIsNumber(DictKeyValue(dict, key))
End Function

Public Function DictKeyIsBoolean(ByVal dict As Variant, ByVal key As String) As Boolean
    DictKeyIsBoolean = DictIsBoolean(DictKeyValue(dict, key))
End Function

Public Function DictKeyIsScalar(ByVal dict As Variant, ByVal key As String) As Boolean
    DictKeyIsScalar = DictIsScalar(DictKeyValue(dict, key))
End Function

Public Function DictKeyIsNull(ByVal dict As Variant, ByVal key As String) As Boolean
    DictKeyIsNull = DictIsNull(DictKeyValue(dict, key))
End Function

Public Function DictKeyIsEmpty(ByVal dict As Variant, ByVal key As String) As Boolean
    DictKeyIsEmpty = DictIsEmpty(DictKeyValue(dict, key))
End Function

Public Function DictKeyIsDict(ByVal dict As Variant, ByVal key As String) As Boolean
    DictKeyIsDict = DictIsDictionary(DictKeyValue(dict, key))
End Function

Public Function DictKeyIsCollection(ByVal dict As Variant, ByVal key As String) As Boolean
    DictKeyIsCollection = DictIsCollection(DictKeyValue(dict, key))
End Function

Public Function DictKeyIsObject(ByVal dict As Variant, ByVal key As String) As Boolean
    DictKeyIsObject = DictIsObject(DictKeyValue(dict, key))
End Function

Public Function DictCopyKey(ByVal target As Object, _
                            ByVal targetKey As String, _
                            ByVal source As Object, _
                            ByVal sourceKey As String) As Boolean
    ' Copy one member from `source` into `target` safely.
    ' Dictionary.cls trap (verified on Word 16): Let-assignment of an
    ' OBJECT-valued member through an As Object variable (v = dict("key"))
    ' raises error 450; Set works, and IsObject(dict("key")) is safe. So
    ' objects are copied with Set, scalars with Let. This is the single
    ' canonical safe way to copy members between dictionaries.
    On Error GoTo ErrHandler
    If target Is Nothing Then Exit Function
    If source Is Nothing Then Exit Function
    If Not source.Exists(sourceKey) Then Exit Function

    If IsObject(source(sourceKey)) Then
        Set target(targetKey) = source(sourceKey)
    Else
        target(targetKey) = source(sourceKey)
    End If
    DictCopyKey = True
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modDict.DictCopyKey"
    DictCopyKey = False
End Function


' --- Internal ---------------------------------------------------------------

Private Function AsDict(ByVal value As Variant) As Object
    ' Coerce a Variant (possibly holding a Dictionary.cls object) to an Object
    ' local. Returns Nothing when value is not an object. All subsequent reads
    ' must go through this Object reference to avoid error 458.
    On Error Resume Next
    Set AsDict = value
    On Error GoTo 0
End Function
