Option Explicit

' ============================================================================
' Module  : testDict
' Purpose : Tests for modDict (safe dictionary wrapper).
'
'           Part of the dev-only test suite under test/ (module name test*).
'           Run the FULL suite via test/Run-Tests.ps1; modTest.RunAllTests
'           aggregates a concise subset (see modTest).
'
' Public API:
'   RunTests() -> String report of [PASS]/[FAIL] lines (no dialogs)
' ============================================================================


Public Function RunTests() As String
    Dim report As String
    report = "testDict" & vbCrLf & String(40, "-") & vbCrLf
    report = report & TestResult("self checks", TestDictSelfChecks()) & vbCrLf
    report = report & TestResult("key access", TestDictKeyAccess()) & vbCrLf
    report = report & TestResult("safe via Variant", TestDictVariantSafety()) & vbCrLf
    report = report & TestResult("object values via Set", TestDictObjectValues()) & vbCrLf
    report = report & TestResult("copy key (Set/Let)", TestDictCopyKey()) & vbCrLf
    report = report & TestResult("ByVal Variant read coerces to Object", TestDictByValVariantRead()) & vbCrLf
    RunTests = report
End Function


Private Function TestDictSelfChecks() As Boolean
    On Error GoTo ErrHandler
    Dim ok As Boolean
    ok = True

    Dim d As Object
    Set d = New Dictionary
    ok = ok And DictIsDictionary(d)
    ok = ok And (Not DictIsCollection(d))
    ok = ok And DictIsObject(d)

    Dim c As Collection
    Set c = New Collection
    ok = ok And DictIsCollection(c)
    ok = ok And (Not DictIsDictionary(c))

    ok = ok And DictIsString("abc")
    ok = ok And (Not DictIsString(123))
    ok = ok And (Not DictIsString(d))
    ok = ok And DictIsNumber(42)
    ok = ok And DictIsNumber(3.14)
    ok = ok And (Not DictIsNumber("x"))
    ok = ok And DictIsBoolean(True)
    ok = ok And (Not DictIsBoolean(1))
    ok = ok And DictIsNull(Null)
    ok = ok And (Not DictIsNull(0))
    ok = ok And DictIsEmpty(Empty)
    ok = ok And (Not DictIsEmpty(0))
    ok = ok And DictIsScalar("x")
    ok = ok And DictIsScalar(42)
    ok = ok And (Not DictIsScalar(d))
    ok = ok And (Not DictIsScalar(Null))
    ok = ok And (Not DictIsDictionary(Nothing))
    ok = ok And (Not DictIsDictionary(42))

    TestDictSelfChecks = ok
    Exit Function

ErrHandler:
    TestDictSelfChecks = False
End Function

Private Function TestDictKeyAccess() As Boolean
    On Error GoTo ErrHandler
    Dim ok As Boolean
    ok = True

    Dim d As Object
    Set d = New Dictionary
    d("name") = "Zhang"
    d("count") = 3
    d("active") = True
    d("nothing") = Null
    Set d("nested") = New Dictionary
    Set d("items") = New Collection

    ok = ok And DictHasKey(d, "name")
    ok = ok And (Not DictHasKey(d, "missing"))
    ok = ok And (DictKeyString(d, "name") = "Zhang")
    ok = ok And (DictKeyString(d, "missing") = "")
    ok = ok And (DictKeyString(d, "missing", "fb") = "fb")
    ok = ok And (DictKeyLong(d, "count") = 3)
    ok = ok And (DictKeyLong(d, "missing") = 0)
    ok = ok And (DictKeyLong(d, "missing", 7) = 7)
    ok = ok And (DictKeyBool(d, "active") = True)
    ok = ok And DictKeyIsString(d, "name")
    ok = ok And (Not DictKeyIsString(d, "count"))
    ok = ok And DictKeyIsNumber(d, "count")
    ok = ok And DictKeyIsBoolean(d, "active")
    ok = ok And DictKeyIsNull(d, "nothing")
    ok = ok And DictKeyIsEmpty(d, "missing")
    ok = ok And DictKeyIsScalar(d, "name")
    ok = ok And (Not DictKeyIsScalar(d, "nested"))
    ok = ok And DictKeyIsDict(d, "nested")
    ok = ok And DictKeyIsCollection(d, "items")
    ok = ok And DictKeyIsObject(d, "nested")
    ok = ok And (Not DictKeyIsObject(d, "name"))
    ok = ok And DictIsDictionary(DictKeyObject(d, "nested"))
    ok = ok And DictIsCollection(DictKeyValue(d, "items"))

    TestDictKeyAccess = ok
    Exit Function

ErrHandler:
    TestDictKeyAccess = False
End Function

Private Function TestDictVariantSafety() As Boolean
    ' Regression: hold a Dictionary.cls instance in a VARIANT (the error-458
    ' trap scenario) and confirm modDict still reads it safely.
    On Error GoTo ErrHandler
    Dim ok As Boolean
    ok = True

    Dim v As Variant
    Set v = New Dictionary
    v("title") = "(Zhang, 2020)"
    v("page") = 12
    v("bold") = True

    ok = ok And (DictKeyString(v, "title") = "(Zhang, 2020)")
    ok = ok And (DictKeyLong(v, "page") = 12)
    ok = ok And (DictKeyBool(v, "bold") = True)
    ok = ok And DictKeyIsString(v, "title")
    ok = ok And DictKeyIsNumber(v, "page")
    ok = ok And DictIsDictionary(v)

    ' Absent / non-object input must degrade to safe defaults, not errors.
    Dim nothingDict As Variant
    nothingDict = Empty
    ok = ok And (Not DictHasKey(nothingDict, "x"))
    ok = ok And (DictKeyString(nothingDict, "x") = "")

    TestDictVariantSafety = ok
    Exit Function

ErrHandler:
    TestDictVariantSafety = False
End Function

Private Function TestDictObjectValues() As Boolean
    ' Regression (2026-08-22): reading an OBJECT-valued member through an
    ' As Object variable must use Set. Let-assignment (v = dict("objKey"))
    ' raises error 450 with Dictionary.cls. modDict.DictKeyObject + Set is the
    ' safe pattern (used by modRefresh.CopyDictionaryValue for cites/params).
    On Error GoTo ErrHandler
    Dim ok As Boolean
    ok = True

    Dim d As Object
    Set d = New Dictionary
    Set d("cites") = New Collection
    d("cites").Add "a"
    Set d("params") = New Dictionary
    d("params")("x") = 1

    ' Safe: read object values via DictKeyObject (Set) from an Object-typed dict.
    Dim col As Object
    Set col = DictKeyObject(d, "cites")
    ok = ok And (Not col Is Nothing)
    ok = ok And DictIsCollection(col)
    ok = ok And (col.Count = 1)

    Dim nested As Object
    Set nested = DictKeyObject(d, "params")
    ok = ok And (Not nested Is Nothing)
    ok = ok And DictIsDictionary(nested)

    ' Key-type helpers on object-valued keys must return False, not crash.
    ok = ok And (DictKeyIsString(d, "cites") = False)
    ok = ok And (DictKeyIsNumber(d, "cites") = False)
    ok = ok And DictKeyIsCollection(d, "cites")
    ok = ok And DictKeyIsDict(d, "params")

    ' Direct Set read through an As Object variable (CopyDictionaryValue pattern).
    Dim col2 As Collection
    Set col2 = d("cites")
    ok = ok And (col2.Count = 1)

    TestDictObjectValues = ok
    Exit Function

ErrHandler:
    TestDictObjectValues = False
End Function

Private Function TestDictCopyKey() As Boolean
    ' Regression (2026-08-22): DictCopyKey is the canonical safe member copy
    ' (Set for object values, Let for scalars). This is what modRefresh uses to
    ' build the refresh request context from the citation source.
    On Error GoTo ErrHandler
    Dim ok As Boolean
    ok = True

    Dim src As Object
    Set src = New Dictionary
    Set src("cites") = New Collection
    src("cites").Add "x"
    src("title") = "T"
    src("count") = 2

    Dim dst As Object
    Set dst = New Dictionary
    ok = ok And DictCopyKey(dst, "cites", src, "cites")   ' object value via Set
    ok = ok And DictCopyKey(dst, "title", src, "title")   ' scalar via Let
    ok = ok And DictCopyKey(dst, "count", src, "count")   ' scalar via Let
    ok = ok And (Not DictCopyKey(dst, "missing", src, "nope"))

    ok = ok And DictKeyIsCollection(dst, "cites")
    Dim col As Collection
    Set col = dst("cites")
    ok = ok And (col.Count = 1)
    ok = ok And (DictKeyString(dst, "title") = "T")
    ok = ok And (DictKeyLong(dst, "count") = 2)

    TestDictCopyKey = ok
    Exit Function

ErrHandler:
    TestDictCopyKey = False
End Function


Private Function TestDictByValVariantRead() As Boolean
    ' Regression (2026-08-23): reads through a ByVal Variant parameter must go
    ' through modDict (see README "Dictionary safety") so the style id is always
    ' extracted - this is what CloneStyle/CloneStyleForRefresh now do.
    On Error GoTo ErrHandler
    Dim ok As Boolean
    ok = True

    Dim style As Object
    Set style = New Dictionary
    style("id") = "the-journal-of-international-studies"
    style("title") = "国际政治研究"
    style("citationType") = "note-citation"

    ' Fixed pattern: ByVal Variant param read via DictKeyString (as the real
    ' CloneStyle / CloneStyleForRefresh do after the root-cause fix).
    Dim fixed As Object
    Set fixed = ReadStyleFixed(style)
    ok = ok And (CStr(fixed("id")) = "the-journal-of-international-studies")
    ok = ok And (CStr(fixed("title")) = "国际政治研究")
    ok = ok And (CStr(fixed("citationType")) = "note-citation")

    ' Missing / non-object input must degrade to "" without crashing.
    Dim missing As Object
    Set missing = New Dictionary
    missing("title") = "国际政治研究"
    Dim fixed2 As Object
    Set fixed2 = ReadStyleFixed(missing)
    ok = ok And (Len(CStr(fixed2("id"))) = 0)
    Dim fixed3 As Object
    Set fixed3 = ReadStyleFixed(Empty)
    ok = ok And (Len(CStr(fixed3("id"))) = 0)

    TestDictByValVariantRead = ok
    Exit Function

ErrHandler:
    TestDictByValVariantRead = False
End Function

' Same read pattern as frmSettings.CloneStyle (via modDict).
Private Function ReadStyleFixed(ByVal source As Variant) As Object
    Dim result As Object
    Set result = New Dictionary
    result("id") = DictKeyString(source, "id")
    result("title") = DictKeyString(source, "title")
    result("citationType") = DictKeyString(source, "citationType")
    Set ReadStyleFixed = result
End Function


Private Function TestResult(ByVal name As String, ByVal passed As Boolean) As String
    If passed Then
        TestResult = "[PASS] " & name
    Else
        TestResult = "[FAIL] " & name
    End If
End Function
