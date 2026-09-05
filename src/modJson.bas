Option Explicit

' ============================================================================
' Module  : modJson
' Purpose : Thin wrapper around VBA-JSON's JsonConverter.
'           https://github.com/VBA-tools/VBA-JSON
'
' Prerequisite: Import JsonConverter.bas before this module.
'
' Public API:
'   JsonParse(jsonString)       -> Variant
'   JsonStringify(value)        -> String
'   JsonStringifyPretty(value)  -> String
' ============================================================================

Public Function JsonParse(ByVal jsonString As String) As Variant
    On Error Resume Next
    Set JsonParse = JsonConverter.ParseJson(jsonString)
    If Err.Number <> 0 Then Set JsonParse = Nothing
    On Error GoTo 0
End Function

Public Function JsonStringify(ByVal value As Variant) As String
    On Error Resume Next
    JsonStringify = JsonConverter.ConvertToJson(value)
    If Err.Number <> 0 Then JsonStringify = vbNullString
    On Error GoTo 0
End Function

Public Function JsonStringifyPretty(ByVal value As Variant, _
                                     Optional ByVal indent As Long = 2) As String
    On Error Resume Next
    JsonStringifyPretty = JsonConverter.ConvertToJson(value, indent)
    If Err.Number <> 0 Then JsonStringifyPretty = vbNullString
    On Error GoTo 0
End Function
