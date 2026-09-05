Option Explicit

' ============================================================================
' Module  : modShell
' Purpose : Minimal cross-platform shell execution support.
'
'   Windows - unsupported; returns ""
'   macOS   - executes commands via popen (libc.dylib)
'
' Public API:
'   ShellExec(command)  -> stdout, or "" when unsupported/on error
'
' Inspired by VBA-Web WebHelpers.ExecuteInShell
' Copyright (c) Tim Hall, MIT license
' ============================================================================

' --- Mac: libc.dylib declarations for popen-based shell execution ---

#If Mac Then
    #If VBA7 Then
        Private Declare PtrSafe Function popen_File Lib "/usr/lib/libc.dylib" Alias "popen" _
            (ByVal cmd As String, ByVal mode As String) As LongPtr
        Private Declare PtrSafe Function popen_Close Lib "/usr/lib/libc.dylib" Alias "pclose" _
            (ByVal file As LongPtr) As LongPtr
        Private Declare PtrSafe Function popen_Read Lib "/usr/lib/libc.dylib" Alias "fread" _
            (ByVal buffer As String, ByVal size As LongPtr, ByVal count As LongPtr, ByVal file As LongPtr) As LongPtr
        Private Declare PtrSafe Function popen_Eof Lib "/usr/lib/libc.dylib" Alias "feof" _
            (ByVal file As LongPtr) As LongPtr
    #Else
        Private Declare Function popen_File Lib "libc.dylib" Alias "popen" _
            (ByVal cmd As String, ByVal mode As String) As Long
        Private Declare Function popen_Close Lib "libc.dylib" Alias "pclose" _
            (ByVal file As Long) As Long
        Private Declare Function popen_Read Lib "libc.dylib" Alias "fread" _
            (ByVal buffer As String, ByVal size As Long, ByVal count As Long, ByVal file As Long) As Long
        Private Declare Function popen_Eof Lib "libc.dylib" Alias "feof" _
            (ByVal file As Long) As Long
    #End If
#End If


' --- ShellExec - Execute a command and return stdout. ---

Public Function ShellExec(ByVal cmd As String) As String
#If Mac Then
    #If VBA7 Then
        Dim file As LongPtr
        Dim bytesRead As LongPtr
    #Else
        Dim file As Long
        Dim bytesRead As Long
    #End If

    Dim output As String
    Dim chunk As String

    On Error GoTo ErrHandler
    file = popen_File(cmd, "r")
    If file = 0 Then Exit Function

    Do While popen_Eof(file) = 0
        chunk = Space$(255)
        bytesRead = popen_Read(chunk, 1, 255, file)
        If bytesRead > 0 Then
            output = output & Left$(chunk, CLng(bytesRead))
        End If
    Loop

    popen_Close file
    file = 0
    ShellExec = output
    Exit Function

ErrHandler:
    DiagnosticsReraiseIfDev "modShell.ShellExec"
    On Error Resume Next
    If file <> 0 Then popen_Close file
    ShellExec = ""
    On Error GoTo 0
#Else
    ShellExec = ""
#End If
End Function