Attribute VB_Name = "modStyleApplyLog"
Option Explicit

' Style-apply chain diagnostics (On Error GoTo Handler + Erl).
' Writes Debug.Print + appends TEMP\VCF_StyleApply.log

Private Const LOG_FILE_NAME As String = "VCF_StyleApply.log"

Public Sub LogError(ByVal ModuleName As String, _
                    ByVal MemberName As String, _
                    ByVal ErrNumber As Long, _
                    ByVal ErrDescription As String, _
                    ByVal LineNumber As Long)
    Dim Line As String
    Dim Path As String
    Dim Fn As Integer

    On Error Resume Next

    Line = Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & _
           ModuleName & "." & MemberName & vbTab & _
           "Err=" & CStr(ErrNumber) & vbTab & _
           "Erl=" & CStr(LineNumber) & vbTab & _
           ErrDescription

    Debug.Print "STYLE-APPLY " & Line

    Path = StyleApplyLogPath()
    If Len(Path) = 0 Then Exit Sub

    Fn = FreeFile
    Open Path For Append As #Fn
    Print #Fn, Line
    Close #Fn
End Sub

' Capture Err, log, re-raise so callers in the chain can log too (or outer Resume Next swallows after log).
Public Sub LogErrorAndReraise(ByVal ModuleName As String, ByVal MemberName As String)
    Dim ErrNumber As Long
    Dim ErrSource As String
    Dim ErrDescription As String
    Dim LineNumber As Long

    ErrNumber = Err.Number
    ErrSource = Err.Source
    ErrDescription = Err.Description
    LineNumber = Erl

    LogError ModuleName, MemberName, ErrNumber, ErrDescription, LineNumber

    If Len(ErrSource) = 0 Then
        Err.Raise ErrNumber, ModuleName & "." & MemberName, ErrDescription
    Else
        Err.Raise ErrNumber, ErrSource, ErrDescription
    End If
End Sub

Public Function StyleApplyLogPath() As String
    Dim Folder As String

    On Error Resume Next
    Folder = Environ$("TEMP")
    If Len(Folder) = 0 Then Folder = Environ$("TMP")
    If Len(Folder) = 0 Then Exit Function
    If Right$(Folder, 1) <> "\" Then Folder = Folder & "\"
    StyleApplyLogPath = Folder & LOG_FILE_NAME
End Function
