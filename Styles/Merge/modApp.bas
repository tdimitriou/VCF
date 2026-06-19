Attribute VB_Name = "modApp"
Option Explicit

Sub Main()

Try:
    On Error GoTo Catch
    
    Dim DL As cDirList
    
    Dim Path As String
    Path = App.Path
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    
    Set DL = New_c.FSO.GetDirList(Path, , "*.xml")
    
    If DL.FilesCount = 0 Then
        MsgBox "No files to merge!", vbExclamation + vbOKOnly
        Exit Sub
    End If
    

    Dim SB As cStringBuilder
    Set SB = New_c.StringBuilder
    
    Dim Index As Long
    
    Dim File As String
    Dim FilesMerged As Long
    
    For Index = 0 To DL.FilesCount - 1
        File = DL.FileName(Index)
        
        If LCase$(File) <> "styles.xml" Then
            FilesMerged = FilesMerged + 1
            SB.AppendNL New_c.FSO.ReadTextContent(Path & File)
        End If
    Next
    
    New_c.FSO.WriteTextContent Path & "Styles.xml", Replace$(SB.ToString, vbCrLf & vbCrLf, vbCrLf)
        
    MsgBox "Successfully merged " & FilesMerged & " files into 'Styles.xml'", vbInformation + vbOKOnly
    Exit Sub
    
Catch:
    
    MsgBox Err.Description, vbExclamation + vbOKOnly

End Sub
