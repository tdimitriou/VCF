Attribute VB_Name = "modXamlResources"
Option Explicit

Public Function LoadXamlFolder() As ObservableDictionary
    Dim Dict As ObservableDictionary
    Set Dict = New ObservableDictionary
    
    Dim Root As String
    Root = App.Path
    If Right$(Root, 1) <> "\" Then Root = Root & "\"
    Root = Root & "Resources\XAML\"
    
    If New_c.FSO.FolderExists(Root) Then AddFolderResources Dict, Root
    
    Set LoadXamlFolder = Dict
End Function

Private Sub AddFolderResources(ByVal Dictionary As ObservableDictionary, ByVal Path As String, Optional Subdirectory As String = "")
    On Error GoTo Done
    
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    If Len(Subdirectory) Then If Right$(Subdirectory, 1) <> "\" Then Subdirectory = Subdirectory & "\"
    
    With New_c.FSO.GetDirList(Path & Subdirectory, , "*.xml")
        Dim Index As Long
        For Index = 0 To .FilesCount - 1
            Dim Key As String
            Key = Subdirectory & Replace$(.FileName(Index), ".xml", "", , , vbTextCompare)
            Dictionary.Item(Key) = New_c.Crypt.UTF8ToVBString(New_c.FSO.ReadByteContent(Path & Subdirectory & .FileName(Index)))
        Next
        
        For Index = 0 To .SubDirsCount - 1
            AddFolderResources Dictionary, Path, Subdirectory & .SubDirName(Index)
        Next
    End With

Done:
End Sub
