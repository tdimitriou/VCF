VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XAMLImagePropertyManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function LoadImage(ByVal Path As String) As String
    If LCase$(Left$(Path, 4)) = "res:" Then
        LoadImage = LoadImageFromResource(Right$(Path, Len(Path) - 4))
    Else
        LoadImage = LoadImageFromFile(Path)
    End If
End Function

Private Function LoadImageFromFile(ByVal Path As String) As String
    If Not New_c.FSO.FileExists(Path) Then Exit Function
    
    Dim Key As String
    'Key = Replace$(Replace$(Path, "\", "_"), ":", "_")
    Key = Path
    
    If Not Cairo.ImageList.Exists(Key) Then
        If New_c.FSO.GetFileExtension(Path) = "ico" Then
            Call Cairo.ImageList.AddIconFromFile(Key, Path)
        Else
            Call Cairo.ImageList.AddImage(Key, Path)
        End If
    End If
    
    LoadImageFromFile = Key
End Function

Private Function LoadImageFromResource(ByVal Key As String) As String
    '
End Function
