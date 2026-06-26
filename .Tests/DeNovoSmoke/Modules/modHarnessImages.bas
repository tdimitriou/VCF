Attribute VB_Name = "modHarnessImages"
Option Explicit

Public Sub RegisterHarnessImages()
    On Error Resume Next
    
    Dim Base As String
    Base = App.Path
    If Right$(Base, 1) <> "\" Then Base = Base & "\"
    Base = Base & "Resources\XAML\Resources\"
    
    RegisterOne "Resources\ClockIn.png", Base & "ClockIn.png"
    RegisterOne "Resources\Reboot.png", Base & "Reboot.png"
    RegisterOne "Resources\Close.png", Base & "Close.png"
End Sub

Private Sub RegisterOne(ByVal Key As String, ByVal FilePath As String)
    If Not New_c.FSO.FileExists(FilePath) Then Exit Sub
    If Cairo.ImageList.Exists(Key) Then Exit Sub
    Cairo.ImageList.AddImage Key, FilePath
End Sub
