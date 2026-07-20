Attribute VB_Name = "modBorderTest"
Option Explicit

' Opens a modeless bordered window (XAML BorderStyle=2) for chrome lifecycle testing.
Public Sub ShowBorderTestWindow()
    On Error GoTo Fail
    
    modBorderChromeDiag.Reset
    
    Dim TestWin As BorderTestWindow
    Set TestWin = New BorderTestWindow
    TestWin.Show vbModeless
    
    Exit Sub

Fail:

    MsgBox Err.Description, vbCritical, "DeNovoSmoke border test"

End Sub
