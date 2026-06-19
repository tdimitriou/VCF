Attribute VB_Name = "modApp"
Option Explicit

Public MyApp As MyApp

'CSEH: ErrMsgBox
Public Sub Start()
Try:
    On Error GoTo Catch
        
    VCF.SetCustomConstructor New ObjectConstructor
    
    Set MyApp = New MyApp

Exit Sub
    
Catch:
    MsgBox Err.Description, , App.FileDescription
    Err.Clear
End Sub
