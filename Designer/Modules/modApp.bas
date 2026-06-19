Attribute VB_Name = "modApp"
Option Explicit

Sub Main()
    On Error Resume Next
    
    ' Show the designer form
    frmDesigner.Show
    
    ' Enter message loop
    Cairo.WidgetForms.EnterMessageLoop
End Sub

