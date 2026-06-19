Attribute VB_Name = "modMain"
Option Explicit

Public Form1 As New Form1
Public MyApp As App1

Sub Main()
    VCF.SetCustomConstructor Nothing
    
    Set MyApp = New App1
    
    Dim WW As IApplication
    Set WW = MyApp
    
    WW.Resources.Add "Hello", "World!"
    WW.Run Form1
'    Form1.Form.Show
    
'    Cairo.WidgetForms.EnterMessageLoop
End Sub
