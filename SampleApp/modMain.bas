Attribute VB_Name = "modMain"
Option Explicit

Public ShellWindow As New ShellWindow
Public MyApp As MyApp

Sub Main()
    VCF.SetCustomConstructor Nothing
    
    Set MyApp = New MyApp
    
    Dim IApp As IApplication
    Set IApp = MyApp
    
    IApp.Resources.Add "Hello", "World!"
    IApp.Run ShellWindow

End Sub
