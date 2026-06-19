Attribute VB_Name = "modApp"
Option Explicit

Public ShellWindow As New ShellWindow
Public MyApp As MyApp

Sub StartApp()
    VCF.SetCustomConstructor Nothing
    
    Set MyApp = New MyApp
    
    Dim IApp As IApplication
    Set IApp = MyApp
    
    IApp.Resources.Add "Hello", "World!"
    

    'IApp.Resources.Clear
    
    MyApp.Base.Run ShellWindow
    
    
    'EnumerateResources
End Sub

Private Function EnumerateResources()
    Dim r
    Dim Dictionary As ObservableDictionary
    Set Dictionary = MyApp.Base.Resources

    
    Debug.Print "Dictionary Count:"; Dictionary.Count
    
    Dim Value As String
    On Error Resume Next
    For Each r In Dictionary
        Value = r
        If Err Then
            Value = ""
            Err.Clear
        End If
        Debug.Print "Type:"; TypeName(r), "Index:"; Dictionary.IndexOf(r), "Key:"; Dictionary.KeyOfIndex(Dictionary.IndexOf(r)), "Value:"; Value
    Next
End Function
