Attribute VB_Name = "modApp"
Option Explicit

Public HarnessApp As AppHost
Public XamlResources As ObservableDictionary
Public LanguageManager As StubLanguageManager
Public AppCommands As StubAppCommands
Public AppProperties As StubAppProperties
Public ActiveLoginViewModel As StubLoginViewModel
Public Shell As ShellWindow

Public Sub Start()
    On Error GoTo Fail
    
    VCF.ClearCustomConstructor
    VCF.StrictXamlLoad = False
    
    Set LanguageManager = New StubLanguageManager
    Set AppCommands = New StubAppCommands
    Set AppProperties = New StubAppProperties
    AppProperties.StartClock
    
    modHarnessImages.RegisterHarnessImages
    
    Set HarnessApp = New AppHost
    HarnessApp.InitializeApplication
    
    Set Shell = New ShellWindow
    Shell.InitializeViews
    
    HarnessApp.Run Shell
    
    Exit Sub
    
Fail:
    MsgBox Err.Description, vbCritical, "DeNovoSmoke"
End Sub
