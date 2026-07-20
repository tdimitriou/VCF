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
    
    ResetVcfGlobals
    modHarnessAppManager.ResetShutdownState
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
    
    ' Run returns when Cairo widget forms are removed (see modHarnessAppManager.Shutdown).

    ResetSession
    
    Exit Sub

Fail:

    MsgBox Err.Description, vbCritical, "DeNovoSmoke"

End Sub

' After Application.Run — IDE session reset only (DeNovo does not null Shell on exit in compiled exe).
Private Sub ResetSession()
    On Error Resume Next
    
    VCF.SetCustomConstructor Nothing
    Set Shell = Nothing
    VCF.ClearApplication
    Set ActiveLoginViewModel = Nothing
    Set AppProperties = Nothing
    Set AppCommands = Nothing
    Set LanguageManager = Nothing
    Set XamlResources = Nothing
    Set HarnessApp = Nothing
End Sub

Private Sub ResetVcfGlobals()

    On Error Resume Next

    VCF.SetCustomConstructor Nothing
    VCF.ClearApplication

End Sub

