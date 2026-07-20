Attribute VB_Name = "modHarnessAppManager"
Option Explicit

Private m_ShutdownInProgress As Boolean

' DeNovo AppManager.Shutdown core — end Cairo message loop only.
' Omits POS VB6 Forms unload / data teardown (not applicable to smoke harness).
Public Sub Shutdown()
    On Error Resume Next
    If m_ShutdownInProgress Then Exit Sub
    m_ShutdownInProgress = True
    
    If Not modApp.Shell Is Nothing Then modApp.Shell.StopTimers
    If Not modApp.AppProperties Is Nothing Then modApp.AppProperties.StopClock
    Cairo.WidgetForms.RemoveAll
End Sub

Public Sub ResetShutdownState()
    m_ShutdownInProgress = False
End Sub
