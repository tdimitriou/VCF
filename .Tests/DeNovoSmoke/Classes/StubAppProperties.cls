VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StubAppProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INotifyPropertyChanged

Private m_Event As VCF.PropertyChangedEvent
Private WithEvents m_Timer As cTimer
Attribute m_Timer.VB_VarHelpID = -1

Public Property Get TerminalID() As Long
    TerminalID = 1
End Property

Public Property Get CurrentUser() As String
    CurrentUser = "Harness User"
End Property

Public Property Get CurrentTime() As Date
    CurrentTime = Now
End Property

Public Sub StartClock()
    If Not m_Timer Is Nothing Then Exit Sub
    Set m_Timer = New_c.Timer(1000, True)
End Sub

Public Sub StopClock()
    Set m_Timer = Nothing
End Sub

Private Function INotifyPropertyChanged_PropertyChangedEvent() As VCF.PropertyChangedEvent
    Set INotifyPropertyChanged_PropertyChangedEvent = m_Event
End Function

Private Sub Class_Initialize()
    Set m_Event = New PropertyChangedEvent
    m_Event.Register Me
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not m_Event Is Nothing Then m_Event.Unregister
End Sub

Private Sub m_Timer_Timer()
    m_Event.OnPropertyChanged "CurrentTime"
End Sub
