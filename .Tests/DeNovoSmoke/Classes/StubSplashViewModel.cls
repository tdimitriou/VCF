VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StubSplashViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INotifyPropertyChanged

Private m_Event As VCF.PropertyChangedEvent
Private m_Progress As String
Private m_InfoMessage As String

Public Property Get Progress() As String
    Progress = m_Progress
End Property

Public Property Let Progress(ByVal Value As String)
    If m_Progress = Value Then Exit Property
    m_Progress = Value
    m_Event.OnPropertyChanged "Progress"
End Property

Public Property Get InfoMessage() As String
    InfoMessage = m_InfoMessage
End Property

Public Property Let InfoMessage(ByVal Value As String)
    If m_InfoMessage = Value Then Exit Property
    m_InfoMessage = Value
    m_Event.OnPropertyChanged "InfoMessage"
End Property

Private Function INotifyPropertyChanged_PropertyChangedEvent() As VCF.PropertyChangedEvent
    Set INotifyPropertyChanged_PropertyChangedEvent = m_Event
End Function

Private Sub Class_Initialize()
    Set m_Event = New PropertyChangedEvent
    m_Event.Register Me
    m_Progress = "..."
    m_InfoMessage = "Loading User Interface"
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    m_Event.Unregister
End Sub
