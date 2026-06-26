VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StubLanguageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INotifyPropertyChanged

Private m_Event As VCF.PropertyChangedEvent
Private m_Text As Object

Public Property Get Text(Identifier) As String
    On Error Resume Next
    If m_Text.Exists(CStr(Identifier)) Then
        Text = m_Text(CStr(Identifier))
    Else
        Text = CStr(Identifier)
    End If
End Property

Private Function INotifyPropertyChanged_PropertyChangedEvent() As VCF.PropertyChangedEvent
    Set INotifyPropertyChanged_PropertyChangedEvent = m_Event
End Function

Private Sub Class_Initialize()
    Set m_Event = New PropertyChangedEvent
    m_Event.Register Me
    Set m_Text = CreateObject("Scripting.Dictionary")
    m_Text.CompareMode = TextCompare
    m_Text("Loading") = "Loading"
    m_Text("Clear") = "Clear"
    m_Text("Login") = "Login"
    m_Text("Restart") = "Restart"
    m_Text("Exit") = "Exit"
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    m_Event.Unregister
End Sub
