VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LanguageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INotifyPropertyChanged

Private m_Event As PropertyChangedEvent

Private m_Language As String
Private Col As Collection

Public Property Get Text(Index) As String
    On Error Resume Next
    Text = Col(Index)
End Property

Public Property Get Language() As String
    Language = m_Language
End Property

Public Property Let Language(ByVal Value As String)
    If Language = Value Then Exit Property
    
    m_Language = Value
    
    SetLanguage
    
    m_Event.OnPropertyChanged "Language"
    m_Event.OnPropertyChanged "Text"
End Property

Private Sub SetLanguage()
    Dim c As Collection
    Set c = New Collection
    
    If m_Language = "EN" Then
    
        c.Add "Text 1", "t1"
        c.Add "Text 2", "t2"
    
    Else
        c.Add "Κείμενο 1", "t1"
        c.Add "Κείμενο 2", "t2"
    
    End If
    
    Set Col = c
    
End Sub

Private Sub Class_Initialize()
    Debug.Print "New LanguageManager"
    
    Set m_Event = New PropertyChangedEvent
    m_Event.Register Me
    
    Language = "EN"
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    m_Event.Unregister
End Sub

Private Function INotifyPropertyChanged_PropertyChangedEvent() As VCF.PropertyChangedEvent
    Set INotifyPropertyChanged_PropertyChangedEvent = m_Event
End Function
