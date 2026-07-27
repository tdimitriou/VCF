VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phase0ViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INotifyPropertyChanged

Private m_Event As VCF.PropertyChangedEvent
Private m_Title As String
Private m_RowIndex As Long

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(ByVal Value As String)
    If m_Title = Value Then Exit Property
    m_Title = Value
    m_Event.OnPropertyChanged "Title"
End Property

Public Property Get RowIndex() As Long
    RowIndex = m_RowIndex
End Property

Public Property Let RowIndex(ByVal Value As Long)
    If m_RowIndex = Value Then Exit Property
    m_RowIndex = Value
    m_Event.OnPropertyChanged "RowIndex"
End Property

Private Function INotifyPropertyChanged_PropertyChangedEvent() As VCF.PropertyChangedEvent
    Set INotifyPropertyChanged_PropertyChangedEvent = m_Event
End Function

Private Sub Class_Initialize()
    Set m_Event = New PropertyChangedEvent
    m_Event.Register Me
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    m_Event.Unregister
End Sub
