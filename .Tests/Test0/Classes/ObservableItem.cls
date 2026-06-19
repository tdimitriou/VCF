VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObservableItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INotifyPropertyChanged

Private m_PropertyChangedEvent As VCF.PropertyChangedEvent

Private m_FirstName As String
Private m_LastName As String

Public Property Get FirstName() As String
    FirstName = m_FirstName
End Property

Public Property Let FirstName(ByVal Value As String)
    If m_FirstName = Value Then Exit Property
    
    m_FirstName = Value
    
    OnPropertyChanged "FirstName"
End Property

Public Property Get LastName() As String
    LastName = m_LastName
End Property

Public Property Let LastName(ByVal Value As String)
    If m_LastName = Value Then Exit Property
    
    m_LastName = Value
    
    OnPropertyChanged "LastName"
End Property

Public Function PropertyChangedEvent() As VCF.PropertyChangedEvent
    Set PropertyChangedEvent = m_PropertyChangedEvent
End Function

Private Sub Class_Initialize()
    Set m_PropertyChangedEvent = New VCF.PropertyChangedEvent
    m_PropertyChangedEvent.Register Me
End Sub

Private Sub Class_Terminate()
    m_PropertyChangedEvent.Unregister
End Sub

Private Function INotifyPropertyChanged_PropertyChangedEvent() As VCF.PropertyChangedEvent
    Set INotifyPropertyChanged_PropertyChangedEvent = PropertyChangedEvent
End Function

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    m_PropertyChangedEvent.OnPropertyChanged PropertyName
End Sub
