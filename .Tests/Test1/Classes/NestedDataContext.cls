VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NestedDataContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INotifyPropertyChanged

Private m_Event As PropertyChangedEvent
Private WithEvents m_Col As ObservableCollection
Attribute m_Col.VB_VarHelpID = -1

Public Child As NestedDataContext

Private Sub Class_Initialize()
    Set m_Event = New PropertyChangedEvent
    m_Event.Register Me
    
    Set m_Col = New ObservableCollection
End Sub

Private Sub Class_Terminate()
    m_Event.Unregister
End Sub

Private Function INotifyPropertyChanged_PropertyChangedEvent() As VCF.PropertyChangedEvent
    Set INotifyPropertyChanged_PropertyChangedEvent = m_Event
End Function

Private Sub m_Col_CollectionChanged(sender As Object, e As VCF.CollectionChangedEventArgs)
    m_Event.OnPropertyChanged "Count"
End Sub

Public Property Get List() As ObservableCollection
    Set List = m_Col
End Property

Public Property Get Count() As Long
    Count = m_Col.Count
End Property
