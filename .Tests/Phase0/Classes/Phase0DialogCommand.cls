VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phase0DialogCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements VCF.ICommand
Implements VCF.INotifyCanExecuteChanged

Private m_LastParameter As Variant
Private m_ExecuteCount As Long
Private m_CanExecute As Boolean
Private m_Event As VCF.PropertyChangedEvent

Public Property Get LastParameter() As Variant
    If IsObject(m_LastParameter) Then
        Set LastParameter = m_LastParameter
    Else
        LastParameter = m_LastParameter
    End If
End Property

Public Property Get ExecuteCount() As Long
    ExecuteCount = m_ExecuteCount
End Property

Public Property Get CanExecute() As Boolean
    CanExecute = m_CanExecute
End Property

Public Property Let CanExecute(ByVal Value As Boolean)
    If m_CanExecute = Value Then Exit Property
    m_CanExecute = Value
    RaiseCanExecuteChanged
End Property

Public Sub Reset()
    m_ExecuteCount = 0
    m_LastParameter = Empty
    m_CanExecute = True
End Sub

Public Sub RaiseCanExecuteChanged()
    If Not m_Event Is Nothing Then m_Event.OnPropertyChanged "CanExecute"
End Sub

Private Function ICommand_CanExecute(Optional Parameter As Variant) As Boolean
    ICommand_CanExecute = m_CanExecute
End Function

Private Sub ICommand_Execute(Optional Parameter As Variant)
    m_ExecuteCount = m_ExecuteCount + 1
    If IsObject(Parameter) Then
        Set m_LastParameter = Parameter
    Else
        m_LastParameter = Parameter
    End If
End Sub

Private Function INotifyCanExecuteChanged_CanExecuteChangedEvent() As VCF.PropertyChangedEvent
    Set INotifyCanExecuteChanged_CanExecuteChangedEvent = m_Event
End Function

Private Sub Class_Initialize()
    m_CanExecute = True
    Set m_Event = New VCF.PropertyChangedEvent
    m_Event.Register Me
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not m_Event Is Nothing Then m_Event.Unregister
End Sub
