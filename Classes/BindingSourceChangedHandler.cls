VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BindingSourceChangedHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPropertyChangedHandler

Private m_BindingPtr As Long
Private m_Notifier As PropertyChangedEvent

Friend Sub Connect(ByVal Binding As Binding, ByVal Notifier As PropertyChangedEvent)
    Disconnect
    If Notifier Is Nothing Then Exit Sub

    m_BindingPtr = ObjPtr(Binding)
    Set m_Notifier = Notifier
    m_Notifier.AddHandler Me
End Sub

Friend Sub Disconnect()
    On Error Resume Next
    If Not m_Notifier Is Nothing Then m_Notifier.RemoveHandler Me
    Set m_Notifier = Nothing
    m_BindingPtr = 0
End Sub

Private Sub IPropertyChangedHandler_OnPropertyChanged(ByVal Sender As Object, ByVal PropertyName As Variant)
    If m_BindingPtr = 0 Then Exit Sub

    Dim BindingObj As Binding
    Set BindingObj = API.ObjFromPtr(m_BindingPtr)
    If BindingObj Is Nothing Then Exit Sub

    BindingObj.NotifySourcePropertyChanged PropertyName
End Sub

Private Sub Class_Terminate()
    Disconnect
End Sub
