VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NestedPropertyChildChangedHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPropertyChangedHandler

Private m_OwnerPtr As Long
Private m_Notifier As PropertyChangedEvent

Friend Sub Connect(ByVal Owner As NestedProperty, ByVal Notifier As PropertyChangedEvent)
    Disconnect
    If Notifier Is Nothing Then Exit Sub

    m_OwnerPtr = ObjPtr(Owner)
    Set m_Notifier = Notifier
    m_Notifier.AddHandler Me
End Sub

Friend Sub Disconnect()
    On Error Resume Next
    If Not m_Notifier Is Nothing Then m_Notifier.RemoveHandler Me
    Set m_Notifier = Nothing
    m_OwnerPtr = 0
End Sub

Private Sub IPropertyChangedHandler_OnPropertyChanged(ByVal Sender As Object, ByVal PropertyName As Variant)
    If m_OwnerPtr = 0 Then Exit Sub

    Dim Owner As NestedProperty
    Set Owner = API.ObjFromPtr(m_OwnerPtr)
    If Owner Is Nothing Then Exit Sub

    Owner.HandleChildPropertyChanged PropertyName
End Sub

Private Sub Class_Terminate()
    Disconnect
End Sub
