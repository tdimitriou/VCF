VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionChangedEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_SenderPtr As Long
Private m_IsRegistered As Boolean

Public Event CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)

Public Property Get IsRegistered() As Boolean
    IsRegistered = m_IsRegistered
End Property

Public Sub Register(Sender As Object)
    If m_IsRegistered Then
        Err.Raise vbObjectError, , "Already Registered"
        Exit Sub
    End If
    
    If Sender Is Nothing Then Exit Sub
    
    m_SenderPtr = ObjPtr(Sender)
    
    m_IsRegistered = True
End Sub

Public Sub Unregister()
    m_IsRegistered = False
    m_SenderPtr = 0
End Sub

Public Sub OnCollectionChanged(ByVal e As CollectionChangedEventArgs)
    If Not m_IsRegistered Then
        Err.Raise vbObjectError, , "Event not Registered"
        Exit Sub
    End If
    
    Dim Sender As Object
    Set Sender = GetSender
    
    If Sender Is Nothing Then Exit Sub
    
    RaiseEvent CollectionChanged(Sender, e)
End Sub

Private Function GetSender() As Object
    Set GetSender = API.ObjFromPtr(m_SenderPtr)
End Function

