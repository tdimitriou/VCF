VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PropertyChangedEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_SenderPtr As Long
Private m_IsRegistered As Boolean
Private m_Handlers As List
Private m_IsNotifying As Boolean

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

Public Sub AddHandler(ByVal Handler As IPropertyChangedHandler)
    If Handler Is Nothing Then Exit Sub
    If m_Handlers.Contains(Handler) Then Exit Sub
    m_Handlers.Add Handler
End Sub

Public Sub RemoveHandler(ByVal Handler As IPropertyChangedHandler)
    On Error Resume Next
    If Handler Is Nothing Then Exit Sub
    If Not m_Handlers.Contains(Handler) Then Exit Sub
    m_Handlers.Remove Handler
End Sub

Public Sub OnPropertyChanged(ByVal PropertyName As String)
    If Not m_IsRegistered Then
        Err.Raise vbObjectError, , "Event not Registered"
        Exit Sub
    End If
    
    If m_IsNotifying Then Exit Sub
    
    Dim Sender As Object
    Set Sender = GetSender
    
    If Sender Is Nothing Then Exit Sub
    
    m_IsNotifying = True
    
    On Error GoTo Finally
    
    Dim Handler As IPropertyChangedHandler
    Dim Item As Variant
    
    For Each Item In m_Handlers
        Set Handler = Item
        If Not Handler Is Nothing Then Handler.OnPropertyChanged Sender, PropertyName
    Next
    
Finally:
    m_IsNotifying = False
End Sub

Private Function GetSender() As Object
    Set GetSender = API.ObjFromPtr(m_SenderPtr)
End Function

Private Sub Class_Initialize()
    Set m_Handlers = New List
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    m_Handlers.Clear
End Sub
