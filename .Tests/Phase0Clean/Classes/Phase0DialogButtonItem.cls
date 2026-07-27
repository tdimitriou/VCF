VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phase0DialogButtonItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INotifyPropertyChanged

Private m_Event As VCF.PropertyChangedEvent
Private m_Text As String
Private m_Value As Variant
Private m_Command As ICommand

Public Property Get Text() As String
    Text = m_Text
End Property

Public Property Let Text(ByVal Value As String)
    If m_Text = Value Then Exit Property
    m_Text = Value
    m_Event.OnPropertyChanged "Text"
End Property

Public Property Get Value() As Variant
    If IsObject(m_Value) Then
        Set Value = m_Value
    Else
        Value = m_Value
    End If
End Property

Public Property Let Value(ByVal NewValue As Variant)
    If IsObject(NewValue) Then
        Set m_Value = NewValue
    Else
        m_Value = NewValue
    End If
    m_Event.OnPropertyChanged "Value"
End Property

Public Property Set Value(ByVal NewValue As Object)
    Set m_Value = NewValue
    m_Event.OnPropertyChanged "Value"
End Property

Public Property Get Command() As ICommand
    Set Command = m_Command
End Property

Public Property Set Command(ByVal Value As ICommand)
    Set m_Command = Value
    m_Event.OnPropertyChanged "Command"
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
