VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PropertyTrigger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_PropertyName As String
Private m_TriggerValue As String
Private m_Setters As cCollection

Public Sub Initialize(ByVal PropertyName As String, ByVal TriggerValue As String)
    m_PropertyName = PropertyName
    m_TriggerValue = TriggerValue
End Sub

Public Property Get PropertyName() As String
    PropertyName = m_PropertyName
End Property

Public Property Get TriggerValue() As String
    TriggerValue = m_TriggerValue
End Property

Public Sub SetSetter(ByVal PropertyName As String, ByVal Value As Variant)
    Dim e As KeyValuePair

    e.Key = PropertyName
    API.CopyVariable Value, e.Value
    m_Setters.Prop(PropertyName) = e
End Sub

Public Function SetterCount() As Long
    SetterCount = m_Setters.Count
End Function

Public Function SetterKeyAt(ByVal Index As Long) As String
    SetterKeyAt = m_Setters.KeyByIndex(Index)
End Function

Public Function SetterValueAt(ByVal Index As Long) As Variant
    Dim e As KeyValuePair
    e = m_Setters.ItemByIndex(Index)
    API.CopyVariable e.Value, SetterValueAt
End Function

Public Function Clone() As PropertyTrigger
    Dim Copy As PropertyTrigger
    Dim i As Long

    Set Copy = New PropertyTrigger
    Copy.Initialize m_PropertyName, m_TriggerValue

    For i = 0 To m_Setters.Count - 1
        Copy.SetSetter SetterKeyAt(i), SetterValueAt(i)
    Next

    Set Clone = Copy
End Function

Private Sub Class_Initialize()
    Set m_Setters = New_c.Collection(False)
End Sub
