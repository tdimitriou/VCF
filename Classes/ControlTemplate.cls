VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ControlTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Children As UIElementCollection

Public TargetType As String
Public Key As String
Public Name As String

Public Property Get Children() As UIElementCollection
    Set Children = m_Children
End Property

Public Function Clone() As ControlTemplate
    Dim Copy As ControlTemplate
    Dim i As Long

    Set Copy = New ControlTemplate
    Copy.TargetType = TargetType
    Copy.Key = Key
    Copy.Name = Name

    For i = 0 To m_Children.Count - 1
        Copy.Children.Add m_Children(i)
    Next

    Set Clone = Copy
End Function

Private Sub Class_Initialize()
    Set m_Children = New UIElementCollection
End Sub
