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
Private m_HasContentAlignment As Boolean
Private m_ContentHAlign As Long
Private m_ContentVAlign As Long

Public TargetType As String
Public Key As String
Public Name As String

Public Property Get Children() As UIElementCollection
    Set Children = m_Children
End Property

' Lookless-prep: explicit ContentPresenter-slot alignment (avoids fragile TypeOf
' on paint-only ContentPresenter across the EXE/DLL boundary).
Public Property Get HasContentAlignmentMarker() As Boolean
    HasContentAlignmentMarker = m_HasContentAlignment
End Property

Public Property Get ContentHorizontalAlignment() As Long
    ContentHorizontalAlignment = m_ContentHAlign
End Property

Public Property Get ContentVerticalAlignment() As Long
    ContentVerticalAlignment = m_ContentVAlign
End Property

Public Sub SetContentAlignmentMarker(ByVal Horizontal As Long, ByVal Vertical As Long)
    m_HasContentAlignment = True
    m_ContentHAlign = Horizontal
    m_ContentVAlign = Vertical
End Sub

Public Function Clone() As ControlTemplate
    Dim Copy As ControlTemplate
    Dim i As Long

    Set Copy = New ControlTemplate
    Copy.TargetType = TargetType
    Copy.Key = Key
    Copy.Name = Name
    If m_HasContentAlignment Then
        Copy.SetContentAlignmentMarker m_ContentHAlign, m_ContentVAlign
    End If

    For i = 0 To m_Children.Count - 1
        Copy.Children.Add m_Children(i)
    Next

    Set Clone = Copy
End Function

Private Sub Class_Initialize()
    Set m_Children = New UIElementCollection
    m_HasContentAlignment = False
    m_ContentHAlign = AlignmentConstants.vbCenter
    m_ContentVAlign = 2
End Sub
