VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XamlLoadException"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Message As String
Private m_ElementName As String
Private m_PropertyName As String
Private m_Line As Long
Private m_Column As Long
Private m_InnerCode As Long

Friend Sub Initialize( _
    ByVal Message As String, _
    ByVal ElementName As String, _
    ByVal PropertyName As String, _
    ByVal Line As Long, _
    ByVal Column As Long, _
    ByVal InnerCode As Long)

    m_Message = Message
    m_ElementName = ElementName
    m_PropertyName = PropertyName
    m_Line = Line
    m_Column = Column
    m_InnerCode = InnerCode
End Sub

Public Property Get Message() As String
Attribute Message.VB_UserMemId = 0
    Message = m_Message
End Property

Public Property Get ElementName() As String
    ElementName = m_ElementName
End Property

Public Property Get PropertyName() As String
    PropertyName = m_PropertyName
End Property

Public Property Get Line() As Long
    Line = m_Line
End Property

Public Property Get Column() As Long
    Column = m_Column
End Property

Public Property Get InnerCode() As Long
    InnerCode = m_InnerCode
End Property

Public Function ToString() As String
    Dim Parts As String
    Parts = m_Message

    If Len(m_ElementName) > 0 Then Parts = Parts & " [Element=" & m_ElementName & "]"
    If Len(m_PropertyName) > 0 Then Parts = Parts & " [Property=" & m_PropertyName & "]"
    If m_Line > 0 Then Parts = Parts & " [Line=" & m_Line & "]"
    If m_Column > 0 Then Parts = Parts & " [Column=" & m_Column & "]"
    If m_InnerCode <> 0 Then Parts = Parts & " [Inner=" & m_InnerCode & "]"

    ToString = Parts
End Function

Public Sub Raise()
    Err.Raise vbXamlLoadError, "VCF.XamlLoadException", ToString()
End Sub
