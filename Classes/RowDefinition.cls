VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RowDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Height As String

Public Property Get Height() As String
    Height = m_Height
End Property

Public Property Let Height(ByVal Value As String)
    m_Height = Trim$(Value)
End Property

Private Sub Class_Initialize()
    m_Height = "*"
End Sub
