VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Children              As UIElementCollection

Public DataType                 As String
Public Key                      As String
Public Name                     As String

Public Property Get Children() As UIElementCollection
    Set Children = m_Children
End Property

Private Sub Class_Initialize()
    Set m_Children = New UIElementCollection
End Sub
