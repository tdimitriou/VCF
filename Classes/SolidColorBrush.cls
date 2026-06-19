VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SolidColorBrush"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Color As String

Public Property Get Color() As String
Attribute Color.VB_UserMemId = 0
    Color = modStaticClasses.Color.FromString(m_Color)
End Property

Public Property Let Color(ByVal Value As String)
    m_Color = Value
End Property
