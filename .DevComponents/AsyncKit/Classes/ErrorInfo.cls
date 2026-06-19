VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Read-Only properties providing full COM Error context
Private m_Number As Long
Private m_Description As String
Private m_Source As String

' Internal friend method used by BackgroundWorker to populate the data
Friend Sub Populate(ByVal ErrNumber As Long, ByVal ErrDescription As String, ByVal ErrSource As String)
    m_Number = ErrNumber
    m_Description = ErrDescription
    m_Source = ErrSource
End Sub

Public Property Get Number() As Long
Attribute Number.VB_UserMemId = 0
    Number = m_Number
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Get Source() As String
    Source = m_Source
End Property

