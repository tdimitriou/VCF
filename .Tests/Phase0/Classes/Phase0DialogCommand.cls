VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phase0DialogCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements VCF.ICommand

Private m_LastParameter As Variant
Private m_ExecuteCount As Long

Public Property Get LastParameter() As Variant
    If IsObject(m_LastParameter) Then
        Set LastParameter = m_LastParameter
    Else
        LastParameter = m_LastParameter
    End If
End Property

Public Property Get ExecuteCount() As Long
    ExecuteCount = m_ExecuteCount
End Property

Public Sub Reset()
    m_ExecuteCount = 0
    m_LastParameter = Empty
End Sub

Private Function ICommand_CanExecute(Optional Parameter As Variant) As Boolean
    ICommand_CanExecute = True
End Function

Private Sub ICommand_Execute(Optional Parameter As Variant)
    m_ExecuteCount = m_ExecuteCount + 1
    If IsObject(Parameter) Then
        Set m_LastParameter = Parameter
    Else
        m_LastParameter = Parameter
    End If
End Sub
