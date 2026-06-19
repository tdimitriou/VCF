VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActionCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

Private m_CanExecute As Variant
Private m_Execute As VCF.Function

Friend Sub Init(Execute As VCF.Function, CanExecute)
    Set m_Execute = Execute
    
    If TypeOf CanExecute Is VCF.Function Then
        Set m_CanExecute = CanExecute
    Else
        m_CanExecute = CanExecute
    End If
End Sub

Private Function ICommand_CanExecute(Optional Parameter As Variant) As Boolean
    If TypeOf m_CanExecute Is VCF.Function Then
        ICommand_CanExecute = m_CanExecute.Execute
    Else
        ICommand_CanExecute = m_CanExecute
    End If
End Function

Private Sub ICommand_Execute(Optional Parameter As Variant)
    m_Execute.Execute Parameter
End Sub
