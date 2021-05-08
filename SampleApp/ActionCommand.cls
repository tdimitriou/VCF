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

Implements VCF.ICommand

Private m_Function As VCF.Function

Private Function ICommand_CanExecute(Optional Parameter As Variant) As Boolean
    ICommand_CanExecute = True
End Function

Private Sub ICommand_Execute(Optional Parameter As Variant)
    m_Function.Execute Parameter
End Sub

Friend Sub Initialize(Action As VCF.Function)
    Set m_Function = Action
End Sub
