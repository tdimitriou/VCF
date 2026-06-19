VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShellExecCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements ICommand

Public Function CanExecute(Optional Parameter As Variant) As Boolean
    CanExecute = True
End Function

Public Sub Execute(Optional Parameter As Variant)
    On Error Resume Next
    
    Shell Parameter
End Sub

Private Function ICommand_CanExecute(Optional Parameter As Variant) As Boolean
    ICommand_CanExecute = CanExecute
End Function

Private Sub ICommand_Execute(Optional Parameter As Variant)
    Call Execute(Parameter)
End Sub


