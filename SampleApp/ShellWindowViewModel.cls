VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShellWindowViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get ButtonClick() As ICommand
    Dim Cmd As ActionCommand
    Set Cmd = New ActionCommand
    
    Cmd.Initialize VCF.NewFunction(Me, "ButtonClickMethod")
    
    Set ButtonClick = Cmd
    
End Property

Public Sub ButtonClickMethod()

    MsgBox "You Clicked Me!"

End Sub

