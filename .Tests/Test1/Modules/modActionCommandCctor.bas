Attribute VB_Name = "modActionCommandCctor"
Option Explicit

Public Function NewActionCommand(Execute As VCF.Function, CanExecute As Variant) As ActionCommand
    Set NewActionCommand = New ActionCommand
    NewActionCommand.Init Execute, CanExecute
End Function
