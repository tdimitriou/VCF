Attribute VB_Name = "modInformation"
Option Explicit

Public Function IsNothing(Expression) As Boolean
    On Error Resume Next
    
    ' If expression is not an object an error will
    ' occur and the function will return false
    IsNothing = (Expression Is Nothing)
End Function

