VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XAMLThicknessConstructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function NewThickness(Optional Args As String) As Thickness
    Dim Params
    
    If Len(Args) = 0 Then
        Set NewThickness = modConstructors.NewThickness
        Exit Function
    End If
    
    Params = Split(Args, ",")
    
    Dim Index As Long
    For Index = 0 To UBound(Params)
        Params(Index) = Val(Trim$(Params(Index)))
    Next
        
    Select Case UBound(Params)
        Case 0: Set NewThickness = modConstructors.NewThickness(Params(0))
        Case 1: Set NewThickness = modConstructors.NewThickness(Params(0), Params(1))
        Case 3: Set NewThickness = modConstructors.NewThickness(Params(0), Params(1), Params(2), Params(3))
    End Select
End Function
