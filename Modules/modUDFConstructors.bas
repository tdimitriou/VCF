Attribute VB_Name = "modUDFConstructors"
Option Explicit

Public Function NewCornerRadius(ByVal Args As String) As CornerRadius
    If Len(Args) = 0 Then Exit Function
    
    Dim Params
    Params = Split(Args, ",")
    
    If UBound(Params) = 0 Then
        With NewCornerRadius
            .TopLeft = Trim$(Params(0))
            .TopRight = Trim$(Params(0))
            .BottomLeft = Trim$(Params(0))
            .BottomRight = Trim$(Params(0))
        End With
    ElseIf UBound(Params) = 3 Then
        With NewCornerRadius
            .TopLeft = Trim$(Params(0))
            .TopRight = Trim$(Params(1))
            .BottomLeft = Trim$(Params(2))
            .BottomRight = Trim$(Params(3))
        End With
    End If
End Function
