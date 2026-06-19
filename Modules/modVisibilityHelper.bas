Attribute VB_Name = "modVisibilityHelper"
Option Explicit

Public Sub SetVisibility(W As cWidgetBase, Value As Visibility)
    Select Case Value
        Case VisibilityVisible
            If Not W.Visible Then W.Visible = True
        Case VisibilityHidden, VisibilityCollapsed
            If W.Visible Then W.Visible = False
    End Select
End Sub

Public Sub ApplyVisibility(W As cWidgetBase, Value As Visibility)
    SetVisibility W, Value
End Sub
