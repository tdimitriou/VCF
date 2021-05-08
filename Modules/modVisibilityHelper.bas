Attribute VB_Name = "modVisibilityHelper"
Option Explicit

Public Sub SetVisibility(W As cWidgetBase, Value As Visibility)
    Dim bVisible As Boolean
    bVisible = (Value = VisibilityVisible)
    If W.Visible <> bVisible Then W.Visible = bVisible
End Sub
