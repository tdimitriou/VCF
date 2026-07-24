Attribute VB_Name = "modVisibilityHelper"
Option Explicit

Public Sub SetVisibility(W As cWidgetBase, Value As Visibility)
    If W Is Nothing Then Exit Sub
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

' WPF UIElement.IsHitTestVisible - False skips this widget in Cairo hit-testing
' (clicks pass through to whatever is behind / the parent).
Public Sub ApplyIsHitTestVisible(W As cWidgetBase, ByVal Value As Boolean)
    If W Is Nothing Then Exit Sub
    W.ImplementsHitTest = Value
End Sub
