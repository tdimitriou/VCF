Attribute VB_Name = "modBindingExpressions"
Option Explicit

Public Sub OnDataContextChanged(ByVal Target As IDependencyObject)
    RefreshTargetBindings Target
End Sub

Public Sub RefreshTargetBindings(ByVal Target As IDependencyObject)
    Dim Bindings As List
    Dim Item As Variant
    Dim Expr As BindingExpression
    Dim B As Binding

    Set Bindings = GetTargetBindings(Target)
    If Bindings Is Nothing Then Exit Sub

    For Each Item In Bindings
        If TypeOf Item Is BindingExpression Then
            Set Expr = Item
            Expr.UpdateTarget
        ElseIf TypeOf Item Is Binding Then
            Set B = Item
            B.RefreshTarget
        End If
    Next
End Sub

Public Sub DetachTargetBindings(ByVal Target As IDependencyObject)
    Dim Bindings As List
    Dim Item As Variant
    Dim Expr As BindingExpression
    Dim B As Binding
    Dim Snapshot As List
    Dim Copy As Variant

    Set Bindings = GetTargetBindings(Target)
    If Bindings Is Nothing Then Exit Sub

    Set Snapshot = New List
    For Each Item In Bindings
        Snapshot.Add Item
    Next

    For Each Copy In Snapshot
        If TypeOf Copy Is BindingExpression Then
            Set Expr = Copy
            Expr.Detach
        ElseIf TypeOf Copy Is Binding Then
            Set B = Copy
            B.DetachBinding
        End If
    Next
End Sub

Private Function GetTargetBindings(ByVal Target As IDependencyObject) As List
    On Error Resume Next
    Set GetTargetBindings = API.CObj(Target).Bindings
End Function
