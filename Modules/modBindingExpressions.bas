Attribute VB_Name = "modBindingExpressions"
Option Explicit

Private m_RefreshBindings As Boolean

Public Sub OnDataContextChanged(ByVal Target As IDependencyObject)
    ' Binding SrcDepObj callbacks already push target updates when DataContext
    ' (DependencyProperty) changes. Avoid RefreshTargetBindings here ? re-entrant
    ' SetValue during DataContext change can recurse through layout/render paths.
End Sub

Public Sub RefreshTargetBindings(ByVal Target As IDependencyObject)
    Dim Bindings As List
    Dim Item As Variant
    Dim Expr As BindingExpression
    Dim B As Binding

    If m_RefreshBindings Then Exit Sub
    m_RefreshBindings = True

    Set Bindings = GetTargetBindings(Target)
    If Bindings Is Nothing Then GoTo Finally

    For Each Item In Bindings
        If TypeOf Item Is BindingExpression Then
            Set Expr = Item
            Expr.UpdateTarget
        ElseIf TypeOf Item Is Binding Then
            Set B = Item
            B.RefreshTarget
        End If
    Next

Finally:
    m_RefreshBindings = False
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

' Detach every binding in the visual tree while targets/sources are still valid.
Public Sub DetachBindingsTree(ByVal Root As Object)
    On Error Resume Next
    If Root Is Nothing Then Exit Sub
    
    DetachBindingsOnNode Root
    
    If TypeOf Root Is IControl Then
        Dim Ctrl As IControl
        Dim Child As Object
        Dim Kids As UIElementCollection
        Set Ctrl = Root
        ' TextBlock/Image etc. return Children = Nothing ? For Each Nothing AVs in VB6/Cairo.
        Set Kids = Ctrl.Children
        If Kids Is Nothing Then Exit Sub
        For Each Child In Kids
            DetachBindingsTree Child
        Next
    End If
End Sub

Private Sub DetachBindingsOnNode(ByVal Root As Object)
    On Error Resume Next
    
    If TypeOf Root Is IUserControl Then
        Dim Uc As IUserControl
        Set Uc = Root
        DetachTargetBindings Uc.Base
    ElseIf TypeOf Root Is IDependencyObject Then
        DetachTargetBindings Root
    End If
End Sub

Private Function GetTargetBindings(ByVal Target As IDependencyObject) As List
    On Error Resume Next
    Set GetTargetBindings = API.CObj(Target).Bindings
End Function

' Flush TwoWay/OneWayToSource bindings with UpdateSourceTrigger=LostFocus or pending delay.
Public Sub FlushLostFocusBindings(ByVal Target As IDependencyObject)
    FlushSourceBindings Target, False
End Sub

' Force-flush TwoWay/OneWayToSource bindings (TextBox Enter / LostFocus before app handlers).
Public Sub FlushSourceBindings(ByVal Target As IDependencyObject, Optional ByVal ForceAll As Boolean = True)
    Dim Bindings As List
    Dim Item As Variant
    Dim Expr As BindingExpression
    Dim B As Binding

    On Error Resume Next

    Set Bindings = GetTargetBindings(Target)
    If Bindings Is Nothing Then Exit Sub

    For Each Item In Bindings
        Set Expr = Nothing
        Set B = Nothing
        If TypeOf Item Is BindingExpression Then
            Set Expr = Item
            If Not Expr.Binding Is Nothing Then
                If ForceAll Or Expr.Binding.NeedsLostFocusFlush Then
                    Expr.Binding.FlushUpdateSource True
                End If
            End If
        ElseIf TypeOf Item Is Binding Then
            Set B = Item
            If ForceAll Or B.NeedsLostFocusFlush Then
                B.FlushUpdateSource True
            End If
        End If
    Next
End Sub
