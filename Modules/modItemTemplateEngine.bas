Attribute VB_Name = "modItemTemplateEngine"
Option Explicit

Private Function ResolveItemTemplate( _
    ByVal ItemTemplate As DataTemplate, _
    ByVal Item As Variant, _
    Optional ByVal ResourceHost As UIElementBase, _
    Optional ByVal ItemTypeName As String = vbNullString) As DataTemplate

    Dim ActiveTemplate As DataTemplate
    Dim Key As String
    Dim ResourceTemplate As Variant

    Set ActiveTemplate = ItemTemplate

    If Not ResourceHost Is Nothing Then
        If Len(ItemTypeName) = 0 Then
            If IsObject(Item) Then ItemTypeName = TypeName(Item)
        End If
        If Len(ItemTypeName) > 0 Then
            Key = "DataTemplate_" & ItemTypeName
            Call API.CopyVariable(ResourceHost.TryFindResource(Key), ResourceTemplate)
            If IsObject(ResourceTemplate) Then
                If TypeOf ResourceTemplate Is DataTemplate Then Set ActiveTemplate = ResourceTemplate
            End If
        End If
    End If

    Set ResolveItemTemplate = ActiveTemplate
End Function

Private Function CloneTemplateChild(ByVal Child As Object) As Object
    Dim Cloner As ICloneable
    Dim TbSrc As TextBlock
    Dim BtnSrc As Button

    If Child Is Nothing Then Exit Function

    If TypeOf Child Is TextBlock Then
        Set TbSrc = Child
        ' Never use TextBlock.Clone here — cProperties.BindTo copies the full COM surface and
        ' destabilizes the IDE under dense ItemTemplate clones. Copy paint props + Bindings only.
        Set CloneTemplateChild = CloneTextBlockWithBindings(TbSrc)
    ElseIf TypeOf Child Is Button Then
        Set BtnSrc = Child
        ' Same rule as TextBlock — no full COM Clone; paint + Bindings only (7c-dialog).
        Set CloneTemplateChild = CloneButtonWithBindings(BtnSrc)
    ElseIf TypeOf Child Is ICloneable Then
        Set Cloner = Child
        Set CloneTemplateChild = Cloner.Clone
    End If
End Function

Private Sub ApplyItemDataContext(ByVal Visual As Object, ByVal Item As Variant)
    Dim Element As IUIElement

    If Not IsObject(Item) Then Exit Sub
    If Visual Is Nothing Then Exit Sub
    If Not TypeOf Visual Is IUIElement Then Exit Sub

    Set Element = Visual
    Set Element.DataContext = Item
End Sub

Public Function CloneItemVisualForItem( _
    ByVal ItemTemplate As DataTemplate, _
    ByVal Item As Variant, _
    Optional ByVal ResourceHost As UIElementBase, _
    Optional ByVal ItemTypeName As String = vbNullString) As Object

    Dim ActiveTemplate As DataTemplate
    Dim Child As Object
    Dim CloneObj As Object

    Set ActiveTemplate = ResolveItemTemplate(ItemTemplate, Item, ResourceHost, ItemTypeName)
    If ActiveTemplate Is Nothing Then Exit Function
    If ActiveTemplate.Children.Count = 0 Then Exit Function

    Set Child = ActiveTemplate.Children(0)
    Set CloneObj = CloneTemplateChild(Child)
    ApplyItemDataContext CloneObj, Item
    Set CloneItemVisualForItem = CloneObj
End Function

Public Function CloneDataTemplateForItem( _
    ByVal ItemTemplate As DataTemplate, _
    ByVal Item As Variant, _
    Optional ByVal ResourceHost As UIElementBase, _
    Optional ByVal ItemTypeName As String = vbNullString) As DataTemplate

    On Error GoTo Fail

    Dim ActiveTemplate As DataTemplate
    Dim i As Long
    Dim Child As Object
    Dim CloneObj As Object

    Set ActiveTemplate = ResolveItemTemplate(ItemTemplate, Item, ResourceHost, ItemTypeName)
    If ActiveTemplate Is Nothing Then Exit Function
    If ActiveTemplate.Children.Count = 0 Then Exit Function

    Set CloneDataTemplateForItem = New DataTemplate

    For i = 0 To ActiveTemplate.Children.Count - 1
        Set Child = ActiveTemplate.Children(i)
        Set CloneObj = CloneTemplateChild(Child)
        If CloneObj Is Nothing Then GoTo NextChild
        ApplyItemDataContext CloneObj, Item
        CloneDataTemplateForItem.Children.Add CloneObj
NextChild:
    Next i

    Exit Function

Fail:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

Private Function CloneTextBlockQuick(ByVal Source As TextBlock) As TextBlock
    Dim Target As TextBlock

    Set Target = New TextBlock
    Target.Text = Source.Text
    Target.ForeColor = Source.ForeColor
    Target.FontName = Source.FontName
    Target.FontSize = Source.FontSize
    Target.FontBold = Source.FontBold
    Target.FontItalic = Source.FontItalic
    Target.FontUnderline = Source.FontUnderline
    Target.FontStrikeThrough = Source.FontStrikeThrough
    Target.HorizontalAlignment = Source.HorizontalAlignment
    Target.VerticalAlignment = Source.VerticalAlignment

    Set CloneTextBlockQuick = Target
End Function

' Paint props + Binding graph (DataContext SrcDepObj or fixed Source). Avoids TextBlock.Clone.
Private Function CloneTextBlockWithBindings(ByVal Source As TextBlock) As TextBlock
    Dim Target As TextBlock
    Dim Item As Variant
    Dim SrcB As Binding
    Dim DstB As Binding

    Set Target = CloneTextBlockQuick(Source)
    If Source.Bindings Is Nothing Then
        Set CloneTextBlockWithBindings = Target
        Exit Function
    End If
    If Source.Bindings.Count = 0 Then
        Set CloneTextBlockWithBindings = Target
        Exit Function
    End If

    On Error GoTo Fail

    For Each Item In Source.Bindings
        If Not TypeOf Item Is Binding Then GoTo NextBinding
        Set SrcB = Item
        If SrcB.TargetProperty Is Nothing Then GoTo NextBinding

        Set DstB = New Binding
        Set DstB.TargetProperty = Target.DependencyProperties.GetProperty(SrcB.TargetProperty.Name)

        If Not SrcB.SrcDepObj Is Nothing Then
            ' Markup default: Source is the DataContext DP (or other SrcDepObj).
            Set DstB.Source = Target.DependencyProperties.GetProperty(SrcB.SrcDepObj.Name)
        ElseIf Not SrcB.Source Is Nothing Then
            Set DstB.Source = SrcB.Source
        End If

        DstB.Path = SrcB.Path
        DstB.Mode = SrcB.Mode
        If Not SrcB.Converter Is Nothing Then Set DstB.Converter = SrcB.Converter
        DstB.StringFormat = SrcB.StringFormat
        Set DstB.Target = Target
        Target.Bindings.Add DstB
NextBinding:
    Next

    Set CloneTextBlockWithBindings = Target
    Exit Function

Fail:
    Err.Raise Err.Number, "modItemTemplateEngine.CloneTextBlockWithBindings", Err.Description
End Function

Private Function CloneButtonQuick(ByVal Source As Button) As Button
    Dim Target As Button
    Dim Cap As Variant

    Set Target = New Button
    Call API.CopyVariable(Source.Content, Cap)
    If IsObject(Cap) Then
        Set Target.Content = Cap
    ElseIf Not IsEmpty(Cap) And Not IsNull(Cap) Then
        Target.Content = Cap
    End If
    Target.BorderWidth = Source.BorderWidth
    Target.CornerRadius = Source.CornerRadius
    Target.Width = Source.Width
    Target.Height = Source.Height

    Set CloneButtonQuick = Target
End Function

' Chrome props + Binding graph for dialog/button ItemTemplates. Avoids full COM Clone.
Private Function CloneButtonWithBindings(ByVal Source As Button) As Button
    Dim Target As Button
    Dim Item As Variant
    Dim SrcB As Binding
    Dim DstB As Binding

    Set Target = CloneButtonQuick(Source)
    If Source.Bindings Is Nothing Then
        Set CloneButtonWithBindings = Target
        Exit Function
    End If
    If Source.Bindings.Count = 0 Then
        Set CloneButtonWithBindings = Target
        Exit Function
    End If

    On Error GoTo Fail

    For Each Item In Source.Bindings
        If Not TypeOf Item Is Binding Then GoTo NextBinding
        Set SrcB = Item
        If SrcB.TargetProperty Is Nothing Then GoTo NextBinding

        Set DstB = New Binding
        Set DstB.TargetProperty = Target.DependencyProperties.GetProperty(SrcB.TargetProperty.Name)

        If Not SrcB.SrcDepObj Is Nothing Then
            Set DstB.Source = Target.DependencyProperties.GetProperty(SrcB.SrcDepObj.Name)
        ElseIf Not SrcB.Source Is Nothing Then
            Set DstB.Source = SrcB.Source
        End If

        DstB.Path = SrcB.Path
        DstB.Mode = SrcB.Mode
        If Not SrcB.Converter Is Nothing Then Set DstB.Converter = SrcB.Converter
        DstB.StringFormat = SrcB.StringFormat
        Set DstB.Target = Target
        Target.Bindings.Add DstB
NextBinding:
    Next

    Set CloneButtonWithBindings = Target
    Exit Function

Fail:
    Err.Raise Err.Number, "modItemTemplateEngine.CloneButtonWithBindings", Err.Description
End Function

Public Sub ValidateItemsSourceValue(Value, ByVal SourceName As String)
    If IsObject(Value) Then
        If Value Is Nothing Then Exit Sub
        If Not TypeOf Value Is ObservableCollection Then
            Err.Raise vbObjectError + 4, SourceName, "ItemsSource must be an ObservableCollection"
        End If
    ElseIf Not IsEmpty(Value) And Not IsNull(Value) Then
        Err.Raise vbObjectError + 4, SourceName, "ItemsSource must be an ObservableCollection"
    End If
End Sub

' Clone ItemsPanelTemplate root for ItemsControl.ItemsHost (empty panel shell).
Public Function CloneItemsPanelRoot(ByVal Source As Object) As Object
    Dim UgSrc As UniformGrid
    Dim Ug As UniformGrid
    Dim SpSrc As StackPanel
    Dim Sp As StackPanel

    If Source Is Nothing Then Exit Function

    If TypeOf Source Is UniformGrid Then
        Set UgSrc = Source
        Set Ug = New UniformGrid
        ' Avoid Rows/Columns/Padding setters here — they call MoveChildren on an unhosted grid.
        Ug.Widget.LockRefresh = True
        Call Ug.ApplyItemsPanelMetrics(UgSrc.Rows, UgSrc.Columns, UgSrc.Padding)
        Ug.Width = UgSrc.Width
        Ug.Height = UgSrc.Height
        On Error Resume Next
        Ug.DependencyProperties.SetCurrentValue "ShowGridLines", UgSrc.DependencyProperties.GetValue("ShowGridLines")
        On Error GoTo 0
        Ug.Widget.LockRefresh = False
        Set CloneItemsPanelRoot = Ug
        Exit Function
    End If

    If TypeOf Source Is StackPanel Then
        Set SpSrc = Source
        Set Sp = New StackPanel
        Sp.Orientation = SpSrc.Orientation
        Sp.Width = SpSrc.Width
        Sp.Height = SpSrc.Height
        Set CloneItemsPanelRoot = Sp
        Exit Function
    End If
End Function
