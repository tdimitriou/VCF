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

    If Child Is Nothing Then Exit Function

    If TypeOf Child Is TextBlock Then
        Dim TbSrc As TextBlock
        Set TbSrc = Child
        Set CloneTemplateChild = CloneTextBlockQuick(TbSrc)
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
