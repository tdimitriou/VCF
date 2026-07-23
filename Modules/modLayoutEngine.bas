Attribute VB_Name = "modLayoutEngine"
Option Explicit

Public Type LayoutRect
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Public Enum GridUnitType
    GridUnitPixel = 0
    GridUnitStar = 1
    GridUnitAuto = 2
End Enum

Public Type GridLength
    Value As Double
    Unit As GridUnitType
End Type

Public Function LayoutRectFromMargin( _
    ByVal Margin As Thickness, _
    ByVal Width As Double, _
    ByVal Height As Double) As LayoutRect

    With LayoutRectFromMargin
        .Left = CSng(Margin.Left)
        .Top = CSng(Margin.Top)
        If Width > 0 Then
            .Width = CSng(Width)
        End If
        If Height > 0 Then
            .Height = CSng(Height)
        End If
    End With
End Function

Public Sub ApplyLayoutRectToElement(ByVal Element As IUIElement, R As LayoutRect)
    Element.Move R.Left, R.Top, R.Width, R.Height
End Sub

Public Function IsLayoutCollapsed(ByVal Value As Visibility) As Boolean
    IsLayoutCollapsed = (Value = VisibilityCollapsed)
End Function

Public Function IsWidgetVisible(ByVal Value As Visibility) As Boolean
    IsWidgetVisible = (Value = VisibilityVisible)
End Function

Public Function MapDesignPropertyAlias(ByVal Dep As IDependencyObject, ByVal Name As String) As String
    MapDesignPropertyAlias = Name
    Select Case LCase$(Name)
        Case "text"
            If Dep.DependencyProperties.Exists("Content") Then MapDesignPropertyAlias = "Content"
    End Select
End Function
Public Function ParseGridLength(ByVal Spec As String) As GridLength
    Dim Text As String
    Dim StarPos As Long

    Text = Trim$(Spec)
    If Len(Text) = 0 Then Text = "*"

    With ParseGridLength
        If LCase$(Text) = "auto" Then
            .Unit = GridUnitAuto
            .Value = 0#
        ElseIf Right$(Text, 1) = "*" Then
            .Unit = GridUnitStar
            StarPos = InStr(1, Text, "*")
            If StarPos = 1 Then
                .Value = 1#
            Else
                .Value = Val(Left$(Text, StarPos - 1))
                If .Value <= 0 Then .Value = 1#
            End If
        Else
            .Unit = GridUnitPixel
            .Value = CDbl(Val(Text))
        End If
    End With
End Function

Public Function ParseOrientation(ByVal Spec As String) As Orientation
    Select Case LCase$(Trim$(Spec))
        Case "horizontal", "1"
            ParseOrientation = OrientationHorizontal
        Case Else
            ParseOrientation = OrientationVertical
    End Select
End Function

Public Function ReadElementVisibility(ByVal Child As Object) As Visibility
    On Error Resume Next
    If TypeOf Child Is IControl Then
        If Child.DependencyProperties.Exists("Visibility") Then
            ReadElementVisibility = CLng(Child.DependencyProperties.GetValue("Visibility"))
            Exit Function
        End If
        If Child.DependencyProperties.Exists("Visible") Then
            ' Visible=False ? Collapsed (WPF BoolToVisibility / POS grid reclaim).
            If CBool(Child.DependencyProperties.GetValue("Visible")) Then
                ReadElementVisibility = VisibilityVisible
            Else
                ReadElementVisibility = VisibilityCollapsed
            End If
            Exit Function
        End If
    End If
    If TypeOf Child Is IUIElement Then
        Select Case TypeName(Child)
            Case "Panel", "UserControl", "StackPanel", "Grid", "ContentControl", "Border", "Button"
                ReadElementVisibility = Child.Visibility
                Exit Function
        End Select
    End If
    ReadElementVisibility = VisibilityVisible
End Function

' Parent must rearrange when a child Visibility toggles Hidden/Collapsed/Visible.
Public Sub InvalidateParentLayout(ByVal Child As Object)
    Dim ParentCtrl As IControl
    Dim ParentObj As Object
    Dim Sp As StackPanel
    Dim G As Grid
    Dim Ug As UniformGrid
    Dim P As Panel
    Dim B As Border
    Dim Cc As ContentControl

    On Error Resume Next

    If Child Is Nothing Then Exit Sub
    If Not TypeOf Child Is IControl Then Exit Sub
    Set ParentCtrl = Child.Parent
    If ParentCtrl Is Nothing Then Exit Sub
    Set ParentObj = ParentCtrl
    Err.Clear
    On Error GoTo 0

    Select Case TypeName(ParentObj)
        Case "StackPanel"
            Set Sp = ParentObj
            Sp.RelayoutChildren
        Case "Grid"
            Set G = ParentObj
            G.RelayoutChildren
        Case "UniformGrid"
            Set Ug = ParentObj
            Ug.ArrangeChildren
        Case "Panel"
            Set P = ParentObj
            P.RelayoutChildren
        Case "Border"
            Set B = ParentObj
            B.RelayoutChildren
        Case "ContentControl"
            Set Cc = ParentObj
            Cc.RelayoutChildren
    End Select
End Sub

Public Function ReadElementMargin(ByVal Child As Object) As Thickness
    On Error Resume Next
    If Child.DependencyProperties.Exists("Margin") Then
        Set ReadElementMargin = Child.DependencyProperties.GetValue("Margin")
    End If
    If ReadElementMargin Is Nothing Then
        Set ReadElementMargin = modConstructors.NewThickness(0, 0, 0, 0)
    End If
End Function

Public Function ReadElementWidth(ByVal Child As Object) As Double
    Dim V As Variant
    On Error Resume Next
    If Child.DependencyProperties.Exists("Width") Then
        Call API.CopyVariable(Child.DependencyProperties.GetValue("Width"), V)
        If Not IsNull(V) And Not IsEmpty(V) And IsNumeric(V) Then ReadElementWidth = CDbl(V)
    End If
End Function

Public Function ReadElementHeight(ByVal Child As Object) As Double
    Dim V As Variant
    On Error Resume Next
    If Child.DependencyProperties.Exists("Height") Then
        Call API.CopyVariable(Child.DependencyProperties.GetValue("Height"), V)
        If Not IsNull(V) And Not IsEmpty(V) And IsNumeric(V) Then ReadElementHeight = CDbl(V)
    End If
End Function

' MaxWidth/MaxHeight = 0 means unbounded. Apply after choosing a candidate size.
Public Function ClampElementWidth(ByVal Child As Object, ByVal Width As Double) As Double
    Dim MinW As Double
    Dim MaxW As Double

    ClampElementWidth = Width
    On Error Resume Next
    If Child Is Nothing Then Exit Function
    If Not TypeOf Child Is IDependencyObject Then Exit Function
    If Child.DependencyProperties Is Nothing Then Exit Function

    If Child.DependencyProperties.Exists("MinWidth") Then
        MinW = CDbl(Child.DependencyProperties.GetValue("MinWidth"))
        If MinW > ClampElementWidth Then ClampElementWidth = MinW
    End If
    If Child.DependencyProperties.Exists("MaxWidth") Then
        MaxW = CDbl(Child.DependencyProperties.GetValue("MaxWidth"))
        If MaxW > 0 And ClampElementWidth > MaxW Then ClampElementWidth = MaxW
    End If
End Function

Public Function ClampElementHeight(ByVal Child As Object, ByVal Height As Double) As Double
    Dim MinH As Double
    Dim MaxH As Double

    ClampElementHeight = Height
    On Error Resume Next
    If Child Is Nothing Then Exit Function
    If Not TypeOf Child Is IDependencyObject Then Exit Function
    If Child.DependencyProperties Is Nothing Then Exit Function

    If Child.DependencyProperties.Exists("MinHeight") Then
        MinH = CDbl(Child.DependencyProperties.GetValue("MinHeight"))
        If MinH > ClampElementHeight Then ClampElementHeight = MinH
    End If
    If Child.DependencyProperties.Exists("MaxHeight") Then
        MaxH = CDbl(Child.DependencyProperties.GetValue("MaxHeight"))
        If MaxH > 0 And ClampElementHeight > MaxH Then ClampElementHeight = MaxH
    End If
End Function

Public Function GetGridAttachedLong(ByVal Child As IUIElement, ByVal Key As String, Optional ByVal DefaultValue As Long = 0) As Long
    Dim Dict As ObservableDictionary

    GetGridAttachedLong = DefaultValue
    On Error Resume Next
    If Child Is Nothing Then Exit Function
    If Not Child.AttachedProperties.ContainsKey("Grid") Then Exit Function
    Set Dict = Child.AttachedProperties("Grid")
    If Dict.ContainsKey(Key) Then GetGridAttachedLong = CLng(Dict(Key))
End Function

Public Sub SetGridAttachedLong(ByVal Child As IUIElement, ByVal Key As String, ByVal Value As Long)
    Dim Dict As ObservableDictionary

    If Child Is Nothing Then Exit Sub
    If Len(Key) = 0 Then Exit Sub

    If Child.AttachedProperties.ContainsKey("Grid") Then
        Set Dict = Child.AttachedProperties("Grid")
    Else
        Set Dict = New ObservableDictionary
        Child.AttachedProperties.Add "Grid", Dict
    End If

    If Dict.ContainsKey(Key) Then
        Dict.Item(Key) = Value
    Else
        Dict.Add Key, Value
    End If
End Sub

Public Sub ApplyChildWidgetVisibility(ByVal Child As Object, ByVal Value As Visibility)
    On Error Resume Next
    If TypeOf Child Is IControl Then
        ApplyVisibility Child.Widget, Value
    End If
End Sub

Public Function ControlWidgetKey(ByVal Child As Object) As String
    ControlWidgetKey = "_" & ObjPtr(Child)
End Function

Public Sub AttachChildWidget( _
    ByVal Child As Object, _
    ByVal HostWidget As cWidgetBase, _
    ByVal ChildVis As Visibility)

    Dim Key As String

    If Not TypeOf Child Is IControl Then Exit Sub
    If Child.Widget Is Nothing Then Exit Sub
    If HostWidget Is Nothing Then Exit Sub

    Key = ControlWidgetKey(Child)
    If HostWidget.Widgets.Exists(Key) Then HostWidget.Widgets.Remove Key
    If Not HostWidget.Widgets.Exists(Key) Then
        HostWidget.Widgets.Add Child, Key, , , , , IsWidgetVisible(ChildVis)
    End If
End Sub

Public Sub DetachCollapsedChild(ByVal Child As Object, ByVal HostWidget As cWidgetBase)
    Dim Key As String

    If Not TypeOf Child Is IControl Then Exit Sub
    If HostWidget Is Nothing Then Exit Sub

    Key = ControlWidgetKey(Child)
    If HostWidget.Widgets.Exists(Key) Then HostWidget.Widgets.Remove Key
End Sub

Public Sub ArrangeStackPanelChildren( _
    ByVal Children As UIElementCollection, _
    ByVal HostWidget As cWidgetBase, _
    ByVal PanelOrientation As Orientation, _
    Optional ByVal OverrideHostWidth As Single = 0, _
    Optional ByVal OverrideHostHeight As Single = 0)

    Dim Child As Object
    Dim ChildUI As IUIElement
    Dim ChildVis As Visibility
    Dim Margin As Thickness
    Dim R As LayoutRect
    Dim HostWidth As Single
    Dim HostHeight As Single
    Dim Offset As Single
    Dim ChildWidth As Double
    Dim ChildHeight As Double

    If Children Is Nothing Then Exit Sub
    If HostWidget Is Nothing Then Exit Sub

    If OverrideHostWidth > 0 Then
        HostWidth = OverrideHostWidth
    Else
        HostWidth = HostWidget.Width
    End If
    If OverrideHostHeight > 0 Then
        HostHeight = OverrideHostHeight
    Else
        HostHeight = HostWidget.Height
    End If

    Offset = 0!

    For Each Child In Children
        If Not TypeOf Child Is IUIElement Then GoTo NextChild
        Set ChildUI = Child

        ChildVis = ReadElementVisibility(Child)
        If IsLayoutCollapsed(ChildVis) Then
            DetachCollapsedChild Child, HostWidget
            GoTo NextChild
        End If

        If Not TypeOf Child Is IControl Then GoTo NextChild

        AttachChildWidget Child, HostWidget, ChildVis

        Set Margin = ReadElementMargin(Child)
        ChildWidth = ReadElementWidth(Child)
        ChildHeight = ReadElementHeight(Child)

        If PanelOrientation = OrientationVertical Then
            R.Left = CSng(Margin.Left)
            R.Top = Offset + CSng(Margin.Top)
            If ChildWidth > 0 Then
                R.Width = CSng(ChildWidth)
            Else
                R.Width = HostWidth - CSng(Margin.Left + Margin.Right)
            End If
            If ChildHeight > 0 Then
                R.Height = CSng(ChildHeight)
            Else
                R.Height = 0!
            End If
            R.Width = CSng(ClampElementWidth(Child, R.Width))
            R.Height = CSng(ClampElementHeight(Child, R.Height))
            Offset = R.Top + R.Height + CSng(Margin.Bottom)
        Else
            R.Left = Offset + CSng(Margin.Left)
            R.Top = CSng(Margin.Top)
            If ChildWidth > 0 Then
                R.Width = CSng(ChildWidth)
            Else
                R.Width = 0!
            End If
            If ChildHeight > 0 Then
                R.Height = CSng(ChildHeight)
            Else
                R.Height = HostHeight - CSng(Margin.Top + Margin.Bottom)
            End If
            R.Width = CSng(ClampElementWidth(Child, R.Width))
            R.Height = CSng(ClampElementHeight(Child, R.Height))
            Offset = R.Left + R.Width + CSng(Margin.Right)
        End If

        ApplyLayoutRectToElement ChildUI, R
        ApplyChildWidgetVisibility Child, ChildVis

NextChild:
    Next
End Sub

Public Sub ArrangeDecoratorChild( _
    ByVal Child As Object, _
    ByVal HostWidget As cWidgetBase, _
    Optional ByVal InsetLeft As Single = 0, _
    Optional ByVal InsetTop As Single = 0, _
    Optional ByVal InsetRight As Single = 0, _
    Optional ByVal InsetBottom As Single = 0)

    Dim ChildUI As IUIElement
    Dim ChildVis As Visibility
    Dim R As LayoutRect

    If Child Is Nothing Then Exit Sub
    If HostWidget Is Nothing Then Exit Sub
    If Not TypeOf Child Is IUIElement Then Exit Sub
    Set ChildUI = Child

    ChildVis = ReadElementVisibility(Child)
    If IsLayoutCollapsed(ChildVis) Then
        DetachCollapsedChild Child, HostWidget
        Exit Sub
    End If

    If Not TypeOf Child Is IControl Then Exit Sub

    AttachChildWidget Child, HostWidget, ChildVis

    R.Left = InsetLeft
    R.Top = InsetTop
    R.Width = HostWidget.Width - InsetLeft - InsetRight
    R.Height = HostWidget.Height - InsetTop - InsetBottom
    If R.Width < 0! Then R.Width = 0!
    If R.Height < 0! Then R.Height = 0!
    R.Width = CSng(ClampElementWidth(Child, R.Width))
    R.Height = CSng(ClampElementHeight(Child, R.Height))

    ApplyLayoutRectToElement ChildUI, R
    ApplyChildWidgetVisibility Child, ChildVis
End Sub

Public Sub ArrangeGridChildren( _
    ByVal Children As UIElementCollection, _
    ByVal HostWidget As cWidgetBase, _
    ByVal RowDefinitions As ObservableCollection, _
    ByVal ColumnDefinitions As ObservableCollection, _
    Optional ByVal OverrideHostWidth As Single = 0, _
    Optional ByVal OverrideHostHeight As Single = 0)

    Dim Child As Object
    Dim ChildUI As IUIElement
    Dim ChildVis As Visibility
    Dim Margin As Thickness
    Dim R As LayoutRect
    Dim HostWidth As Single
    Dim HostHeight As Single
    Dim RowCount As Long
    Dim ColCount As Long
    Dim RowSizes() As Single
    Dim ColSizes() As Single
    Dim RowOffsets() As Single
    Dim ColOffsets() As Single

    If Children Is Nothing Then Exit Sub
    If HostWidget Is Nothing Then Exit Sub

    If OverrideHostWidth > 0 Then
        HostWidth = OverrideHostWidth
    Else
        HostWidth = HostWidget.Width
    End If
    If OverrideHostHeight > 0 Then
        HostHeight = OverrideHostHeight
    Else
        HostHeight = HostWidget.Height
    End If

    RowCount = GridTrackCount(RowDefinitions, Children, True)
    ColCount = GridTrackCount(ColumnDefinitions, Children, False)
    If RowCount < 1 Then RowCount = 1
    If ColCount < 1 Then ColCount = 1

    ComputeGridTracks RowDefinitions, RowCount, HostHeight, Children, True, RowSizes, RowOffsets
    ComputeGridTracks ColumnDefinitions, ColCount, HostWidth, Children, False, ColSizes, ColOffsets

    For Each Child In Children
        If Not TypeOf Child Is IUIElement Then GoTo NextChild
        Set ChildUI = Child

        ChildVis = ReadElementVisibility(Child)
        If IsLayoutCollapsed(ChildVis) Then
            DetachCollapsedChild Child, HostWidget
            GoTo NextChild
        End If

        If Not TypeOf Child Is IControl Then GoTo NextChild

        AttachChildWidget Child, HostWidget, ChildVis

        Dim Row As Long
        Dim Col As Long
        Dim RowSpan As Long
        Dim ColSpan As Long
        Dim i As Long
        Dim CellWidth As Single
        Dim CellHeight As Single

        Row = GetGridAttachedLong(ChildUI, "Row", 0)
        Col = GetGridAttachedLong(ChildUI, "Column", 0)
        RowSpan = GetGridAttachedLong(ChildUI, "RowSpan", 1)
        ColSpan = GetGridAttachedLong(ChildUI, "ColumnSpan", 1)
        If RowSpan < 1 Then RowSpan = 1
        If ColSpan < 1 Then ColSpan = 1
        If Row >= RowCount Then Row = RowCount - 1
        If Col >= ColCount Then Col = ColCount - 1
        If Row + RowSpan > RowCount Then RowSpan = RowCount - Row
        If Col + ColSpan > ColCount Then ColSpan = ColCount - Col

        Set Margin = ReadElementMargin(Child)

        R.Left = ColOffsets(Col) + CSng(Margin.Left)
        R.Top = RowOffsets(Row) + CSng(Margin.Top)

        CellWidth = 0!
        For i = Col To Col + ColSpan - 1
            CellWidth = CellWidth + ColSizes(i)
        Next
        CellHeight = 0!
        For i = Row To Row + RowSpan - 1
            CellHeight = CellHeight + RowSizes(i)
        Next

        R.Width = CellWidth - CSng(Margin.Left + Margin.Right)
        R.Height = CellHeight - CSng(Margin.Top + Margin.Bottom)
        If R.Width < 0! Then R.Width = 0!
        If R.Height < 0! Then R.Height = 0!
        R.Width = CSng(ClampElementWidth(Child, R.Width))
        R.Height = CSng(ClampElementHeight(Child, R.Height))

        ApplyLayoutRectToElement ChildUI, R
        ApplyChildWidgetVisibility Child, ChildVis

NextChild:
    Next
End Sub

Private Function GridTrackCount( _
    ByVal Definitions As ObservableCollection, _
    ByVal Children As UIElementCollection, _
    ByVal IsRow As Boolean) As Long

    Dim Child As Object
    Dim ChildUI As IUIElement
    Dim Track As Long
    Dim Span As Long
    Dim MaxTrack As Long

    MaxTrack = 0
    If Not Definitions Is Nothing Then MaxTrack = Definitions.Count

    For Each Child In Children
        If Not TypeOf Child Is IUIElement Then GoTo NextChild
        Set ChildUI = Child
        If IsLayoutCollapsed(ReadElementVisibility(Child)) Then GoTo NextChild

        If IsRow Then
            Track = GetGridAttachedLong(ChildUI, "Row", 0)
            Span = GetGridAttachedLong(ChildUI, "RowSpan", 1)
        Else
            Track = GetGridAttachedLong(ChildUI, "Column", 0)
            Span = GetGridAttachedLong(ChildUI, "ColumnSpan", 1)
        End If
        If Span < 1 Then Span = 1
        If Track + Span > MaxTrack Then MaxTrack = Track + Span

NextChild:
    Next

    GridTrackCount = MaxTrack
End Function

Private Sub ComputeGridTracks( _
    ByVal Definitions As ObservableCollection, _
    ByVal TrackCount As Long, _
    ByVal Available As Single, _
    ByVal Children As UIElementCollection, _
    ByVal IsRow As Boolean, _
    ByRef Sizes() As Single, _
    ByRef Offsets() As Single)

    Dim Lengths() As GridLength
    Dim i As Long
    Dim Def As Object
    Dim Spec As String
    Dim FixedTotal As Single
    Dim AutoTotal As Single
    Dim StarWeight As Double
    Dim StarTotal As Double
    Dim Remaining As Single
    Dim Offset As Single

    ReDim Lengths(0 To TrackCount - 1)
    ReDim Sizes(0 To TrackCount - 1)
    ReDim Offsets(0 To TrackCount - 1)

    For i = 0 To TrackCount - 1
        Spec = "*"
        If Not Definitions Is Nothing Then
            If i < Definitions.Count Then
                Set Def = Definitions(i)
                If IsRow Then
                    Spec = Def.Height
                Else
                    Spec = Def.Width
                End If
            End If
        End If
        Lengths(i) = ParseGridLength(Spec)
    Next

    FixedTotal = 0!
    AutoTotal = 0!
    StarTotal = 0#
    For i = 0 To TrackCount - 1
        Select Case Lengths(i).Unit
            Case GridUnitPixel
                Sizes(i) = CSng(Lengths(i).Value)
                FixedTotal = FixedTotal + Sizes(i)
            Case GridUnitAuto
                Sizes(i) = GridAutoTrackSize(Children, i, IsRow)
                AutoTotal = AutoTotal + Sizes(i)
            Case GridUnitStar
                StarTotal = StarTotal + Lengths(i).Value
        End Select
    Next

    Remaining = Available - FixedTotal - AutoTotal
    If Remaining < 0! Then Remaining = 0!

    For i = 0 To TrackCount - 1
        If Lengths(i).Unit = GridUnitStar Then
            If StarTotal > 0 Then
                Sizes(i) = CSng(Remaining * (Lengths(i).Value / StarTotal))
            Else
                Sizes(i) = CSng(Remaining / TrackCount)
            End If
        End If
    Next

    Offset = 0!
    For i = 0 To TrackCount - 1
        Offsets(i) = Offset
        Offset = Offset + Sizes(i)
    Next
End Sub

Private Function GridAutoTrackSize( _
    ByVal Children As UIElementCollection, _
    ByVal TrackIndex As Long, _
    ByVal IsRow As Boolean) As Single

    Dim Child As Object
    Dim ChildUI As IUIElement
    Dim Track As Long
    Dim Span As Long
    Dim Margin As Thickness
    Dim Desired As Single
    Dim MaxDesired As Single

    MaxDesired = 0!

    For Each Child In Children
        If Not TypeOf Child Is IUIElement Then GoTo NextChild
        Set ChildUI = Child
        If IsLayoutCollapsed(ReadElementVisibility(Child)) Then GoTo NextChild

        If IsRow Then
            Track = GetGridAttachedLong(ChildUI, "Row", 0)
            Span = GetGridAttachedLong(ChildUI, "RowSpan", 1)
            If TrackIndex < Track Or TrackIndex >= Track + Span Then GoTo NextChild
            Set Margin = ReadElementMargin(Child)
            Desired = CSng(ClampElementHeight(Child, ReadElementHeight(Child)) + Margin.Top + Margin.Bottom)
        Else
            Track = GetGridAttachedLong(ChildUI, "Column", 0)
            Span = GetGridAttachedLong(ChildUI, "ColumnSpan", 1)
            If TrackIndex < Track Or TrackIndex >= Track + Span Then GoTo NextChild
            Set Margin = ReadElementMargin(Child)
            Desired = CSng(ClampElementWidth(Child, ReadElementWidth(Child)) + Margin.Left + Margin.Right)
        End If
        If Desired > MaxDesired Then MaxDesired = Desired

NextChild:
    Next

    GridAutoTrackSize = MaxDesired
End Function
