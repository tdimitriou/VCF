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

' WPF FrameworkElement alignment in panel slots (Grid cells).
Public Enum LayoutHorizontalAlignment
    LayoutHAlignLeft = 0
    LayoutHAlignCenter = 1
    LayoutHAlignRight = 2
    LayoutHAlignStretch = 3
End Enum

Public Enum LayoutVerticalAlignment
    LayoutVAlignTop = 0
    LayoutVAlignCenter = 1
    LayoutVAlignBottom = 2
    LayoutVAlignStretch = 3
End Enum

Public Type GridLength
    Value As Double
    Unit As GridUnitType
End Type

Public Function ParseDock(ByVal Spec As String) As Dock
    Select Case LCase$(Trim$(Spec))
        Case "top", "1"
            ParseDock = DockTop
        Case "right", "2"
            ParseDock = DockRight
        Case "bottom", "3"
            ParseDock = DockBottom
        Case Else
            ' left, 0, empty, unknown
            ParseDock = DockLeft
    End Select
End Function

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
            Case "Panel", "UserControl", "StackPanel", "Grid", "DockPanel", "Canvas", "ContentControl", "Border", "Button"
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
    Dim Dp As DockPanel
    Dim Cv As Canvas
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
        Case "DockPanel"
            Set Dp = ParentObj
            Dp.RelayoutChildren
        Case "Canvas"
            Set Cv = ParentObj
            Cv.RelayoutChildren
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
    Dim Dep As IDependencyObject
    Dim FullName As String
    Dim V As Variant

    GetGridAttachedLong = DefaultValue
    On Error Resume Next
    If Child Is Nothing Then Exit Function

    FullName = "Grid." & Key
    If TypeOf Child Is IDependencyObject Then
        Set Dep = Child
        If Not Dep.DependencyProperties Is Nothing Then
            If Dep.DependencyProperties.Exists(FullName) Then
                V = Dep.DependencyProperties.GetValue(FullName)
                If IsNumeric(V) Then GetGridAttachedLong = CLng(V)
                Exit Function
            End If
        End If
    End If

    If Not Child.AttachedProperties.ContainsKey("Grid") Then Exit Function
    Set Dict = Child.AttachedProperties("Grid")
    If Dict.ContainsKey(Key) Then GetGridAttachedLong = CLng(Dict(Key))
End Function

Public Function ParseLayoutHorizontalAlignment(ByVal Spec As String) As LayoutHorizontalAlignment
    Select Case LCase$(Trim$(Spec))
        Case "left": ParseLayoutHorizontalAlignment = LayoutHAlignLeft
        Case "center": ParseLayoutHorizontalAlignment = LayoutHAlignCenter
        Case "right": ParseLayoutHorizontalAlignment = LayoutHAlignRight
        Case Else: ParseLayoutHorizontalAlignment = LayoutHAlignStretch
    End Select
End Function

Public Function ParseLayoutVerticalAlignment(ByVal Spec As String) As LayoutVerticalAlignment
    Select Case LCase$(Trim$(Spec))
        Case "top": ParseLayoutVerticalAlignment = LayoutVAlignTop
        Case "center": ParseLayoutVerticalAlignment = LayoutVAlignCenter
        Case "bottom": ParseLayoutVerticalAlignment = LayoutVAlignBottom
        Case Else: ParseLayoutVerticalAlignment = LayoutVAlignStretch
    End Select
End Function

' TextBlock.HorizontalAlignment / VerticalAlignment are text-box aligns, not
' FrameworkElement layout aligns. Always Stretch in panel slots (WPF: layout
' Stretch by default; TextAlignment is separate).
Public Function ReadLayoutHorizontalAlignment(ByVal Child As Object) As LayoutHorizontalAlignment
    Dim V As Variant

    ReadLayoutHorizontalAlignment = LayoutHAlignStretch
    On Error Resume Next
    If Child Is Nothing Then Exit Function
    If TypeOf Child Is TextBlock Then Exit Function
    If Not TypeOf Child Is IDependencyObject Then Exit Function
    If Child.DependencyProperties Is Nothing Then Exit Function
    If Not Child.DependencyProperties.Exists("HorizontalAlignment") Then Exit Function

    V = Child.DependencyProperties.GetValue("HorizontalAlignment")
    If VarType(V) = vbString Then
        ReadLayoutHorizontalAlignment = ParseLayoutHorizontalAlignment(CStr(V))
    ElseIf IsNumeric(V) Then
        Select Case CLng(V)
            Case LayoutHAlignLeft, LayoutHAlignCenter, LayoutHAlignRight, LayoutHAlignStretch
                ReadLayoutHorizontalAlignment = CLng(V)
            Case Else
                ReadLayoutHorizontalAlignment = LayoutHAlignStretch
        End Select
    End If
End Function

Public Function ReadLayoutVerticalAlignment(ByVal Child As Object) As LayoutVerticalAlignment
    Dim V As Variant

    ReadLayoutVerticalAlignment = LayoutVAlignStretch
    On Error Resume Next
    If Child Is Nothing Then Exit Function
    If TypeOf Child Is TextBlock Then Exit Function
    If Not TypeOf Child Is IDependencyObject Then Exit Function
    If Child.DependencyProperties Is Nothing Then Exit Function
    If Not Child.DependencyProperties.Exists("VerticalAlignment") Then Exit Function

    V = Child.DependencyProperties.GetValue("VerticalAlignment")
    If VarType(V) = vbString Then
        ReadLayoutVerticalAlignment = ParseLayoutVerticalAlignment(CStr(V))
    ElseIf IsNumeric(V) Then
        Select Case CLng(V)
            Case LayoutVAlignTop, LayoutVAlignCenter, LayoutVAlignBottom, LayoutVAlignStretch
                ReadLayoutVerticalAlignment = CLng(V)
            Case Else
                ReadLayoutVerticalAlignment = LayoutVAlignStretch
        End Select
    End If
End Function

Public Function AlignElementInSlot( _
    ByVal Child As Object, _
    ByVal SlotLeft As Single, _
    ByVal SlotTop As Single, _
    ByVal SlotWidth As Single, _
    ByVal SlotHeight As Single, _
    ByVal HAlign As LayoutHorizontalAlignment, _
    ByVal VAlign As LayoutVerticalAlignment) As LayoutRect

    Dim Measured As LayoutRect
    Dim ChildW As Double
    Dim ChildH As Double
    Dim AvailW As Double
    Dim AvailH As Double

    AvailW = CDbl(SlotWidth)
    AvailH = CDbl(SlotHeight)
    If AvailW < 0# Then AvailW = 0#
    If AvailH < 0# Then AvailH = 0#

    AlignElementInSlot.Left = SlotLeft
    AlignElementInSlot.Top = SlotTop
    AlignElementInSlot.Width = SlotWidth
    AlignElementInSlot.Height = SlotHeight

    ' Stretch + unset size → fill slot. Stretch + explicit Width/Height → keep
    ' author size (WPF); do not expand fixed-size children to the cell.
    If HAlign = LayoutHAlignStretch Then
        ChildW = ReadElementWidth(Child)
        If ChildW <= 0# Then
            AlignElementInSlot.Width = CSng(ClampElementWidth(Child, AvailW))
        Else
            ChildW = ClampElementWidth(Child, ChildW)
            If ChildW > AvailW Then ChildW = AvailW
            AlignElementInSlot.Width = CSng(ChildW)
            AlignElementInSlot.Left = SlotLeft
        End If
    Else
        Measured = MeasureElementSize(Child, AvailW, 0#)
        ChildW = CDbl(Measured.Width)
        If ChildW <= 0# Then ChildW = ReadElementWidth(Child)
        If ChildW <= 0# Then ChildW = AvailW
        ChildW = ClampElementWidth(Child, ChildW)
        If ChildW > AvailW Then ChildW = AvailW
        AlignElementInSlot.Width = CSng(ChildW)
        Select Case HAlign
            Case LayoutHAlignCenter
                AlignElementInSlot.Left = SlotLeft + CSng((AvailW - ChildW) / 2#)
            Case LayoutHAlignRight
                AlignElementInSlot.Left = SlotLeft + CSng(AvailW - ChildW)
            Case Else
                AlignElementInSlot.Left = SlotLeft
        End Select
    End If

    If VAlign = LayoutVAlignStretch Then
        ChildH = ReadElementHeight(Child)
        If ChildH <= 0# Then
            AlignElementInSlot.Height = CSng(ClampElementHeight(Child, AvailH))
        Else
            ChildH = ClampElementHeight(Child, ChildH)
            If ChildH > AvailH Then ChildH = AvailH
            AlignElementInSlot.Height = CSng(ChildH)
            AlignElementInSlot.Top = SlotTop
        End If
    Else
        Measured = MeasureElementSize(Child, 0#, AvailH)
        ChildH = CDbl(Measured.Height)
        If ChildH <= 0# Then ChildH = ReadElementHeight(Child)
        If ChildH <= 0# Then ChildH = AvailH
        ChildH = ClampElementHeight(Child, ChildH)
        If ChildH > AvailH Then ChildH = AvailH
        AlignElementInSlot.Height = CSng(ChildH)
        Select Case VAlign
            Case LayoutVAlignCenter
                AlignElementInSlot.Top = SlotTop + CSng((AvailH - ChildH) / 2#)
            Case LayoutVAlignBottom
                AlignElementInSlot.Top = SlotTop + CSng(AvailH - ChildH)
            Case Else
                AlignElementInSlot.Top = SlotTop
        End Select
    End If
End Function

Public Sub SetGridAttachedLong(ByVal Child As IUIElement, ByVal Key As String, ByVal Value As Long)
    Dim Dict As ObservableDictionary
    Dim Dep As IDependencyObject
    Dim FullName As String

    If Child Is Nothing Then Exit Sub
    If Len(Key) = 0 Then Exit Sub

    ' Nested-dict shim (XAML writer / legacy readers).
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

    ' Per-element DP bag (ClearValue / GetValue / future binding).
    On Error Resume Next
    If TypeOf Child Is IDependencyObject Then
        Set Dep = Child
        FullName = "Grid." & Key
        EnsureAttachedProperty Dep, FullName
        If Dep.DependencyProperties.Exists(FullName) Then
            Dep.DependencyProperties.SetValue FullName, Value
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Public Function GetDockAttachedLong(ByVal Child As IUIElement, Optional ByVal DefaultValue As Long = DockLeft) As Long
    Dim Dict As ObservableDictionary
    Dim Dep As IDependencyObject
    Dim FullName As String
    Dim V As Variant

    GetDockAttachedLong = DefaultValue
    On Error Resume Next
    If Child Is Nothing Then Exit Function

    FullName = "DockPanel.Dock"
    If TypeOf Child Is IDependencyObject Then
        Set Dep = Child
        If Not Dep.DependencyProperties Is Nothing Then
            If Dep.DependencyProperties.Exists(FullName) Then
                V = Dep.DependencyProperties.GetValue(FullName)
                If VarType(V) = vbString Then
                    GetDockAttachedLong = ParseDock(CStr(V))
                ElseIf IsNumeric(V) Then
                    GetDockAttachedLong = CLng(V)
                End If
                Exit Function
            End If
        End If
    End If

    If Not Child.AttachedProperties.ContainsKey("DockPanel") Then Exit Function
    Set Dict = Child.AttachedProperties("DockPanel")
    If Not Dict.ContainsKey("Dock") Then Exit Function
    V = Dict("Dock")
    If VarType(V) = vbString Then
        GetDockAttachedLong = ParseDock(CStr(V))
    ElseIf IsNumeric(V) Then
        GetDockAttachedLong = CLng(V)
    End If
End Function

Public Sub SetDockAttachedLong(ByVal Child As IUIElement, ByVal Value As Long)
    Dim Dict As ObservableDictionary
    Dim Dep As IDependencyObject
    Dim FullName As String

    If Child Is Nothing Then Exit Sub

    If Child.AttachedProperties.ContainsKey("DockPanel") Then
        Set Dict = Child.AttachedProperties("DockPanel")
    Else
        Set Dict = New ObservableDictionary
        Child.AttachedProperties.Add "DockPanel", Dict
    End If

    If Dict.ContainsKey("Dock") Then
        Dict.Item("Dock") = Value
    Else
        Dict.Add "Dock", Value
    End If

    On Error Resume Next
    If TypeOf Child Is IDependencyObject Then
        Set Dep = Child
        FullName = "DockPanel.Dock"
        EnsureAttachedProperty Dep, FullName
        If Dep.DependencyProperties.Exists(FullName) Then
            Dep.DependencyProperties.SetValue FullName, Value
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub

Public Function GetCanvasAttachedDouble(ByVal Child As IUIElement, ByVal Key As String, Optional ByVal DefaultValue As Double = 0#) As Double
    Dim Dict As ObservableDictionary
    Dim Dep As IDependencyObject
    Dim FullName As String
    Dim V As Variant

    GetCanvasAttachedDouble = DefaultValue
    On Error Resume Next
    If Child Is Nothing Then Exit Function

    FullName = "Canvas." & Key
    If TypeOf Child Is IDependencyObject Then
        Set Dep = Child
        If Not Dep.DependencyProperties Is Nothing Then
            If Dep.DependencyProperties.Exists(FullName) Then
                V = Dep.DependencyProperties.GetValue(FullName)
                If IsNumeric(V) Then GetCanvasAttachedDouble = CDbl(V)
                Exit Function
            End If
        End If
    End If

    If Not Child.AttachedProperties.ContainsKey("Canvas") Then Exit Function
    Set Dict = Child.AttachedProperties("Canvas")
    If Not Dict.ContainsKey(Key) Then Exit Function
    V = Dict(Key)
    If IsNumeric(V) Then GetCanvasAttachedDouble = CDbl(V)
End Function

Public Sub SetCanvasAttachedDouble(ByVal Child As IUIElement, ByVal Key As String, ByVal Value As Double)
    Dim Dict As ObservableDictionary
    Dim Dep As IDependencyObject
    Dim FullName As String

    If Child Is Nothing Then Exit Sub
    If Len(Key) = 0 Then Exit Sub

    If Child.AttachedProperties.ContainsKey("Canvas") Then
        Set Dict = Child.AttachedProperties("Canvas")
    Else
        Set Dict = New ObservableDictionary
        Child.AttachedProperties.Add "Canvas", Dict
    End If

    If Dict.ContainsKey(Key) Then
        Dict.Item(Key) = Value
    Else
        Dict.Add Key, Value
    End If

    On Error Resume Next
    If TypeOf Child Is IDependencyObject Then
        Set Dep = Child
        FullName = "Canvas." & Key
        EnsureAttachedProperty Dep, FullName
        If Dep.DependencyProperties.Exists(FullName) Then
            Dep.DependencyProperties.SetValue FullName, Value
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Sub

' Lazy-register attached DP on the target (do not eager-register on every instance).
Public Sub EnsureAttachedProperty(ByVal Target As IDependencyObject, ByVal FullName As String)
    Dim Reg As DependencyPropertyRegistry
    Dim Def As Variant
    Dim Meta As DependencyPropertyMetadata
    Dim PropType As VbVarType

    If Target Is Nothing Then Exit Sub
    If Len(FullName) = 0 Then Exit Sub
    If Target.DependencyProperties Is Nothing Then Exit Sub
    If Target.DependencyProperties.Exists(FullName) Then Exit Sub

    Set Reg = modStaticClasses.DependencyPropertyRegistry
    Reg.EnsureBuiltInTypes
    If Not Reg.IsAttachedRegistered(FullName) Then Exit Sub

    Def = Reg.GetAttachedDefault(FullName)
    PropType = Reg.GetAttachedPropertyType(FullName)
    Set Meta = NewDependencyPropertyMetadata(True, False, False, OneWay, Def)
    Target.DependencyProperties.Register FullName, PropType, , , , Meta
End Sub

' Binding / XAML: if TargetProperty is a registered attached name (Grid.Row, …),
' ensure it exists on the bag before Exists/GetProperty.
Public Sub EnsureAttachedTargetIfRegistered(ByVal Target As IDependencyObject, ByVal FullName As String)
    Dim Reg As DependencyPropertyRegistry

    If Target Is Nothing Then Exit Sub
    If InStr(FullName, ".") = 0 Then Exit Sub

    Set Reg = modStaticClasses.DependencyPropertyRegistry
    Reg.EnsureBuiltInTypes
    If Reg.IsAttachedRegistered(FullName) Then
        EnsureAttachedProperty Target, FullName
    End If
End Sub

' Attached DPs that affect parent panel geometry (binding / SetValue must RelayoutChildren).
Public Function IsAttachedLayoutProperty(ByVal FullName As String) As Boolean
    Select Case LCase$(FullName)
        Case "grid.row", "grid.column", "grid.rowspan", "grid.columnspan", _
             "dockpanel.dock", "canvas.left", "canvas.top"
            IsAttachedLayoutProperty = True
        Case Else
            IsAttachedLayoutProperty = False
    End Select
End Function

Public Sub ApplyChildWidgetVisibility(ByVal Child As Object, ByVal Value As Visibility)
    Dim ChildControl As IControl

    On Error Resume Next
    If Not TypeOf Child Is IControl Then Exit Sub
    Set ChildControl = Child
    ApplyVisibility ChildControl.Widget, Value
End Sub

Public Function ControlWidgetKey(ByVal Child As Object) As String
    ControlWidgetKey = "_" & ObjPtr(Child)
End Function

Public Sub AttachChildWidget( _
    ByVal Child As Object, _
    ByVal HostWidget As cWidgetBase, _
    ByVal ChildVis As Visibility)

    Dim Key As String
    Dim ChildControl As IControl

    If Not TypeOf Child Is IControl Then Exit Sub
    Set ChildControl = Child
    If ChildControl.Widget Is Nothing Then Exit Sub
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

Public Function MeasureElementSize( _
    ByVal Child As Object, _
    ByVal AvailableWidth As Double, _
    ByVal AvailableHeight As Double) As LayoutRect

    Dim W As Double
    Dim H As Double

    MeasureElementSize.Left = 0!
    MeasureElementSize.Top = 0!
    MeasureElementSize.Width = 0!
    MeasureElementSize.Height = 0!

    If Child Is Nothing Then Exit Function

    ' Prefer MeasureOverride (WPF name); fall back to MeasureLayout alias.
    On Error Resume Next
    CallByName Child, "MeasureOverride", VbMethod, AvailableWidth, AvailableHeight
    If Err.Number <> 0 Then
        Err.Clear
        CallByName Child, "MeasureLayout", VbMethod, AvailableWidth, AvailableHeight
    End If
    If Err.Number = 0 Then
        Err.Clear
        W = CDbl(CallByName(Child, "DesiredWidth", VbGet))
        If Err.Number <> 0 Then W = 0#
        Err.Clear
        H = CDbl(CallByName(Child, "DesiredHeight", VbGet))
        If Err.Number <> 0 Then H = 0#
        Err.Clear
        On Error GoTo 0
        If W <= 0# Then W = ReadElementWidth(Child)
        If H <= 0# Then H = ReadElementHeight(Child)
        MeasureElementSize.Width = CSng(ClampElementWidth(Child, W))
        MeasureElementSize.Height = CSng(ClampElementHeight(Child, H))
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0

    ' Leaf / no MeasureOverride: explicit Width/Height DPs only (unset stays 0 —
    ' do not expand to Available*, which would break StackPanel Auto sizing).
    W = ReadElementWidth(Child)
    H = ReadElementHeight(Child)
    MeasureElementSize.Width = CSng(ClampElementWidth(Child, W))
    MeasureElementSize.Height = CSng(ClampElementHeight(Child, H))
End Function

' Content-driven stack measure (no widget Move). Vertical: children get height=0
' available so unset Height stays 0; cross-axis uses AvailableWidth as stretch slot.
Public Sub MeasureStackPanelContent( _
    ByVal Children As UIElementCollection, _
    ByVal PanelOrientation As Orientation, _
    ByVal AvailableWidth As Double, _
    ByVal AvailableHeight As Double, _
    ByRef OutContentWidth As Double, _
    ByRef OutContentHeight As Double)

    Dim Child As Object
    Dim ChildVis As Visibility
    Dim Margin As Thickness
    Dim Measured As LayoutRect
    Dim ContentW As Double
    Dim ContentH As Double
    Dim SlotW As Double
    Dim SlotH As Double
    Dim Offset As Double

    OutContentWidth = 0#
    OutContentHeight = 0#
    ContentW = 0#
    ContentH = 0#
    Offset = 0#

    If Children Is Nothing Then Exit Sub

    For Each Child In Children
        ChildVis = ReadElementVisibility(Child)
        If IsLayoutCollapsed(ChildVis) Then GoTo NextMeasureChild

        Set Margin = ReadElementMargin(Child)

        If PanelOrientation = OrientationVertical Then
            SlotW = AvailableWidth - Margin.Left - Margin.Right
            If SlotW < 0# Then SlotW = 0#
            Measured = MeasureElementSize(Child, SlotW, 0#)
            Offset = Offset + Margin.Top + CDbl(Measured.Height) + Margin.Bottom
            If (CDbl(Measured.Width) + Margin.Left + Margin.Right) > ContentW Then
                ContentW = CDbl(Measured.Width) + Margin.Left + Margin.Right
            End If
            ContentH = Offset
        Else
            SlotH = AvailableHeight - Margin.Top - Margin.Bottom
            If SlotH < 0# Then SlotH = 0#
            Measured = MeasureElementSize(Child, 0#, SlotH)
            Offset = Offset + Margin.Left + CDbl(Measured.Width) + Margin.Right
            ContentW = Offset
            If (CDbl(Measured.Height) + Margin.Top + Margin.Bottom) > ContentH Then
                ContentH = CDbl(Measured.Height) + Margin.Top + Margin.Bottom
            End If
        End If

NextMeasureChild:
    Next

    OutContentWidth = ContentW
    OutContentHeight = ContentH
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
    Dim Measured As LayoutRect
    Dim HostWidth As Single
    Dim HostHeight As Single
    Dim Offset As Single
    Dim ChildWidth As Double
    Dim ChildHeight As Double
    Dim SlotW As Double
    Dim SlotH As Double

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

        If PanelOrientation = OrientationVertical Then
            SlotW = CDbl(HostWidth) - Margin.Left - Margin.Right
            If SlotW < 0# Then SlotW = 0#
            ' Infinite-height measure for stack children: pass 0 available height
            ' so MeasureLayout/leaf path uses explicit Height only.
            Measured = MeasureElementSize(Child, SlotW, 0#)
            ChildWidth = Measured.Width
            ChildHeight = Measured.Height

            R.Left = CSng(Margin.Left)
            R.Top = Offset + CSng(Margin.Top)
            If ChildWidth > 0 Then
                R.Width = CSng(ChildWidth)
            Else
                R.Width = HostWidth - CSng(Margin.Left + Margin.Right)
            End If
            R.Height = CSng(ChildHeight)
            R.Width = CSng(ClampElementWidth(Child, R.Width))
            R.Height = CSng(ClampElementHeight(Child, R.Height))
            Offset = R.Top + R.Height + CSng(Margin.Bottom)
        Else
            SlotH = CDbl(HostHeight) - Margin.Top - Margin.Bottom
            If SlotH < 0# Then SlotH = 0#
            Measured = MeasureElementSize(Child, 0#, SlotH)
            ChildWidth = Measured.Width
            ChildHeight = Measured.Height

            R.Left = Offset + CSng(Margin.Left)
            R.Top = CSng(Margin.Top)
            R.Width = CSng(ChildWidth)
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

' Single-child Border/ContentControl measure: child Desired + insets (child Margin).
Public Sub MeasureDecoratorContent( _
    ByVal Child As Object, _
    ByVal AvailableWidth As Double, _
    ByVal AvailableHeight As Double, _
    ByVal InsetLeft As Double, _
    ByVal InsetTop As Double, _
    ByVal InsetRight As Double, _
    ByVal InsetBottom As Double, _
    ByRef OutContentWidth As Double, _
    ByRef OutContentHeight As Double)

    Dim SlotW As Double
    Dim SlotH As Double
    Dim Measured As LayoutRect
    Dim ChildVis As Visibility

    OutContentWidth = 0#
    OutContentHeight = 0#

    If Child Is Nothing Then Exit Sub

    ChildVis = ReadElementVisibility(Child)
    If IsLayoutCollapsed(ChildVis) Then Exit Sub

    SlotW = AvailableWidth - InsetLeft - InsetRight
    SlotH = AvailableHeight - InsetTop - InsetBottom
    If SlotW < 0# Then SlotW = 0#
    If SlotH < 0# Then SlotH = 0#

    Measured = MeasureElementSize(Child, SlotW, SlotH)
    OutContentWidth = CDbl(Measured.Width) + InsetLeft + InsetRight
    OutContentHeight = CDbl(Measured.Height) + InsetTop + InsetBottom
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

        CellWidth = 0!
        For i = Col To Col + ColSpan - 1
            CellWidth = CellWidth + ColSizes(i)
        Next
        CellHeight = 0!
        For i = Row To Row + RowSpan - 1
            CellHeight = CellHeight + RowSizes(i)
        Next

        Dim SlotL As Single
        Dim SlotT As Single
        Dim SlotW As Single
        Dim SlotH As Single
        Dim HAlign As LayoutHorizontalAlignment
        Dim VAlign As LayoutVerticalAlignment

        SlotL = ColOffsets(Col) + CSng(Margin.Left)
        SlotT = RowOffsets(Row) + CSng(Margin.Top)
        SlotW = CellWidth - CSng(Margin.Left + Margin.Right)
        SlotH = CellHeight - CSng(Margin.Top + Margin.Bottom)
        If SlotW < 0! Then SlotW = 0!
        If SlotH < 0! Then SlotH = 0!

        HAlign = ReadLayoutHorizontalAlignment(Child)
        VAlign = ReadLayoutVerticalAlignment(Child)
        R = AlignElementInSlot(Child, SlotL, SlotT, SlotW, SlotH, HAlign, VAlign)

        ApplyLayoutRectToElement ChildUI, R
        ApplyChildWidgetVisibility Child, ChildVis

NextChild:
    Next
End Sub

' Content-driven grid measure (no widget Move). Reuses ComputeGridTracks so
' Desired* matches what arrange would allocate for the same Available*.
Public Sub MeasureGridContent( _
    ByVal Children As UIElementCollection, _
    ByVal RowDefinitions As ObservableCollection, _
    ByVal ColumnDefinitions As ObservableCollection, _
    ByVal AvailableWidth As Double, _
    ByVal AvailableHeight As Double, _
    ByRef OutContentWidth As Double, _
    ByRef OutContentHeight As Double)

    Dim RowCount As Long
    Dim ColCount As Long
    Dim RowSizes() As Single
    Dim ColSizes() As Single
    Dim RowOffsets() As Single
    Dim ColOffsets() As Single
    Dim i As Long
    Dim ContentW As Double
    Dim ContentH As Double

    OutContentWidth = 0#
    OutContentHeight = 0#
    ContentW = 0#
    ContentH = 0#

    If Children Is Nothing Then Exit Sub

    RowCount = GridTrackCount(RowDefinitions, Children, True)
    ColCount = GridTrackCount(ColumnDefinitions, Children, False)
    If RowCount < 1 Then RowCount = 1
    If ColCount < 1 Then ColCount = 1

    ComputeGridTracks RowDefinitions, RowCount, CSng(AvailableHeight), Children, True, RowSizes, RowOffsets
    ComputeGridTracks ColumnDefinitions, ColCount, CSng(AvailableWidth), Children, False, ColSizes, ColOffsets

    For i = 0 To RowCount - 1
        ContentH = ContentH + CDbl(RowSizes(i))
    Next
    For i = 0 To ColCount - 1
        ContentW = ContentW + CDbl(ColSizes(i))
    Next

    OutContentWidth = ContentW
    OutContentHeight = ContentH
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
    Dim Measured As LayoutRect
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
            ' Infinite available on both axes: Auto uses Desired (explicit DP or MeasureLayout).
            Measured = MeasureElementSize(Child, 0#, 0#)
            Desired = CSng(ClampElementHeight(Child, CDbl(Measured.Height)) + Margin.Top + Margin.Bottom)
        Else
            Track = GetGridAttachedLong(ChildUI, "Column", 0)
            Span = GetGridAttachedLong(ChildUI, "ColumnSpan", 1)
            If TrackIndex < Track Or TrackIndex >= Track + Span Then GoTo NextChild
            Set Margin = ReadElementMargin(Child)
            Measured = MeasureElementSize(Child, 0#, 0#)
            Desired = CSng(ClampElementWidth(Child, CDbl(Measured.Width)) + Margin.Left + Margin.Right)
        End If
        If Desired > MaxDesired Then MaxDesired = Desired

NextChild:
    Next

    GridAutoTrackSize = MaxDesired
End Function

' Content-driven DockPanel measure (no widget Move).
' Width  = Left+Right docked + max(Top/Bottom/fill widths)
' Height = Top+Bottom docked + max(Left/Right/fill heights)
Public Sub MeasureDockPanelContent( _
    ByVal Children As UIElementCollection, _
    ByVal LastChildFill As Boolean, _
    ByVal AvailableWidth As Double, _
    ByVal AvailableHeight As Double, _
    ByRef OutContentWidth As Double, _
    ByRef OutContentHeight As Double)

    Dim Child As Object
    Dim ChildUI As IUIElement
    Dim Kids() As Object
    Dim N As Long
    Dim i As Long
    Dim ChildVis As Visibility
    Dim Margin As Thickness
    Dim Measured As LayoutRect
    Dim ChildW As Double
    Dim ChildH As Double
    Dim AccLR As Double
    Dim AccTB As Double
    Dim MaxCrossLR As Double
    Dim MaxCrossTB As Double
    Dim FillW As Double
    Dim FillH As Double
    Dim DockSide As Dock
    Dim IsFill As Boolean
    Dim SlotW As Double
    Dim SlotH As Double

    OutContentWidth = 0#
    OutContentHeight = 0#
    AccLR = 0#
    AccTB = 0#
    MaxCrossLR = 0#
    MaxCrossTB = 0#
    FillW = 0#
    FillH = 0#
    N = 0

    If Children Is Nothing Then Exit Sub

    For Each Child In Children
        ReDim Preserve Kids(0 To N)
        Set Kids(N) = Child
        N = N + 1
    Next

    For i = 0 To N - 1
        Set Child = Kids(i)
        If Not TypeOf Child Is IUIElement Then GoTo NextMeasureDock
        Set ChildUI = Child
        ChildVis = ReadElementVisibility(Child)
        If IsLayoutCollapsed(ChildVis) Then GoTo NextMeasureDock

        Set Margin = ReadElementMargin(Child)
        IsFill = LastChildFill And (i = N - 1)

        SlotW = AvailableWidth - Margin.Left - Margin.Right
        SlotH = AvailableHeight - Margin.Top - Margin.Bottom
        If SlotW < 0# Then SlotW = 0#
        If SlotH < 0# Then SlotH = 0#
        ' Unconstrained when available is 0: leaf uses explicit Width/Height only.
        Measured = MeasureElementSize(Child, SlotW, SlotH)
        ChildW = CDbl(Measured.Width) + Margin.Left + Margin.Right
        ChildH = CDbl(Measured.Height) + Margin.Top + Margin.Bottom

        If IsFill Then
            FillW = ChildW
            FillH = ChildH
        Else
            DockSide = GetDockAttachedLong(ChildUI, DockLeft)
            Select Case DockSide
                Case DockTop, DockBottom
                    AccTB = AccTB + ChildH
                    If ChildW > MaxCrossTB Then MaxCrossTB = ChildW
                Case Else
                    AccLR = AccLR + ChildW
                    If ChildH > MaxCrossLR Then MaxCrossLR = ChildH
            End Select
        End If

NextMeasureDock:
    Next

    If FillW > MaxCrossTB Then
        OutContentWidth = AccLR + FillW
    Else
        OutContentWidth = AccLR + MaxCrossTB
    End If

    If FillH > MaxCrossLR Then
        OutContentHeight = AccTB + FillH
    Else
        OutContentHeight = AccTB + MaxCrossLR
    End If
End Sub

Public Sub ArrangeDockPanelChildren( _
    ByVal Children As UIElementCollection, _
    ByVal HostWidget As cWidgetBase, _
    ByVal LastChildFill As Boolean, _
    Optional ByVal OverrideHostWidth As Single = 0, _
    Optional ByVal OverrideHostHeight As Single = 0)

    Dim Child As Object
    Dim ChildUI As IUIElement
    Dim Kids() As Object
    Dim N As Long
    Dim i As Long
    Dim ChildVis As Visibility
    Dim Margin As Thickness
    Dim R As LayoutRect
    Dim Measured As LayoutRect
    Dim HostWidth As Single
    Dim HostHeight As Single
    Dim RemL As Double
    Dim RemT As Double
    Dim RemW As Double
    Dim RemH As Double
    Dim ChildW As Double
    Dim ChildH As Double
    Dim SlotW As Double
    Dim SlotH As Double
    Dim DockSide As Dock
    Dim IsFill As Boolean

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

    RemL = 0#
    RemT = 0#
    RemW = CDbl(HostWidth)
    RemH = CDbl(HostHeight)
    N = 0

    For Each Child In Children
        ReDim Preserve Kids(0 To N)
        Set Kids(N) = Child
        N = N + 1
    Next

    For i = 0 To N - 1
        Set Child = Kids(i)
        If Not TypeOf Child Is IUIElement Then GoTo NextArrangeDock
        Set ChildUI = Child

        ChildVis = ReadElementVisibility(Child)
        If IsLayoutCollapsed(ChildVis) Then
            DetachCollapsedChild Child, HostWidget
            GoTo NextArrangeDock
        End If

        If Not TypeOf Child Is IControl Then GoTo NextArrangeDock

        AttachChildWidget Child, HostWidget, ChildVis
        Set Margin = ReadElementMargin(Child)
        IsFill = LastChildFill And (i = N - 1)

        SlotW = RemW - Margin.Left - Margin.Right
        SlotH = RemH - Margin.Top - Margin.Bottom
        If SlotW < 0# Then SlotW = 0#
        If SlotH < 0# Then SlotH = 0#

        If IsFill Then
            R.Left = CSng(RemL + Margin.Left)
            R.Top = CSng(RemT + Margin.Top)
            ChildW = ReadElementWidth(Child)
            ChildH = ReadElementHeight(Child)
            If ChildW <= 0# Then ChildW = SlotW Else If ChildW > SlotW Then ChildW = SlotW
            If ChildH <= 0# Then ChildH = SlotH Else If ChildH > SlotH Then ChildH = SlotH
            R.Width = CSng(ClampElementWidth(Child, ChildW))
            R.Height = CSng(ClampElementHeight(Child, ChildH))
            ApplyLayoutRectToElement ChildUI, R
            ApplyChildWidgetVisibility Child, ChildVis
            GoTo NextArrangeDock
        End If

        DockSide = GetDockAttachedLong(ChildUI, DockLeft)
        Measured = MeasureElementSize(Child, SlotW, SlotH)
        ChildW = CDbl(Measured.Width)
        ChildH = CDbl(Measured.Height)

        Select Case DockSide
            Case DockTop
                If ChildW <= 0# Then ChildW = SlotW
                If ChildH <= 0# Then ChildH = SlotH
                If ChildW > SlotW Then ChildW = SlotW
                If ChildH > SlotH Then ChildH = SlotH
                ' Stretch horizontally in remaining width.
                ChildW = SlotW
                R.Left = CSng(RemL + Margin.Left)
                R.Top = CSng(RemT + Margin.Top)
                R.Width = CSng(ClampElementWidth(Child, ChildW))
                R.Height = CSng(ClampElementHeight(Child, ChildH))
                RemT = RemT + Margin.Top + CDbl(R.Height) + Margin.Bottom
                RemH = RemH - (Margin.Top + CDbl(R.Height) + Margin.Bottom)
            Case DockBottom
                If ChildH <= 0# Then ChildH = SlotH
                If ChildH > SlotH Then ChildH = SlotH
                ChildW = SlotW
                ChildH = ClampElementHeight(Child, ChildH)
                R.Width = CSng(ClampElementWidth(Child, ChildW))
                R.Height = CSng(ChildH)
                R.Left = CSng(RemL + Margin.Left)
                R.Top = CSng(RemT + RemH - Margin.Bottom - CDbl(R.Height))
                RemH = RemH - (Margin.Top + CDbl(R.Height) + Margin.Bottom)
            Case DockRight
                If ChildW <= 0# Then ChildW = SlotW
                If ChildW > SlotW Then ChildW = SlotW
                ChildH = SlotH
                ChildW = ClampElementWidth(Child, ChildW)
                R.Width = CSng(ChildW)
                R.Height = CSng(ClampElementHeight(Child, ChildH))
                R.Left = CSng(RemL + RemW - Margin.Right - CDbl(R.Width))
                R.Top = CSng(RemT + Margin.Top)
                RemW = RemW - (Margin.Left + CDbl(R.Width) + Margin.Right)
            Case Else ' DockLeft
                If ChildW <= 0# Then ChildW = SlotW
                If ChildW > SlotW Then ChildW = SlotW
                ChildH = SlotH
                ChildW = ClampElementWidth(Child, ChildW)
                R.Left = CSng(RemL + Margin.Left)
                R.Top = CSng(RemT + Margin.Top)
                R.Width = CSng(ChildW)
                R.Height = CSng(ClampElementHeight(Child, ChildH))
                RemL = RemL + Margin.Left + CDbl(R.Width) + Margin.Right
                RemW = RemW - (Margin.Left + CDbl(R.Width) + Margin.Right)
        End Select

        If RemW < 0# Then RemW = 0#
        If RemH < 0# Then RemH = 0#

        ApplyLayoutRectToElement ChildUI, R
        ApplyChildWidgetVisibility Child, ChildVis

NextArrangeDock:
    Next
End Sub

' Canvas measure: bounding box of children at Canvas.Left/Top + size.
Public Sub MeasureCanvasContent( _
    ByVal Children As UIElementCollection, _
    ByVal AvailableWidth As Double, _
    ByVal AvailableHeight As Double, _
    ByRef OutContentWidth As Double, _
    ByRef OutContentHeight As Double)

    Dim Child As Object
    Dim ChildUI As IUIElement
    Dim ChildVis As Visibility
    Dim Margin As Thickness
    Dim Measured As LayoutRect
    Dim L As Double
    Dim T As Double
    Dim ExtR As Double
    Dim ExtB As Double
    Dim MaxR As Double
    Dim MaxB As Double

    OutContentWidth = 0#
    OutContentHeight = 0#
    MaxR = 0#
    MaxB = 0#

    If Children Is Nothing Then Exit Sub

    For Each Child In Children
        If Not TypeOf Child Is IUIElement Then GoTo NextMeasureCanvas
        Set ChildUI = Child
        ChildVis = ReadElementVisibility(Child)
        If IsLayoutCollapsed(ChildVis) Then GoTo NextMeasureCanvas

        Set Margin = ReadElementMargin(Child)
        L = GetCanvasAttachedDouble(ChildUI, "Left", 0#)
        T = GetCanvasAttachedDouble(ChildUI, "Top", 0#)
        Measured = MeasureElementSize(Child, 0#, 0#)
        ExtR = L + Margin.Left + CDbl(Measured.Width) + Margin.Right
        ExtB = T + Margin.Top + CDbl(Measured.Height) + Margin.Bottom
        If ExtR > MaxR Then MaxR = ExtR
        If ExtB > MaxB Then MaxB = ExtB

NextMeasureCanvas:
    Next

    OutContentWidth = MaxR
    OutContentHeight = MaxB
    If AvailableWidth > 0 And OutContentWidth > AvailableWidth Then OutContentWidth = AvailableWidth
    If AvailableHeight > 0 And OutContentHeight > AvailableHeight Then OutContentHeight = AvailableHeight
End Sub

Public Sub ArrangeCanvasChildren( _
    ByVal Children As UIElementCollection, _
    ByVal HostWidget As cWidgetBase, _
    Optional ByVal OverrideHostWidth As Single = 0, _
    Optional ByVal OverrideHostHeight As Single = 0)

    Dim Child As Object
    Dim ChildUI As IUIElement
    Dim ChildVis As Visibility
    Dim Margin As Thickness
    Dim R As LayoutRect
    Dim Measured As LayoutRect
    Dim L As Double
    Dim T As Double
    Dim ChildW As Double
    Dim ChildH As Double

    If Children Is Nothing Then Exit Sub
    If HostWidget Is Nothing Then Exit Sub

    For Each Child In Children
        If Not TypeOf Child Is IUIElement Then GoTo NextArrangeCanvas
        Set ChildUI = Child

        ChildVis = ReadElementVisibility(Child)
        If IsLayoutCollapsed(ChildVis) Then
            DetachCollapsedChild Child, HostWidget
            GoTo NextArrangeCanvas
        End If

        If Not TypeOf Child Is IControl Then GoTo NextArrangeCanvas

        AttachChildWidget Child, HostWidget, ChildVis
        Set Margin = ReadElementMargin(Child)
        L = GetCanvasAttachedDouble(ChildUI, "Left", 0#)
        T = GetCanvasAttachedDouble(ChildUI, "Top", 0#)
        Measured = MeasureElementSize(Child, 0#, 0#)
        ChildW = CDbl(Measured.Width)
        ChildH = CDbl(Measured.Height)
        If ChildW <= 0# Then ChildW = ReadElementWidth(Child)
        If ChildH <= 0# Then ChildH = ReadElementHeight(Child)
        ChildW = ClampElementWidth(Child, ChildW)
        ChildH = ClampElementHeight(Child, ChildH)

        R.Left = CSng(L + Margin.Left)
        R.Top = CSng(T + Margin.Top)
        R.Width = CSng(ChildW)
        R.Height = CSng(ChildH)

        ApplyLayoutRectToElement ChildUI, R
        ApplyChildWidgetVisibility Child, ChildVis

NextArrangeCanvas:
    Next
End Sub
